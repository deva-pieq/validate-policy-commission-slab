#!/usr/bin/env python3
"""
validate_apl_vs_db_after_import.py

Post-Import APL vs Database Validation Engine (Read-Only)
=========================================================
Runs after `import_policy_commission.py` finishes updating the database.
Verifies that actual persisted PostgreSQL database values match the original
APL input file, evaluates update execution outcomes, and generates an audit-grade
multi-tab Excel validation workbook.

Strictly Read-Only: NEVER executes INSERT, UPDATE, DELETE, or schema modifications.
"""

import os
import sys
import re
import json
import copy
import difflib
import logging
import argparse
import datetime
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List, Tuple, Optional, Any, Union

# Ensure UTF-8 output encoding on Windows terminals
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import pandas as pd

# OpenPyXL for professional Excel workbook generation
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# Check PostgreSQL driver availability
DB_DRIVER_AVAILABLE = False
try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    DB_DRIVER_AVAILABLE = True
except ImportError:
    try:
        import importlib
        psycopg2 = importlib.import_module("psycopg2_binary")
        RealDictCursor = getattr(importlib.import_module("psycopg2_binary.extras"), "RealDictCursor", None)
        DB_DRIVER_AVAILABLE = True
    except (ImportError, AttributeError):
        DB_DRIVER_AVAILABLE = False


# ---------------------------------------------------------------------------
# Logging Setup
# ---------------------------------------------------------------------------
def setup_logger(verbose: bool = False) -> logging.Logger:
    logger = logging.getLogger("PostImportValidator")
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter(
            fmt="[%(asctime)s] [%(levelname)s] %(message)s", datefmt="%H:%M:%S"
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return logger


# ---------------------------------------------------------------------------
# Load Environment Configuration (.env)
# ---------------------------------------------------------------------------
def load_env_file(env_file_path: str = ".env") -> int:
    """Loads environment variables from a .env file into os.environ."""
    target_path = env_file_path
    if not os.path.isabs(target_path) and not os.path.exists(target_path):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        alt_path = os.path.join(script_dir, env_file_path)
        if os.path.exists(alt_path):
            target_path = alt_path

    if not os.path.exists(target_path):
        return 0

    count = 0
    try:
        with open(target_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, val = line.split("=", 1)
                    key = key.strip()
                    val = val.strip().strip("'").strip('"')
                    if key and key not in os.environ:
                        os.environ[key] = val
                        count += 1
    except Exception:
        pass
    return count


# ---------------------------------------------------------------------------
# Formula Injection Protection & Sanitization
# ---------------------------------------------------------------------------
def sanitize_for_excel(val: Any) -> Any:
    """
    Prevents CSV / Excel Formula Injection vulnerabilities.
    If a string starts with '=', '+', '-', or '@', prepend a single quote.
    """
    if val is None:
        return ""
    if isinstance(val, (int, float, Decimal, bool)):
        return val
    s = str(val)
    if s and s[0] in ("=", "+", "-", "@"):
        return f"'{s}"
    return s


# ---------------------------------------------------------------------------
# Parsing & Normalization Utilities
# ---------------------------------------------------------------------------
def clean_policy_number(val: Any) -> str:
    """Strips quotes, apostrophes, and surrounding whitespace."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return s.strip("'\"").strip()


def clean_agent_code(val: Any) -> str:
    """Cleans agent code: strips quotes, whitespace, and trailing alphabetic suffixes."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    v = str(val).strip().lstrip("'\"")
    v = re.sub(r"[A-Za-z]+$", "", v).strip()
    return v


def classify_contract_role(contract_val: Any) -> str:
    """Maps contract column string to canonical role: Agency, Sales Leader, Agent, Sub Agent."""
    c = str(contract_val or "").strip().lower()
    if c in ("agency", ""):
        return "Agency"
    if "sales" in c or "leader" in c:
        return "Sales Leader"
    if "sub" in c or re.match(r"^sa\d+$", c):
        return "Sub Agent"
    return "Agent"


def _month_number(val: Any, default: int) -> int:
    """Sheet 0 and database NULL both mean an open-ended month."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    text = str(val).strip()
    if not text or text.lower() in ("none", "null", "nan"):
        return default
    try:
        return int(float(text))
    except (TypeError, ValueError):
        return default


def _band_key(month_from: Any, month_to: Any) -> Tuple[int, int]:
    return (_month_number(month_from, 1), _month_number(month_to, 0))


def _read_apl_field(row: Dict[str, Any], *names: str) -> Any:
    lowered = {str(k).strip().lower(): v for k, v in row.items()}
    for name in names:
        if name.lower() in lowered:
            return lowered[name.lower()]
    return None


def _prepared_apl_rows(apl_records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Align sheet rows with the slabs the importer actually stored.

    Open ranges such as Agency 13-0 are sliced to 13-60, 61-120, and 121-0
    before the sequence roll-up, matching normalize_year_bands + rollup.
    """
    if not apl_records:
        return []
    has_share = any(_read_apl_field(row, "Normalized_Split_Share_%") not in (None, "") for row in apl_records)
    has_month = any(_read_apl_field(row, "From Month", "from_month", "month from") not in (None, "") for row in apl_records)
    if has_share or not has_month:
        prepared = []
        for row in apl_records:
            item = dict(row)
            item["_split_share"] = parse_float_safe(
                _read_apl_field(row, "Normalized_Split_Share_%", "Rate", "Rate %", "rate"),
                0.0,
            )
            prepared.append(item)
        return prepared

    frame = pd.DataFrame(apl_records)
    rate_col = next((c for c in frame.columns if str(c).strip().lower() in ("rate", "rate %", "commission %")), None)
    if rate_col:
        frame["_raw_rate"] = frame[rate_col]
    try:
        from validate_policy_commission_db import (
            fix_inclusive_month_band_overlaps,
            normalize_year_bands,
            rollup_commission_splits,
        )
        fixed, _ = fix_inclusive_month_band_overlaps(frame)
        sliced, _ = normalize_year_bands(fixed)
        rolled = rollup_commission_splits(sliced)
    except Exception:
        rolled = frame
        rolled["_split_share"] = [
            parse_float_safe(_read_apl_field(row, "Rate", "Rate %", "rate"), 0.0)
            for row in apl_records
        ]
        return rolled.to_dict(orient="records")

    if "Normalized_Split_Share_%" in rolled.columns:
        rolled["_split_share"] = rolled["Normalized_Split_Share_%"]
    else:
        rolled["_split_share"] = 0.0
    return rolled.to_dict(orient="records")


def _db_splits_for_record(record: Dict[str, Any]) -> List[Dict[str, Any]]:
    cs = record.get("commission_split")
    if isinstance(cs, list):
        return [item for item in cs if isinstance(item, dict)]
    if isinstance(cs, str):
        try:
            parsed = json.loads(cs)
            if isinstance(parsed, list):
                return [item for item in parsed if isinstance(item, dict)]
        except Exception:
            return []
    return []


def parse_float_safe(val: Any, default: float = 0.0) -> float:
    """Parses numeric string or number into float safely."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    s = str(val).replace("$", "").replace("%", "").replace(",", "").strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return default
    try:
        return float(s)
    except ValueError:
        return default


def compare_decimals(val1: Any, val2: Any, tolerance: float = 0.01) -> bool:
    """Compares two numeric values within a defined decimal tolerance.

    Comparison is done in integer cents. Floor-rounding a split can leave a
    one-cent gap (65.99 vs 66.00). A raw float subtract of that pair is
    slightly above 0.01, so a 0.01 tolerance would reject a match.
    """
    f1 = parse_float_safe(val1, 0.0)
    f2 = parse_float_safe(val2, 0.0)
    cents_1 = int(round(f1 * 100))
    cents_2 = int(round(f2 * 100))
    allowed = int(round(tolerance * 100))
    return abs(cents_1 - cents_2) <= allowed


def clean_name_ignoring_id(name: Any) -> str:
    """Strips agent IDs, numbers, or codes in parentheses, prefixes like '8307933 - Gregory Teipel', and generational suffixes (Jr, Sr, II, III, IV, V)."""
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    s = str(name).strip()
    s = re.sub(r"^\s*[\w\d]+\s*-\s*", "", s).strip()
    s = re.sub(r"\s*\([^)]*\)$", "", s).strip()
    s = re.sub(r"\s+\b(jr|sr|[ivx]+)\b\.?\s*$", "", s, flags=re.IGNORECASE).strip()
    return s


def normalize_level(val: Any) -> str:
    """
    Normalizes level / contract strings into canonical level forms:
    - 'LVL 5', 'LVL5', 'LVL-5', 'Level 5' -> 'Level 5'
    - 'LVL 1', 'LVL1', 'Level 1' -> 'Level 1'
    - 'LVL 2', 'LVL2', 'Level 2' -> 'Level 2'
    - 'LVL 3', 'LVL3', 'Level 3' -> 'Level 3'
    - 'LVL 4', 'LVL4', 'Level 4' -> 'Level 4'
    - 'SUBAGT1', 'SUBAGT 1', 'SA1', 'SA 1' -> 'SA1'
    - 'SUBAGT2', 'SUBAGT 2', 'SA2', 'SA 2' -> 'SA2'
    - 'SUBAGT<N>' -> 'SA<N>'
    - 'Sales Leader', 'SL' -> 'Sales Leader'
    - 'Agency', '', None -> 'Agency'
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if not s or s.lower() in ("none", "nan", "null", "-", "n/a"):
        return ""

    if s.lower() == "agency":
        return "Agency"

    if re.search(r"sales\s*leader|\bsl\b", s, re.I):
        return "Sales Leader"

    m_lvl = re.match(r"^(?:lvl|level)\s*[-_]?\s*(\d+)$", s, re.I)
    if m_lvl:
        return f"Level {m_lvl.group(1)}"

    m_sa = re.match(r"^(?:subagt|sub\s*agent|sa)\s*[-_]?\s*(\d+)$", s, re.I)
    if m_sa:
        return f"SA{m_sa.group(1)}"

    return s


def compare_levels(apl_contract_or_lvl: Any, db_lvl: Any, role: str = "") -> Dict[str, Any]:
    """
    Zero-compromise level comparator:
    Compares APL Contract/Level (e.g. 'LVL5', 'SUBAGT1', 'Sales Leader', 'Agency')
    against DB split 'level' (e.g. 'Level 5', 'SA1', 'Sales Leader', None/empty).
    """
    norm_apl = normalize_level(apl_contract_or_lvl)
    norm_db = normalize_level(db_lvl)
    r = str(role or "").strip().lower()

    # Agency level: DB stores None/empty or Agency
    if r == "agency" or norm_apl == "Agency":
        if norm_db in ("Agency", "", "None"):
            return {
                "match": True,
                "status": "MATCH",
                "apl_level": norm_apl or "Agency",
                "db_level": norm_db or "Agency",
                "disp": "MATCH (Agency)",
            }

    # Sales Leader: DB may store 'Sales Leader', 'Level 1', 'Level 5', or empty
    if r == "sales leader" or norm_apl == "Sales Leader":
        if norm_db in ("Sales Leader", "", "None", "Level 1", "Level 2", "Level 3", "Level 4", "Level 5") or norm_db == norm_apl:
            return {
                "match": True,
                "status": "MATCH",
                "apl_level": norm_apl or "Sales Leader",
                "db_level": norm_db or "Sales Leader",
                "disp": f"MATCH ({norm_db or 'Sales Leader'})",
            }

    # Broker / Agent contract level: DB may store 'Agent', 'Broker', or empty
    if norm_apl.lower() in ("broker", "agent") and norm_db.lower() in ("broker", "agent", ""):
        return {
            "match": True,
            "status": "MATCH",
            "apl_level": norm_apl,
            "db_level": norm_db or "Agent",
            "disp": f"MATCH ({norm_db or 'Agent'})",
        }

    # Direct canonical match (e.g. 'LVL5' == 'Level 5', 'SA1' == 'SA1')
    if norm_apl and norm_db and norm_apl.lower() == norm_db.lower():
        return {
            "match": True,
            "status": "MATCH",
            "apl_level": norm_apl,
            "db_level": norm_db,
            "disp": f"MATCH ({norm_db})",
        }

    return {
        "match": False,
        "status": "MISMATCH",
        "apl_level": norm_apl or str(apl_contract_or_lvl),
        "db_level": norm_db or str(db_lvl),
        "disp": f"MISMATCH (File: '{norm_apl or apl_contract_or_lvl}' vs DB: '{norm_db or db_lvl}')",
    }


# ---------------------------------------------------------------------------
# Person Name Matching Engine
# ---------------------------------------------------------------------------
def compare_person_names(apl_name: Any, db_name: Any) -> Dict[str, Any]:
    """
    Dedicated 4-stage person-name matching engine.
    Supports extra middle-name tokens in both directions while preserving order
    and requiring matching first and last tokens.
    Automatically strips trailing parentheses with IDs/codes (e.g. 'Heidi Berthelsen (19617796)').
    """
    orig_apl = str(apl_name or "").strip()
    orig_db = str(db_name or "").strip()

    if not orig_apl and not orig_db:
        return {
            "status": "MISMATCH",
            "method": "NONE",
            "similarity": 0.0,
            "normalized_apl": "",
            "normalized_db": "",
            "reason": "Both names empty",
        }
    if not orig_apl or not orig_db:
        return {
            "status": "MISMATCH",
            "method": "NONE",
            "similarity": 0.0,
            "normalized_apl": orig_apl,
            "normalized_db": orig_db,
            "reason": "One name is empty",
        }

    # Clean IDs/codes from parentheses for accurate name validation
    clean_apl = clean_name_ignoring_id(orig_apl)
    clean_db = clean_name_ignoring_id(orig_db)

    # Stage 1: Exact Match (either raw strings or clean names)
    if orig_apl == orig_db or clean_apl == clean_db:
        method = "EXACT" if orig_apl == orig_db else "BASIC_NORMALIZATION"
        return {
            "status": "EXACT_MATCH" if orig_apl == orig_db else "NORMALIZED_MATCH",
            "method": method,
            "similarity": 1.0,
            "normalized_apl": clean_apl,
            "normalized_db": clean_db,
            "reason": "Exact string equality" if orig_apl == orig_db else "Exact name match (ignoring ID in parentheses)",
        }

    # Stage 2: Basic Normalization
    norm_apl = " ".join(clean_apl.strip().strip("'\"").lower().split())
    norm_db = " ".join(clean_db.strip().strip("'\"").lower().split())

    if norm_apl == norm_db:
        return {
            "status": "NORMALIZED_MATCH",
            "method": "BASIC_NORMALIZATION",
            "similarity": 1.0,
            "normalized_apl": norm_apl,
            "normalized_db": norm_db,
            "reason": "Normalized match (case, whitespace, quotes, ignoring ID in parentheses)",
        }

    # Agency Alias Match (Main Line Benefits <-> MLB)
    if ("main line" in norm_apl or "mlb" in norm_apl) and ("main line" in norm_db or "mlb" in norm_db):
        return {
            "status": "NORMALIZED_MATCH",
            "method": "AGENCY_ALIAS",
            "similarity": 1.0,
            "normalized_apl": norm_apl,
            "normalized_db": norm_db,
            "reason": "Agency alias match (Main Line Benefits / MLB)",
        }

    # Nickname Normalization (e.g. Greg <-> Gregory)
    norm_apl_nick = re.sub(r"\bgreg\b", "gregory", norm_apl)
    norm_db_nick = re.sub(r"\bgreg\b", "gregory", norm_db)
    if norm_apl_nick == norm_db_nick:
        return {
            "status": "NORMALIZED_MATCH",
            "method": "NICKNAME_NORMALIZATION",
            "similarity": 1.0,
            "normalized_apl": norm_apl,
            "normalized_db": norm_db,
            "reason": "Normalized nickname match (Greg / Gregory)",
        }

    t_apl = norm_apl.split()
    t_db = norm_db.split()

    # Stage 3: Extra Middle-Name Tokens
    # Bidirectional: APL has extra middle tokens or DB has extra middle tokens
    t_short, t_long = (t_apl, t_db) if len(t_apl) <= len(t_db) else (t_db, t_apl)

    if len(t_short) >= 2 and len(t_long) > len(t_short):
        # First name and last name must match
        if t_short[0] == t_long[0] and t_short[-1] == t_long[-1]:
            # Subsequence check: All tokens of t_short must appear in t_long in preserved order
            it = iter(t_long)
            if all(tok in it for tok in t_short):
                sim = round(len(t_short) / len(t_long), 3)
                return {
                    "status": "NORMALIZED_MATCH",
                    "method": "EXTRA_MIDDLE_NAME",
                    "similarity": sim,
                    "normalized_apl": norm_apl,
                    "normalized_db": norm_db,
                    "reason": "Extra middle-name tokens preserved in order between shared first and last names",
                }

    # Equal token count (>= 3) with matching first and last names but different middle tokens
    # e.g., 'John Walk Smith' vs 'John Michael Smith'
    if len(t_apl) >= 3 and len(t_db) >= 3 and t_apl[0] == t_db[0] and t_apl[-1] == t_db[-1]:
        return {
            "status": "NORMALIZED_MATCH",
            "method": "EXTRA_MIDDLE_NAME",
            "similarity": 0.85,
            "normalized_apl": norm_apl,
            "normalized_db": norm_db,
            "reason": "First and last names match exactly with middle-name variations",
        }

        # Surnames differ
        if t_apl[-1] != t_db[-1]:
            sur_ratio = difflib.SequenceMatcher(None, t_apl[-1], t_db[-1]).ratio()
            # If first name matches and surname has small edit distance (e.g., Smith vs Smyth or Smythe)
            if sur_ratio >= 0.70 and t_apl[0] == t_db[0]:
                return {
                    "status": "CLOSE_MATCH_REVIEW_REQUIRED",
                    "method": "FUZZY_SPELLING_VARIATION",
                    "similarity": round(sur_ratio, 3),
                    "normalized_apl": norm_apl,
                    "normalized_db": norm_db,
                    "reason": f"First name matches but surname has possible spelling variation ({t_apl[-1]} vs {t_db[-1]})",
                }
            return {
                "status": "MISMATCH",
                "method": "NONE",
                "similarity": round(sur_ratio, 3),
                "normalized_apl": norm_apl,
                "normalized_db": norm_db,
                "reason": f"Different surnames ({t_apl[-1]} vs {t_db[-1]})",
            }
        elif t_apl[0] != t_db[0]:
            # First names differ
            first_ratio = difflib.SequenceMatcher(None, t_apl[0], t_db[0]).ratio()
            if first_ratio >= 0.70 and t_apl[-1] == t_db[-1]:
                return {
                    "status": "CLOSE_MATCH_REVIEW_REQUIRED",
                    "method": "FUZZY_SPELLING_VARIATION",
                    "similarity": round(first_ratio, 3),
                    "normalized_apl": norm_apl,
                    "normalized_db": norm_db,
                    "reason": f"Surname matches but first name has possible spelling variation ({t_apl[0]} vs {t_db[0]})",
                }
            return {
                "status": "MISMATCH",
                "method": "NONE",
                "similarity": round(first_ratio, 3),
                "normalized_apl": norm_apl,
                "normalized_db": norm_db,
                "reason": f"Different first names ({t_apl[0]} vs {t_db[0]})",
            }

    # Reordered tokens (e.g. 'Smith John' vs 'John Smith')
    if set(t_apl) == set(t_db) and len(t_apl) > 1 and t_apl != t_db:
        return {
            "status": "MISMATCH",
            "method": "NONE",
            "similarity": 0.7,
            "normalized_apl": norm_apl,
            "normalized_db": norm_db,
            "reason": "Reordered name tokens do not match positional order",
        }

    # Overall string similarity fallback
    overall_ratio = difflib.SequenceMatcher(None, norm_apl, norm_db).ratio()
    if overall_ratio >= 0.85:
        return {
            "status": "CLOSE_MATCH_REVIEW_REQUIRED",
            "method": "FUZZY_SPELLING_VARIATION",
            "similarity": round(overall_ratio, 3),
            "normalized_apl": norm_apl,
            "normalized_db": norm_db,
            "reason": "High string similarity spelling variation requiring review",
        }

    return {
        "status": "MISMATCH",
        "method": "NONE",
        "similarity": round(overall_ratio, 3),
        "normalized_apl": norm_apl,
        "normalized_db": norm_db,
        "reason": "Names do not match",
    }


# ---------------------------------------------------------------------------
# Policy Identifier Matching Engine
# ---------------------------------------------------------------------------
def match_policy_identifier(apl_policy: str, db_candidates: List[str]) -> Dict[str, Any]:
    """
    Conservative policy identifier matching engine across all policy types.
    Stage A: Exact match
    Stage B: Normalized match (case, whitespace, quote stripping)
    Stage C: Close-match candidate search (extra char, missing char, typo, suffix)
    """
    raw_apl = str(apl_policy or "").strip()
    if not raw_apl or raw_apl.lower() == "nan":
        return {
            "status": "NO_MATCH",
            "method": "NONE",
            "similarity": 0.0,
            "reason": "Empty APL policy number",
            "apl_policy": raw_apl,
            "db_policy": None,
        }

    # Stage A: Exact Match
    for db in db_candidates:
        if raw_apl == str(db):
            return {
                "status": "EXACT_MATCH",
                "method": "EXACT",
                "similarity": 1.0,
                "reason": "Exact string equality",
                "apl_policy": raw_apl,
                "db_policy": str(db),
            }

    # Stage B: Normalized Match
    norm_apl = clean_policy_number(raw_apl).upper()
    norm_apl = " ".join(norm_apl.split())

    for db in db_candidates:
        norm_db = clean_policy_number(db).upper()
        norm_db = " ".join(norm_db.split())
        if norm_apl == norm_db:
            return {
                "status": "NORMALIZED_MATCH",
                "method": "BASIC_NORMALIZATION",
                "similarity": 1.0,
                "reason": "Normalized match (stripped quotes/whitespace/case)",
                "apl_policy": raw_apl,
                "db_policy": str(db),
            }

    # Stage C: Close-Match Candidate Search
    matches_by_db: Dict[str, Dict[str, Any]] = {}

    for db in db_candidates:
        db_str = str(db)
        norm_db = clean_policy_number(db_str).upper()
        norm_db = " ".join(norm_db.split())

        # 1. Extra character in APL
        if len(norm_apl) == len(norm_db) + 1:
            for i in range(len(norm_apl)):
                if norm_apl[:i] + norm_apl[i + 1:] == norm_db:
                    matches_by_db[db_str] = {
                        "status": "CLOSE_MATCH_EXTRA_CHARACTERS",
                        "method": "EXTRA_CHARACTER",
                        "similarity": 0.95,
                        "reason": f"APL has extra character '{norm_apl[i]}' at position {i + 1}",
                        "db_policy": db_str,
                    }
                    break

        # 2. Missing character in APL
        elif len(norm_apl) + 1 == len(norm_db):
            for i in range(len(norm_db)):
                if norm_db[:i] + norm_db[i + 1:] == norm_apl:
                    matches_by_db[db_str] = {
                        "status": "CLOSE_MATCH_MISSING_CHARACTERS",
                        "method": "MISSING_CHARACTER",
                        "similarity": 0.95,
                        "reason": f"APL missing character '{norm_db[i]}' at position {i + 1}",
                        "db_policy": db_str,
                    }
                    break

        # 3. Possible typo / single character substitution
        elif len(norm_apl) == len(norm_db) and norm_apl:
            diffs = [(i, c1, c2) for i, (c1, c2) in enumerate(zip(norm_apl, norm_db)) if c1 != c2]
            if len(diffs) == 1:
                idx, c1, c2 = diffs[0]
                matches_by_db[db_str] = {
                    "status": "CLOSE_MATCH_POSSIBLE_TYPO",
                    "method": "POSSIBLE_TYPO",
                    "similarity": 0.92,
                    "reason": f"Single character substitution: '{c1}' vs '{c2}' at position {idx + 1}",
                    "db_policy": db_str,
                }

        # 4. Trailing alphabetic suffix variation (e.g. OVR, SL)
        if db_str not in matches_by_db:
            stripped_apl = re.sub(r"[A-Za-z]+$", "", norm_apl)
            stripped_db = re.sub(r"[A-Za-z]+$", "", norm_db)
            if stripped_apl and stripped_db and stripped_apl == stripped_db and (stripped_apl != norm_apl or stripped_db != norm_db):
                matches_by_db[db_str] = {
                    "status": "CLOSE_MATCH_EXTRA_CHARACTERS",
                    "method": "SUFFIX_VARIATION",
                    "similarity": 0.90,
                    "reason": f"Suffix variation stripped: {norm_apl} vs {norm_db}",
                    "db_policy": db_str,
                }

    if len(matches_by_db) > 1:
        cands = list(matches_by_db.keys())
        return {
            "status": "AMBIGUOUS_MATCH",
            "method": "MULTIPLE_CANDIDATES",
            "similarity": 0.8,
            "reason": f"Multiple plausible DB candidates satisfy similarity rules: {cands}",
            "apl_policy": raw_apl,
            "db_policy": None,
        }

    if len(matches_by_db) == 1:
        match_info = next(iter(matches_by_db.values()))
        match_info["apl_policy"] = raw_apl
        return match_info

    return {
        "status": "NO_MATCH",
        "method": "NONE",
        "similarity": 0.0,
        "reason": "No matching database policy found",
        "apl_policy": raw_apl,
        "db_policy": None,
    }


# ---------------------------------------------------------------------------
# Read-Only Database Manager
# ---------------------------------------------------------------------------
class ReadOnlyDatabaseManager:
    """Manages read-only PostgreSQL queries with strictly enforced read-only sessions."""

    def __init__(
        self,
        db_url: str = None,
        host: str = None,
        port: int = 5432,
        dbname: str = None,
        user: str = None,
        password: str = None,
        logger: Optional[logging.Logger] = None,
    ):
        self.db_url = db_url or os.getenv("DATABASE_URL")
        self.user = user or os.getenv("DB_USER")
        self.password = password or os.getenv("DB_PASSWORD")
        self.host = host or os.getenv("DB_HOST", "localhost")
        self.port = port or int(os.getenv("DB_PORT", 5432))
        self.dbname = dbname or os.getenv("DB_NAME", "prod_icm_db")
        self.logger = logger or logging.getLogger("PostImportValidator")
        self.conn = None
        self.user_display_by_uuid: Dict[str, Dict[str, str]] = {}

    def connect(self) -> bool:
        """Establishes read-only PostgreSQL connection with timeout and sslmode."""
        if not DB_DRIVER_AVAILABLE:
            self.logger.warning("psycopg2 driver not installed. Database queries unavailable.")
            return False

        try:
            sslmode = os.getenv("DB_SSLMODE") or ("require" if "rds.amazonaws.com" in str(self.db_url or self.host) else "prefer")
            if self.db_url:
                clean_url = str(self.db_url).strip()
                if clean_url.lower().startswith("jdbc:"):
                    clean_url = clean_url[5:]
                conn_kwargs = {
                    "dsn": clean_url,
                    "connect_timeout": 6,
                    "options": "-c statement_timeout=180000",
                    "sslmode": sslmode,
                }
                if self.user:
                    conn_kwargs["user"] = self.user
                if self.password:
                    conn_kwargs["password"] = self.password
                self.conn = psycopg2.connect(**conn_kwargs)
            else:
                self.conn = psycopg2.connect(
                    host=self.host,
                    port=self.port,
                    dbname=self.dbname,
                    user=self.user,
                    password=self.password,
                    connect_timeout=6,
                    options="-c statement_timeout=180000",
                    sslmode=sslmode,
                )

            # Enforce read-only transaction and autocommit
            self.conn.set_session(readonly=True, autocommit=True)
            self.logger.info("Connected to PostgreSQL database in strictly READ-ONLY mode.")
            return True
        except Exception as e:
            self.logger.error(f"Could not connect to PostgreSQL database: {e}")
            self.conn = None
            return False

    def close(self):
        if self.conn:
            try:
                self.conn.close()
            except Exception:
                pass
            self.conn = None

    def fetch_all_users_and_contacts(self, agency_id: int) -> Dict[str, str]:
        """Maps (agent_code / agent_name) -> user_uuid and populates user_display_by_uuid."""
        if not self.conn:
            return {}
        user_map: Dict[str, str] = {}
        self.user_display_by_uuid = {}
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                # 1. Users table
                cur.execute(
                    """
                    SELECT uuid, code 
                    FROM users 
                    WHERE (tenant_id = %s OR agency_id = %s) AND uuid IS NOT NULL;
                """,
                    (agency_id, agency_id),
                )
                for r in cur.fetchall():
                    u_uuid = str(r.get("uuid") or "").strip()
                    code = str(r.get("code") or "").strip()
                    if u_uuid:
                        if code:
                            user_map[code.lower()] = u_uuid
                            clean_c = clean_agent_code(code).lower()
                            user_map[clean_c] = u_uuid
                            self.user_display_by_uuid[u_uuid.lower()] = {"name": "", "code": code}

                # 2. Contacts table
                cur.execute(
                    """
                    SELECT relationship_id, first_name, last_name 
                    FROM contacts 
                    WHERE agency_id = %s AND relationship_id IS NOT NULL;
                """,
                    (agency_id,),
                )
                for r in cur.fetchall():
                    rel_id = str(r.get("relationship_id") or "").strip()
                    if not rel_id:
                        continue
                    fname = str(r.get("first_name") or "").strip()
                    lname = str(r.get("last_name") or "").strip()
                    full = " ".join(f"{fname} {lname}".lower().split())
                    if full:
                        user_map[full] = rel_id
                        display = self.user_display_by_uuid.setdefault(rel_id.lower(), {"name": "", "code": ""})
                        display["name"] = " ".join(f"{fname} {lname}".split())
        except Exception as e:
            self.logger.warning(f"Warning: Could not fetch user/contact directory: {e}")
        return user_map

    def fetch_policies_and_commissions(
        self, agency_id: int, policy_numbers: List[str]
    ) -> Tuple[Dict[str, List[Dict[str, Any]]], List[str]]:
        """
        Batch fetches policies, products, and policy_commission records.
        Returns:
            (records_by_policy_number, all_existing_db_policy_numbers)
        """
        if not self.conn:
            return {}, []

        clean_pols = [p for p in {clean_policy_number(x) for x in policy_numbers} if p]
        if not clean_pols:
            return {}, []

        results: Dict[str, List[Dict[str, Any]]] = {}
        all_db_numbers: List[str] = []

        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                # 1. Fetch policies and products
                cur.execute(
                    """
                    SELECT 
                        p.uuid AS product_uuid,
                        p.name AS product_name,
                        pol.uuid AS policy_uuid,
                        pol.number AS policy_number,
                        pol.writing_agent_user_id,
                        pol.pay_to_agent_user_id
                    FROM products p
                    JOIN policies pol ON (pol.agency_id = p.agency_id AND pol.product_id = p.uuid)
                    WHERE p.agency_id = %s
                      AND pol.agency_id = %s
                      AND pol.number = ANY(%s);
                """,
                    (agency_id, agency_id, clean_pols),
                )
                policy_rows = cur.fetchall()

                # Collect all policy numbers found
                for row in policy_rows:
                    p_num = str(row["policy_number"]).strip()
                    if p_num not in all_db_numbers:
                        all_db_numbers.append(p_num)

                # 2. Fetch policy_commission records
                cur.execute(
                    """
                    SELECT 
                        pc.uuid AS commission_uuid,
                        pc.policy_uuid,
                        pc.rule_name,
                        pc.version,
                        pc.commission_type,
                        pc.payment_frequency,
                        pc.month_from,
                        pc.month_to,
                        pc.payout_method,
                        pc.fee_amount,
                        pc.premium_amount,
                        pc.commission_split
                    FROM policy_commission pc
                    WHERE pc.agency_id = %s
                      AND pc.policy_uuid IN (
                          SELECT uuid FROM policies WHERE agency_id = %s AND number = ANY(%s)
                      );
                """,
                    (agency_id, agency_id, clean_pols),
                )
                comm_rows = cur.fetchall()

                # Group commissions by policy_uuid
                comm_map: Dict[str, List[Dict[str, Any]]] = {}
                for c in comm_rows:
                    pol_uuid = str(c.get("policy_uuid") or "").strip()
                    if pol_uuid:
                        comm_map.setdefault(pol_uuid, []).append(dict(c))

                # Combine policies with matching commission rules
                for p in policy_rows:
                    p_num = str(p["policy_number"]).strip()
                    pol_uuid = str(p["policy_uuid"]).strip()
                    matching_comms = comm_map.get(pol_uuid, [])

                    if not matching_comms:
                        rec = dict(p)
                        rec.update({
                            "commission_uuid": None,
                            "rule_name": None,
                            "commission_type": None,
                            "payment_frequency": None,
                            "month_from": None,
                            "month_to": None,
                            "payout_method": None,
                            "fee_amount": None,
                            "premium_amount": None,
                            "commission_split": [],
                        })
                        results.setdefault(p_num, []).append(rec)
                    else:
                        for comm in matching_comms:
                            rec = dict(p)
                            rec.update(comm)
                            results.setdefault(p_num, []).append(rec)

        except Exception as e:
            self.logger.error(f"Error querying database for policies: {e}")

        return results, all_db_numbers

    def format_user_display(self, user_id: Any) -> str:
        """Formats display string for a user UUID."""
        uid = str(user_id or "").strip().lower()
        if not uid or uid in ("none", "null"):
            return ""
        info = self.user_display_by_uuid.get(uid, {})
        name = info.get("name", "")
        code = info.get("code", "")
        if name and code:
            return f"{name} ({code})"
        return name or (f"Code: {code}" if code else uid)


# ---------------------------------------------------------------------------
# Execution Status Determination
# ---------------------------------------------------------------------------
def load_execution_report(report_path: Optional[str]) -> Dict[str, Dict[str, Any]]:
    """Loads optional execution report (JSON, CSV, or text log) into policy-keyed dict."""
    if not report_path or not os.path.exists(report_path):
        return {}

    exec_map: Dict[str, Dict[str, Any]] = {}
    try:
        _, ext = os.path.splitext(report_path.lower())
        if ext == ".json":
            with open(report_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    for item in data:
                        p_num = clean_policy_number(
                            item.get("policyNumber") or item.get("policy_number") or item.get("policy")
                        )
                        if p_num:
                            exec_map[p_num] = item
                elif isinstance(data, dict):
                    for k, v in data.items():
                        p_num = clean_policy_number(k)
                        if p_num:
                            exec_map[p_num] = v if isinstance(v, dict) else {"status": v}
        elif ext in (".csv", ".xlsx", ".xls"):
            df = pd.read_excel(report_path) if ext in (".xlsx", ".xls") else pd.read_csv(report_path)
            pol_col = next((c for c in df.columns if "policy" in str(c).lower()), None)
            if pol_col:
                for _, row in df.iterrows():
                    p_num = clean_policy_number(row[pol_col])
                    if p_num:
                        exec_map[p_num] = row.to_dict()
    except Exception as e:
        print(f"⚠️ Could not load execution report from '{report_path}': {e}")
    return exec_map


def determine_execution_status(
    policy_number: str,
    exec_report_map: Dict[str, Dict[str, Any]],
    pipeline_output_dir: Optional[str] = None,
) -> Tuple[str, str]:
    """
    Determines execution outcome based on verifiable project evidence.
    Statuses: EXECUTED_SUCCESSFULLY, EXECUTED_WITH_ERRORS, SKIPPED, REJECTED, FAILED, UNKNOWN.
    """
    pol = clean_policy_number(policy_number)
    if not pol:
        return "UNKNOWN", "Empty policy number"

    # 1. Check explicit execution report map if provided
    if pol in exec_report_map:
        info = exec_report_map[pol]
        st = str(info.get("status") or info.get("execution_status") or "").strip().upper()
        msg = str(info.get("message") or info.get("error") or info.get("reason") or "Report entry").strip()
        if "SUCCESS" in st or st == "1" or st == "OK" or "CREATED" in st:
            return "EXECUTED_SUCCESSFULLY", f"Report: {msg}"
        if "REJECT" in st:
            return "REJECTED", f"Report: {msg}"
        if "SKIP" in st:
            return "SKIPPED", f"Report: {msg}"
        if "FAIL" in st or "ERROR" in st:
            return "FAILED", f"Report error: {msg}"

    # 2. Check run directory artifacts if available
    if pipeline_output_dir and os.path.exists(pipeline_output_dir):
        # Look for rejected payloads json
        for fname in os.listdir(pipeline_output_dir):
            if fname.startswith("rejected_missing_user_uuid_payloads") and fname.endswith(".json"):
                fpath = os.path.join(pipeline_output_dir, fname)
                try:
                    with open(fpath, "r", encoding="utf-8") as rf:
                        rej_data = json.load(rf)
                        for item in rej_data:
                            if clean_policy_number(item.get("policyNumber") or item.get("policy_number")) == pol:
                                return "REJECTED", f"Found in artifact '{fname}': Missing required user UUID"
                except Exception:
                    pass

        # Look for auto-updated verification excel
        for fname in os.listdir(pipeline_output_dir):
            if fname.startswith("db_verification_auto_updated") and fname.endswith(".xlsx"):
                fpath = os.path.join(pipeline_output_dir, fname)
                try:
                    df_up = pd.read_excel(fpath)
                    pcol = next((c for c in df_up.columns if "policy" in str(c).lower()), None)
                    if pcol and pol in [clean_policy_number(x) for x in df_up[pcol]]:
                        return "EXECUTED_SUCCESSFULLY", f"Found in artifact '{fname}': Successfully updated via API"
                except Exception:
                    pass

    return "UNKNOWN", "No per-policy execution report or artifact evidence provided"


# ---------------------------------------------------------------------------
# Overall Status Combiner Matrix
# ---------------------------------------------------------------------------
def compute_overall_status(exec_status: str, val_status: str) -> str:
    """Computes final auditable overall status combining execution and data validation."""
    if val_status == "DB_RECORD_NOT_FOUND":
        return "DB_RECORD_NOT_FOUND"
    if val_status == "AMBIGUOUS_MATCH":
        return "AMBIGUOUS_MATCH"
    if val_status == "NOT_VALIDATED":
        return "NOT_VALIDATED"

    if exec_status == "REJECTED":
        return "REJECTED"
    if exec_status == "FAILED":
        return "EXECUTION_FAILED"
    if exec_status == "SKIPPED":
        return "EXECUTION_SKIPPED"
    if exec_status == "EXECUTED_WITH_ERRORS":
        return "EXECUTION_REVIEW_REQUIRED"

    if exec_status == "EXECUTED_SUCCESSFULLY":
        if val_status == "EXACT_DATA_MATCH":
            return "SUCCESS_EXACT_MATCH"
        if val_status == "MATCH_AFTER_NORMALIZATION":
            return "SUCCESS_NORMALIZED_MATCH"
        if val_status == "CLOSE_MATCH_REVIEW_REQUIRED":
            return "REVIEW_REQUIRED"
        if val_status == "DATA_MISMATCH":
            return "EXECUTED_BUT_DATA_MISMATCH"

    if exec_status == "UNKNOWN":
        if val_status in ("EXACT_DATA_MATCH", "MATCH_AFTER_NORMALIZATION"):
            return "EXECUTION_STATUS_UNKNOWN"
        if val_status == "CLOSE_MATCH_REVIEW_REQUIRED":
            return "REVIEW_REQUIRED"
        if val_status == "DATA_MISMATCH":
            return "DATA_MISMATCH_EXECUTION_UNKNOWN"

    return "REVIEW_REQUIRED"


# ---------------------------------------------------------------------------
# Core Comparison & Validation Engine
# ---------------------------------------------------------------------------
def validate_policy_records(
    apl_records: List[Dict[str, Any]],
    db_records: List[Dict[str, Any]],
    policy_match: Dict[str, Any],
    exec_status: str,
    exec_evidence: str,
    user_display_fn: Any,
) -> Dict[str, Any]:
    """
    Compares a single APL policy's rows against its persisted DB policy_commission records.
    Returns:
        {
            "summary": Dict,
            "field_comparisons": List[Dict],
            "split_comparisons": List[Dict]
        }
    """
    field_comps: List[Dict[str, Any]] = []
    split_comps: List[Dict[str, Any]] = []

    apl_policy = policy_match["apl_policy"]
    db_policy = policy_match["db_policy"]
    p_status = policy_match["status"]

    primary_apl = apl_records[0] if apl_records else {}
    apl_prod = str(primary_apl.get("Product") or primary_apl.get("product") or "").strip()

    # Determine writing agent and agent names in APL
    apl_wa_name = ""
    apl_agent_name = ""
    for r in apl_records:
        role = classify_contract_role(r.get("Contract") or r.get("contract"))
        name = str(r.get("Agent Name") or r.get("agent_name") or "").strip()
        if role in ("Sub Agent", "Agent") and not apl_wa_name:
            apl_wa_name = name
        if role == "Agent" and not apl_agent_name:
            apl_agent_name = name
    if not apl_agent_name and apl_wa_name:
        apl_agent_name = apl_wa_name

    matched_fields = 0
    mismatched_fields = 0
    review_fields = 0
    mismatch_names: List[str] = []
    review_reasons: List[str] = []

    # If DB record not found
    if not db_records:
        if p_status == "AMBIGUOUS_MATCH":
            val_status = "AMBIGUOUS_MATCH"
            review_reasons.append("Ambiguous policy candidate match in DB")
        else:
            val_status = "DB_RECORD_NOT_FOUND"
            review_reasons.append("Policy not found in database")

        overall_status = compute_overall_status(exec_status, val_status)
        summary = {
            "apl_policy": apl_policy,
            "db_policy": db_policy or "",
            "policy_match_status": p_status,
            "policy_match_method": policy_match.get("method", "NONE"),
            "policy_similarity": policy_match.get("similarity", 0.0),
            "product_name": apl_prod,
            "carrier": "MLB",
            "agency": "Main Line Benefits (5)",
            "apl_agent_name": apl_agent_name,
            "db_agent_name": "",
            "agent_match_status": "NOT_APPLICABLE",
            "apl_writing_agent_name": apl_wa_name,
            "db_writing_agent_name": "",
            "writing_agent_match_status": "NOT_APPLICABLE",
            "execution_status": exec_status,
            "execution_evidence": exec_evidence,
            "data_validation_status": val_status,
            "overall_status": overall_status,
            "matched_fields_count": 0,
            "mismatched_fields_count": 0,
            "review_required_fields_count": 1,
            "mismatch_field_names": "",
            "review_required": True,
            "review_reason": "; ".join(review_reasons),
        }
        return {
            "summary": summary,
            "field_comparisons": field_comps,
            "split_comparisons": split_comps,
        }

    # Policy identifier field comparison
    pol_rev = p_status.startswith("CLOSE_MATCH") or p_status == "AMBIGUOUS_MATCH"
    if pol_rev:
        review_fields += 1
        review_reasons.append(f"Policy identifier match requires review ({p_status})")
    elif p_status in ("EXACT_MATCH", "NORMALIZED_MATCH"):
        matched_fields += 1

    field_comps.append({
        "apl_policy": apl_policy,
        "db_policy": db_policy,
        "field_name": "policy_number",
        "apl_value": apl_policy,
        "db_value": db_policy,
        "status": p_status,
        "method": policy_match.get("method", "EXACT"),
        "diff_reason": policy_match.get("reason", ""),
        "review_required": pol_rev,
    })

    # Primary DB record
    primary_db = db_records[0]
    db_prod = str(primary_db.get("product_name") or "").strip()

    # Product Name comparison
    prod_norm_apl = " ".join(apl_prod.lower().split())
    prod_norm_db = " ".join(db_prod.lower().split())
    if apl_prod == db_prod:
        prod_st = "EXACT_MATCH"
        matched_fields += 1
    elif prod_norm_apl == prod_norm_db or prod_norm_apl in prod_norm_db or prod_norm_db in prod_norm_apl:
        prod_st = "NORMALIZED_MATCH"
        matched_fields += 1
    else:
        prod_st = "DATA_MISMATCH"
        mismatched_fields += 1
        mismatch_names.append("product_name")

    field_comps.append({
        "apl_policy": apl_policy,
        "db_policy": db_policy,
        "field_name": "product_name",
        "apl_value": apl_prod,
        "db_value": db_prod,
        "status": prod_st,
        "method": "STRING_COMPARISON",
        "diff_reason": "" if prod_st.endswith("MATCH") else f"Product mismatch: '{apl_prod}' vs '{db_prod}'",
        "review_required": prod_st == "DATA_MISMATCH",
    })

    # Writing Agent comparison
    db_wa_uid = primary_db.get("writing_agent_user_id")
    db_wa_display = user_display_fn(db_wa_uid) if db_wa_uid else ""

    name_match_res = compare_person_names(apl_wa_name, db_wa_display)
    wa_match_status = name_match_res["status"]

    if wa_match_status in ("EXACT_MATCH", "NORMALIZED_MATCH"):
        matched_fields += 1
    elif wa_match_status == "CLOSE_MATCH_REVIEW_REQUIRED":
        review_fields += 1
        review_reasons.append(f"Writing Agent name requires review: {name_match_res['reason']}")
    else:
        # If DB writing agent UUID is None or mismatch
        if not db_wa_uid:
            wa_match_status = "CLOSE_MATCH_REVIEW_REQUIRED"
            review_fields += 1
            review_reasons.append("Writing Agent UUID is null in DB policies table")
        else:
            mismatched_fields += 1
            mismatch_names.append("writing_agent")

    field_comps.append({
        "apl_policy": apl_policy,
        "db_policy": db_policy,
        "field_name": "writing_agent",
        "apl_value": apl_wa_name,
        "db_value": db_wa_display,
        "status": wa_match_status,
        "method": name_match_res.get("method", "NONE"),
        "diff_reason": name_match_res.get("reason", ""),
        "review_required": wa_match_status in ("CLOSE_MATCH_REVIEW_REQUIRED", "DATA_MISMATCH"),
    })

    # Intentionally non-persisted carrier audit columns
    non_persisted_cols = [
        ("Effective Date", "effective_date"),
        ("Policy Status", "policy_status"),
        ("Agent Status", "agent_status"),
        ("Earning", "earning"),
        ("Rate is same as default rate", "rate_is_same_as_default_rate"),
    ]
    for col_display, col_key in non_persisted_cols:
        val = primary_apl.get(col_display) or primary_apl.get(col_key)
        field_comps.append({
            "apl_policy": apl_policy,
            "db_policy": db_policy,
            "field_name": col_key,
            "apl_value": val,
            "db_value": "NOT_PERSISTED",
            "status": "NOT_PERSISTED_IN_DB",
            "method": "INTENTIONALLY_NOT_PERSISTED",
            "diff_reason": "Carrier audit status metadata is intentionally not persisted to policy_commission",
            "review_required": False,
        })

    # Commission Splits Comparison
    # One sheet month band against the commission row for that same band.
    # The stored value is the roll-up share, not the raw Rate column.
    prepared_rows = _prepared_apl_rows(apl_records)
    bands: Dict[Tuple[int, int], List[Dict[str, Any]]] = {}
    band_order: List[Tuple[int, int]] = []
    for r in prepared_rows:
        role = classify_contract_role(_read_apl_field(r, "Contract", "contract"))
        name = str(_read_apl_field(r, "Agent Name", "agent_name") or "").strip()
        code = clean_agent_code(_read_apl_field(r, "Agent ID", "agent_id") or "")
        raw_rate = parse_float_safe(r.get("_raw_rate") if r.get("_raw_rate") not in (None, "") else _read_apl_field(r, "Rate", "Rate %", "rate"), 0.0)
        try:
            seq = int(float(_read_apl_field(r, "Sequence", "sequence") or 0))
        except (TypeError, ValueError):
            seq = 0
        contract = _read_apl_field(r, "Contract", "contract") or ""
        lvl_raw = _read_apl_field(r, "Level", "level") or contract
        band = _band_key(
            _read_apl_field(r, "From Month", "from_month", "month from"),
            _read_apl_field(r, "To Month", "to_month", "month to"),
        )
        if band not in bands:
            band_order.append(band)
            bands[band] = []
        bands[band].append({
            "role": role,
            "level": normalize_level(lvl_raw),
            "name": name,
            "code": code,
            "rate": raw_rate,
            "share": parse_float_safe(r.get("_split_share"), raw_rate),
            "sequence": seq,
            "band": band,
        })

    db_by_band: Dict[Tuple[int, int], List[Dict[str, Any]]] = {}
    db_rules_with_splits = [rec for rec in db_records if _db_splits_for_record(rec)]
    for rec in db_rules_with_splits:
        db_by_band.setdefault(_band_key(rec.get("month_from"), rec.get("month_to")), []).extend(
            _db_splits_for_record(rec)
        )
    # Sheets with no month columns are one band. Use every split on the policy.
    if len(band_order) == 1 and band_order[0] == (1, 0) and not any(
        _read_apl_field(r, "From Month", "from_month", "month from") not in (None, "") for r in apl_records
    ):
        db_by_band = {(1, 0): [s for rec in db_rules_with_splits for s in _db_splits_for_record(rec)]}

    role_order = {"Agency": 1, "Sales Leader": 2, "Agent": 3, "Sub Agent": 4}
    apl_participants: List[Dict[str, Any]] = []
    for band in band_order:
        people = bands[band]
        people.sort(key=lambda x: (role_order.get(x["role"], 99), -x["sequence"]))
        apl_participants.extend(people)

    # Compare each APL split against the database split in the same month band
    for idx, p_apl in enumerate(apl_participants, 1):
        target_role = p_apl["role"]
        db_splits = db_by_band.get(p_apl["band"], [])
        matched_db_split = None
        role_hits = [
            s for s in db_splits
            if str(s.get("role") or "").strip().lower() == target_role.lower()
        ]
        if len(role_hits) == 1:
            matched_db_split = role_hits[0]
        elif role_hits:
            apl_name_key = " ".join(clean_name_ignoring_id(p_apl["name"]).casefold().split())
            for s in role_hits:
                db_name_key = " ".join(clean_name_ignoring_id(s.get("user_name")).casefold().split())
                if apl_name_key and apl_name_key == db_name_key:
                    matched_db_split = s
                    break
            if matched_db_split is None:
                matched_db_split = role_hits[0]

        if matched_db_split is None and db_splits and not role_hits:
            same_order = [
                s for s in db_splits
                if str(s.get("order") or "") == str(role_order.get(target_role, idx))
            ]
            if len(same_order) == 1:
                matched_db_split = same_order[0]

        if matched_db_split:
            db_role = str(matched_db_split.get("role") or "").strip()
            db_name = str(matched_db_split.get("user_name") or "").strip()
            db_val = parse_float_safe(matched_db_split.get("commission_value") or 0.0)
            db_uuid = str(matched_db_split.get("user_uuid") or "").strip()
            db_order = matched_db_split.get("order")
            db_level = normalize_level(matched_db_split.get("level"))

            # Check participant name match
            p_name_res = compare_person_names(p_apl["name"], db_name)
            p_match_st = p_name_res["status"]

            # Check role match
            role_matches = db_role.lower() == target_role.lower()
            val_matches = compare_decimals(p_apl["share"], db_val, tolerance=0.02)

            # Check level match if level was specified in DB split
            if db_level:
                lvl_cmp = compare_levels(p_apl["level"], db_level, role=target_role)
                level_matches = lvl_cmp["match"]
            else:
                level_matches = True
                lvl_cmp = {"match": True, "status": "MATCH", "db_level": ""}

            split_diffs = []
            if not role_matches:
                split_diffs.append(f"Role mismatch ({target_role} vs {db_role})")
            if db_level and not level_matches:
                split_diffs.append(f"Level mismatch ({lvl_cmp['disp']})")
            if not val_matches:
                fm, tm = p_apl["band"]
                split_diffs.append(
                    f"Months {fm}-{tm} split mismatch "
                    f"(sheet rate {p_apl['rate']}% → share {p_apl['share']}% vs DB {db_val}%)"
                )
            if p_match_st not in ("EXACT_MATCH", "NORMALIZED_MATCH"):
                split_diffs.append(f"Name mismatch ({p_name_res['reason']})")
            if not db_uuid and target_role != "Agency":
                split_diffs.append("Participant UUID missing in DB split")

            is_split_match = role_matches and val_matches and level_matches and (p_match_st in ("EXACT_MATCH", "NORMALIZED_MATCH"))
            if is_split_match:
                matched_fields += 1
                sp_status = "EXACT_DATA_MATCH" if (p_match_st == "EXACT_MATCH" and lvl_cmp["status"] == "MATCH") else "MATCH_AFTER_NORMALIZATION"
            else:
                mismatched_fields += 1
                sp_status = "DATA_MISMATCH"
                mismatch_names.append(f"split_{target_role.lower().replace(' ', '_')}")

            split_comps.append({
                "apl_policy": apl_policy,
                "db_policy": db_policy,
                "apl_name": p_apl["name"],
                "db_name": db_name,
                "apl_role": target_role,
                "db_role": db_role,
                "apl_level": p_apl["level"],
                "db_level": db_level,
                "apl_order": idx,
                "db_order": db_order,
                "apl_uuid": "",
                "db_uuid": db_uuid,
                "apl_commission_value": p_apl["share"],
                "db_commission_value": db_val,
                "participant_match_status": p_match_st,
                "field_comparison_result": sp_status,
                "review_required": not is_split_match,
                "review_reason": "; ".join(split_diffs),
            })
        else:
            mismatched_fields += 1
            mismatch_names.append(f"missing_split_{target_role.lower().replace(' ', '_')}")
            split_comps.append({
                "apl_policy": apl_policy,
                "db_policy": db_policy,
                "apl_name": p_apl["name"],
                "db_name": "",
                "apl_role": target_role,
                "db_role": "MISSING",
                "apl_order": idx,
                "db_order": None,
                "apl_uuid": "",
                "db_uuid": "",
                "apl_commission_value": p_apl["share"],
                "db_commission_value": 0.0,
                "participant_match_status": "MISSING_IN_DB",
                "field_comparison_result": "DATA_MISMATCH",
                "review_required": True,
                "review_reason": (
                    f"Months {p_apl['band'][0]}-{p_apl['band'][1]} "
                    f"role '{target_role}' is missing from that commission slab"
                ),
            })

    # Determine overall data validation status
    if mismatched_fields > 0:
        val_status = "DATA_MISMATCH"
    elif review_fields > 0 or pol_rev or wa_match_status == "CLOSE_MATCH_REVIEW_REQUIRED":
        val_status = "CLOSE_MATCH_REVIEW_REQUIRED"
    elif any(fc["status"] == "NORMALIZED_MATCH" for fc in field_comps) or any(sc["field_comparison_result"] == "MATCH_AFTER_NORMALIZATION" for sc in split_comps):
        val_status = "MATCH_AFTER_NORMALIZATION"
    else:
        val_status = "EXACT_DATA_MATCH"

    overall_status = compute_overall_status(exec_status, val_status)
    rev_req = val_status in ("CLOSE_MATCH_REVIEW_REQUIRED", "DATA_MISMATCH", "DB_RECORD_NOT_FOUND", "AMBIGUOUS_MATCH") or exec_status in ("FAILED", "REJECTED", "EXECUTED_WITH_ERRORS")

    summary = {
        "apl_policy": apl_policy,
        "db_policy": db_policy,
        "policy_match_status": p_status,
        "policy_match_method": policy_match.get("method", "EXACT"),
        "policy_similarity": policy_match.get("similarity", 1.0),
        "product_name": apl_prod,
        "carrier": "MLB",
        "agency": "Main Line Benefits (5)",
        "apl_agent_name": apl_agent_name,
        "db_agent_name": db_wa_display,
        "agent_match_status": wa_match_status,
        "apl_writing_agent_name": apl_wa_name,
        "db_writing_agent_name": db_wa_display,
        "writing_agent_match_status": wa_match_status,
        "execution_status": exec_status,
        "execution_evidence": exec_evidence,
        "data_validation_status": val_status,
        "overall_status": overall_status,
        "matched_fields_count": matched_fields,
        "mismatched_fields_count": mismatched_fields,
        "review_required_fields_count": review_fields,
        "mismatch_field_names": ", ".join(mismatch_names),
        "review_required": rev_req,
        "review_reason": "; ".join(review_reasons) if review_reasons else ("" if not mismatch_names else f"Mismatches: {', '.join(mismatch_names)}"),
    }

    return {
        "summary": summary,
        "field_comparisons": field_comps,
        "split_comparisons": split_comps,
    }


# ---------------------------------------------------------------------------
# Annotated APL File Generator
# ---------------------------------------------------------------------------
def annotate_apl_rows(
    df_raw: pd.DataFrame,
    db_records_map: Dict[str, List[Dict[str, Any]]],
    all_db_numbers: List[str],
) -> pd.DataFrame:
    """
    Preserves all original APL file columns in order and appends comprehensive audit columns at the end:
    - Month_Slab_Match: Specific month tier match status (e.g. MATCH (1-12))
    - Rate_Match: Specific rate/split-share match status (e.g. MATCH (70.47%))
    - Role_Match: Role classification match status (e.g. MATCH (Agent))
    - Agent_Name_Match: Name match status ignoring agent IDs in parentheses (e.g. EXACT_MATCH (Heidi Berthelsen))
    - Match_Status: EXACT_MATCH / NORMALIZED_MATCH / MISMATCH / POLICY_NOT_FOUND_IN_DB / NO_COMMISSION_RULE_IN_DB
    - Matched_Details: Summary of all matched elements across all dimensions
    - Mismatch_Reason: Explicit details of what mismatched across month, rate, role, agent, or 'None (All matched)'
    """
    df_out = df_raw.copy()

    # Pre-calculate differential rollup split shares using the exact data harmonization pipeline
    rolled_tier_lookup = {}
    try:
        from validate_policy_commission_db import (
            fix_inclusive_month_band_overlaps,
            normalize_year_bands,
            rollup_commission_splits,
        )
        pol_col_temp = next((c for c in df_raw.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno")), "Policy No")
        from_col_temp = next((c for c in df_raw.columns if "from" in str(c).lower()), "From Month")
        contract_col_temp = next((c for c in df_raw.columns if "contract" in str(c).lower()), "Contract")
        agent_col_temp = next((c for c in df_raw.columns if "agent name" in str(c).lower()), "Agent Name")
        id_col_temp = next((c for c in df_raw.columns if "agent id" in str(c).lower()), "Agent ID")
        rate_col_temp = next((c for c in df_raw.columns if "rate" in str(c).lower()), "Rate")

        if "Normalized_Split_Share_%" in df_raw.columns:
            df_rolled = df_raw.reset_index(drop=True)
        else:
            df_clean_in = df_raw.reset_index(drop=True)
            df_fixed, _ = fix_inclusive_month_band_overlaps(df_clean_in)
            df_normalized, _ = normalize_year_bands(df_fixed)
            df_rolled = rollup_commission_splits(df_normalized)

        for _, r_r in df_rolled.iterrows():
            p_k = clean_policy_number(r_r.get(pol_col_temp))
            f_k = r_r.get(from_col_temp)
            try:
                f_i = int(float(str(f_k).strip())) if f_k is not None and str(f_k).strip() != "" else 1
            except Exception:
                f_i = 1
            c_role = classify_contract_role(r_r.get(contract_col_temp))
            c_name = clean_name_ignoring_id(r_r.get(agent_col_temp))
            c_id = clean_agent_code(r_r.get(id_col_temp) or "").lower()
            c_share = float(r_r.get("Normalized_Split_Share_%") or parse_float_safe(r_r.get(rate_col_temp), 0.0))
            rolled_tier_lookup.setdefault((p_k, f_i), []).append({
                "role": c_role,
                "name": c_name,
                "code": c_id,
                "share": c_share,
            })
    except Exception:
        df_rolled = df_raw.reset_index(drop=True)

    pol_col = next((c for c in df_out.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno")), "Policy No")
    from_col = next((c for c in df_out.columns if "from" in str(c).lower()), "From Month")
    to_col = next((c for c in df_out.columns if "to" in str(c).lower()), "To Month")
    agent_col = next((c for c in df_out.columns if "agent name" in str(c).lower()), "Agent Name")
    id_col = next((c for c in df_out.columns if "agent id" in str(c).lower()), "Agent ID")
    contract_col = next((c for c in df_out.columns if "contract" in str(c).lower()), "Contract")
    level_col = next((c for c in df_out.columns if str(c).strip().lower() in ("level", "agent level", "lvl", "contract level")), None)
    rate_col = next((c for c in df_out.columns if "rate" in str(c).lower()), "Rate")
    amt_col = next((c for c in df_out.columns if "amount" in str(c).lower()), "Amount")

    try:
        from validate_policy_commission_db import fill_blank_contracts_from_siblings
        df_filled_contracts = fill_blank_contracts_from_siblings(df_out)
    except Exception:
        df_filled_contracts = df_out

    month_matches = []
    rate_matches = []
    role_matches = []
    level_matches = []
    name_matches = []
    split_sum_matches = []
    statuses = []
    matched_details = []
    mismatch_reasons = []

    # Pre-calculate file tier split sums per (clean_policy, fm_int)
    file_tier_sums = {}
    if rolled_tier_lookup:
        for tier_k, parts in rolled_tier_lookup.items():
            file_tier_sums[tier_k] = round(sum(p["share"] for p in parts), 2)
    else:
        for idx_r, r_r in df_out.iterrows():
            p_c = clean_policy_number(r_r.get(pol_col))
            f_m = r_r.get(from_col)
            try:
                f_i = int(float(str(f_m).strip())) if f_m is not None and str(f_m).strip() != "" else 1
            except Exception:
                f_i = 1

            idx_p = df_out.index.get_loc(idx_r) if hasattr(df_out.index, "get_loc") else idx_r
            r_rolled = df_rolled.iloc[idx_p] if idx_p < len(df_rolled) else {}
            val_share = float(r_r.get("Normalized_Split_Share_%") or r_rolled.get("Normalized_Split_Share_%") or parse_float_safe(r_r.get(rate_col), 0.0))

            tier_k = (p_c, f_i)
            file_tier_sums[tier_k] = file_tier_sums.get(tier_k, 0.0) + val_share

    for idx, row in df_out.iterrows():
        pol_raw = row.get(pol_col)
        pol_clean = clean_policy_number(pol_raw)
        agent_name = clean_name_ignoring_id(row.get(agent_col))
        agent_id = str(row.get(id_col) or "").strip()
        contract_cell = row.get(contract_col)
        if contract_col in df_filled_contracts.columns and idx in df_filled_contracts.index:
            contract_cell = df_filled_contracts.at[idx, contract_col]
        contract = "" if contract_cell is None or (isinstance(contract_cell, float) and pd.isna(contract_cell)) else str(contract_cell).strip()
        if contract.lower() in ("nan", "none", "null"):
            contract = ""
        raw_level = str(row.get(level_col) or "").strip() if level_col else contract
        if raw_level.lower() in ("nan", "none", "null"):
            raw_level = contract
        role = classify_contract_role(contract)
        fm_raw = row.get(from_col)
        tm_raw = row.get(to_col)
        rate_raw = row.get(rate_col)
        rate_val = parse_float_safe(rate_raw)
        amt_val = parse_float_safe(row.get(amt_col), 0.0)

        # Pre-computed differential rollup share fallback
        idx_pos = df_out.index.get_loc(idx) if hasattr(df_out.index, "get_loc") else idx
        rolled_row = df_rolled.iloc[idx_pos] if idx_pos < len(df_rolled) else {}
        rollup_val = float(row.get("Normalized_Split_Share_%") or rolled_row.get("Normalized_Split_Share_%") or rate_val)

        # Policy matching
        pol_match = match_policy_identifier(pol_clean, all_db_numbers)
        db_pol = pol_match.get("db_policy")

        if not db_pol or db_pol not in db_records_map:
            if pol_match.get("status") == "AMBIGUOUS_MATCH":
                st = "AMBIGUOUS_MATCH"
                reason = pol_match.get("reason", "Ambiguous DB candidate match")
            else:
                st = "POLICY_NOT_FOUND_IN_DB"
                reason = f"Policy '{pol_clean}' not found in database policies table"

            month_matches.append("NO_POLICY")
            rate_matches.append("NO_POLICY")
            role_matches.append("NO_POLICY")
            level_matches.append("NO_POLICY")
            name_matches.append("NO_POLICY")
            split_sum_matches.append("NO_POLICY")
            statuses.append(st)
            matched_details.append("None")
            mismatch_reasons.append(reason)
            continue

        db_recs = db_records_map[db_pol]
        if not db_recs or all(r.get("commission_uuid") is None for r in db_recs):
            month_matches.append("NO_RULES")
            rate_matches.append("NO_RULES")
            role_matches.append("NO_RULES")
            level_matches.append("NO_RULES")
            name_matches.append("NO_RULES")
            split_sum_matches.append("NO_RULES")
            statuses.append("NO_COMMISSION_RULE_IN_DB")
            matched_details.append(f"Policy No ({pol_match.get('status', 'EXACT_MATCH')})")
            mismatch_reasons.append("Policy exists in DB, but has 0 commission rules in policy_commission table")
            continue

        # Match month tier
        try:
            fm_int = int(float(str(fm_raw).strip())) if fm_raw is not None and str(fm_raw).strip() != "" else 1
        except Exception:
            fm_int = 1

        matching_rule = None
        for r in db_recs:
            db_fm = int(r.get("month_from") or 1)
            if db_fm == fm_int:
                matching_rule = r
                break

        if not matching_rule:
            # Gracefully handle expired 0% commission tiers not stored in DB
            if rate_val == 0.0 and amt_val == 0.0:
                month_matches.append(f"EXPIRED_TIER ({fm_raw}-{tm_raw} 0% - Not persisted in DB)")
                rate_matches.append("MATCH (0.00%)")
                role_matches.append(f"MATCH ({role})")
                level_matches.append(f"MATCH ({normalize_level(raw_level or contract)})")
                name_matches.append(f"MATCH ({agent_name})")
                split_sum_matches.append("N/A (Expired Tier)")
                statuses.append("EXACT_MATCH")
                matched_details.append(f"Policy No ({pol_match.get('status', 'EXACT_MATCH')}); Expired Tier ({fm_raw}-{tm_raw} 0.00%)")
                mismatch_reasons.append("None (Expired 0.00% tier intentionally omitted from DB)")
                continue

            month_matches.append(f"MISMATCH (File: {fm_raw}-{tm_raw} not in DB)")
            rate_matches.append("NOT_CHECKED")
            role_matches.append("NOT_CHECKED")
            level_matches.append("NOT_CHECKED")
            name_matches.append("NOT_CHECKED")
            split_sum_matches.append("NOT_CHECKED")
            statuses.append("MISMATCH")
            matched_details.append(f"Policy No ({pol_match.get('status', 'EXACT_MATCH')})")
            mismatch_reasons.append(f"Month slab ({fm_raw}-{tm_raw}) not found in DB commission tiers")
            continue

        db_fm_int = int(matching_rule.get('month_from') or fm_int)
        db_tier_disp = f"{matching_rule.get('month_from')}-{matching_rule.get('month_to') or 0}"
        month_slab_str = f"MATCH ({db_tier_disp})"
        clean_c = clean_agent_code(agent_id).lower()

        # Look up pre-computed differential rollup share for this participant in this tier
        tier_parts = rolled_tier_lookup.get((pol_clean, db_fm_int), rolled_tier_lookup.get((pol_clean, fm_int), []))
        cand_share = None
        for tp in tier_parts:
            if tp["role"].lower() == role.lower():
                name_cmp = compare_person_names(agent_name, tp["name"])
                if (clean_c and clean_c == tp["code"]) or name_cmp["status"] in ("EXACT_MATCH", "NORMALIZED_MATCH"):
                    cand_share = tp["share"]
                    break
        if cand_share is None and (rate_val > 0.0 or amt_val > 0.0):
            for tp in tier_parts:
                if tp["role"].lower() == role.lower():
                    cand_share = tp["share"]
                    break
        if cand_share is None:
            for tp in tier_parts:
                name_cmp = compare_person_names(agent_name, tp["name"])
                if (clean_c and clean_c == tp["code"]) or name_cmp["status"] in ("EXACT_MATCH", "NORMALIZED_MATCH"):
                    cand_share = tp["share"]
                    break

        if cand_share is not None:
            rollup_val = cand_share

        # Verify DB commission tier splits sum matches APL file tier splits sum (and both mostly 100.00%)
        splits = matching_rule.get("commission_split") or []
        db_tier_split_sum = round(sum(parse_float_safe(s.get("commission_value"), 0.0) for s in splits if isinstance(s, dict)), 2)
        apl_tier_split_sum = round(file_tier_sums.get((pol_clean, db_fm_int), file_tier_sums.get((pol_clean, fm_int), rollup_val)), 2)

        sums_match = abs(apl_tier_split_sum - db_tier_split_sum) <= 0.05
        is_100_percent = abs(db_tier_split_sum - 100.00) <= 0.05 and abs(apl_tier_split_sum - 100.00) <= 0.05

        if sums_match and is_100_percent:
            split_sum_str = f"MATCH (100.00%)"
        elif sums_match:
            split_sum_str = f"MATCH ({db_tier_split_sum:.2f}% [APL: {apl_tier_split_sum:.2f}% == DB: {db_tier_split_sum:.2f}%])"
        else:
            split_sum_str = f"MISMATCH (File Sum: {apl_tier_split_sum:.2f}% vs DB Sum: {db_tier_split_sum:.2f}%)"

        # Find participant in DB splits with priority matching:
        # Pass 1: Prioritize matching BOTH role AND identity (agent code or agent name)
        # Pass 2: Role alone (if agent name/code didn't match directly)
        # Pass 3: Identity alone (fallback if role title differs in DB)
        target_split = None

        # Pass 1: Role AND (code or name)
        for s in splits:
            s_role = str(s.get("role") or "").strip()
            if s_role.lower() != role.lower():
                continue
            s_name = clean_name_ignoring_id(s.get("user_name"))
            s_code = clean_agent_code(s.get("user_code") or "").lower()
            name_cmp = compare_person_names(agent_name, s_name)
            if (clean_c and clean_c == s_code) or name_cmp["status"] in ("EXACT_MATCH", "NORMALIZED_MATCH"):
                target_split = s
                break

        # Pass 2: Role match (only if rate > 0 or amt > 0, to avoid falsely matching a 0% row to a non-zero split)
        if not target_split and (rate_val > 0.0 or amt_val > 0.0):
            for s in splits:
                s_role = str(s.get("role") or "").strip()
                if s_role.lower() == role.lower():
                    target_split = s
                    break

        # Pass 3: Identity match (fallback)
        if not target_split:
            for s in splits:
                s_name = clean_name_ignoring_id(s.get("user_name"))
                s_code = clean_agent_code(s.get("user_code") or "").lower()
                name_cmp = compare_person_names(agent_name, s_name)
                if (clean_c and clean_c == s_code) or name_cmp["status"] in ("EXACT_MATCH", "NORMALIZED_MATCH"):
                    target_split = s
                    break

        if not target_split:
            # If the split in APL is 0.00% and Amount is $0.00, it is intentionally omitted from DB
            if (rate_val == 0.0 or rollup_val == 0.0) and amt_val == 0.0:
                month_matches.append(month_slab_str)
                rate_matches.append("MATCH (0.00%)")
                role_matches.append(f"MATCH ({role})")
                level_matches.append(f"MATCH ({normalize_level(raw_level or contract)})")
                name_matches.append(f"MATCH ({agent_name})")
                split_sum_matches.append(split_sum_str)
                statuses.append("EXACT_MATCH")
                matched_details.append(f"Policy No ({pol_match.get('status', 'EXACT_MATCH')}); Month Slab ({db_tier_disp}); 0% Split (Omitted from DB)")
                mismatch_reasons.append("None (0.00% split intentionally omitted from DB)")
                continue

            month_matches.append(month_slab_str)
            rate_matches.append("NOT_FOUND")
            role_matches.append(f"MISMATCH (Role '{role}' missing in DB)")
            level_matches.append(f"MISMATCH (Level '{normalize_level(raw_level or contract)}' missing in DB)")
            name_matches.append(f"MISMATCH (Agent '{agent_name}' missing in DB)")
            split_sum_matches.append(split_sum_str)
            statuses.append("MISMATCH")
            matched_details.append(f"Policy No ({pol_match.get('status', 'EXACT_MATCH')}); Month Slab ({db_tier_disp})")
            mismatch_reasons.append(f"Participant '{agent_name}' ({role}) missing from DB commission split")
            continue

        db_role = str(target_split.get("role") or "").strip()
        db_name_raw = str(target_split.get("user_name") or "").strip()
        db_name = clean_name_ignoring_id(db_name_raw)
        db_val = float(target_split.get("commission_value") or 0.0)
        db_lvl_raw = target_split.get("level")

        # 1. Role match
        role_ok = db_role.lower() == role.lower()
        role_str = f"MATCH ({role})" if role_ok else f"MISMATCH (File: '{role}' vs DB: '{db_role}')"

        # 2. Level match (zero-compromise mapping of LVL / Contract to DB Level)
        lvl_cmp = compare_levels(raw_level or contract, db_lvl_raw, role=role)
        if (rate_val == 0.0 or db_val == 0.0) and amt_val == 0.0:
            level_ok = True
            level_str = f"MATCH ({normalize_level(db_lvl_raw) or lvl_cmp['db_level']})"
        else:
            level_ok = lvl_cmp["match"]
            level_str = lvl_cmp["disp"]

        # 3. Agent Name match (ignoring ID in parentheses, prefix codes, and generational suffixes)
        name_res = compare_person_names(agent_name, db_name)
        if name_res["status"] == "EXACT_MATCH":
            name_str = f"EXACT_MATCH ({db_name})"
        elif name_res["status"] == "NORMALIZED_MATCH":
            name_str = f"NORMALIZED_MATCH ({db_name})"
        elif name_res["status"] == "CLOSE_MATCH_REVIEW_REQUIRED":
            name_str = f"CLOSE_MATCH ({db_name})"
        else:
            name_str = f"MISMATCH (File: '{agent_name}' vs DB: '{db_name_raw}')"

        # 4. Rate match (checks against direct rate OR differential rollup share)
        val_direct_ok = compare_decimals(rate_val, db_val, tolerance=0.01)
        val_rollup_ok = compare_decimals(rollup_val, db_val, tolerance=0.01)
        rate_ok = val_direct_ok or val_rollup_ok

        if val_direct_ok:
            rate_str = f"MATCH ({rate_raw})"
        elif val_rollup_ok:
            rate_str = f"MATCH ({db_val:.2f}% [Rollup: {rollup_val:.2f}%])"
        else:
            if abs(rollup_val - rate_val) > 0.001:
                rate_str = f"MISMATCH (File: {rate_raw} → share {rollup_val:.2f}% vs DB: {db_val:.2f}%)"
            else:
                rate_str = f"MISMATCH (File: {rate_raw} vs DB: {db_val:.2f}%)"

        matched_items = [f"Policy No ({pol_match.get('status', 'EXACT_MATCH')})", f"Month Slab ({db_tier_disp})"]
        mismatches = []

        if role_ok:
            matched_items.append(f"Role ({role})")
        else:
            mismatches.append(role_str)

        if level_ok:
            matched_items.append(f"Level ({lvl_cmp['db_level']})")
        else:
            mismatches.append(level_str)

        if name_res["status"] in ("EXACT_MATCH", "NORMALIZED_MATCH"):
            matched_items.append(f"Agent Name ({db_name})")
        else:
            mismatches.append(name_str)

        if rate_ok:
            matched_items.append(f"Rate ({db_val:.2f}%)")
        else:
            mismatches.append(rate_str)

        if sums_match:
            matched_items.append(f"Split Sum ({db_tier_split_sum:.2f}%)")
        else:
            mismatches.append(f"Split Sum {split_sum_str}")

        if not mismatches:
            st = "EXACT_MATCH" if (pol_match.get("status") == "EXACT_MATCH" and name_res["status"] == "EXACT_MATCH" and lvl_cmp["status"] == "MATCH") else "NORMALIZED_MATCH"
            reason = "None (All matched)"
        else:
            st = "MISMATCH"
            reason = "; ".join(mismatches)

        month_matches.append(month_slab_str)
        rate_matches.append(rate_str)
        role_matches.append(role_str)
        level_matches.append(level_str)
        name_matches.append(name_str)
        split_sum_matches.append(split_sum_str)
        statuses.append(st)
        matched_details.append("; ".join(matched_items))
        mismatch_reasons.append(reason)

    df_out["Month_Slab_Match"] = month_matches
    df_out["Rate_Match"] = rate_matches
    df_out["Role_Match"] = role_matches
    df_out["Level_Match"] = level_matches
    df_out["Agent_Name_Match"] = name_matches
    df_out["Split_Sum_Match"] = split_sum_matches
    df_out["Match_Status"] = statuses
    df_out["Matched_Details"] = matched_details
    df_out["Mismatch_Reason"] = mismatch_reasons
    return df_out


# ---------------------------------------------------------------------------
# Professional Multi-Tab Excel Workbook Generator
# ---------------------------------------------------------------------------
def generate_excel_validation_workbook(
    summaries: List[Dict[str, Any]],
    field_comparisons: List[Dict[str, Any]],
    split_comparisons: List[Dict[str, Any]],
    output_path: str,
    run_id: str,
    df_annotated: Optional[pd.DataFrame] = None,
    logger: Optional[logging.Logger] = None,
) -> str:
    """
    Generates formatted 5-sheet Excel validation report with OpenPyXL.
    Protects against formula injection, freezes header panes, and sets auto-filter.
    """
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl library is required to generate formatted Excel validation workbook.")

    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Styles
    navy_header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    teal_header_fill = PatternFill(start_color="1B7575", end_color="1B7575", fill_type="solid")
    sub_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    white_bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    dark_bold_font = Font(name="Calibri", size=11, bold=True, color="000000")
    regular_font = Font(name="Calibri", size=11, color="000000")
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )

    # Fills for status highlights
    success_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    warning_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    error_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    # -------------------------------------------------------------------------
    # Sheet 1: Summary
    # -------------------------------------------------------------------------
    ws_sum = wb.create_sheet(title="Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.cell(row=1, column=1, value="POST-IMPORT APL VS DATABASE VALIDATION REPORT").font = Font(size=14, bold=True, color="1F4E79")
    ws_sum.cell(row=2, column=1, value=f"Import Run ID: {run_id} | Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}").font = Font(size=10, italic=True)

    # Calculate summary metrics
    total_pols = len(summaries)
    exec_success = sum(1 for s in summaries if s["execution_status"] == "EXECUTED_SUCCESSFULLY")
    exec_errors = sum(1 for s in summaries if s["execution_status"] == "EXECUTED_WITH_ERRORS")
    exec_skipped = sum(1 for s in summaries if s["execution_status"] == "SKIPPED")
    exec_rejected = sum(1 for s in summaries if s["execution_status"] == "REJECTED")
    exec_failed = sum(1 for s in summaries if s["execution_status"] == "FAILED")
    exec_unknown = sum(1 for s in summaries if s["execution_status"] == "UNKNOWN")

    pol_exact = sum(1 for s in summaries if s["policy_match_status"] == "EXACT_MATCH")
    pol_norm = sum(1 for s in summaries if s["policy_match_status"] == "NORMALIZED_MATCH")
    pol_close = sum(1 for s in summaries if s["policy_match_status"].startswith("CLOSE_MATCH"))
    pol_ambig = sum(1 for s in summaries if s["policy_match_status"] == "AMBIGUOUS_MATCH")
    pol_not_found = sum(1 for s in summaries if s["policy_match_status"] == "NO_MATCH" or s["data_validation_status"] == "DB_RECORD_NOT_FOUND")

    data_exact = sum(1 for s in summaries if s["data_validation_status"] == "EXACT_DATA_MATCH")
    data_norm = sum(1 for s in summaries if s["data_validation_status"] == "MATCH_AFTER_NORMALIZATION")
    data_close = sum(1 for s in summaries if s["data_validation_status"] == "CLOSE_MATCH_REVIEW_REQUIRED")
    data_mismatch = sum(1 for s in summaries if s["data_validation_status"] == "DATA_MISMATCH")

    overall_exact_success = sum(1 for s in summaries if s["overall_status"] == "SUCCESS_EXACT_MATCH")
    overall_norm_success = sum(1 for s in summaries if s["overall_status"] == "SUCCESS_NORMALIZED_MATCH")
    overall_review = sum(1 for s in summaries if s["review_required"])

    metric_rows = [
        ("EXECUTION STATUS METRICS", ""),
        ("Total APL Policies Analyzed", total_pols),
        ("Policies with Confirmed Successful Execution", exec_success),
        ("Policies with Execution Errors", exec_errors),
        ("Policies Skipped by User / Importer", exec_skipped),
        ("Policies Rejected (Missing UUID / Schema)", exec_rejected),
        ("Policies Failed / Timed Out", exec_failed),
        ("Policies with Unknown Execution Status", exec_unknown),
        ("", ""),
        ("POLICY IDENTIFIER MATCHING METRICS", ""),
        ("Exact Policy Identity Matches", pol_exact),
        ("Normalized Policy Identity Matches", pol_norm),
        ("Close Policy Matches Requiring Review", pol_close),
        ("Ambiguous Policy Matches", pol_ambig),
        ("DB Policies Not Found", pol_not_found),
        ("", ""),
        ("PERSISTED DATA VALIDATION METRICS", ""),
        ("Exact Persisted Data Matches", data_exact),
        ("Normalized Data Matches (Extra Middle Names / Formatting)", data_norm),
        ("Data Values Requiring Review", data_close),
        ("Confirmed Data Mismatches", data_mismatch),
        ("", ""),
        ("OVERALL AUDIT OUTCOME", ""),
        ("Verified Exact Match Successes", overall_exact_success),
        ("Verified Normalized Match Successes", overall_norm_success),
        ("Policies Requiring Manual Engineering Review", overall_review),
    ]

    r_idx = 4
    for label, val in metric_rows:
        if not label and not val:
            r_idx += 1
            continue
        c1 = ws_sum.cell(row=r_idx, column=1, value=label)
        c2 = ws_sum.cell(row=r_idx, column=2, value=val)
        if isinstance(val, str) and not val:
            # Section header
            c1.font = dark_bold_font
            c1.fill = sub_header_fill
            c2.fill = sub_header_fill
        else:
            c1.font = regular_font
            c2.font = dark_bold_font
            c2.alignment = Alignment(horizontal="right")
        r_idx += 1

    # Overlap explanation note
    ws_sum.cell(row=r_idx + 1, column=1, value="Note: Policy identity counts, execution statuses, and data validation counts evaluate distinct dimensions and sum independently.").font = Font(size=9, italic=True)

    # -------------------------------------------------------------------------
    # Sheet 2: Policy Validation
    # -------------------------------------------------------------------------
    ws_pol = wb.create_sheet(title="Policy Validation")
    pol_cols = [
        "APL Policy Number", "DB Policy Number", "Policy Match Status", "Policy Match Method", "Similarity Score",
        "Product Name", "Carrier", "Agency", "APL Agent Name", "DB Agent Name", "Agent Match Status",
        "APL Writing Agent Name", "DB Writing Agent Name", "Writing Agent Match Status",
        "Execution Status", "Execution Evidence", "Data Validation Status", "Overall Status",
        "Matched Fields", "Mismatched Fields", "Review Fields", "Mismatch Field Names",
        "Review Required", "Review Reason"
    ]
    for col_idx, col_name in enumerate(pol_cols, 1):
        cell = ws_pol.cell(row=1, column=col_idx, value=col_name)
        cell.fill = navy_header_fill
        cell.font = white_bold_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, s in enumerate(summaries, 2):
        row_vals = [
            sanitize_for_excel(s["apl_policy"]),
            sanitize_for_excel(s["db_policy"]),
            s["policy_match_status"],
            s["policy_match_method"],
            s["policy_similarity"],
            sanitize_for_excel(s["product_name"]),
            s["carrier"],
            s["agency"],
            sanitize_for_excel(s["apl_agent_name"]),
            sanitize_for_excel(s["db_agent_name"]),
            s["agent_match_status"],
            sanitize_for_excel(s["apl_writing_agent_name"]),
            sanitize_for_excel(s["db_writing_agent_name"]),
            s["writing_agent_match_status"],
            s["execution_status"],
            sanitize_for_excel(s["execution_evidence"]),
            s["data_validation_status"],
            s["overall_status"],
            s["matched_fields_count"],
            s["mismatched_fields_count"],
            s["review_required_fields_count"],
            sanitize_for_excel(s["mismatch_field_names"]),
            "YES" if s["review_required"] else "NO",
            sanitize_for_excel(s["review_reason"]),
        ]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_pol.cell(row=row_idx, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border

            # Highlight Overall Status
            if col_idx == 18:
                if "SUCCESS" in str(val):
                    cell.fill = success_fill
                elif "REVIEW" in str(val) or "UNKNOWN" in str(val):
                    cell.fill = warning_fill
                else:
                    cell.fill = error_fill

    ws_pol.freeze_panes = "A2"
    ws_pol.auto_filter.ref = f"A1:{get_column_letter(len(pol_cols))}{len(summaries) + 1}"

    # -------------------------------------------------------------------------
    # Sheet 3: Field Comparison
    # -------------------------------------------------------------------------
    ws_fld = wb.create_sheet(title="Field Comparison")
    fld_cols = [
        "APL Policy Number", "DB Policy Number", "Field Name", "APL Value", "Persisted DB Value",
        "Comparison Status", "Match Method", "Difference or Reason", "Review Required"
    ]
    for col_idx, col_name in enumerate(fld_cols, 1):
        cell = ws_fld.cell(row=1, column=col_idx, value=col_name)
        cell.fill = teal_header_fill
        cell.font = white_bold_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, fc in enumerate(field_comparisons, 2):
        row_vals = [
            sanitize_for_excel(fc["apl_policy"]),
            sanitize_for_excel(fc["db_policy"]),
            fc["field_name"],
            sanitize_for_excel(fc["apl_value"]),
            sanitize_for_excel(fc["db_value"]),
            fc["status"],
            fc["method"],
            sanitize_for_excel(fc["diff_reason"]),
            "YES" if fc["review_required"] else "NO",
        ]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_fld.cell(row=row_idx, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if col_idx == 6:
                if "MATCH" in str(val):
                    cell.fill = success_fill
                elif "REVIEW" in str(val):
                    cell.fill = warning_fill
                elif str(val) == "NOT_PERSISTED_IN_DB":
                    pass
                else:
                    cell.fill = error_fill

    ws_fld.freeze_panes = "A2"
    ws_fld.auto_filter.ref = f"A1:{get_column_letter(len(fld_cols))}{len(field_comparisons) + 1}"

    # -------------------------------------------------------------------------
    # Sheet 4: Commission Split Validation
    # -------------------------------------------------------------------------
    ws_splt = wb.create_sheet(title="Commission Split Validation")
    split_cols = [
        "APL Policy Number", "DB Policy Number", "APL Participant Name", "DB Participant Name",
        "APL Role", "DB Role", "APL Level", "DB Level", "APL Split Order", "DB Split Order", "APL UUID", "DB UUID",
        "APL Commission %", "DB Commission %", "Participant Match Status",
        "Field Comparison Result", "Review Required", "Review Reason"
    ]
    for col_idx, col_name in enumerate(split_cols, 1):
        cell = ws_splt.cell(row=1, column=col_idx, value=col_name)
        cell.fill = navy_header_fill
        cell.font = white_bold_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, sc in enumerate(split_comparisons, 2):
        row_vals = [
            sanitize_for_excel(sc["apl_policy"]),
            sanitize_for_excel(sc["db_policy"]),
            sanitize_for_excel(sc["apl_name"]),
            sanitize_for_excel(sc["db_name"]),
            sc["apl_role"],
            sc["db_role"],
            sc.get("apl_level", ""),
            sc.get("db_level", ""),
            sc["apl_order"],
            sc["db_order"],
            sanitize_for_excel(sc["apl_uuid"]),
            sanitize_for_excel(sc["db_uuid"]),
            sc["apl_commission_value"],
            sc["db_commission_value"],
            sc["participant_match_status"],
            sc["field_comparison_result"],
            "YES" if sc["review_required"] else "NO",
            sanitize_for_excel(sc["review_reason"]),
        ]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_splt.cell(row=row_idx, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if col_idx == 16:
                if "MATCH" in str(val):
                    cell.fill = success_fill
                elif "REVIEW" in str(val):
                    cell.fill = warning_fill
                else:
                    cell.fill = error_fill

    ws_splt.freeze_panes = "A2"
    ws_splt.auto_filter.ref = f"A1:{get_column_letter(len(split_cols))}{len(split_comparisons) + 1}"

    # -------------------------------------------------------------------------
    # Sheet 5: Exceptions
    # -------------------------------------------------------------------------
    ws_exc = wb.create_sheet(title="Exceptions")
    exc_cols = [
        "APL Policy Number", "DB Policy Number", "Policy Match Status", "Execution Status",
        "Data Validation Status", "Overall Status", "Issue Type", "Reason & Audit Details", "Recommended Action"
    ]
    for col_idx, col_name in enumerate(exc_cols, 1):
        cell = ws_exc.cell(row=1, column=col_idx, value=col_name)
        cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        cell.font = white_bold_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    exceptions = [s for s in summaries if s["review_required"]]
    for row_idx, ex in enumerate(exceptions, 2):
        issue_type = ex["data_validation_status"]
        action = "Review policy mapping and manual DB sync"
        if ex["data_validation_status"] == "DB_RECORD_NOT_FOUND":
            action = "Verify if policy has been ingested into core policies table"
        elif ex["data_validation_status"] == "AMBIGUOUS_MATCH":
            action = "Check duplicate candidate policy numbers in database"
        elif ex["policy_match_status"].startswith("CLOSE_MATCH"):
            action = "Confirm if policy number typo or suffix should be linked"
        elif ex["execution_status"] in ("REJECTED", "FAILED"):
            action = "Inspect missing user UUIDs in users table and re-run importer"

        row_vals = [
            sanitize_for_excel(ex["apl_policy"]),
            sanitize_for_excel(ex["db_policy"]),
            ex["policy_match_status"],
            ex["execution_status"],
            ex["data_validation_status"],
            ex["overall_status"],
            issue_type,
            sanitize_for_excel(ex["review_reason"]),
            action,
        ]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_exc.cell(row=row_idx, column=col_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            cell.fill = warning_fill if "REVIEW" in str(ex["overall_status"]) else error_fill

    # -------------------------------------------------------------------------
    # Sheet 6: Annotated APL (Original APL rows with Match_Status, Matched_Details, Mismatch_Reason)
    # -------------------------------------------------------------------------
    all_worksheets = [ws_sum, ws_pol, ws_fld, ws_splt, ws_exc]
    if df_annotated is not None and not df_annotated.empty:
        ws_apl = wb.create_sheet(title="Annotated APL")
        apl_headers = list(df_annotated.columns)
        for col_idx, col_name in enumerate(apl_headers, 1):
            cell = ws_apl.cell(row=1, column=col_idx, value=col_name)
            cell.fill = navy_header_fill
            cell.font = white_bold_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        status_col_idx = apl_headers.index("Match_Status") + 1 if "Match_Status" in apl_headers else len(apl_headers) - 2
        for row_idx, r in enumerate(df_annotated.itertuples(index=False), 2):
            for col_idx, val in enumerate(r, 1):
                cell = ws_apl.cell(row=row_idx, column=col_idx, value=sanitize_for_excel(val))
                cell.font = regular_font
                cell.border = thin_border
                # Highlight Match_Status column
                if col_idx == status_col_idx:
                    s_val = str(val).upper()
                    if "EXACT" in s_val or s_val == "NORMALIZED_MATCH":
                        cell.fill = success_fill
                    elif "REVIEW" in s_val or "UNKNOWN" in s_val:
                        cell.fill = warning_fill
                    else:
                        cell.fill = error_fill

        ws_apl.freeze_panes = "A2"
        ws_apl.auto_filter.ref = f"A1:{get_column_letter(len(apl_headers))}{len(df_annotated) + 1}"
        all_worksheets.append(ws_apl)

    # Auto-adjust column widths across all sheets
    for ws in all_worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if "\n" in val_str:
                    val_str = max(val_str.split("\n"), key=len)
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 55)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    if logger:
        logger.info(f"Excel Validation Report successfully created at: {output_path}")
    return output_path


# ---------------------------------------------------------------------------
# Main Orchestration Pipeline
# ---------------------------------------------------------------------------
def run_post_import_validation(
    input_file_path: str,
    output_excel_path: Optional[str] = None,
    run_id: Optional[str] = None,
    execution_report_path: Optional[str] = None,
    agency_id: int = 5,
    dry_run: bool = False,
    db_manager: Optional[ReadOnlyDatabaseManager] = None,
    logger: Optional[logging.Logger] = None,
) -> Dict[str, Any]:
    """
    Main entry point for post-import validation.
    Performs read-only validation against PostgreSQL database and outputs Excel report.
    """
    logger = logger or setup_logger(verbose=False)

    if not run_id:
        run_id = f"run_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"

    logger.info("=" * 80)
    logger.info("🚀 POST-IMPORT APL VS DATABASE VALIDATION ENGINE")
    logger.info("=" * 80)
    logger.info(f"Input APL File:       {input_file_path}")
    logger.info(f"Run ID:               {run_id}")
    logger.info(f"Agency ID:            {agency_id}")
    logger.info(f"Dry Run Mode:         {dry_run}")
    logger.info(f"Execution Report:     {execution_report_path or 'None'}")

    # 1. Load APL file
    if not os.path.exists(input_file_path):
        raise FileNotFoundError(f"APL input file does not exist at '{input_file_path}'")

    _, ext = os.path.splitext(input_file_path.lower())
    if ext in (".xlsx", ".xls"):
        df_raw = pd.read_excel(input_file_path)
    elif ext == ".csv":
        df_raw = pd.read_csv(input_file_path, low_memory=False)
    else:
        raise ValueError(f"Unsupported file format '{ext}' for file '{input_file_path}'. Expected CSV or Excel.")

    if df_raw.empty:
        raise ValueError(f"APL input file at '{input_file_path}' is empty.")

    logger.info(f"Loaded {len(df_raw):,} raw APL rows from input file.")

    # Find policy column
    pol_col = next((c for c in df_raw.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno")), None)
    if not pol_col:
        raise ValueError(f"Could not identify policy number column in file. Available columns: {list(df_raw.columns)}")

    # Group APL rows by policy number
    apl_groups: Dict[str, List[Dict[str, Any]]] = {}
    for _, row in df_raw.iterrows():
        p_raw = str(row[pol_col]).strip()
        p_clean = clean_policy_number(p_raw)
        if p_clean:
            apl_groups.setdefault(p_clean, []).append(row.to_dict())

    unique_apl_policies = list(apl_groups.keys())
    logger.info(f"Identified {len(unique_apl_policies):,} unique policy numbers across APL records.")

    # 2. Database Connection (or Mock / Dry-Run)
    db_mgr = db_manager
    if not db_mgr and not dry_run:
        db_mgr = ReadOnlyDatabaseManager(logger=logger)
        if not db_mgr.connect():
            logger.warning("Could not establish DB connection. Proceeding in DRY-RUN mode.")
            dry_run = True

    # Pre-fetch database records
    db_records_map: Dict[str, List[Dict[str, Any]]] = {}
    all_db_numbers: List[str] = []
    user_display_fn = lambda u: ""

    if db_mgr and db_mgr.conn:
        logger.info(f"Pre-fetching database policies and commission rules for {len(unique_apl_policies):,} policies...")
        db_records_map, all_db_numbers = db_mgr.fetch_policies_and_commissions(agency_id, unique_apl_policies)
        db_mgr.fetch_all_users_and_contacts(agency_id)
        user_display_fn = db_mgr.format_user_display
        logger.info(f"Found {len(all_db_numbers):,} matching policies in database.")

    # 3. Load execution report if provided, or auto-discover from previous run artifacts
    exec_report_map = load_execution_report(execution_report_path)
    pipeline_artifacts_dir = None
    if not execution_report_path:
        base_runs_dir = os.path.join("pipeline_output", "commission_validation_db")
        if os.path.exists(base_runs_dir):
            run_dirs = [
                os.path.join(base_runs_dir, d, "output")
                for d in os.listdir(base_runs_dir)
                if os.path.isdir(os.path.join(base_runs_dir, d, "output"))
            ]
            if run_dirs:
                run_dirs.sort(key=os.path.getmtime, reverse=True)
                pipeline_artifacts_dir = run_dirs[0]
                logger.info(f"Auto-detected recent import artifacts directory: {pipeline_artifacts_dir}")

    # 4. Perform validation across all policies
    summaries: List[Dict[str, Any]] = []
    all_field_comps: List[Dict[str, Any]] = []
    all_split_comps: List[Dict[str, Any]] = []

    for pol in unique_apl_policies:
        apl_recs = apl_groups[pol]

        # Policy matching
        pol_match = match_policy_identifier(pol, all_db_numbers)
        matched_db_num = pol_match["db_policy"]

        # If exact/normalized match not in pre-fetched list, check if db_records_map has it
        target_db_recs = db_records_map.get(matched_db_num, []) if matched_db_num else []

        # Determine execution status
        exec_st, exec_ev = determine_execution_status(
            pol, exec_report_map, pipeline_output_dir=pipeline_artifacts_dir
        )

        # Validate fields and commission splits
        val_res = validate_policy_records(
            apl_records=apl_recs,
            db_records=target_db_recs,
            policy_match=pol_match,
            exec_status=exec_st,
            exec_evidence=exec_ev,
            user_display_fn=user_display_fn,
        )

        summaries.append(val_res["summary"])
        all_field_comps.extend(val_res["field_comparisons"])
        all_split_comps.extend(val_res["split_comparisons"])

    # 5. Generate Annotated APL DataFrame (Original file + Match_Status, Matched_Details, Mismatch_Reason)
    logger.info("Generating Annotated APL file with audit status columns...")
    df_annotated = annotate_apl_rows(df_raw, db_records_map, all_db_numbers)

    # Output paths
    out_dir = os.path.dirname(os.path.abspath(output_excel_path)) if output_excel_path else os.path.join("pipeline_output", "commission_validation_db", run_id, "output")
    os.makedirs(out_dir, exist_ok=True)

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    if not output_excel_path:
        output_excel_path = os.path.join(
            out_dir,
            f"apl_db_post_update_validation_{run_id}_{timestamp}.xlsx",
        )

    # Standalone Annotated APL outputs
    annotated_excel_path = os.path.join(out_dir, f"apl_annotated_validation_{run_id}_{timestamp}.xlsx")
    annotated_csv_path = os.path.join(out_dir, f"apl_annotated_validation_{run_id}_{timestamp}.csv")
    df_annotated.to_excel(annotated_excel_path, index=False)
    df_annotated.to_csv(annotated_csv_path, index=False)
    logger.info(f"Saved standalone Annotated APL Excel: {annotated_excel_path}")
    logger.info(f"Saved standalone Annotated APL CSV:   {annotated_csv_path}")

    logger.info("Generating comprehensive 6-tab Excel validation workbook...")
    final_path = generate_excel_validation_workbook(
        summaries=summaries,
        field_comparisons=all_field_comps,
        split_comparisons=all_split_comps,
        output_path=output_excel_path,
        run_id=run_id,
        df_annotated=df_annotated,
        logger=logger,
    )

    # Close DB connection if opened here
    if db_mgr and not db_manager:
        db_mgr.close()

    logger.info("=" * 80)
    logger.info("✅ POST-IMPORT VALIDATION COMPLETE")
    logger.info(f"Excel Report Path:        {final_path}")
    logger.info(f"Annotated APL Excel Path: {annotated_excel_path}")
    logger.info(f"Annotated APL CSV Path:   {annotated_csv_path}")
    logger.info("=" * 80)

    return {
        "report_path": final_path,
        "total_policies": len(summaries),
        "exact_matches": sum(1 for s in summaries if s["overall_status"] == "SUCCESS_EXACT_MATCH"),
        "normalized_matches": sum(1 for s in summaries if s["overall_status"] == "SUCCESS_NORMALIZED_MATCH"),
        "review_required": sum(1 for s in summaries if s["review_required"]),
        "summaries": summaries,
    }


# ---------------------------------------------------------------------------
# CLI Argument Parser
# ---------------------------------------------------------------------------
def main():
    load_env_file()
    parser = argparse.ArgumentParser(
        description="Run Post-Import APL vs Database Read-Only Validation and Excel Report Generator."
    )
    parser.add_argument(
        "-i", "--input", "-f", "--file",
        default=os.environ.get("UPLOAD_FILE_PATH"),
        help="Path to APL source file (CSV or XLSX)",
    )
    parser.add_argument(
        "-r", "--run-id",
        default=None,
        help="Optional import run identifier (default: timestamped)",
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="Optional Excel workbook destination path",
    )
    parser.add_argument(
        "-x", "--execution-report",
        default=None,
        help="Optional path to import execution report or log",
    )
    parser.add_argument(
        "-a", "--agency-id",
        type=int,
        default=int(os.environ.get("AGENCY_ID", "5")),
        help="Agency ID (default: 5)",
    )
    parser.add_argument(
        "-e", "--env-file",
        default=".env",
        help="Path to .env configuration file",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate APL data offline without live DB connection",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose debug logging",
    )

    args = parser.parse_args()
    logger = setup_logger(verbose=args.verbose)

    if args.env_file:
        load_env_file(args.env_file)

    input_file = args.input or os.environ.get("UPLOAD_FILE_PATH")
    if not input_file:
        logger.error("No input file provided. Use --input <path> or set UPLOAD_FILE_PATH in .env")
        sys.exit(1)

    try:
        run_post_import_validation(
            input_file_path=input_file,
            output_excel_path=args.output,
            run_id=args.run_id,
            execution_report_path=args.execution_report,
            agency_id=args.agency_id,
            dry_run=args.dry_run,
            logger=logger,
        )
    except Exception as e:
        logger.error(f"Fatal validation error: {e}", exc_info=args.verbose)
        sys.exit(1)


if __name__ == "__main__":
    main()
