#!/usr/bin/env python3
"""
2-Stage Policy Commission Import Validation & DB Verification Engine
====================================================================
Stage 1 (First Validation & Split Normalization):
  - Normalizes product names & updates agent IDs.
  - Performs sequence rate/amount differential roll-ups (normalize_commission_splits.py).
  - Calculates line-item normalized split share percentages (summing to 100.00%).
  - Populates total Carrier Commission and Payout Method (PERCENTAGE / FIXED FEE).
  - Validates total Carrier Commission against master rate sheet (validate_extracted_commission.py).
  - EXPORTS PHASE 1 REPORTS:
      * file_validated_clean_<timestamp>.xlsx
      * file_validated_errors_<timestamp>.xlsx

Stage 2 (Database Verification for agency_id = 5):
  - Takes Phase 1 Valid records and queries PostgreSQL database:
      1. products table (agency_id=5, name={product_name}) -> obtained_product_uuid
      2. policies table (agency_id=5, number={policynumber}, product_id={obtained_product_uuid}) -> obtained_policy_uuid
      3. policy_commission table (agency_id=5, policy_uuid={obtained_policy_uuid}) -> DB commission rules
  - Verifies DB rules against file records.
  - EXPORTS PHASE 2 REPORTS:
      * db_matched_commission_<timestamp>.xlsx
      * db_unmatched_commission_<timestamp>.xlsx
      * updated_policies_<timestamp>.csv / .txt (unique policies successfully POST-recreated)

Usage:
  python validate_policy_commission_db.py --file /path/to/commission_upload.csv
  python validate_policy_commission_db.py --dry-run
  python validate_policy_commission_db.py --env-file .env --agency-id 5
"""

import argparse
import datetime
import functools
import math
import os
import re
import shutil
import sys
import time
from typing import Dict, List, Optional, Set, Tuple, Union

# Ensure UTF-8 output encoding on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import numpy as np
import pandas as pd
import json

# Import tabulate with fallback
try:
    from tabulate import tabulate
except ImportError:
    def tabulate(data, headers=None, tablefmt=None):
        """Very simple markdown table renderer fallback.
        data: list of rows (iterables)
        headers: list of column names
        """
        if headers:
            lines = ["| " + " | ".join(headers) + " |", "|" + "---|" * len(headers)]
        else:
            lines = []
        for row in data:
            lines.append("| " + " | ".join(str(v) for v in row) + " |")
        return "\n".join(lines)


def are_codes_matching(code1: str, code2: str) -> bool:
    """Check if two agent user codes are equivalent, ignoring suffixes like AV6old, SL, OVR, etc."""
    if not code1 or not code2:
        return False
    c1 = str(code1).strip().lower()
    c2 = str(code2).strip().lower()
    if c1 == c2:
        return True
    if c1.startswith(c2) or c2.startswith(c1) or c1 in c2 or c2 in c1:
        return True
    d1 = re.sub(r'[^0-9]', '', c1)
    d2 = re.sub(r'[^0-9]', '', c2)
    if d1 and d2 and len(d1) >= 4 and len(d2) >= 4 and d1 == d2:
        return True
    return False


def clean_name_ignoring_id(name) -> str:
    """Strips agent IDs, numbers, or codes in parentheses, prefixes like '8307933 - Gregory Teipel', and generational suffixes (Jr, Sr, II, III, IV, V)."""
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    s = str(name).strip()
    s = re.sub(r"^\s*[\w\d]+\s*-\s*", "", s).strip()
    s = re.sub(r"\s*\([^)]*\)$", "", s).strip()
    s = re.sub(r"\s+\b(jr|sr|[ivx]+)\b\.?\s*$", "", s, flags=re.IGNORECASE).strip()
    return s


def are_names_matching(name1: str, name2: str) -> bool:
    """Check if two agent names match, collapsing extra whitespace, normalizing case, aliases & extra middle names."""
    if not name1 or not name2:
        return False
    c1 = clean_name_ignoring_id(name1)
    c2 = clean_name_ignoring_id(name2)
    try:
        from validate_apl_vs_db_after_import import compare_person_names
        res = compare_person_names(c1, c2)
        return res["status"] in ("EXACT_MATCH", "NORMALIZED_MATCH")
    except ImportError:
        n1 = " ".join(c1.lower().split())
        n2 = " ".join(c2.lower().split())
        if n1 == n2:
            return True
        if ("main line" in n1 or "mlb" in n1) and ("main line" in n2 or "mlb" in n2):
            return True
        return False



# Check psycopg2 / DB driver availability
DB_DRIVER_AVAILABLE = False
try:
    import psycopg2
    from psycopg2 import sql
    from psycopg2.extras import RealDictCursor
    DB_DRIVER_AVAILABLE = True
except ImportError:
    try:
        import importlib
        psycopg2 = importlib.import_module("psycopg2_binary")
        sql = getattr(psycopg2, "sql", None)
        RealDictCursor = getattr(importlib.import_module("psycopg2_binary.extras"), "RealDictCursor", None)
        DB_DRIVER_AVAILABLE = True
    except (ImportError, AttributeError):
        DB_DRIVER_AVAILABLE = False





# ---------------------------------------------------------------------------
# Helper: Load .env File
# ---------------------------------------------------------------------------
def load_env_file(env_file_path: str = ".env", override: bool = False) -> int:
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
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip().strip("'").strip('"')
                    if key:
                        if override or key not in os.environ:
                            os.environ[key] = value
                            count += 1
    except Exception as e:
        print(f"⚠️ Warning: Could not load env file '{target_path}': {e}")
    return count


def resolve_sso_credentials() -> Tuple[Optional[str], Optional[str]]:
    """
    Resolve Keycloak username/password without picking up the Windows USERNAME
    (e.g. LENOVO) when .env has USERNAME=email@....
    Prefer SSO_USERNAME / SSO_PASSWORD; only use USERNAME if it looks like an email.
    """
    username = (
        os.environ.get("SSO_USERNAME")
        or os.environ.get("SSO_USER")
        or os.environ.get("API_USERNAME")
    )
    raw_user = os.environ.get("USERNAME") or os.environ.get("USER")
    if not username and raw_user and "@" in raw_user:
        username = raw_user
    password = (
        os.environ.get("SSO_PASSWORD")
        or os.environ.get("API_PASSWORD")
        or os.environ.get("PASSWORD")
    )
    return (username.strip() if username else None), (password if password else None)


# ---------------------------------------------------------------------------
# Formatting & Parsing Helpers
# ---------------------------------------------------------------------------
def clean_key(series: pd.Series) -> pd.Series:
    """Normalizes string keys (strips quotes/whitespace, uppercase, collapses spaces)."""
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.strip("'")
        .str.strip('"')
        .str.upper()
        .apply(lambda s: " ".join(s.split()))
    )


def clean_agent_code(val: str) -> str:
    """Cleans an agent code / user_code value:
    - Strips leading/trailing whitespace and leading quote characters (e.g. '21758900 -> 21758900).
    - Removes any trailing non-numeric alphabetic suffix (e.g. 21758900SL -> 21758900).
    """
    if not val:
        return val
    # Strip whitespace
    v = str(val).strip()
    # Remove leading quote characters (', ")
    v = v.lstrip("'\"")
    # Remove trailing alphabetic-only suffix (e.g. SL, SL1 not stripped; only pure alpha like SL)
    import re as _re
    v = _re.sub(r'[A-Za-z]+$', '', v).strip()
    return v


def policy_suffix_candidates(policy_number: str) -> List[str]:
    """Alternate policy numbers to try when the full number is not in the policies table.

    Tries, in order:
    - trailing alphabetic suffix removed (OVR, SL, and any other trailing letters)
    - short letter-led suffix removed (A8, K1, K67)
    - one trailing character removed (8526429013 -> 852642901)
    - two trailing characters removed, only if the one-character strip is not in the DB
    """
    pol = str(policy_number or "").strip().strip("'\"").strip()
    if not pol or pol.lower() == "nan":
        return []

    found: List[str] = []

    def add(value: str) -> None:
        value = str(value or "").strip()
        if value and value != pol and value not in found and len(value) >= 6:
            found.append(value)

    # Trailing letters, including a multi-letter suffix such as OVR or a single letter.
    add(re.sub(r"[A-Za-z]+$", "", pol))
    # Letter plus digits left on the end of a numeric policy number, such as A8, K1, or K67.
    token_stripped = re.sub(r"[A-Za-z]\d{1,3}$", "", pol)
    if token_stripped.isdigit():
        add(token_stripped)
    # One character, then two. Includes a numeric suffix such as 8526429013 -> 852642901.
    if len(pol) > 1:
        add(pol[:-1])
    if len(pol) > 2:
        add(pol[:-2])
    return found


def lookup_policy_db_records(chunk_db_map: dict, policy_number: str, clean_prod_name: str) -> list:
    """Return bulk-fetched policy rows for an exact policy number, preferring the product match."""
    db_records = chunk_db_map.get((policy_number, clean_prod_name), [])
    if db_records:
        return db_records
    pol_records = chunk_db_map.get(policy_number, [])
    filtered = [
        r for r in pol_records
        if clean_prod_name in " ".join(str(r.get("product_name", "")).lower().split())
        or " ".join(str(r.get("product_name", "")).lower().split()) in clean_prod_name
    ]
    if filtered:
        return filtered
    return list(pol_records) if pol_records else []


def pick_policy_by_product(rows: list, normalized_name: str, file_name: str) -> Tuple[list, str]:
    """Narrow DB rows that belong to several policies down to one policy by product.

    Tries the normalized product name (from the normalization list) first. When the
    product has no entry in that list (normalized name is the same as the file name)
    or the normalized name finds nothing, the file product name is used. Each name is
    tried as an exact match, then as a contains match.

    Returns (rows of the single matching policy, "") or ([], reason) when the product
    still points at more than one policy, or at none.
    """
    def _clean(s) -> str:
        return " ".join(str(s or "").lower().split())

    rows = [r for r in (rows or []) if isinstance(r, dict)]
    names = []
    for name in (normalized_name, file_name):
        cn = _clean(name)
        if cn and cn != "nan" and cn not in names:
            names.append(cn)
    if not names:
        return [], "no product name in the file to tell them apart"

    for cn in names:
        for mode in ("exact", "contains"):
            if mode == "exact":
                hits = [r for r in rows if _clean(r.get("product_name")) == cn]
            else:
                hits = [
                    r for r in rows
                    if _clean(r.get("product_name"))
                    and (cn in _clean(r.get("product_name")) or _clean(r.get("product_name")) in cn)
                ]
            uuids = {str(r.get("policy_uuid") or "").strip() for r in hits}
            uuids.discard("")
            if len(uuids) == 1:
                return hits, ""
            if len(uuids) > 1:
                return [], f"product '{cn}' still matches {len(uuids)} policies ({_db_products_text(rows)})"
    return [], (
        f"none of them has product {' / '.join(repr(n) for n in names)} "
        f"({_db_products_text(rows)})"
    )


def _db_products_text(rows: list) -> str:
    """'DB policies: <number> = <product>; ...' for the candidate rows (one entry per policy)."""
    seen = {}
    for r in rows or []:
        if not isinstance(r, dict):
            continue
        key = str(r.get("policy_uuid") or "").strip() or str(r.get("policy_number") or "")
        if key in seen:
            continue
        num = str(r.get("policy_number") or "").strip() or "?"
        prod = str(r.get("product_name") or "").strip() or "no product"
        seen[key] = f"{num} = '{prod}'"
    return "DB policies: " + "; ".join(seen.values()) if seen else "DB policies: none"


def parse_amount(val) -> float:
    """Parses currency strings or numeric values into float."""
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace("$", "").replace(",", "").strip()
    if not val_str:
        return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0


def parse_rate(val) -> float:
    """Parses percentage strings or numeric values into float percentage scale (e.g. 70.0).
    Handles Excel percentage cell formatting where 125% is read as float 1.25.
    """
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip()
    has_pct = "%" in val_str
    val_str = val_str.replace("%", "").strip()
    if not val_str:
        return 0.0
    try:
        v = float(val_str)
        # Excel stores percentage formats as decimal fractions (1.0 = 100%, 1.25 = 125%, 0.70 = 70%).
        # If the original string did NOT have an explicit '%' sign and 0 < v <= 2.5, scale up to percentage (0-100+).
        if not has_pct and 0 < v <= 2.5:
            v = round(v * 100.0, 4)
        return v
    except ValueError:
        return 0.0


def format_amount(val: float) -> str:
    """Formats float values to standard currency strings."""
    if val < 0:
        return f"-${abs(val):,.2f}"
    return f"${val:,.2f}"


def format_rate(val: float) -> str:
    """Formats float values into percentage strings."""
    return f"{val:.2f} %"


def normalize_product_str(p: str) -> str:
    """Strips punctuation (/ , - *) and collapses whitespace for product matching."""
    if not p or pd.isna(p):
        return ""
    clean = str(p).replace("/", " ").replace("-", " ").replace(",", " ").replace("*", " ")
    return " ".join(clean.lower().split())


def load_file(file_path: str) -> pd.DataFrame:
    """Loads CSV or Excel file (or directory of files) into a unified DataFrame."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File or directory not found: '{file_path}'")

    if os.path.isdir(file_path):
        dfs = []
        for root, _, files in os.walk(file_path):
            for f in sorted(files):
                if f.startswith("~$") or f.startswith("."):
                    continue
                ext = os.path.splitext(f.lower())[1]
                if ext in (".csv", ".xlsx", ".xls"):
                    full_p = os.path.join(root, f)
                    try:
                        frame = load_file(full_p)
                        print(f"📄 Reading '{full_p}' ({len(frame):,} rows)")
                        dfs.append(frame)
                    except Exception as err:
                        print(f"⚠️ Warning: Could not read file '{full_p}': {err}")
        if not dfs:
            raise ValueError(f"No valid CSV or Excel files found in directory '{file_path}'.")
        print(f"📄 Combined {len(dfs)} file(s) from '{file_path}'.")
        return pd.concat(dfs, ignore_index=True)

    _, ext = os.path.splitext(file_path.lower())
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(file_path)
    elif ext == ".csv":
        return pd.read_csv(file_path, low_memory=False)
    else:
        raise ValueError(f"Unsupported file format '{ext}' for file '{file_path}'.")


def archive_upload_input(src_path: str, dest_dir: str) -> None:
    """Copy the upload file, or every spreadsheet in an upload folder, into the run input folder."""
    if not os.path.isdir(src_path):
        shutil.copy2(src_path, os.path.join(dest_dir, os.path.basename(src_path)))
        return

    copied = 0
    for root, _, files in os.walk(src_path):
        for name in sorted(files):
            if name.startswith("~$") or name.startswith("."):
                continue
            ext = os.path.splitext(name.lower())[1]
            if ext not in (".csv", ".xlsx", ".xls"):
                continue
            full = os.path.join(root, name)
            rel = os.path.relpath(full, src_path)
            dest = os.path.join(dest_dir, rel)
            parent = os.path.dirname(dest)
            if parent:
                os.makedirs(parent, exist_ok=True)
            shutil.copy2(full, dest)
            copied += 1
    if copied == 0:
        print(f"⚠️ No spreadsheet files archived from '{src_path}'.")
    else:
        print(f"📁 Archived {copied} input file(s) from the upload folder.")


def load_product_normalization(norm_path: str = None) -> Dict[str, str]:
    """Loads product normalization map (Raw Product -> Normalized Product)."""
    if not norm_path or not os.path.exists(norm_path):
        return {}

    try:
        df_norm = load_file(norm_path)
        df_norm = df_norm.rename(columns=lambda c: str(c).strip())

        prod_col = next((c for c in df_norm.columns if c.lower() == "product"), None)
        target_col = next((c for c in df_norm.columns if "new product" in c.lower()), None)

        if not prod_col or not target_col:
            return {}

        norm_dict = {}
        for _, row in df_norm.iterrows():
            p_val = str(row[prod_col]).strip()
            t_val = str(row[target_col]).strip()
            if p_val and t_val and t_val.lower() not in ("nan", "none", ""):
                clean_p = " ".join(p_val.lower().split())
                norm_dict[clean_p] = t_val
                norm_dict[normalize_product_str(p_val)] = t_val
        return norm_dict
    except Exception as e:
        print(f"⚠️ Warning: Could not load product normalization file: {e}")
        return {}


def load_agent_roster(roster_path: str = None) -> Dict[str, str]:
    """Loads Agent Roster mapping dictionary (Cleaned Agent Name -> Agent ID)."""
    if not roster_path or not os.path.exists(roster_path):
        return {}

    try:
        df_roster = load_file(roster_path)
        df_roster = df_roster.rename(columns=lambda c: str(c).strip())

        name_col = next((c for c in df_roster.columns if "agent name" in c.lower() or "name" in c.lower()), None)
        id_col = next((c for c in df_roster.columns if "agent id" in c.lower() or "id" in c.lower()), None)

        if not name_col or not id_col:
            return {}

        roster_map = {}
        for _, row in df_roster.iterrows():
            name_val = str(row[name_col]).strip()
            id_val = str(row[id_col]).strip()
            if name_val and id_val and id_val.lower() not in ("nan", "none", ""):
                clean_n = " ".join(name_val.lower().split())
                roster_map[clean_n] = id_val
        return roster_map
    except Exception as e:
        print(f"⚠️ Warning: Could not load agent roster file: {e}")
        return {}


def unpivot_product_commission_rates(excel_path: str) -> pd.DataFrame:
    """Unpivots month tier column groups from ProductWise Commission Rate.xlsx into rules DataFrame."""
    df_raw = load_file(excel_path)
    suffixes = [""] + [f".{i}" for i in range(1, 17)]
    rules = []

    for idx, row in df_raw.iterrows():
        prod_id = row.get("Product ID")
        prod_desc = str(row.get("Product Description", "")).strip()
        contract = str(row.get("Contract", "")).strip()

        if not prod_desc or prod_desc.lower() in ("nan", "none", ""):
            continue

        for s in suffixes:
            fm_col = f"From-To Months{s}"
            pct_col = f"Rate (Percentage){s}"
            dollar_col = f"Rate (Dollar Amount){s}"
            fixed_col = f"Column Rate is Fixed{s}"

            if fm_col in row and pd.notna(row[fm_col]):
                fm_val = str(row[fm_col]).strip()
                if not fm_val or fm_val.lower() in ("nan", "none", ""):
                    continue

                parts = fm_val.split("-")
                try:
                    from_m = int(parts[0].strip())
                    to_m = int(parts[1].strip()) if len(parts) > 1 else from_m
                    if to_m == 0:
                        to_m = 9999
                except Exception:
                    continue

                pct_val = float(row.get(pct_col, 0.0)) if pd.notna(row.get(pct_col)) else 0.0
                dollar_val = float(row.get(dollar_col, 0.0)) if pd.notna(row.get(dollar_col)) else 0.0
                is_fixed = str(row.get(fixed_col, "N")).strip().upper()

                rate_pct_scaled = pct_val * 100.0 if (0.0 < pct_val <= 1.0) else pct_val

                rules.append({
                    "Product_ID": prod_id,
                    "Product_Description": prod_desc,
                    "Contract": contract,
                    "From_Month": from_m,
                    "To_Month": to_m,
                    "Rate_Percentage": rate_pct_scaled,
                    "Rate_Dollar": dollar_val,
                    "Is_Fixed": is_fixed,
                    "_prod_clean": " ".join(prod_desc.lower().split()),
                    "_prod_norm": " ".join(" ".join(prod_desc.replace("/", " ").replace("-", " ").replace(",", " ").replace("*", " ").split()).lower().split()),
                    "_contract_clean": " ".join(contract.lower().split()),
                })

    return pd.DataFrame(rules)


# ---------------------------------------------------------------------------
# YEAR-WISE FROM-TO NORMALIZATION + OVERRIDE 3-ROW EXPANSION
# ---------------------------------------------------------------------------
_OPEN_MONTH_SENTINEL = 10**9


def _month_to_int(val) -> int:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0
    try:
        return int(float(str(val).strip()))
    except (TypeError, ValueError):
        return 0


def _interval_bounds(fm, tm) -> Tuple[int, int]:
    start = max(1, _month_to_int(fm) or 1)
    end_raw = _month_to_int(tm)
    end = _OPEN_MONTH_SENTINEL if end_raw <= 0 else end_raw
    if end < start:
        end = start
    return start, end


def _band_to_from_to(start: int, end: int) -> Tuple[int, int]:
    return start, 0 if end >= _OPEN_MONTH_SENTINEL else end


def classify_contract_role(contract_val: str) -> str:
    """Map Contract column to agency | sales_leader | agent | sub_agent."""
    c = str(contract_val or "").strip().lower()
    if c in ("agency", ""):
        return "agency"
    if "sales" in c or "leader" in c:
        return "sales_leader"
    # SUBAGT1/2, Sub Agent, SA1/SA2 — keep separate so writing agent is not dropped vs LVL*
    if "sub" in c or re.match(r"^sa\d+$", c):
        return "sub_agent"
    return "agent"


def _find_col(df: pd.DataFrame, names: Tuple[str, ...]) -> Optional[str]:
    lowered = {str(c).strip().lower(): c for c in df.columns}
    for n in names:
        if n in lowered:
            return lowered[n]
    return None


def _format_rate_like(sample, value: float):
    """Preserve sheet rate style (e.g. '100.00%') when writing synthetic rows."""
    if sample is None or (isinstance(sample, float) and pd.isna(sample)):
        return f"{value:.2f}%"
    s = str(sample).strip()
    if "%" in s:
        return f"{value:.2f}%"
    try:
        float(s.replace("%", "").strip())
        return value
    except ValueError:
        return f"{value:.2f}%"


def _is_mlb_or_agency_person(name: str = "", code: str = "", item: Optional[dict] = None) -> bool:
    """True when the person is Main Line Benefits / agency entity (must not fill SL or Agent)."""
    n = " ".join(str(name or "").lower().split())
    if "main line" in n or n in ("mlb", "mlb parallel run", "mainline benefits"):
        return True
    if item:
        entity = str(item.get("entity") or "").strip().lower()
        role = str(item.get("role") or "").strip().lower()
        if entity == "agency" or role == "agency":
            return True
        uuid = str(item.get("user_uuid") or item.get("uuid") or "").strip().lower()
        if uuid == "62a97a01-2801-4c7c-9cd0-cf38f2e0a4f4":
            return True
    return False


def _parse_split_people(commission_split) -> Dict[str, dict]:
    """Extract Sales Leader / Agent from DB commission_split.

    Live OVR bug pattern to ignore: Agency missing, Sales Leader = MLB, Agent = Broker.
    Never promote MLB/agency into SL or Agent. Broker maps to Agent.
    """
    people: Dict[str, dict] = {}
    if commission_split is None:
        return people
    splits = commission_split
    if isinstance(splits, str):
        try:
            splits = json.loads(splits)
        except Exception:
            return people
    if not isinstance(splits, list):
        return people

    broker_as_agent: Optional[dict] = None
    for item in splits:
        if not isinstance(item, dict):
            continue
        role = str(item.get("role") or item.get("level") or "").strip().lower()
        name = str(item.get("user_name") or item.get("name") or "").strip()
        code = clean_agent_code(str(item.get("user_code") or item.get("code") or "").strip())
        if not name and not code:
            continue
        # Skip Agency / MLB entirely — they belong only on Agency order-1
        if role == "agency" or _is_mlb_or_agency_person(name, code, item):
            continue
        person = {"name": name, "code": code}
        if "sales" in role or "leader" in role:
            people["sales_leader"] = person
        elif "sub" in role or role == "sub_agent":
            people["sub_agent"] = person
        elif role == "broker":
            # Sheet Broker = writing agent; also candidate for SL when app stuffed MLB into SL
            broker_as_agent = person
            if "agent" not in people:
                people["agent"] = person
        elif role == "agent" or role.startswith("lvl") or role.startswith("level"):
            if "agent" not in people or role == "agent":
                people["agent"] = person

    # If SL was only MLB (filtered out) but Broker/Agent exists → SL = same person @ 0% later
    if "sales_leader" not in people:
        if "agent" in people:
            people["sales_leader"] = dict(people["agent"])
        elif broker_as_agent:
            people["sales_leader"] = dict(broker_as_agent)
            people.setdefault("agent", dict(broker_as_agent))
    return people


def fix_inclusive_month_band_overlaps(df: pd.DataFrame) -> Tuple[pd.DataFrame, dict]:
    """
    Remove inclusive From/To overlaps on the same policy+product+person.

    Example (Agency):
        13-72 then 72-120  ->  13-72 then 73-120
    so month 72 is not claimed by two contiguous bands (which creates a bad
    single-month 72-72 slice after year-band re-slicing).

    Rule: for each person sorted by From Month, if next.From <= prev.To
    (and prev.To > 0), bump next.From to prev.To + 1. Open-ended bands
    (To Month = 0) are left untouched as the prior end.
    """
    if df is None or df.empty:
        return df, {"rows_fixed": 0, "examples": []}

    df_out = df.copy()
    pol_col = _find_col(df_out, ("policy no", "policy number", "policy", "policyno", "policy_number"))
    prod_col = _find_col(df_out, ("product", "product name", "product_description", "product_name"))
    fm_col = _find_col(df_out, ("from month", "from_month", "month", "month from", "month_from"))
    tm_col = _find_col(df_out, ("to month", "to_month", "month to", "month_to"))
    agent_col = _find_col(df_out, ("agent id", "agent_id", "agentid", "eacid"))
    seq_col = _find_col(df_out, ("sequence", "seq"))
    contract_col = _find_col(df_out, ("contract", "level", "role"))

    if not all([pol_col, fm_col, tm_col]):
        return df_out, {"rows_fixed": 0, "examples": [], "skipped": True}

    group_keys = [pol_col]
    if prod_col:
        group_keys.append(prod_col)
    # Prefer Agent ID; fall back to Sequence+Contract so Agency-only rows still group
    if agent_col:
        group_keys.append(agent_col)
    else:
        if seq_col:
            group_keys.append(seq_col)
        if contract_col:
            group_keys.append(contract_col)

    rows_fixed = 0
    examples: List[dict] = []

    for _, group in df_out.groupby(group_keys, sort=False):
        ordered_idx = sorted(
            group.index.tolist(),
            key=lambda i: (_month_to_int(df_out.at[i, fm_col]), _month_to_int(df_out.at[i, tm_col]), i),
        )
        prev_tm = None
        for idx in ordered_idx:
            fm = _month_to_int(df_out.at[idx, fm_col])
            tm = _month_to_int(df_out.at[idx, tm_col])
            if prev_tm is not None and prev_tm > 0 and fm > 0 and fm <= prev_tm:
                new_fm = prev_tm + 1
                # Only bump when the band would still be valid (or open-ended)
                if tm <= 0 or new_fm <= tm:
                    if len(examples) < 20:
                        examples.append({
                            "policy": str(df_out.at[idx, pol_col]),
                            "old_from": fm,
                            "new_from": new_fm,
                            "to": tm,
                        })
                    df_out.at[idx, fm_col] = new_fm
                    fm = new_fm
                    rows_fixed += 1
            # Track end of this band for the next comparison (use possibly bumped from)
            if tm > 0:
                prev_tm = tm if prev_tm is None else max(prev_tm, tm)
            # open-ended (tm==0): do not constrain later bands via prev_tm beyond current

    return df_out, {"rows_fixed": rows_fixed, "examples": examples}


def _is_blank_contract(val: Any) -> bool:
    """True when Contract was left empty on a continuation month band."""
    if val is None:
        return True
    try:
        if pd.isna(val):
            return True
    except (TypeError, ValueError):
        pass
    return str(val).strip().lower() in ("", "nan", "none", "null")


def _person_token(val: Any) -> str:
    """Agent id with quotes removed. Keeps an SL suffix so Agent and Sales Leader stay distinct."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    token = str(val).strip().strip("'\"").upper()
    if token.endswith(".0") and token[:-2].isdigit():
        token = token[:-2]
    return token


def _sequence_token(val: Any) -> str:
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(val).strip()
    if text.lower() in ("", "nan", "none", "null"):
        return ""
    try:
        return str(int(float(text)))
    except (TypeError, ValueError):
        return text.upper()


def _inferred_contract(agent_id: Any, agent_name: Any) -> Optional[str]:
    """Role for a blank Contract, using the person rather than the empty cell.

    Main Line Benefits is Agency. An agent id ending in SL is the Sales Leader.
    The writing agent is whoever is left.
    """
    name = ""
    if agent_name is not None:
        try:
            blank_name = pd.isna(agent_name)
        except (TypeError, ValueError):
            blank_name = False
        if not blank_name:
            name = " ".join(str(agent_name).strip().lower().split())
    if "main line" in name or name in ("mlb", "mlb parallel run"):
        return "Agency"
    if _person_token(agent_id).endswith("SL"):
        return "Sales Leader"
    return None


def fill_blank_contracts_from_siblings(df: pd.DataFrame) -> pd.DataFrame:
    """Fill a blank Contract from the same person, or from their id and name.

    Sheets often fill Contract only on the first band, or leave it empty on every
    band. A blank contract is classified as one role, so year-band normalize keeps
    a single row and Stage 1 then reports that band as an invalid 1-row split.
    """
    if df is None or df.empty:
        return df

    df_out = df.copy()
    pol_col = _find_col(df_out, ("policy no", "policy number", "policy", "policyno", "policy_number"))
    prod_col = _find_col(df_out, ("product", "product name", "product_description", "product_name"))
    contract_col = _find_col(df_out, ("contract", "level", "role"))
    agent_col = _find_col(df_out, ("agent id", "agent_id", "agentid"))
    name_col = next(
        (c for c in df_out.columns if "agent name" in str(c).strip().lower() or str(c).strip().lower() == "agent_name"),
        None,
    )
    seq_col = _find_col(df_out, ("sequence", "seq"))
    fm_col = _find_col(df_out, ("from month", "from_month", "month", "month from", "month_from"))
    if not contract_col or not pol_col:
        return df_out
    df_out[contract_col] = df_out[contract_col].astype("object")

    known: Dict[Tuple[str, str, str, str], Tuple[int, Any]] = {}
    for _, row in df_out.iterrows():
        if _is_blank_contract(row.get(contract_col)):
            continue
        key = (
            str(row.get(pol_col, "")).strip().strip("'\"").upper(),
            str(row.get(prod_col, "")).strip().upper() if prod_col else "",
            _person_token(row.get(agent_col)) if agent_col else "",
            _sequence_token(row.get(seq_col)) if seq_col else "",
        )
        from_month = _month_to_int(row.get(fm_col)) if fm_col else 0
        previous = known.get(key)
        if previous is None or from_month < previous[0]:
            known[key] = (from_month, row.get(contract_col))

    def _row_key(row: pd.Series) -> Tuple[str, str, str, str]:
        return (
            str(row.get(pol_col, "")).strip().strip("'\"").upper(),
            str(row.get(prod_col, "")).strip().upper() if prod_col else "",
            _person_token(row.get(agent_col)) if agent_col else "",
            _sequence_token(row.get(seq_col)) if seq_col else "",
        )

    for idx, row in df_out.iterrows():
        if not _is_blank_contract(row.get(contract_col)):
            continue
        hit = known.get(_row_key(row))
        if hit is not None:
            df_out.at[idx, contract_col] = hit[1]
            continue
        inferred = _inferred_contract(
            row.get(agent_col) if agent_col else "",
            row.get(name_col) if name_col else "",
        )
        if inferred:
            df_out.at[idx, contract_col] = inferred

    # One remaining blank person on the policy is the writing agent.
    leftovers: Dict[Tuple[str, str], Set[str]] = {}
    for _, row in df_out.iterrows():
        if not _is_blank_contract(row.get(contract_col)):
            continue
        pol_prod = _row_key(row)[:2]
        leftovers.setdefault(pol_prod, set()).add(_row_key(row)[2])
    single_agent = {pol_prod for pol_prod, people in leftovers.items() if len(people) == 1 and "" not in people}
    for idx, row in df_out.iterrows():
        if not _is_blank_contract(row.get(contract_col)):
            continue
        if _row_key(row)[:2] in single_agent:
            df_out.at[idx, contract_col] = "Agent"
    return df_out


def normalize_year_bands(df: pd.DataFrame) -> Tuple[pd.DataFrame, dict]:
    """
    Re-slice From-To so every person shares the finest common month bands for a policy+product.

    Two people with the same role stay as two rows. StreetZA and StreetZC are both
    agents; sequence 2 must not be dropped because sequence 1 is already an agent.

    Example: Agency 1-48 + SL/Agent 1-12 and 13-48
          -> Agency/SL/Agent for 1-12 and Agency/SL/Agent for 13-48.
    """
    if df is None or df.empty:
        return df, {"groups_touched": 0, "rows_in": 0, "rows_out": 0}

    df = fill_blank_contracts_from_siblings(df)

    pol_col = _find_col(df, ("policy no", "policy number", "policy", "policyno", "policy_number"))
    prod_col = _find_col(df, ("product", "product name", "product_description", "product_name"))
    fm_col = _find_col(df, ("from month", "from_month", "month", "month from", "month_from"))
    tm_col = _find_col(df, ("to month", "to_month", "month to", "month_to"))
    contract_col = _find_col(df, ("contract", "level", "role"))
    seq_col = _find_col(df, ("sequence", "seq"))
    agent_col = _find_col(df, ("agent id", "agent_id", "agentid"))
    name_col = next(
        (c for c in df.columns if "agent name" in str(c).strip().lower() or str(c).strip().lower() == "agent_name"),
        None,
    )

    if not all([pol_col, prod_col, fm_col, tm_col, contract_col]):
        print("⚠️ Year-band normalize skipped: missing Policy/Product/From/To/Contract columns.")
        return df.copy(), {"groups_touched": 0, "rows_in": len(df), "rows_out": len(df), "skipped": True}

    def _person_key(row: pd.Series) -> Tuple[str, str, str]:
        seq = _sequence_token(row.get(seq_col)) if seq_col else ""
        agent = _person_token(row.get(agent_col)) if agent_col else ""
        if not agent and name_col:
            agent = " ".join(str(row.get(name_col) or "").split()).upper()
        role = classify_contract_role(row.get(contract_col, ""))
        return (seq, agent, role)

    def _covers(row: pd.Series, b_start: int, b_end: int) -> bool:
        start, end = _interval_bounds(row.get(fm_col), row.get(tm_col))
        return start <= b_start and end >= b_end

    out_rows: List[dict] = []
    groups_touched = 0

    for (_, _), group in df.groupby([pol_col, prod_col], sort=False):
        cuts: Set[int] = set()
        person_rows: Dict[Tuple[str, str, str], List[pd.Series]] = {}
        person_order: List[Tuple[str, str, str]] = []
        for _, row in group.iterrows():
            key = _person_key(row)
            if key not in person_rows:
                person_rows[key] = []
                person_order.append(key)
            person_rows[key].append(row)
            start, end = _interval_bounds(row.get(fm_col), row.get(tm_col))
            cuts.add(start)
            if end < _OPEN_MONTH_SENTINEL:
                cuts.add(end + 1)

        if not cuts:
            for _, row in group.iterrows():
                out_rows.append(row.to_dict())
            continue

        cuts_sorted = sorted(cuts)
        # Ensure open-ended coverage when any person uses To Month = 0
        has_open = any(
            _interval_bounds(r.get(fm_col), r.get(tm_col))[1] >= _OPEN_MONTH_SENTINEL
            for rows in person_rows.values() for r in rows
        )
        if has_open and (not cuts_sorted or cuts_sorted[-1] < _OPEN_MONTH_SENTINEL):
            cuts_sorted.append(_OPEN_MONTH_SENTINEL)

        bands: List[Tuple[int, int]] = []
        for i in range(len(cuts_sorted) - 1):
            b_start = cuts_sorted[i]
            next_cut = cuts_sorted[i + 1]
            b_end = _OPEN_MONTH_SENTINEL if next_cut >= _OPEN_MONTH_SENTINEL else next_cut - 1
            if b_end >= b_start:
                bands.append((b_start, b_end))

        # Detect whether slicing actually changes anything
        original_keys = {
            (_person_key(r), _month_to_int(r.get(fm_col)), _month_to_int(r.get(tm_col)))
            for _, r in group.iterrows()
        }
        preview_keys = set()
        for b_start, b_end in bands:
            fm_out, tm_out = _band_to_from_to(b_start, b_end)
            for key, rows in person_rows.items():
                if any(_covers(r, b_start, b_end) for r in rows):
                    preview_keys.add((key, fm_out, tm_out))
        if preview_keys != original_keys:
            groups_touched += 1

        for b_start, b_end in bands:
            fm_out, tm_out = _band_to_from_to(b_start, b_end)
            for key in person_order:
                covering = [r for r in person_rows[key] if _covers(r, b_start, b_end)]
                if not covering:
                    continue
                # Prefer the tightest covering interval, then the first row for that person
                covering.sort(
                    key=lambda r: (
                        _interval_bounds(r.get(fm_col), r.get(tm_col))[1]
                        - _interval_bounds(r.get(fm_col), r.get(tm_col))[0]
                    )
                )
                base = covering[0].to_dict()
                base[fm_col] = fm_out
                base[tm_col] = tm_out
                out_rows.append(base)

    df_out = pd.DataFrame(out_rows)
    # Preserve original column order when possible
    if not df_out.empty:
        ordered = [c for c in df.columns if c in df_out.columns]
        extras = [c for c in df_out.columns if c not in ordered]
        df_out = df_out[ordered + extras]
    summary = {
        "groups_touched": groups_touched,
        "rows_in": len(df),
        "rows_out": len(df_out),
    }
    return df_out.reset_index(drop=True), summary


def _is_override_group(group: pd.DataFrame, prod_col: str, contract_col: str, rate_col: Optional[str]) -> bool:
    prod = str(group.iloc[0].get(prod_col, ""))
    if "ovr" in prod.lower():
        return True
    if not rate_col:
        return False
    agency_rate = 0.0
    other_rate = 0.0
    for _, row in group.iterrows():
        role = classify_contract_role(row.get(contract_col, ""))
        rate = parse_rate(row.get(rate_col))
        if role == "agency":
            agency_rate = max(agency_rate, rate)
        else:
            other_rate += rate
    return agency_rate >= 99.99 and other_rate <= 0.01


def is_override_commission_split(splits: list) -> bool:
    """True when Agency takes ~100% and all other roles are ~0% (OVERRIDE, not REGULAR)."""
    agency_pct = 0.0
    other_pct = 0.0
    for s in splits or []:
        if not isinstance(s, dict):
            continue
        role = str(s.get("role") or "").strip().lower()
        entity = str(s.get("entity") or "").strip().lower()
        try:
            val = float(s.get("commission_value") or 0.0)
        except (TypeError, ValueError):
            val = 0.0
        if role == "agency" or entity == "agency":
            agency_pct = max(agency_pct, val)
        else:
            other_pct += val
    return agency_pct >= 99.99 and other_pct <= 0.01


def resolve_commission_type(
    splits: list,
    product_name: str = "",
    db_commission_type=None,
) -> int:
    """
    1=REGULAR, 2=RENEWAL, 3=BONUS, 4=OVERRIDE.
    Agency 100% (or product name contains OVR) => OVERRIDE, even if DB had type 1.
    """
    if "ovr" in str(product_name or "").lower():
        return 4
    if is_override_commission_split(splits):
        return 4
    try:
        db_t = int(db_commission_type)
        if db_t in (1, 2, 3, 4):
            return db_t
    except (TypeError, ValueError):
        pass
    return 1


def commission_rule_name_for_type(commission_type: int) -> str:
    if int(commission_type) == 4:
        return "Carrier - Product - OVERRIDE"
    return "Carrier - Product - COMMISSION"


# Overrides (product name with OVR) are ON HOLD: when False, override rows are set aside before
# Stage 1 - no override check, no DELETE and no POST for them - and saved to their own Excel.
# Set to True to enable override validation and updates again.
OVERRIDE_UPDATES_ENABLED = False

# Helper columns written by the override builder and read by Stage 1 / Stage 2.
OVR_INVALID_COL = "_override_invalid_reason"
OVR_SL_FROM_FILE_COL = "_override_sl_from_file"
SPLIT_LEVEL_COL = "_split_level"
SPLIT_USER_UUID_COL = "_split_user_uuid"


def _clean_helper_text(val) -> str:
    """Helper-column value as text; blank for None/NaN."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(val).strip()
    return "" if text.lower() in ("nan", "none", "null") else text


def _split_list(raw) -> list:
    """commission_split as a list of dicts (DB returns list or JSON text)."""
    if raw is None:
        return []
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except Exception:
            return []
    if not isinstance(raw, list):
        return []
    return [s for s in raw if isinstance(s, dict)]


def rule_effective_type(rec: dict) -> int:
    """The DB rule's stored commission_type (default 1).

    An override is a rule stored as type 4. An Agency-100% split on a type-1 rule is a
    regular commission (e.g. a renewal band where only the Agency is paid), not an override.
    """
    try:
        return int(rec.get("commission_type") or 1)
    except (TypeError, ValueError):
        return 1


def is_override_product(product_name) -> bool:
    """A file row is an override when its product name carries OVR (e.g. 'Allstate OVR')."""
    return bool(re.search(r"\bovr\b", str(product_name or "").lower()))


def is_override_family(ctype) -> bool:
    try:
        return int(ctype) == 4
    except (TypeError, ValueError):
        return False


def _hierarchy_people(splits: list) -> Dict[str, dict]:
    """Sales Leader / Agent / Sub Agent from one DB split, with level, code and user uuid.

    MLB is allowed as Sales Leader (legacy data has it there). MLB is never the Agent
    or Sub Agent. Empty placeholder slots are skipped.
    """
    people: Dict[str, dict] = {}

    def _order(item: dict) -> int:
        try:
            return int(item.get("order") or 999)
        except (TypeError, ValueError):
            return 999

    for item in sorted(_split_list(splits), key=_order):
        role = str(item.get("role") or "").strip().lower()
        entity = str(item.get("entity") or "").strip().lower()
        if role == "agency" or entity == "agency":
            continue
        name = " ".join(str(item.get("user_name") or "").split())
        code = clean_agent_code(str(item.get("user_code") or "").strip())
        uuid = str(item.get("user_uuid") or "").strip()
        if not name and not code and not uuid:
            continue
        person = {
            "name": name,
            "code": code,
            "uuid": uuid,
            "level": str(item.get("level") or "").strip(),
        }
        if "sales" in role or "leader" in role:
            people.setdefault("sales_leader", person)
            continue
        if _is_mlb_or_agency_person(name, code, item):
            continue
        if "sub" in role:
            people.setdefault("sub_agent", person)
        else:
            people.setdefault("agent", person)
    return people


def _resolve_policy_records(result_map: dict, policy_number: str) -> Tuple[Optional[str], list, str]:
    """DB rows for one policy: the exact number first, then legacy suffix candidates.

    A suffix candidate is used only when it points at exactly one policy, the same rule
    Stage 2 uses. People are never merged across candidates.
    """
    rows = result_map.get(policy_number) or []
    if rows:
        return policy_number, list(rows), ""
    for cand in policy_suffix_candidates(policy_number):
        rows = result_map.get(cand) or []
        if not rows:
            continue
        uuids = {str(r.get("policy_uuid") or "").strip() for r in rows} - {""}
        if len(uuids) > 1:
            return None, [], f"policy number '{cand}' (suffix removed) matches {len(uuids)} policies"
        return cand, list(rows), ""
    return None, [], ""


def _pick_people_source(records: list, band_from: int) -> Tuple[Dict[str, dict], str, str]:
    """People for an override band, taken from one DB rule of the same policy.

    Order: a regular (non-override) rule, preferring the same From Month; then an
    existing DB override that already names a Sales Leader and Agent.
    Returns (people, source_label, source_months).
    """
    with_split = [r for r in records if _split_list(r.get("commission_split"))]
    regular = [r for r in with_split if rule_effective_type(r) != 4]
    override = [r for r in with_split if rule_effective_type(r) == 4]

    def _from(r) -> int:
        return _month_to_int(r.get("month_from")) or 1

    def _agency_only_paid(r) -> bool:
        return is_override_commission_split(_split_list(r.get("commission_split")))

    for label, pool in (("regular", regular), ("db_override", override)):
        # Prefer a regular rule that actually pays the agents over an Agency-100% band
        for rec in sorted(pool, key=lambda r: (_agency_only_paid(r), _from(r) != band_from, _from(r))):
            people = _hierarchy_people(rec.get("commission_split"))
            if "sales_leader" in people and "agent" in people:
                to_m = rec.get("month_to")
                months = f"{_from(rec)}-{to_m if to_m is not None else 'open'}"
                return people, label, months
    return {}, "", ""


def ensure_override_three_rows(
    df: pd.DataFrame,
    db_mgr: Optional["DatabaseManager"] = None,
    agency_id: int = 5,
) -> Tuple[pd.DataFrame, dict]:
    """
    Make every override (OVR / Agency 100%) month band carry Agency + Sales Leader + Agent
    (+ Sub Agent when there is one). Never leaves an Agency-only override.

    People, per role:
      - Agent:       the file's Agent/Broker row, else the DB rule's Agent
      - Sales Leader: the file's Sales Leader row, else the DB rule's Sales Leader (MLB allowed)
      - Sub Agent:   the file's Sub Agent row, else the DB rule's named Sub Agent
    The DB rule is one rule on the same policy (exact number, then a legacy suffix
    candidate): a regular rule first, then an existing DB override.

    Levels and user uuids come from that DB rule. Sequences follow the hierarchy:
    Sub Agent < Agent < Sales Leader < Agency, so Agency keeps the 100%.

    When a band cannot get an Agency, Sales Leader and Agent, every row of that
    policy+product gets OVR_INVALID_COL and Stage 1 sends it to the invalid report.
    Without a DB connection the old file-only fallback is used (Broker as Sales Leader).
    """
    empty_summary = {
        "groups_expanded": 0, "rows_appended": 0, "db_lookups": 0, "db_misses": 0,
        "db_miss_policies": [], "invalid_groups": 0, "invalid_policies": [],
    }
    if df is None or df.empty:
        return df, empty_summary

    pol_col = _find_col(df, ("policy no", "policy number", "policy", "policyno", "policy_number"))
    prod_col = _find_col(df, ("product", "product name", "product_description", "product_name"))
    fm_col = _find_col(df, ("from month", "from_month", "month", "month from", "month_from"))
    tm_col = _find_col(df, ("to month", "to_month", "month to", "month_to"))
    contract_col = _find_col(df, ("contract", "level", "role"))
    rate_col = _find_col(df, ("rate", "rate %", "commission %"))
    agent_col = _find_col(df, ("agent id", "agent_id", "agentid"))
    agent_name_col = next(
        (c for c in df.columns if "agent name" in str(c).strip().lower() or str(c).strip().lower() == "agent_name"),
        None,
    )
    seq_col = _find_col(df, ("sequence", "seq"))
    eacid_col = _find_col(df, ("eacid",))
    status_col = _find_col(df, ("agent status", "agent_status"))
    amt_col = _find_col(df, ("amount", "fee amount", "fee_amount"))

    if not all([pol_col, prod_col, fm_col, tm_col, contract_col]):
        print("⚠️ Override 3-row expand skipped: missing required columns.")
        out = df.copy()
        empty_summary["skipped"] = True
        return out, empty_summary

    df = df.copy()
    for col in (OVR_INVALID_COL, OVR_SL_FROM_FILE_COL, SPLIT_LEVEL_COL, SPLIT_USER_UUID_COL):
        if col not in df.columns:
            df[col] = ""
    if not seq_col:
        df["Sequence"] = 1
        seq_col = "Sequence"

    override_keys = [
        key for key, group in df.groupby([pol_col, prod_col], sort=False)
        if is_override_product(key[1])
    ]
    db_available = db_mgr is not None and bool(getattr(db_mgr, "conn", None))

    # One DB fetch for every override policy (+ suffix candidates), all chunks merged
    # before resolving, so a policy whose rows land in a later chunk is not a miss.
    result_map: dict = {}
    db_lookups = 0
    db_misses = 0
    miss_log: List[str] = []
    if override_keys and db_available:
        pol_list = sorted({str(pol).strip().strip("'\"") for pol, _ in override_keys})
        lookup_numbers: List[str] = []
        seen_numbers: Set[str] = set()
        for p in pol_list:
            for cand in [p] + policy_suffix_candidates(p):
                if cand not in seen_numbers:
                    seen_numbers.add(cand)
                    lookup_numbers.append(cand)
        print(f"🔍 Override expand: reading DB rules for {len(pol_list):,} override policies...")
        for i in range(0, len(lookup_numbers), 500):
            chunk_map = db_mgr.batch_fetch_policy_commission_rules(agency_id, lookup_numbers[i : i + 500], [])
            for k, rows in (chunk_map or {}).items():
                if isinstance(k, str):
                    result_map.setdefault(k, []).extend(rows)

    out_frames: List[pd.DataFrame] = []
    groups_expanded = 0
    rows_appended = 0
    invalid_log: List[str] = []
    override_key_set = set(override_keys)

    for (pol, prod), group in df.groupby([pol_col, prod_col], sort=False):
        if (pol, prod) not in override_key_set:
            out_frames.append(group)
            continue

        pol_key = str(pol).strip().strip("'\"")
        records: list = []
        resolve_note = ""
        if db_available:
            db_lookups += 1
            _, records, resolve_note = _resolve_policy_records(result_map, pol_key)
            if not records:
                db_misses += 1
                miss_log.append(f"{pol_key} | {prod}")

        new_parts: List[pd.DataFrame] = []
        group_changed = False
        group_invalid_reason = ""

        for (fm, tm), band in group.groupby([fm_col, tm_col], sort=False):
            band_out = band.copy()
            roles = band_out[contract_col].apply(classify_contract_role)

            # Broker on an override sheet is the writing agent
            broker_mask = band_out[contract_col].astype(str).str.strip().str.lower() == "broker"
            if broker_mask.any():
                band_out.loc[broker_mask, contract_col] = "Agent"
                group_changed = True

            def _first(role_key: str) -> Optional[pd.Series]:
                hits = band_out[roles == role_key]
                return hits.iloc[0] if len(hits) else None

            agency_row = _first("agency")
            sl_row = _first("sales_leader")
            agent_row = _first("agent")
            sub_row = _first("sub_agent")

            if agency_row is None:
                group_invalid_reason = group_invalid_reason or (
                    f"Override months {fm}-{tm}: no Agency row in the file."
                )
                new_parts.append(band_out)
                continue

            band_from = _month_to_int(fm) or 1
            people, source, source_months = ({}, "", "")
            if db_available:
                people, source, source_months = _pick_people_source(records, band_from)

            def _file_person(row: Optional[pd.Series]) -> Optional[dict]:
                if row is None:
                    return None
                name = str(row.get(agent_name_col, "")).strip() if agent_name_col else ""
                code = clean_agent_code(str(row.get(agent_col, "")).strip()) if agent_col else ""
                return {"name": name, "code": code, "uuid": "", "level": ""}

            file_agent = _file_person(agent_row)
            if file_agent and _is_mlb_or_agency_person(file_agent["name"], file_agent["code"]):
                file_agent = None
            file_sl = _file_person(sl_row)
            file_sub = _file_person(sub_row)

            if db_available:
                agent_p = file_agent or people.get("agent")
                sl_p = file_sl or people.get("sales_leader")
                sub_p = file_sub or people.get("sub_agent")
                # A Sub Agent who is the same person as the Agent is not a 4th level
                if sub_p and not file_sub and agent_p and _same_split_person(sub_p, agent_p):
                    sub_p = None
                # File rows for the same person as the DB rule take its level and user uuid
                for f_row, f_person, role_key in (
                    (agent_row, file_agent, "agent"),
                    (sl_row, file_sl, "sales_leader"),
                    (sub_row, file_sub, "sub_agent"),
                ):
                    db_person = people.get(role_key)
                    if f_row is None or not f_person or not db_person:
                        continue
                    if _same_split_person(f_person, db_person):
                        band_out.loc[f_row.name, SPLIT_LEVEL_COL] = db_person.get("level", "")
                        band_out.loc[f_row.name, SPLIT_USER_UUID_COL] = db_person.get("uuid", "")
            else:
                # No DB: old file-only fallback so dry runs still produce 3 rows
                agent_p = file_agent
                sl_p = file_sl or (dict(file_agent) if file_agent else None)
                sub_p = file_sub

            if not agent_p or not sl_p:
                missing = [lbl for lbl, p in (("Sales Leader", sl_p), ("Agent", agent_p)) if not p]
                if not db_available:
                    why = "no DB connection to look up the regular rule"
                elif resolve_note:
                    why = resolve_note
                elif not records:
                    why = "policy not found in DB"
                else:
                    why = "no regular (non-override) rule with a Sales Leader and Agent on this policy"
                group_invalid_reason = group_invalid_reason or (
                    f"Override months {fm}-{tm}: missing {' and '.join(missing)} ({why})."
                )
                new_parts.append(band_out)
                continue

            template = agency_row

            def _synth_row(contract_name: str, person: dict) -> dict:
                row = template.to_dict()
                row[contract_col] = contract_name
                row[fm_col] = fm
                row[tm_col] = tm
                if rate_col:
                    row[rate_col] = _format_rate_like(template.get(rate_col), 0.0)
                if agent_name_col:
                    row[agent_name_col] = person.get("name", "")
                if agent_col:
                    row[agent_col] = person.get("code", "")
                if eacid_col:
                    row[eacid_col] = person.get("code", "")
                if status_col:
                    row[status_col] = "Active"
                if amt_col:
                    row[amt_col] = ""
                row[SPLIT_LEVEL_COL] = person.get("level", "")
                row[SPLIT_USER_UUID_COL] = person.get("uuid", "")
                row[OVR_INVALID_COL] = ""
                return row

            appended: List[dict] = []
            if agent_row is None:
                appended.append(_synth_row("Agent", agent_p))
            if sl_row is None:
                appended.append(_synth_row("Sales Leader", sl_p))
            if sub_row is None and sub_p:
                appended.append(_synth_row("Sub Agent", sub_p))

            band_full = pd.concat([band_out, pd.DataFrame(appended)], ignore_index=True) if appended else band_out
            band_full[OVR_SL_FROM_FILE_COL] = "Y" if sl_row is not None else ""

            # Sequence by hierarchy: Sub Agent < Agent < Sales Leader < Agency
            rank = {"sub_agent": 0, "agent": 1, "sales_leader": 2, "agency": 3}
            full_roles = band_full[contract_col].apply(classify_contract_role)
            present = sorted({rank.get(r, 1) for r in full_roles})
            seq_of = {r: i + 1 for i, r in enumerate(present)}
            band_full[seq_col] = [seq_of[rank.get(r, 1)] for r in full_roles]

            if appended:
                rows_appended += len(appended)
                group_changed = True
            new_parts.append(band_full)

        group_out = pd.concat(new_parts, ignore_index=True) if new_parts else group
        if group_invalid_reason:
            group_out[OVR_INVALID_COL] = group_invalid_reason
            invalid_log.append(f"{pol_key} | {prod} | {group_invalid_reason}")
        elif group_changed:
            groups_expanded += 1
        out_frames.append(group_out)

    df_out = pd.concat(out_frames, ignore_index=True) if out_frames else df.copy()
    if invalid_log:
        print(f"⚠️ Override expand: {len(invalid_log)} override policy(ies) sent to the invalid report "
              f"(no Agency + Sales Leader + Agent could be built).")
        for line in invalid_log[:20]:
            print(f"   - {line}")
        if len(invalid_log) > 20:
            print(f"   ... and {len(invalid_log) - 20} more")
    if override_keys and not db_available:
        print("⚠️ Override expand: no DB connection, so Sales Leader/Agent were not checked "
              "against the regular rule (Broker used for both).")

    summary = {
        "groups_expanded": groups_expanded,
        "rows_appended": rows_appended,
        "db_lookups": db_lookups,
        "db_misses": db_misses,
        "db_miss_policies": miss_log,
        "invalid_groups": len(invalid_log),
        "invalid_policies": invalid_log,
    }
    return df_out.reset_index(drop=True), summary


# ---------------------------------------------------------------------------
# SEQUENCE DIFFERENTIAL ROLL-UP & SPLIT NORMALIZATION (normalize_commission_splits.py)
# ---------------------------------------------------------------------------
def rollup_commission_splits(df: pd.DataFrame) -> pd.DataFrame:
    """
    Performs sequence differential roll-up and percentage split normalization.
    Computes rate differentials (Seq 1=10%, Seq 2=20%-10%=10%, Seq 3=30%-20%=10%) or
    amount differentials (Seq 1=$20, Seq 2=$50-$20=$30, Seq 3=$80-$50=$30),
    normalizes split shares so they sum to 100.00%, and populates total Carrier Commission.
    """
    df_out = df.copy()

    pol_col = next((c for c in df_out.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
    prod_col = next((c for c in df_out.columns if str(c).strip().lower() in ("product", "product name", "product_description", "product_name")), None)
    fm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("from month", "from_month", "month")), None)
    seq_col = next((c for c in df_out.columns if str(c).strip().lower() in ("sequence", "seq")), None)
    rate_col = next((c for c in df_out.columns if str(c).strip().lower() in ("rate", "rate %", "commission %")), None)
    amt_col = next((c for c in df_out.columns if str(c).strip().lower() in ("amount", "fee amount", "fee_amount")), None)

    if not pol_col:
        return df_out

    if "Carrier Commission" not in df_out.columns:
        df_out["Carrier Commission"] = "$0.00"
    if "Payout Method" not in df_out.columns:
        df_out["Payout Method"] = ""

    df_out["_amt_num"] = df_out[amt_col].apply(parse_amount) if amt_col else 0.0
    df_out["_rate_num"] = df_out[rate_col].apply(parse_rate) if rate_col else 0.0
    df_out["_seq_num"] = pd.to_numeric(df_out[seq_col], errors="coerce").fillna(1).astype(int) if seq_col else 1

    df_out["_pol_clean"] = clean_key(df_out[pol_col])
    df_out["_prod_clean"] = clean_key(df_out[prod_col]) if prod_col else ""
    df_out["_fm_clean"] = clean_key(df_out[fm_col]) if fm_col else ""

    group_cols = ["_pol_clean"]
    if prod_col:
        group_cols.append("_prod_clean")
    if fm_col:
        group_cols.append("_fm_clean")

    grouped = df_out.groupby(group_cols, sort=False)

    res_amt = list(df_out[amt_col]) if amt_col else []
    res_rate = list(df_out[rate_col]) if rate_col else []
    res_carrier_comm = ["$0.00"] * len(df_out)
    res_payout_method = [""] * len(df_out)
    res_split_share = [0.0] * len(df_out)
    res_diff_val = [""] * len(df_out)

    for _, group in grouped:
        indices = group.index.tolist()
        rates = group["_rate_num"].tolist()
        amounts = group["_amt_num"].tolist()
        seqs = group["_seq_num"].tolist()

        has_nonzero_rate = any(r > 0 for r in rates)

        seq_amt_map = {}
        seq_rate_map = {}
        for s, a, r in zip(seqs, amounts, rates):
            if s not in seq_amt_map:
                seq_amt_map[s] = a
            if s not in seq_rate_map:
                seq_rate_map[s] = r

        max_seq = max(seqs) if seqs else 1

        if not has_nonzero_rate:
            # FIXED FEE ($) — Differential Amounts & 100% Split Shares
            adj_amounts = []
            for s, a in zip(seqs, amounts):
                if a > 0:
                    prev_seq = s - 1
                    if prev_seq in seq_amt_map:
                        prev_amt = seq_amt_map[prev_seq]
                        # Round to cents: 20.02 - 18.48 gives 1.5399999999999991 in float, not 1.54
                        adj_amounts.append(max(0.0, round(a - prev_amt, 2)))
                    else:
                        adj_amounts.append(a)
                else:
                    adj_amounts.append(a)

            total_amt = seq_amt_map.get(max_seq, max(amounts) if amounts else 0.0)

            rates_calc = []
            if total_amt > 0:
                non_highest_rate_sum = 0.0
                highest_seq_indices = [i for i, s in enumerate(seqs) if s == max_seq]
                primary_highest_idx = highest_seq_indices[-1]

                for i, adj_a in enumerate(adj_amounts):
                    if i != primary_highest_idx:
                        raw_pct = (adj_a / total_amt) * 100.0
                        # Small margin so 6.9999999999 is cut to 7.00, not 6.99
                        trunc_pct = math.floor(raw_pct * 100.0 + 1e-6) / 100.0
                        rates_calc.append(trunc_pct)
                        non_highest_rate_sum += trunc_pct
                    else:
                        rates_calc.append(None)

                highest_rate = round(100.00 - non_highest_rate_sum, 2)
                rates_calc[primary_highest_idx] = highest_rate
            else:
                rates_calc = [0.0] * len(indices)

            carrier_comm_str = format_amount(total_amt)
            for global_idx, adj_a, r_val in zip(indices, adj_amounts, rates_calc):
                if amt_col:
                    res_amt[global_idx] = format_amount(adj_a)
                if rate_col:
                    res_rate[global_idx] = format_rate(r_val)
                res_diff_val[global_idx] = format_amount(adj_a)
                res_split_share[global_idx] = r_val
                res_carrier_comm[global_idx] = carrier_comm_str
                res_payout_method[global_idx] = "FIXED FEE"

        else:
            # PERCENTAGE (%) — Differential Rates & 100% Split Shares
            rate_diffs = []
            for s, r in zip(seqs, rates):
                prev_seq = s - 1
                if prev_seq in seq_rate_map:
                    prev_r = seq_rate_map[prev_seq]
                    rate_diffs.append(max(0.0, round(r - prev_r, 4)) if r > prev_r else r)
                else:
                    rate_diffs.append(r)

            total_rate = seq_rate_map.get(max_seq, max(rates) if rates else 0.0)

            rates_calc = []
            if total_rate > 0:
                non_highest_rate_sum = 0.0
                highest_seq_indices = [i for i, s in enumerate(seqs) if s == max_seq]
                primary_highest_idx = highest_seq_indices[-1]

                for i, r_diff in enumerate(rate_diffs):
                    if i != primary_highest_idx:
                        raw_pct = (r_diff / total_rate) * 100.0
                        # Small margin so 6.9999999999 is cut to 7.00, not 6.99
                        trunc_pct = math.floor(raw_pct * 100.0 + 1e-6) / 100.0
                        rates_calc.append(trunc_pct)
                        non_highest_rate_sum += trunc_pct
                    else:
                        rates_calc.append(None)

                highest_rate = round(100.00 - non_highest_rate_sum, 2)
                rates_calc[primary_highest_idx] = highest_rate
            else:
                rates_calc = [0.0] * len(indices)

            carrier_comm_str = format_rate(total_rate)
            for global_idx, r_diff, r_val in zip(indices, rate_diffs, rates_calc):
                if rate_col:
                    res_rate[global_idx] = format_rate(r_val)
                res_diff_val[global_idx] = format_rate(r_diff)
                res_split_share[global_idx] = r_val
                res_carrier_comm[global_idx] = carrier_comm_str
                res_payout_method[global_idx] = "PERCENTAGE"

    if amt_col:
        df_out[amt_col] = res_amt
    if rate_col:
        df_out[rate_col] = res_rate
    df_out["Normalized_Split_Share_%"] = res_split_share
    df_out["Differential_Value"] = res_diff_val
    df_out["Carrier Commission"] = res_carrier_comm
    df_out["Payout Method"] = res_payout_method

    df_out = df_out.drop(columns=["_amt_num", "_rate_num", "_seq_num", "_pol_clean", "_prod_clean", "_fm_clean"], errors="ignore")
    return df_out


# ---------------------------------------------------------------------------
# STAGE 1: File Pre-Validation Engine
# ---------------------------------------------------------------------------
def stage1_pre_validate_dataset(
    df_comm: pd.DataFrame,
    df_rules: Optional[pd.DataFrame],
    norm_dict: Dict[str, str] = None,
    agent_roster_dict: Dict[str, str] = None,
    tolerance: float = 0.01,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict]:
    """
    Stage 1: Applies sequence rollups and validates carrier rates against master rate sheet.
    Returns: (df_stage1_valid, df_stage1_invalid, df_stage1_zero, summary_dict)
    """
    df_norm_splits = rollup_commission_splits(df_comm)
    df_out = df_norm_splits.copy()

    # Enrich / normalize Agent ID whenever Agent Name matches in agent_roster_dict
    agent_name_col = next((c for c in df_out.columns if "agent name" in str(c).strip().lower() or str(c).strip().lower() == "agent_name"), None)
    
    if agent_roster_dict:
        agent_id_col = next((c for c in df_out.columns if str(c).strip().lower() in ("agent id", "agent_id", "agentid")), None)
        if agent_name_col:
            if not agent_id_col:
                df_out["Agent ID"] = ""
                agent_id_col = "Agent ID"
            for idx in df_out.index:
                raw_n = str(df_out.loc[idx, agent_name_col]).strip()
                clean_n = " ".join(raw_n.lower().split())
                if clean_n in agent_roster_dict:
                    df_out.loc[idx, agent_id_col] = agent_roster_dict[clean_n]

    pol_col = next((c for c in df_out.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
    prod_col = next((c for c in df_out.columns if str(c).strip().lower() in ("product", "product name", "product_description", "product_name")), None)
    fm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("from month", "from_month", "month", "month from", "month_from")), None)
    tm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("to month", "to_month", "month to", "month_to")), None)
    contract_col_s1 = _find_col(df_out, ("contract", "level", "role"))
    carrier_comm_col = "Carrier Commission"
    payout_method_col = "Payout Method"

    if not prod_col:
        raise KeyError("Missing Product column in input file.")

    if not fm_col:
        df_out["From_Month"] = 1
        fm_col = "From_Month"
        
    if not tm_col:
        df_out["To_Month"] = 0
        tm_col = "To_Month"

    norm_dict = norm_dict or {}

    validation_statuses = [""] * len(df_out)
    expected_values = [""] * len(df_out)
    variances = [0.0] * len(df_out)
    error_reasons = [""] * len(df_out)
    normalized_product_names = [""] * len(df_out)

    valid_count = 0
    mismatch_val_count = 0
    unmatched_prod_count = 0
    unmatched_month_count = 0

    df_out["_pol_clean"] = clean_key(df_out[pol_col]) if pol_col else pd.Series(["ROW_" + str(i) for i in range(len(df_out))], index=df_out.index)
    df_out["_prod_clean"] = clean_key(df_out[prod_col])
    df_out["_fm_clean"] = clean_key(df_out[fm_col])
    df_out["_tm_clean"] = clean_key(df_out[tm_col])

    grouped = df_out.groupby(["_pol_clean", "_prod_clean", "_fm_clean", "_tm_clean"], sort=False)

    for _, group in grouped:
        indices = group.index.tolist()
        seq_1_idx = indices[0]
        seq_1_row = group.loc[seq_1_idx]

        pol_val = str(seq_1_row.get(pol_col, "")).strip() if pol_col else ""
        raw_prod = str(seq_1_row.get(prod_col, "")).strip()
        raw_fm = seq_1_row.get(fm_col, 1)

        clean_prod = " ".join(raw_prod.lower().split())
        norm_prod = normalize_product_str(raw_prod)

        final_norm_prod = raw_prod
        if clean_prod in norm_dict:
            final_norm_prod = norm_dict[clean_prod]
            clean_prod = " ".join(final_norm_prod.lower().split())
            norm_prod = normalize_product_str(final_norm_prod)
        elif norm_prod in norm_dict:
            final_norm_prod = norm_dict[norm_prod]
            clean_prod = " ".join(final_norm_prod.lower().split())

        if not pol_val or pol_val.lower() in ("nan", "none", "", "null"):
            for idx in indices:
                validation_statuses[idx] = "MISSING_POLICY_NUMBER"
                error_reasons[idx] = "Policy number is missing/blank in input upload file."
                expected_values[idx] = "Valid Policy Number"
                variances[idx] = 0.0
                normalized_product_names[idx] = final_norm_prod
            continue

        for idx in indices:
            normalized_product_names[idx] = final_norm_prod

        try:
            from_month_num = int(float(str(raw_fm).strip().replace("'", "").replace('"', "")))
        except (ValueError, TypeError):
            from_month_num = 1

        extracted_comm_str = str(seq_1_row.get(carrier_comm_col, ""))
        payout_method = str(seq_1_row.get(payout_method_col, "")).strip().upper()

        extracted_comm_amt = parse_amount(extracted_comm_str)
        extracted_comm_pct = parse_rate(extracted_comm_str)

        # -------------------------------------------------------------------
        # CHECK OVR: Override bands need Agency + Sales Leader + Agent.
        # Runs before the MLB exception so an Agency-only override is never VALID.
        # -------------------------------------------------------------------
        ovr_reason = ""
        if OVR_INVALID_COL in group.columns:
            ovr_reason = next(
                (_clean_helper_text(v) for v in group[OVR_INVALID_COL] if _clean_helper_text(v)), ""
            )
        # Override = OVR product. An Agency-only renewal band on a regular product is not one.
        is_override_band = is_override_product(raw_prod) or is_override_product(final_norm_prod)
        if is_override_band and not ovr_reason and contract_col_s1:
            present_roles = {classify_contract_role(group.loc[i, contract_col_s1]) for i in indices}
            missing_roles = [
                label for key, label in (("agency", "Agency"), ("sales_leader", "Sales Leader"), ("agent", "Agent"))
                if key not in present_roles
            ]
            if missing_roles:
                ovr_reason = f"Override band is missing {', '.join(missing_roles)} (needs Agency + Sales Leader + Agent)."
        if ovr_reason:
            for idx in indices:
                validation_statuses[idx] = "INVALID_OVERRIDE_LEVELS"
                expected_values[idx] = "Agency + Sales Leader + Agent"
                variances[idx] = float(len(indices))
                error_reasons[idx] = ovr_reason
            continue

        # -------------------------------------------------------------------
        # CHECK 0: MLB Direct Exceptions (Bypass Stage 1 Validation)
        # -------------------------------------------------------------------
        if agent_name_col:
            agent_name_raw = " ".join(str(seq_1_row.get(agent_name_col, "")).lower().split())
            if agent_name_raw in ("main line benefits", "mlb parallel run", "mlb"):
                valid_count += len(indices)
                for idx in indices:
                    validation_statuses[idx] = "VALID"
                    expected_values[idx] = "MLB_EXCEPTION"
                    variances[idx] = 0.0
                    error_reasons[idx] = ""
                continue

        # -------------------------------------------------------------------
        # CHECK 0.5: Negative Split Share Verification
        # -------------------------------------------------------------------
        group_split_shares = [float(group.loc[idx, "Normalized_Split_Share_%"]) for idx in indices]
        has_negative_split = any(s < 0.0 for s in group_split_shares)
        if has_negative_split:
            for idx in indices:
                validation_statuses[idx] = "NEGATIVE_SPLIT_SHARE"
                expected_values[idx] = ">= 0.00 %"
                variances[idx] = min(group_split_shares)
                error_reasons[idx] = f"Group contains negative split share ({min(group_split_shares)}%)."
            continue

        # -------------------------------------------------------------------
        # CHECK 1: Group Split Shares Sum Verification (100.00%)
        # -------------------------------------------------------------------
        group_split_shares = [float(group.loc[idx, "Normalized_Split_Share_%"]) for idx in indices]
        split_sum = round(sum(group_split_shares), 2)
        if abs(split_sum - 100.00) > 0.05:
            group_val = extracted_comm_amt if payout_method == "FIXED FEE" else extracted_comm_pct
            
            if group_val == 0.0:
                for idx in indices:
                    validation_statuses[idx] = "ZERO_COMM_IGNORED"
                    expected_values[idx] = "100.00 %"
                    variances[idx] = round(abs(split_sum - 100.00), 2)
                    error_reasons[idx] = f"Ignored: Commission is 0 and split sum is {split_sum}%"
            else:
                for idx in indices:
                    validation_statuses[idx] = "INVALID_SPLIT_SUM"
                    expected_values[idx] = "100.00 %"
                    variances[idx] = round(abs(split_sum - 100.00), 2)
                    error_reasons[idx] = f"Group split share sum is {split_sum}% instead of expected 100.00%."
            continue

        # -------------------------------------------------------------------
        # CHECK 2: Policy Split Row Count Verification (3 to 4 rows per policy)
        # -------------------------------------------------------------------
        row_count = len(indices)
        if row_count < 3 or row_count > 4:
            for idx in indices:
                validation_statuses[idx] = "INVALID_POLICY_ROW_COUNT"
                expected_values[idx] = "3 to 4 split rows"
                variances[idx] = float(row_count)
                error_reasons[idx] = f"Policy group has {row_count} split rows instead of expected 3 to 4 rows."
            continue

        if df_rules is None or df_rules.empty:
            valid_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "VALID"
                expected_values[idx] = "SKIPPED_RATE_SHEET"
                variances[idx] = 0.0
                error_reasons[idx] = ""
            continue

        p_rules = df_rules[df_rules["Product_Description"] == raw_prod]
        if p_rules.empty:
            p_rules = df_rules[df_rules["Product_Description"] == final_norm_prod]
        if p_rules.empty:
            p_rules = df_rules[df_rules["_prod_clean"] == clean_prod]
        if p_rules.empty and "_prod_norm" in df_rules.columns:
            p_rules = df_rules[df_rules["_prod_norm"] == norm_prod]
        if p_rules.empty:
            p_rules = df_rules[df_rules["_prod_clean"].apply(lambda p: p in clean_prod or clean_prod in p)]

        if p_rules.empty:
            unmatched_prod_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "UNMATCHED_PRODUCT_IN_RATE_SHEET"
                expected_values[idx] = "N/A"
                variances[idx] = 0.0
                error_reasons[idx] = f"Product '{raw_prod}' not found in master rate sheet."
            continue

        m_rules = p_rules[(p_rules["From_Month"] <= from_month_num) & (from_month_num <= p_rules["To_Month"])]
        if m_rules.empty:
            unmatched_month_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "UNMATCHED_MONTH_RANGE"
                expected_values[idx] = "N/A"
                variances[idx] = 0.0
                error_reasons[idx] = f"Month {from_month_num} is out of rate tier ranges for product '{raw_prod}'."
            continue

        agency_rules = m_rules[m_rules["_contract_clean"] == "agency"]
        if agency_rules.empty or (agency_rules["Rate_Dollar"].max() == 0.0 and agency_rules["Rate_Percentage"].max() == 0.0):
            agency_rules = m_rules.sort_values(by=["Rate_Dollar", "Rate_Percentage"], ascending=False)

        agency_rule = agency_rules.sort_values(by=["Rate_Dollar", "Rate_Percentage"], ascending=False).iloc[0]
        exp_agency_pct = agency_rule["Rate_Percentage"]
        exp_agency_dollar = agency_rule["Rate_Dollar"]

        group_valid = True
        group_var = 0.0
        group_exp_str = ""
        group_reason = ""
        group_status = "VALID"

        if payout_method == "PERCENTAGE" or (exp_agency_pct > 0.0 and exp_agency_dollar == 0.0):
            group_exp_str = format_rate(exp_agency_pct)
            group_var = round(abs(extracted_comm_pct - exp_agency_pct), 4)

            if group_var > tolerance:
                group_status = "VALUE_MISMATCH"
                group_reason = f"Carrier Commission mismatch: Extracted benchmark rate '{format_rate(extracted_comm_pct)}' vs expected Agency benchmark '{group_exp_str}'"
                group_valid = False
        elif payout_method == "FIXED FEE" or (exp_agency_dollar > 0.0):
            group_exp_str = format_amount(exp_agency_dollar)
            group_var = round(abs(extracted_comm_amt - exp_agency_dollar), 2)

            if group_var > tolerance:
                group_status = "VALUE_MISMATCH"
                group_reason = f"Carrier Commission mismatch: Extracted benchmark amount '{format_amount(extracted_comm_amt)}' vs expected Agency benchmark '{group_exp_str}'"
                group_valid = False

        if group_valid:
            valid_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = "VALID"
                expected_values[idx] = group_exp_str
                variances[idx] = 0.0
                error_reasons[idx] = ""
        else:
            if group_status == "VALUE_MISMATCH":
                mismatch_val_count += len(indices)
            for idx in indices:
                validation_statuses[idx] = group_status
                expected_values[idx] = group_exp_str
                variances[idx] = group_var
                error_reasons[idx] = group_reason

    df_out["Normalized_Product_Name"] = normalized_product_names
    df_out["Stage1_Status"] = validation_statuses
    df_out["Stage1_Expected_Commission"] = expected_values
    df_out["Stage1_Error_Reason"] = error_reasons

    df_out = df_out.drop(columns=["_pol_clean", "_prod_clean", "_fm_clean"], errors="ignore")

    valid_mask = df_out["Stage1_Status"] == "VALID"
    zero_mask = df_out["Stage1_Status"] == "ZERO_COMM_IGNORED"
    
    invalid_mask = ~(valid_mask | zero_mask)
    
    df_stage1_valid = df_out[valid_mask].copy()
    df_stage1_zero = df_out[zero_mask].copy()
    
    # Grab the original raw rows from df_comm for errors
    df_stage1_invalid = df_comm[invalid_mask].copy()
    
    # Attach just the error context columns back to the raw rows
    df_stage1_invalid["Stage1_Status"] = df_out.loc[invalid_mask, "Stage1_Status"]
    df_stage1_invalid["Stage1_Expected_Commission"] = df_out.loc[invalid_mask, "Stage1_Expected_Commission"]
    df_stage1_invalid["Stage1_Error_Reason"] = df_out.loc[invalid_mask, "Stage1_Error_Reason"]

    summary = {
        "total_input_records": len(df_comm),
        "stage1_valid_records": len(df_stage1_valid),
        "stage1_invalid_records": len(df_stage1_invalid),
        "stage1_zero_ignored": len(df_stage1_zero),
        "value_mismatches": mismatch_val_count,
        "unmatched_products": unmatched_prod_count,
        "unmatched_month_ranges": unmatched_month_count,
    }

    return df_stage1_valid, df_stage1_invalid, df_stage1_zero, summary


# ---------------------------------------------------------------------------
# STAGE 2: Database Verification Engine (PostgreSQL / Read-Only)
# ---------------------------------------------------------------------------
class DatabaseManager:
    """Manages read-only PostgreSQL queries for products, policies, and policy_commission tables."""

    def __init__(self, db_url: str = None, host: str = None, port: int = 5432, dbname: str = None, user: str = None, password: str = None):
        self.db_url = db_url or os.getenv("DATABASE_URL")
        self.user = user or os.getenv("DB_USER")
        self.password = password or os.getenv("DB_PASSWORD")
        self.host = host or os.getenv("DB_HOST", "localhost")
        self.port = port or int(os.getenv("DB_PORT", 5432))
        self.dbname = dbname or os.getenv("DB_NAME", "pieq_db")
        self.conn = None
        self.user_display_by_uuid = {}

    def connect(self) -> bool:
        if not DB_DRIVER_AVAILABLE:
            print("⚠️ Notice: psycopg2 driver not installed. Database verification running in DRY-RUN mode.", flush=True)
            return False

        try:
            sslmode = os.getenv("DB_SSLMODE") or ("require" if "rds.amazonaws.com" in str(self.db_url or self.host) else "prefer")
            target_info = ""
            if self.db_url:
                clean_url = str(self.db_url).strip()
                if clean_url.lower().startswith("jdbc:"):
                    clean_url = clean_url[5:]
                conn_kwargs = {"dsn": clean_url, "connect_timeout": 5, "options": "-c statement_timeout=180000", "sslmode": sslmode}
                if self.user:
                    conn_kwargs["user"] = self.user
                if self.password:
                    conn_kwargs["password"] = self.password

                host_match = re.search(r"@([^:/]+)", clean_url)
                db_match = re.search(r"/([^/?]+)", clean_url)
                log_host = host_match.group(1) if host_match else "remote host"
                log_db = db_match.group(1) if db_match else "db"
                target_info = f"Host: '{log_host}', DB: '{log_db}', User: '{self.user or 'default'}', SSL: '{sslmode}'"

                print(f"   🔌 Connecting to PostgreSQL ({target_info}) [Timeout: 5s]...", flush=True)
                self.conn = psycopg2.connect(**conn_kwargs)
            else:
                target_info = f"Host: '{self.host}:{self.port}', DB: '{self.dbname}', User: '{self.user}', SSL: '{sslmode}'"
                print(f"   🔌 Connecting to PostgreSQL ({target_info}) [Timeout: 5s]...", flush=True)
                self.conn = psycopg2.connect(
                    host=self.host, port=self.port, dbname=self.dbname, user=self.user, password=self.password, connect_timeout=5, options="-c statement_timeout=180000", sslmode=sslmode
                )

            self.conn.set_session(readonly=True, autocommit=True)

            db_version = "PostgreSQL"
            try:
                with self.conn.cursor() as cur:
                    cur.execute("SELECT version();")
                    row = cur.fetchone()
                    if row and row[0]:
                        ver_str = str(row[0])
                        db_version = ver_str.split(",")[0] if "," in ver_str else ver_str.split("\n")[0]
            except Exception:
                pass

            print(f"   🟢 Connected to PostgreSQL ({target_info})", flush=True)
            print(f"   ℹ️ Server Info: {db_version}", flush=True)
            return True

        except Exception as e:
            print(f"   ❌ Could not connect to PostgreSQL database: {e}", flush=True)
            self.conn = None
            return False

    def close(self):
        if self.conn:
            try:
                self.conn.close()
            except Exception:
                pass
            self.conn = None

    @functools.lru_cache(maxsize=4096)
    def get_product_uuid(self, agency_id: int, product_name: str) -> Optional[str]:
        """Queries products table: SELECT * FROM products WHERE agency_id = %s AND (name = %s OR lower(name) = lower(%s))"""
        if not self.conn:
            return None
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                query = """
                    SELECT * FROM products 
                    WHERE agency_id = %s 
                      AND (name = %s OR lower(name) = lower(%s))
                    LIMIT 1;
                """
                cur.execute(query, (agency_id, product_name, product_name))
                row = cur.fetchone()
                if row:
                    val = row.get("uuid") or row.get("id")
                    return str(val) if val is not None else None
        except Exception as e:
            print(f"⚠️ Error querying products for '{product_name}': {e}")
        return None

    @functools.lru_cache(maxsize=16384)
    def get_policy_uuid(self, agency_id: int, policy_number: str, product_uuid: str) -> Optional[str]:
        """Queries policies table: SELECT * FROM policies WHERE agency_id = %s AND number = %s AND product_id::text = %s"""
        if not self.conn:
            return None
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                query = """
                    SELECT * FROM policies 
                    WHERE agency_id = %s 
                      AND number = %s 
                      AND product_id::text = %s
                    LIMIT 1;
                """
                cur.execute(query, (agency_id, policy_number, str(product_uuid)))
                row = cur.fetchone()
                if row:
                    val = row.get("uuid") or row.get("id")
                    return str(val) if val is not None else None
        except Exception as e:
            print(f"⚠️ Error querying policies for number '{policy_number}': {e}")
        return None

    def get_policy_commission(self, agency_id: int, policy_uuid: str) -> List[dict]:
        """Fetch policy_commission rows with created/updated audit + contact display names."""
        if not self.conn:
            return []
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                query = """
                    SELECT
                        pc.*,
                        TRIM(CONCAT(COALESCE(c_created.first_name, ''), ' ', COALESCE(c_created.last_name, ''))) AS created_by_name,
                        TRIM(CONCAT(COALESCE(c_updated.first_name, ''), ' ', COALESCE(c_updated.last_name, ''))) AS updated_by_name
                    FROM policy_commission pc
                    LEFT JOIN contacts c_created
                        ON c_created.relationship_id = pc.created_by
                       AND c_created.agency_id = pc.agency_id
                    LEFT JOIN contacts c_updated
                        ON c_updated.relationship_id = pc.updated_by
                       AND c_updated.agency_id = pc.agency_id
                    WHERE pc.agency_id = %s
                      AND pc.policy_uuid::text = %s;
                """
                cur.execute(query, (agency_id, str(policy_uuid)))
                return [dict(r) for r in cur.fetchall()]
        except Exception as e:
            print(f"⚠️ Error querying policy_commission for policy_uuid '{policy_uuid}': {e}")
        return []

    @functools.lru_cache(maxsize=8192)
    def get_user_uuid(self, agency_id: int, user_code: str = None, user_name: str = None) -> Optional[str]:
        """Find user UUID by code (users) or name (contacts joined to users on relationship_id = uuid)."""
        if not self.conn:
            return None
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                if user_code:
                    clean_code = clean_agent_code(str(user_code))
                    # Match by code on tenant_id or agency_id
                    cur.execute(
                        "SELECT uuid FROM users WHERE (tenant_id = %s OR agency_id = %s) AND (code = %s OR code = %s) LIMIT 1;",
                        (agency_id, agency_id, clean_code, str(user_code).strip())
                    )
                    row = cur.fetchone()
                    if row and row.get("uuid"):
                        return str(row["uuid"])
                if user_name:
                    norm_name = " ".join(clean_name_ignoring_id(user_name).lower().split())
                    # Contacts name match is valid only when relationship_id joins users.uuid
                    contact_name_sql = """
                        SELECT c.relationship_id
                        FROM contacts c
                        INNER JOIN users u ON u.uuid = c.relationship_id
                        WHERE c.agency_id = %s
                          AND c.relationship_id IS NOT NULL
                          AND (u.tenant_id = %s OR u.agency_id = %s)
                          AND (
                              lower(concat(c.first_name, ' ', c.last_name)) = %s
                              OR lower(c.first_name) = %s
                              OR lower(c.last_name) = %s
                          )
                        LIMIT 1;
                    """
                    cur.execute(
                        contact_name_sql,
                        (agency_id, agency_id, agency_id, norm_name, norm_name, norm_name),
                    )
                    row = cur.fetchone()
                    if row and row.get("relationship_id"):
                        return str(row["relationship_id"])
                    # Check common nicknames e.g. Greg -> Gregory
                    if norm_name.startswith("greg "):
                        alt_name = "gregory " + norm_name[5:]
                        cur.execute(
                            """
                            SELECT c.relationship_id
                            FROM contacts c
                            INNER JOIN users u ON u.uuid = c.relationship_id
                            WHERE c.agency_id = %s
                              AND c.relationship_id IS NOT NULL
                              AND (u.tenant_id = %s OR u.agency_id = %s)
                              AND lower(concat(c.first_name, ' ', c.last_name)) = %s
                            LIMIT 1;
                            """,
                            (agency_id, agency_id, agency_id, alt_name),
                        )
                        row = cur.fetchone()
                        if row and row.get("relationship_id"):
                            return str(row["relationship_id"])
        except Exception as e:
            pass
        return None

    def fetch_all_users_map(self, agency_id: int) -> Dict[str, str]:
        """Map code/name -> user_uuid from users, plus contacts names that join users (relationship_id = uuid)."""
        if not self.conn:
            return {}
        user_map = {}
        self.user_display_by_uuid = {}
        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                # 1. Fetch from users table by tenant_id or agency_id
                cur.execute("""
                    SELECT uuid, code 
                    FROM users 
                    WHERE (tenant_id = %s OR agency_id = %s) AND uuid IS NOT NULL;
                """, (agency_id, agency_id))
                rows = cur.fetchall()
                for r in rows:
                    u_uuid = str(r.get("uuid", "")).strip()
                    if not u_uuid:
                        continue
                    raw_code = str(r.get("code") or "").strip()
                    if raw_code:
                        clean_c = clean_agent_code(raw_code).lower()
                        user_map[raw_code.lower()] = u_uuid
                        user_map[clean_c] = u_uuid
                        self.user_display_by_uuid[u_uuid.lower()] = {"name": "", "code": raw_code}

                # 2. Contacts name -> relationship_id only when join to users.uuid succeeds
                try:
                    cur.execute("""
                        SELECT c.relationship_id, c.first_name, c.last_name
                        FROM contacts c
                        INNER JOIN users u ON u.uuid = c.relationship_id
                        WHERE c.agency_id = %s
                          AND c.relationship_id IS NOT NULL
                          AND (u.tenant_id = %s OR u.agency_id = %s);
                    """, (agency_id, agency_id, agency_id))
                    c_rows = cur.fetchall()
                    for r in c_rows:
                        rel_id = str(r.get("relationship_id", "")).strip()
                        if not rel_id:
                            continue
                        fname = str(r.get("first_name", "") or "").strip()
                        lname = str(r.get("last_name", "") or "").strip()
                        full = " ".join(f"{fname} {lname}".lower().split())
                        if full and full not in user_map:
                            user_map[full] = rel_id
                        # Nickname variations (Greg <-> Gregory)
                        if fname.lower() == "gregory":
                            user_map[" ".join(f"greg {lname}".lower().split())] = rel_id
                        elif fname.lower() == "greg":
                            user_map[" ".join(f"gregory {lname}".lower().split())] = rel_id

                        full_display = " ".join(f"{fname} {lname}".split())
                        if full_display:
                            display = self.user_display_by_uuid.setdefault(rel_id.lower(), {"name": "", "code": ""})
                            display["name"] = full_display
                except Exception as ce:
                    pass
        except Exception as e:
            print(f"⚠️ Warning: Could not fetch user map from database: {e}")
        return user_map

    def format_writing_agent(self, writing_agent_user_id) -> str:
        """Display name for policies.writing_agent_user_id, from contacts then users."""
        uid = str(writing_agent_user_id or "").strip()
        if not uid or uid.lower() in ("none", "null"):
            return ""
        info = self.user_display_by_uuid.get(uid.lower(), {})
        name = str(info.get("name") or "").strip()
        code = str(info.get("code") or "").strip()
        if name and code:
            return f"{name} (Code: {code})"
        return name or (f"(Code: {code})" if code else "")

    def fetch_all_agency_commission_rules(self, agency_id: int, policy_numbers: Optional[List[str]] = None) -> Dict[Union[str, Tuple[str, str]], List[dict]]:
        """
        Executes two fast, indexed SQL queries to pre-fetch:
        1) Products & policies for the specified policy numbers (or entire agency).
        2) Matching policy_commission rules.
        Merges them in memory to eliminate Cartesian join bloat and complete in sub-second time.
        """
        if not self.conn:
            return {}
        result_map = {}
        try:
            clean_pols = [str(p).strip() for p in policy_numbers if str(p).strip()] if policy_numbers else []
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                # Query 1: Fetch policies + product details for the agency
                if clean_pols:
                    query_policies = """
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
                    """
                    cur.execute(query_policies, (agency_id, agency_id, clean_pols))
                else:
                    query_policies = """
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
                          AND pol.agency_id = %s;
                    """
                    cur.execute(query_policies, (agency_id, agency_id))
                policy_rows = cur.fetchall()

                # Query 2: Fetch policy_commission records for the matching policies
                # Include audit timestamps + creator/updater display names via contacts.relationship_id
                _comm_select = """
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
                            pc.commission_split,
                            pc.created_at,
                            pc.updated_at,
                            pc.created_by,
                            pc.updated_by,
                            TRIM(CONCAT(COALESCE(c_created.first_name, ''), ' ', COALESCE(c_created.last_name, ''))) AS created_by_name,
                            TRIM(CONCAT(COALESCE(c_updated.first_name, ''), ' ', COALESCE(c_updated.last_name, ''))) AS updated_by_name
                        FROM policy_commission pc
                        LEFT JOIN contacts c_created
                            ON c_created.relationship_id = pc.created_by
                           AND c_created.agency_id = pc.agency_id
                        LEFT JOIN contacts c_updated
                            ON c_updated.relationship_id = pc.updated_by
                           AND c_updated.agency_id = pc.agency_id
                """
                if clean_pols:
                    query_commissions = _comm_select + """
                        WHERE pc.agency_id = %s
                          AND pc.policy_uuid IN (
                              SELECT uuid FROM policies WHERE agency_id = %s AND number = ANY(%s)
                          );
                    """
                    cur.execute(query_commissions, (agency_id, agency_id, clean_pols))
                else:
                    query_commissions = _comm_select + """
                        WHERE pc.agency_id = %s;
                    """
                    cur.execute(query_commissions, (agency_id,))
                commission_rows = cur.fetchall()

                # Group commission records by policy_uuid
                comm_map = {}
                for c in commission_rows:
                    pol_uuid = str(c.get("policy_uuid", "")).strip()
                    if pol_uuid:
                        if pol_uuid not in comm_map:
                            comm_map[pol_uuid] = []
                        comm_map[pol_uuid].append(dict(c))

                # Combine policies with matching commission records
                for p in policy_rows:
                    p_num = str(p["policy_number"]).strip()
                    p_name = " ".join(str(p["product_name"]).lower().split())
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
                            "commission_split": None,
                            "created_at": None,
                            "updated_at": None,
                            "created_by": None,
                            "updated_by": None,
                            "created_by_name": None,
                            "updated_by_name": None,
                        })
                        recs = [rec]
                    else:
                        recs = []
                        for c in matching_comms:
                            rec = dict(p)
                            rec.update(c)
                            recs.append(rec)

                    # 1. Composite key index for exact product-policy pair matching
                    comp_key = (p_num, p_name)
                    if comp_key not in result_map:
                        result_map[comp_key] = []
                    result_map[comp_key].extend(recs)

                    # 2. Simple policy number index for fallback matching
                    if p_num not in result_map:
                        result_map[p_num] = []
                    result_map[p_num].extend(recs)

            return result_map
        except Exception as e:
            print(f"⚠️ Warning: Could not bulk fetch agency commission rules: {e}", flush=True)
            return {}


    def batch_fetch_policy_commission_rules(
        self, agency_id: int, exact_pols: List[str], similar_prefixes: List[str]
    ) -> Dict[Union[str, Tuple[str, str]], List[dict]]:
        """
        Executes a single Partition-First SQL Batch JOIN Query for a chunk of policy numbers.
        Returns dict mapping (policy_number, clean_product_name) -> List of joined DB record dicts.
        """
        if not self.conn or (not exact_pols and not similar_prefixes):
            return {}

        clean_exact = [str(p).strip() for p in exact_pols if str(p).strip()]
        clean_similar = [str(p).strip() for p in similar_prefixes if str(p).strip()]
        
        if not clean_exact and not clean_similar:
            return {}

        try:
            with self.conn.cursor(cursor_factory=RealDictCursor) as cur:
                query_conditions = []
                params = [agency_id, agency_id]
                
                if clean_exact:
                    placeholders = ", ".join(["%s"] * len(clean_exact))
                    query_conditions.append(f"pol.number IN ({placeholders})")
                    params.extend(clean_exact)
                    
                if clean_similar:
                    like_clauses = " OR ".join(["pol.number LIKE %s"] * len(clean_similar))
                    query_conditions.append(f"({like_clauses})")
                    params.extend([f"{p}%" for p in clean_similar])
                    
                where_clause = " OR ".join(query_conditions)
                
                query = f"""
                    SELECT
                        p.uuid AS product_uuid,
                        p.name AS product_name,
                        pol.uuid AS policy_uuid,
                        pol.number AS policy_number,
                        pol.writing_agent_user_id,
                        pol.pay_to_agent_user_id,
                        pc.uuid AS commission_uuid,
                        pc.rule_name,
                        pc.version,
                        pc.commission_type,
                        pc.payment_frequency,
                        pc.month_from,
                        pc.month_to,
                        pc.payout_method,
                        pc.fee_amount,
                        pc.premium_amount,
                        pc.commission_split,
                        pc.created_at,
                        pc.updated_at,
                        pc.created_by,
                        pc.updated_by,
                        TRIM(CONCAT(COALESCE(c_created.first_name, ''), ' ', COALESCE(c_created.last_name, ''))) AS created_by_name,
                        TRIM(CONCAT(COALESCE(c_updated.first_name, ''), ' ', COALESCE(c_updated.last_name, ''))) AS updated_by_name
                    FROM products p
                    JOIN policies pol ON (pol.agency_id = p.agency_id AND pol.product_id::text = p.uuid::text)
                    LEFT JOIN policy_commission pc ON (pc.agency_id = p.agency_id AND pc.policy_uuid::text = pol.uuid::text)
                    LEFT JOIN contacts c_created
                        ON c_created.relationship_id = pc.created_by
                       AND c_created.agency_id = pc.agency_id
                    LEFT JOIN contacts c_updated
                        ON c_updated.relationship_id = pc.updated_by
                       AND c_updated.agency_id = pc.agency_id
                    WHERE p.agency_id = %s
                      AND pol.agency_id = %s
                      AND ({where_clause});
                """
                cur.execute(query, params)
                rows = cur.fetchall()

                result_map = {}
                for r in rows:
                    p_num = str(r["policy_number"]).strip()
                    p_name = " ".join(str(r["product_name"]).lower().split())
                    rec_dict = dict(r)

                    # 1. Composite key index for exact product-policy pair matching
                    comp_key = (p_num, p_name)
                    if comp_key not in result_map:
                        result_map[comp_key] = []
                    result_map[comp_key].append(rec_dict)

                    # 2. Simple policy number index for fallback matching
                    if p_num not in result_map:
                        result_map[p_num] = []
                    result_map[p_num].append(rec_dict)

                return result_map

        except Exception as e:
            print(f"⚠️ Error executing batch policy JOIN query: {e}")
            return {}


def normalize_level(level_val: str) -> str:
    """
    Normalizes level strings to standard format:
    SUBAGT1 -> SA1, SUBAGT2 -> SA2, SUBAGT3 -> SA3 (and SUBAGT<N> -> SA<N>)
    LVL1 -> Level 1, LVL2 -> Level 2 (and LVL<N> -> Level <N>)
    """
    if not level_val:
        return ""
    val_upper = str(level_val).strip().upper()
    
    # SUBAGT<N> -> SA<N>
    subagt_match = re.match(r"^SUBAGT(\d+)$", val_upper)
    if subagt_match:
        return f"SA{subagt_match.group(1)}"
        
    # LVL<N> -> Level <N>
    lvl_match = re.match(r"^LVL(\d+)$", val_upper)
    if lvl_match:
        return f"Level {lvl_match.group(1)}"
        
    return str(level_val).strip()


def build_file_commission_split(group: pd.DataFrame, seq_col, agent_col, agent_name_col, contract_col, split_col) -> list:
    """
    Builds a commission_split JSON array from file rows (same group = same policy+product),
    in the exact same format as the import script's payload builder.
    Reads role/level from the 'Contract' column directly.
    """
    raw_items = []
    seen_keys = set()
    for idx in group.index:
        row = group.loc[idx]
        seq = int(row.get(seq_col, 0)) if seq_col and pd.notna(row.get(seq_col)) else 0
        raw_agent_id = str(row.get(agent_col, "")).strip() if agent_col else ""
        agent_id = clean_agent_code(raw_agent_id)  # Strip leading quotes and trailing alpha suffixes (e.g. '21758900 -> 21758900, 21758900SL -> 21758900)
        agent_name = str(row.get(agent_name_col, "")).strip() if agent_name_col else raw_agent_id
        contract = str(row.get(contract_col, "")).strip() if contract_col else ""
        commission_value = float(row.get(split_col, 0.0)) if split_col else 0.0
        # Override rows built from the regular DB rule carry its level and user uuid
        level_override = _clean_helper_text(row.get(SPLIT_LEVEL_COL))
        uuid_override = _clean_helper_text(row.get(SPLIT_USER_UUID_COL))

        # Deduplicate by (sequence, agent_code) — file has multiple From Month rows per agent
        dedup_key = (seq, agent_id)
        if dedup_key in seen_keys:
            continue
        seen_keys.add(dedup_key)

        raw_items.append({
            "sequence": seq,
            "user_name": agent_name,
            "user_code": agent_id if agent_id else None,
            "commission_value": commission_value,
            "contract": contract,
            "level_override": level_override,
            "uuid_override": uuid_override,
        })

    # Role-first order so writing Sub Agent is always highest order (4), not dropped behind LVL
    def _role_rank(item: dict) -> int:
        role = classify_contract_role(item.get("contract", ""))
        return {"agency": 0, "sales_leader": 1, "agent": 2, "sub_agent": 3}.get(role, 2)

    raw_items.sort(key=lambda x: (_role_rank(x), -x["sequence"]))

    commission_split = []
    for order_idx, item in enumerate(raw_items):
        order = order_idx + 1
        contract_val = item["contract"]
        norm_level = normalize_level(contract_val)
        role_key = classify_contract_role(contract_val)

        if role_key == "agency":
            commission_split.append({
                "order": 1,
                "role": "Agency",
                "user_name": item["user_name"],
                "user_code": "",
                "commission_value": round(item["commission_value"], 2),
                "entity": "agency",
                "level": "",
                "user_uuid": "62a97a01-2801-4c7c-9cd0-cf38f2e0a4f4",
            })
        elif role_key == "sales_leader":
            commission_split.append({
                "order": order,
                "role": "Sales Leader",
                "user_name": item["user_name"],
                "user_code": item["user_code"],
                "commission_value": round(item["commission_value"], 2),
                "entity": "agent",
                "level": norm_level or "Sales Leader",
            })
        elif role_key == "sub_agent":
            commission_split.append({
                "order": order,
                "role": "Sub Agent",
                "user_name": item["user_name"],
                "user_code": item["user_code"],
                "commission_value": round(item["commission_value"], 2),
                "entity": "agent",
                "level": norm_level or "SA1",
            })
        else:
            commission_split.append({
                "order": order,
                "role": "Agent",
                "user_name": item["user_name"],
                "user_code": item["user_code"],
                "commission_value": round(item["commission_value"], 2),
                "entity": "agent",
                "level": norm_level,
            })

        if role_key != "agency":
            if item.get("level_override"):
                commission_split[-1]["level"] = normalize_level(item["level_override"])
            if item.get("uuid_override"):
                commission_split[-1]["user_uuid"] = item["uuid_override"]

    # Re-number orders 1..n in list order (Agency first … Sub Agent last = writing agent)
    for i, split in enumerate(commission_split):
        split["order"] = i + 1
        if split.get("role") == "Agency":
            split["order"] = 1

    # Ensure Agency stays order 1 even if sort glitched
    commission_split.sort(
        key=lambda s: (
            0 if s.get("role") == "Agency" else
            1 if s.get("role") == "Sales Leader" else
            2 if s.get("role") == "Agent" else 3
        )
    )
    for i, split in enumerate(commission_split):
        split["order"] = i + 1

    return commission_split


def format_commission_audit_fields(db_rec: Optional[dict]) -> Tuple[str, str, str, str]:
    """
    Extract created_at / updated_at / created_by / updated_by for Stage 2 reports.
    created_by / updated_by prefer contacts first+last name (created_by_name / updated_by_name);
    fall back to raw UUID when the contact join had no name.
    """
    if not db_rec:
        return "", "", "", ""

    def _ts(val) -> str:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ""
        return str(val).strip()

    def _name(name_key: str, uuid_key: str) -> str:
        name = " ".join(str(db_rec.get(name_key) or "").split())
        if name:
            return name
        return str(db_rec.get(uuid_key) or "").strip()

    return (
        _ts(db_rec.get("created_at")),
        _ts(db_rec.get("updated_at")),
        _name("created_by_name", "created_by"),
        _name("updated_by_name", "updated_by"),
    )


_ISSUE_PARTY_LABELS = {
    "agency": "Agency",
    "sales_leader": "Sales Leader",
    "agent": "Agent",
    "sub_agent": "Sub Agent",
}


def _format_difference_value(value: float) -> str:
    """Compact difference display: 0.5, 0.12, 67.57 (no trailing zeros)."""
    try:
        num = abs(float(value))
    except (TypeError, ValueError):
        return ""
    if abs(num - round(num)) < 1e-9:
        return str(int(round(num)))
    text = f"{num:.4f}".rstrip("0").rstrip(".")
    return text


def _issue_person_display(name: str = "", code: str = "", role_key: str = "", item: Optional[dict] = None) -> str:
    """Agency/MLB → MLB; people → first name (e.g. Jason)."""
    if role_key == "agency" or _is_mlb_or_agency_person(name, code, item):
        return "MLB"
    cleaned = " ".join(str(name or "").split())
    if cleaned:
        return cleaned.split()[0].title()
    code_s = str(code or "").strip()
    return code_s


def build_unmatched_issue_summary(
    file_splits: Optional[list],
    db_splits: Optional[list],
    tolerance: float = 0.02,
    fee_file: Optional[float] = None,
    fee_db: Optional[float] = None,
) -> Tuple[str, str, str]:
    """
    Build unmatched report helpers:
      Difference      → 0.5|0.12
      Issue_Party     → Agent | Agency
      Affected_Person → Jason | MLB
    """
    file_list = file_splits if isinstance(file_splits, list) else []
    db_list = db_splits if isinstance(db_splits, list) else []

    def _role_key_of(item: dict) -> str:
        c_val = str(item.get("contract") or item.get("role") or item.get("level") or "").strip().lower()
        return classify_contract_role(c_val)

    def _norm_name(item: dict) -> str:
        return " ".join(clean_name_ignoring_id(item.get("user_name")).lower().split())

    used_db: Set[int] = set()
    matched_db_by_fidx: Dict[int, dict] = {}
    for f_idx, f_split in enumerate(file_list):
        f_role = _role_key_of(f_split)
        f_code = clean_agent_code(str(f_split.get("user_code") or "").strip().lower())
        f_name = _norm_name(f_split)

        best_match_idx = None
        for db_idx, db_item in enumerate(db_list):
            if db_idx in used_db:
                continue
            db_role = _role_key_of(db_item)
            if db_role != f_role:
                continue
            db_code = clean_agent_code(str(db_item.get("user_code") or "").strip().lower())
            db_name = _norm_name(db_item)
            if f_role == "agency" and db_role == "agency":
                best_match_idx = db_idx
                break
            if f_code and db_code and (f_code == db_code or are_codes_matching(f_code, db_code)):
                best_match_idx = db_idx
                break
            if f_name and db_name and (f_name == db_name or are_names_matching(f_name, db_name)):
                best_match_idx = db_idx
                break
        if best_match_idx is None:
            role_candidates = [
                db_idx for db_idx, db_item in enumerate(db_list)
                if db_idx not in used_db and _role_key_of(db_item) == f_role
            ]
            if len(role_candidates) == 1:
                best_match_idx = role_candidates[0]
        if best_match_idx is not None:
            used_db.add(best_match_idx)
            matched_db_by_fidx[f_idx] = db_list[best_match_idx]

    diff_parts: List[str] = []
    party_parts: List[str] = []
    person_parts: List[str] = []

    for f_idx, f_split in enumerate(file_list):
        role_key = _role_key_of(f_split)
        party = _ISSUE_PARTY_LABELS.get(role_key, role_key.replace("_", " ").title() or "Agent")
        person = _issue_person_display(
            str(f_split.get("user_name") or ""),
            str(f_split.get("user_code") or ""),
            role_key=role_key,
            item=f_split,
        )
        file_val = float(f_split.get("commission_value") or 0.0)
        db_s = matched_db_by_fidx.get(f_idx)
        if db_s is None:
            # Missing on DB side — treat file value as the gap (vs 0)
            if abs(file_val) > tolerance or not db_list:
                diff_parts.append(_format_difference_value(file_val))
                party_parts.append(party)
                person_parts.append(person or party)
            continue
        db_val = float(db_s.get("commission_value") or 0.0)
        gap = abs(file_val - db_val)
        if gap > tolerance:
            diff_parts.append(_format_difference_value(gap))
            party_parts.append(party)
            person_parts.append(person or party)

    if fee_file is not None and fee_db is not None:
        try:
            fee_gap = abs(float(fee_file) - float(fee_db))
        except (TypeError, ValueError):
            fee_gap = 0.0
        if fee_gap > tolerance:
            fee_txt = _format_difference_value(fee_gap)
            if "Agency" not in party_parts:
                diff_parts.append(fee_txt)
                party_parts.append("Agency")
                person_parts.append("MLB")
            elif fee_txt not in diff_parts:
                diff_parts.append(fee_txt)
                party_parts.append("Agency")
                person_parts.append("MLB")

    if not diff_parts:
        return "", "", ""
    return (
        "|".join(diff_parts),
        " | ".join(party_parts),
        " | ".join(person_parts),
    )


def _same_split_person(a: dict, b: dict) -> bool:
    """True when two split entries name the same person (uuid, code, or name; MLB aliases match)."""
    a_uuid = str(a.get("user_uuid") or a.get("uuid") or "").strip().lower()
    b_uuid = str(b.get("user_uuid") or b.get("uuid") or "").strip().lower()
    if a_uuid and b_uuid and a_uuid == b_uuid:
        return True
    a_code = clean_agent_code(str(a.get("user_code") or a.get("code") or "").strip().lower())
    b_code = clean_agent_code(str(b.get("user_code") or b.get("code") or "").strip().lower())
    if a_code and b_code and (a_code == b_code or are_codes_matching(a_code, b_code)):
        return True
    a_name = " ".join(clean_name_ignoring_id(a.get("user_name") or a.get("name")).lower().split())
    b_name = " ".join(clean_name_ignoring_id(b.get("user_name") or b.get("name")).lower().split())
    if _is_mlb_or_agency_person(a_name) and _is_mlb_or_agency_person(b_name):
        return True
    return bool(a_name and b_name and (a_name == b_name or are_names_matching(a_name, b_name)))


def db_override_accepts_file(db_splits: list, file_splits: list, sl_from_file: bool) -> bool:
    """A DB override that already has Agency 100% and a named Sales Leader + Agent is VALID.

    Fee amount and payout method are not compared for overrides.
    - Agent: must be the file's Agent; a different person means replace.
    - Sales Leader: kept when it is the same person. When the DB has someone else
      (often MLB), it is replaced only if the file sheet itself names a Sales Leader;
      otherwise the DB Sales Leader is kept.
    """
    db_splits = _split_list(db_splits)
    if not is_override_commission_split(db_splits):
        return False
    db_people = _hierarchy_people(db_splits)
    if "sales_leader" not in db_people or "agent" not in db_people:
        return False

    def _file_role(role_name: str) -> Optional[dict]:
        return next((s for s in file_splits or [] if str(s.get("role") or "") == role_name), None)

    f_agent = _file_role("Agent")
    f_sl = _file_role("Sales Leader")
    if f_agent and not _same_split_person(f_agent, db_people["agent"]):
        return False
    if f_sl and not _same_split_person(f_sl, db_people["sales_leader"]) and sl_from_file:
        return False
    return True


def find_db_slabs_inside_file_months(
    comm_records: list,
    file_bands: Set[Tuple[int, Optional[int]]],
    file_is_override: bool,
    skip_uuids: Set[str],
) -> list:
    """Same-type DB slabs that the file's months cover but that no file slab pairs with.

    Example: DB 1-12, 13-24 and file 1-11, 12-36. DB 1-12 pairs with file 1-11 (same
    From Month) and is replaced as a mismatch; DB 13-24 starts inside the file months
    but no file slab starts at 13, so it is returned here and deleted.

    Kept: DB slabs fully outside the file months, slabs with the exact months of a
    file slab (matched, already replaced, or a duplicate), and the other type
    (an override never overlaps a regular commission).
    """
    if not file_bands:
        return []
    file_intervals = [_interval_bounds(fm, tm or 0) for fm, tm in file_bands]
    found = []
    for rec in comm_records or []:
        uid = str(rec.get("commission_uuid") or "").strip()
        if not uid or uid in skip_uuids:
            continue
        if is_override_family(rule_effective_type(rec)) != file_is_override:
            continue
        db_fm = _month_to_int(rec.get("month_from")) or 1
        db_tm_raw = rec.get("month_to")
        db_tm = None if db_tm_raw is None or _month_to_int(db_tm_raw) <= 0 else _month_to_int(db_tm_raw)
        if (db_fm, db_tm) in file_bands:
            continue
        start, end = _interval_bounds(db_fm, db_tm or 0)
        if any(start <= f_end and f_start <= end for f_start, f_end in file_intervals):
            found.append(rec)
    return found


def stage2_db_verify_dataset(
    df_valid: pd.DataFrame,
    db_mgr: DatabaseManager,
    agency_id: int = 5,
    dry_run: bool = False,
    chunk_size: int = 2000,
    tolerance: float = 0.02,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, list, list, list, list]:
    """
    Stage 2: Database Verification for Stage-1 Valid Records.
    Uses chunked Partition-First SQL Batch JOIN Queries (2000 policy numbers per chunk).
    Returns: (matched_df, ambiguous_df, unmatched_df, missing_policy_payloads, mismatch_policy_payloads, invalid_user_uuid_payloads, writing_agent_corrections)
    """
    if df_valid.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), [], [], [], []

    df_out = df_valid.copy().reset_index(drop=True)

    pol_col = next((c for c in df_out.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
    prod_col = next((c for c in df_out.columns if str(c).strip().lower() in ("normalized_product_name", "product name", "product")), None)
    # Raw sheet product (OVR marker survives here even if the normalized name drops it)
    prod_col_raw_s2 = next((c for c in df_out.columns if str(c).strip().lower() in ("product", "product name")), None)
    rate_col = next((c for c in df_out.columns if str(c).strip().lower() in ("rate", "rate %", "commission %")), None)
    agent_col = next((c for c in df_out.columns if str(c).strip().lower() in ("agent id", "agent_id", "agentid")), None)
    agent_name_col = next((c for c in df_out.columns if "agent name" in str(c).strip().lower() or str(c).strip().lower() == "agent_name"), None)
    seq_col = next((c for c in df_out.columns if str(c).strip().lower() in ("sequence", "seq")), None)
    contract_col = next((c for c in df_out.columns if str(c).strip().lower() in ("contract", "level", "role")), None)
    split_col = "Normalized_Split_Share_%" if "Normalized_Split_Share_%" in df_out.columns else None
    fm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("from month", "from_month", "month", "month from", "month_from")), None)
    tm_col = next((c for c in df_out.columns if str(c).strip().lower() in ("to month", "to_month", "month to", "month_to")), None)
    amt_col = next((c for c in df_out.columns if str(c).strip().lower() in ("amount", "fee amount", "fee_amount", "carrier commission")), None)

    obtained_product_uuids = [""] * len(df_out)
    obtained_policy_uuids = [""] * len(df_out)
    stage2_statuses = [""] * len(df_out)
    unmatched_reasons = [""] * len(df_out)
    db_vals = [""] * len(df_out)
    file_vals = [""] * len(df_out)
    created_ats = [""] * len(df_out)
    updated_ats = [""] * len(df_out)
    created_bys = [""] * len(df_out)
    updated_bys = [""] * len(df_out)
    differences = [""] * len(df_out)
    issue_parties = [""] * len(df_out)
    affected_persons = [""] * len(df_out)

    def _set_audit(row_idx: int, db_rec: Optional[dict]):
        ca, ua, cb, ub = format_commission_audit_fields(db_rec)
        created_ats[row_idx] = ca
        updated_ats[row_idx] = ua
        created_bys[row_idx] = cb
        updated_bys[row_idx] = ub

    def _set_issue_summary(
        row_idx: int,
        file_splits: Optional[list],
        db_splits: Optional[list],
        fee_file: Optional[float] = None,
        fee_db: Optional[float] = None,
    ):
        diff_s, party_s, person_s = build_unmatched_issue_summary(
            file_splits,
            db_splits,
            tolerance=tolerance,
            fee_file=fee_file,
            fee_db=fee_db,
        )
        differences[row_idx] = diff_s
        issue_parties[row_idx] = party_s
        affected_persons[row_idx] = person_s

    matched_count = 0
    slabs_inside_deleted = 0
    ambiguous_count = 0
    unmatched_count = 0
    missing_policy_payloads = []
    mismatch_policy_payloads = []
    invalid_user_uuid_payloads = []

    grouped = df_out.groupby([pol_col, prod_col], sort=False) if pol_col and prod_col else df_out.groupby(df_out.index)
    group_items = list(grouped)
    total_groups = len(group_items)

    print(f"🔍 Starting Stage 2 Chunked DB Verification for {total_groups:,} policy groups ({len(df_out):,} rows)...")

    import_script_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "policy_commission_import_validation_script"))
    if import_script_dir not in sys.path:
        sys.path.append(import_script_dir)
    from import_policy_commission import commission_payload_issues, _person_name_key

    # Extract unique policy numbers for batch SQL chunking
    unique_exact_policy_numbers = []
    unique_similar_prefixes = []
    
    for _, group in group_items:
        if pol_col:
            seq_1_row = group.loc[group.index[0]]
            pol = str(seq_1_row[pol_col]).strip()
            is_mapped = bool(seq_1_row.get("_is_mapped", False)) if "_is_mapped" in group.columns else False
            
            if pol not in unique_exact_policy_numbers:
                unique_exact_policy_numbers.append(pol)
            if not is_mapped:
                for candidate in policy_suffix_candidates(pol):
                    if candidate not in unique_exact_policy_numbers:
                        unique_exact_policy_numbers.append(candidate)
            if not is_mapped and len(pol) > 11:
                prefix = pol[:11]
                if prefix not in unique_similar_prefixes:
                    unique_similar_prefixes.append(prefix)

    # Pre-fetch all user UUIDs mapping for agency
    db_user_map = db_mgr.fetch_all_users_map(agency_id) if not dry_run and db_mgr.conn else {}
    # File highest-order agent is the intended writing agent; log when DB differs
    writing_agent_corrections: List[dict] = []
    _wa_correction_seen: Set[str] = set()

    def writing_agent_name_for(record) -> str:
        if not isinstance(record, dict):
            return ""
        return db_mgr.format_writing_agent(record.get("writing_agent_user_id"))

    def _highest_order_person(splits: list) -> Optional[dict]:
        """Highest order in Agency→…→Sub Agent hierarchy = writing agent on the sheet."""
        best = None
        best_order = -1
        for s in splits or []:
            if not isinstance(s, dict):
                continue
            role = str(s.get("role") or "").strip().casefold()
            if role == "agency":
                continue
            if (
                role == "sub agent"
                and not str(s.get("user_name") or "").strip()
                and not str(s.get("user_uuid") or "").strip()
            ):
                continue
            try:
                order = int(s.get("order"))
            except (TypeError, ValueError):
                continue
            if order > best_order:
                best_order = order
                best = s
        return best

    def resolve_writing_agent_for_payload(
        sanitized_splits: list,
        db_record: Optional[dict],
        *,
        pol_number: str,
        pol_uuid: str,
    ) -> str:
        """
        Prefer the file's highest-order agent as writingAgentName (sheet is source of truth).

        When DB policies.writing_agent_user_id names a different person (e.g. Jacob Yale vs
        Michael Dunavant), record a correction so we can emit UPDATE SQL for senior eng.
        """
        highest = _highest_order_person(sanitized_splits)
        file_name = str((highest or {}).get("user_name") or "").strip()
        file_uuid = str((highest or {}).get("user_uuid") or "").strip()
        file_code = str((highest or {}).get("user_code") or "").strip()
        db_name = writing_agent_name_for(db_record) if db_record else ""

        if file_name and db_name and _person_name_key(file_name) != _person_name_key(db_name):
            dedup = str(pol_uuid or pol_number)
            if dedup and dedup not in _wa_correction_seen:
                _wa_correction_seen.add(dedup)
                writing_agent_corrections.append({
                    "policyNumber": pol_number,
                    "policyUuid": pol_uuid,
                    "dbWritingAgentName": db_name,
                    "fileWritingAgentName": file_name,
                    "fileWritingAgentUuid": file_uuid,
                    "fileWritingAgentCode": file_code,
                    "dbWritingAgentUserId": str((db_record or {}).get("writing_agent_user_id") or ""),
                })

        # Sheet hierarchy wins so commission POST is not blocked by a stale DB writing agent
        return file_name or db_name

    # Pre-fetch ALL products, policies, and policy_commission rules in 1 single bulk SQL query
    chunk_db_map = {}
    prefix_db_map = {}
    if not dry_run and db_mgr.conn:
        print(f" ⚡ Pre-fetching agency policies & commission rules for {len(unique_exact_policy_numbers):,} policies in bulk...")
        chunk_db_map = db_mgr.fetch_all_agency_commission_rules(agency_id, policy_numbers=unique_exact_policy_numbers)
        print(f"    ✅ Bulk loaded {len(chunk_db_map):,} policy/product indexes into RAM.")
        
        # Pre-index 11-char policy prefixes to eliminate O(N) linear set scans inside verification loop
        for k in chunk_db_map.keys():
            if isinstance(k, str) and len(k) >= 11:
                pref = k[:11]
                if pref not in prefix_db_map:
                    prefix_db_map[pref] = set()
                prefix_db_map[pref].add(k)

    not_found_products = set()
    not_found_policies = set()
    not_found_commissions = set()

    for group_idx, (_, group) in enumerate(group_items, start=1):
        indices = group.index.tolist()
        seq_1_row = group.loc[indices[0]]

        pol_number = str(seq_1_row.get(pol_col, "")).strip()
        prod_name = str(seq_1_row.get(prod_col, "")).strip()
        clean_prod_name = " ".join(prod_name.lower().split())
        # Per-group tracking for the "DB slab inside file months" delete pass
        _mismatch_start = len(mismatch_policy_payloads)
        _blocked_start = len(invalid_user_uuid_payloads)
        _group_bands: Set[Tuple[int, Optional[int]]] = set()
        _group_ctypes: Set[int] = set()
        _group_comm_records: list = []
        _group_pol_uuid = None

        if dry_run or not db_mgr.conn:
            mock_prod_uuid = f"mock-prod-{abs(hash(prod_name)) % 1000000:06d}"
            mock_pol_uuid = f"mock-pol-{abs(hash(pol_number)) % 1000000:06d}"

            for idx in indices:
                obtained_product_uuids[idx] = mock_prod_uuid
                obtained_policy_uuids[idx] = mock_pol_uuid
                stage2_statuses[idx] = "MATCHED"
                unmatched_reasons[idx] = ""
                db_vals[idx] = "DRY_RUN_MOCK"
                file_vals[idx] = "DRY_RUN_MOCK"
                matched_count += 1
        else:
            is_mapped = bool(seq_1_row.get("_is_mapped", False)) if "_is_mapped" in seq_1_row else False
            is_ambiguous = False
            ambiguous_reason = ""

            # Product names used to break a tie when a shortened number hits several policies:
            # normalized name (normalization list) first, then the file's own product name.
            _norm_prod_tb = str(seq_1_row.get("Normalized_Product_Name", "") or "").strip()
            _file_prod_tb = str(seq_1_row.get(prod_col_raw_s2, "") or "").strip() if prod_col_raw_s2 else prod_name

            db_records = lookup_policy_db_records(chunk_db_map, pol_number, clean_prod_name)
            if not db_records and not is_mapped:
                for candidate in policy_suffix_candidates(pol_number):
                    candidate_rows = chunk_db_map.get(candidate) or []
                    if not isinstance(candidate_rows, list):
                        candidate_rows = []
                    distinct_uuids = {
                        str(r.get("policy_uuid") or "").strip()
                        for r in candidate_rows
                        if isinstance(r, dict)
                    }
                    distinct_uuids.discard("")
                    if len(distinct_uuids) > 1:
                        picked, pick_reason = pick_policy_by_product(candidate_rows, _norm_prod_tb, _file_prod_tb)
                        if picked:
                            db_records = picked
                            break
                        is_ambiguous = True
                        ambiguous_reason = (
                            f"Policy number '{pol_number}' matched {len(distinct_uuids)} policies "
                            f"after removing the suffix ('{candidate}'); {pick_reason}. Not updated."
                        )
                        db_records = []
                        break
                    if len(distinct_uuids) == 1:
                        db_records = lookup_policy_db_records(chunk_db_map, candidate, clean_prod_name)
                        break

            if not db_records and not is_ambiguous and not is_mapped and len(pol_number) > 11:
                prefix = pol_number[:11]
                matched_full_pols = prefix_db_map.get(prefix, set())

                if len(matched_full_pols) > 1:
                    prefix_rows = []
                    for full_pol in sorted(matched_full_pols):
                        rows_k = chunk_db_map.get(full_pol) or []
                        if isinstance(rows_k, list):
                            prefix_rows.extend(rows_k)
                    picked, pick_reason = pick_policy_by_product(prefix_rows, _norm_prod_tb, _file_prod_tb)
                    if picked:
                        db_records = picked
                    else:
                        is_ambiguous = True
                        ambiguous_reason = (
                            f"Policy number '{pol_number[:11]}...' matched {len(matched_full_pols)} policies in DB; "
                            f"{pick_reason}. Not updated."
                        )
                        db_records = []
                elif len(matched_full_pols) == 1:
                    found_pol = next(iter(matched_full_pols))
                    db_records = lookup_policy_db_records(chunk_db_map, found_pol, clean_prod_name)

            # Fallback DB Queries if not in chunk map
            found_prod_uuid = db_records[0].get("product_uuid") if db_records else None
            found_pol_uuid = db_records[0].get("policy_uuid") if db_records else None

            if not found_prod_uuid and prod_name not in not_found_products:
                found_prod_uuid = db_mgr.get_product_uuid(agency_id, prod_name)
                if not found_prod_uuid:
                    not_found_products.add(prod_name)
            if not found_pol_uuid and found_prod_uuid and not is_ambiguous and pol_number not in not_found_policies:
                lookup_numbers = [pol_number]
                if not is_mapped:
                    lookup_numbers.extend(policy_suffix_candidates(pol_number))
                for lookup_number in lookup_numbers:
                    found_pol_uuid = db_mgr.get_policy_uuid(agency_id, lookup_number, found_prod_uuid)
                    if found_pol_uuid:
                        break
                if not found_pol_uuid:
                    not_found_policies.add(pol_number)

            if is_ambiguous:
                ambiguous_count += len(indices)
                for idx in indices:
                    obtained_product_uuids[idx] = None
                    obtained_policy_uuids[idx] = None
                    stage2_statuses[idx] = "AMBIGUOUS_SIMILAR_MATCH"
                    unmatched_reasons[idx] = ambiguous_reason or f"Policy number '{pol_number[:11]}...' matched multiple policies in DB."
                    db_vals[idx] = "N/A"
                    file_vals[idx] = "N/A"
            elif not found_prod_uuid or not found_pol_uuid:
                unmatched_count += len(indices)
                for idx in indices:
                    obtained_product_uuids[idx] = found_prod_uuid
                    obtained_policy_uuids[idx] = found_pol_uuid
                    stage2_statuses[idx] = "NO_POLICY_FOUND_IN_DB"
                    unmatched_reasons[idx] = f"Policy number '{pol_number}' not found in policies table."
                    db_vals[idx] = "N/A"
                    file_vals[idx] = "N/A"
            else:
                # Obtain all DB commission rules for policy
                comm_records = [r for r in db_records if r.get("commission_uuid")]
                if not comm_records and found_pol_uuid and found_pol_uuid not in not_found_commissions:
                    comm_records = db_mgr.get_policy_commission(agency_id, found_pol_uuid)
                    if not comm_records:
                        not_found_commissions.add(found_pol_uuid)
                _group_comm_records = comm_records or []
                _group_pol_uuid = found_pol_uuid

                if not comm_records:
                    unmatched_count += len(indices)
                    for idx in indices:
                        row = group.loc[idx]
                        obtained_product_uuids[idx] = found_prod_uuid
                        obtained_policy_uuids[idx] = found_pol_uuid
                        stage2_statuses[idx] = "NO_COMMISSION_RULE_FOUND_IN_DB"
                        unmatched_reasons[idx] = f"No commission rules in policy_commission table for policy_uuid '{found_pol_uuid}'."
                        db_vals[idx] = "N/A"
                        file_vals[idx] = ""
                        
                        raw_fm = row.get(fm_col, 1) if fm_col else 1
                        raw_tm = row.get(tm_col, 0) if tm_col else 0
                        try:
                            file_fm = int(float(str(raw_fm).strip().replace("'", "").replace('"', "")))
                        except:
                            file_fm = 1
                        try:
                            file_tm = int(float(str(raw_tm).strip().replace("'", "").replace('"', "")))
                        except:
                            file_tm = 0
                            
                        expected_db_tm = None if file_tm == 0 else file_tm
                        
                        if fm_col and tm_col:
                            tier_group = group[(group[fm_col] == raw_fm) & (group[tm_col] == raw_tm)]
                        elif fm_col:
                            tier_group = group[group[fm_col] == raw_fm]
                        elif tm_col:
                            tier_group = group[group[tm_col] == raw_tm]
                        else:
                            tier_group = group
                            
                        file_commission_split = build_file_commission_split(
                            tier_group, seq_col, agent_col, agent_name_col, contract_col, split_col
                        )
                        _set_issue_summary(idx, file_commission_split, [])
                        carrier_comm_raw = str(row.get("Carrier Commission", "") or row.get(rate_col, "") or row.get(amt_col, "")).strip()
                        file_payout_method_str = str(row.get("Payout Method", "")).strip().upper()
                        
                        fee_val = 0.0
                        prem_val = 0.0
                        if "FIXED FEE" in file_payout_method_str:
                            fee_val = parse_amount(carrier_comm_raw)
                            payout_method_code = 1
                        else:
                            # Default to PERCENTAGE payout method (code 2)
                            fee_val = parse_rate(carrier_comm_raw)
                            payout_method_code = 2

                        # Enrich file_commission_split items with user_uuid and sanitize values
                        sanitized_splits = []
                        missing_uuid_in_group = False
                        for split_item in file_commission_split:
                            s_item = split_item.copy()
                            c_val = float(s_item.get("commission_value") or 0.0)
                            s_item["commission_value"] = max(0.0, c_val) # Ensure no negative commission values
                            
                            if not s_item.get("user_uuid"):
                                u_code = str(s_item.get("user_code") or "").strip().lower()
                                u_name = " ".join(str(s_item.get("user_name") or "").strip().lower().split())
                                fetched_uuid = db_user_map.get(u_code) if u_code else None
                                if not fetched_uuid and u_name:
                                    fetched_uuid = db_user_map.get(u_name)
                                if not fetched_uuid and (u_code or u_name):
                                    fetched_uuid = db_mgr.get_user_uuid(agency_id, user_code=s_item.get("user_code"), user_name=s_item.get("user_name"))
                                    if u_code:
                                        db_user_map[u_code] = fetched_uuid
                                    if u_name:
                                        db_user_map[u_name] = fetched_uuid
                                if fetched_uuid:
                                    s_item["user_uuid"] = fetched_uuid
                            
                            # Check: ALL non-agency splits MUST have user_uuid (regardless of commission_value)
                            if s_item.get("entity") != "agency" and not s_item.get("user_uuid"):
                                missing_uuid_in_group = True
                            
                            sanitized_splits.append(s_item)

                        _ctype = resolve_commission_type(sanitized_splits, product_name=prod_name)
                        missing_payload = {
                            "policyUuid": found_pol_uuid,
                            "policyNumber": pol_number,
                            "writingAgentName": resolve_writing_agent_for_payload(
                                sanitized_splits,
                                db_records[0] if db_records else None,
                                pol_number=pol_number,
                                pol_uuid=found_pol_uuid,
                            ),
                            "ruleName": commission_rule_name_for_type(_ctype),
                            "version": 1,
                            "commissionType": _ctype,
                            "paymentFrequency": 1,
                            "payoutMethod": payout_method_code,
                            "monthFrom": file_fm,
                            "monthTo": expected_db_tm,
                            "feeAmount": fee_val,
                            "premiumAmount": prem_val,
                            "commissionSplit": sanitized_splits
                        }
                        
                        split_issues = commission_payload_issues(missing_payload)
                        if split_issues:
                            missing_payload["_issue_reason"] = "; ".join(split_issues)
                            invalid_user_uuid_payloads.append(missing_payload)
                        else:
                            missing_policy_payloads.append(missing_payload)
                else:
                    for idx in indices:
                        row = group.loc[idx]
                        obtained_product_uuids[idx] = found_prod_uuid
                        obtained_policy_uuids[idx] = found_pol_uuid

                        # Extract File Months
                        raw_fm = row.get(fm_col, 1) if fm_col else 1
                        raw_tm = row.get(tm_col, 0) if tm_col else 0
                        try:
                            file_fm = int(float(str(raw_fm).strip().replace("'", "").replace('"', "")))
                        except:
                            file_fm = 1
                        try:
                            file_tm = int(float(str(raw_tm).strip().replace("'", "").replace('"', "")))
                        except:
                            file_tm = 0
                            
                        # Normalize File To Month = 0 to DB expected NULL
                        expected_db_tm = None if file_tm == 0 else file_tm

                        # Build file commission_split JSON for this specific month tier
                        if fm_col and tm_col:
                            tier_group = group[(group[fm_col] == raw_fm) & (group[tm_col] == raw_tm)]
                        elif fm_col:
                            tier_group = group[group[fm_col] == raw_fm]
                        elif tm_col:
                            tier_group = group[group[tm_col] == raw_tm]
                        else:
                            tier_group = group
                            
                        file_commission_split = build_file_commission_split(
                            tier_group, seq_col, agent_col, agent_name_col, contract_col, split_col
                        )
                        file_split_json = json.dumps(file_commission_split)

                        row_contract = str(row.get(contract_col, "")).strip() if contract_col else ""
                        row_role = classify_contract_role(row_contract)

                        row_rate = parse_rate(row.get(rate_col, 0.0)) if rate_col else 0.0
                        row_split = float(row.get("Normalized_Split_Share_%", 0.0))
                        row_agent = str(row.get(agent_col, "")).strip() if agent_col else ""
                        row_agent_name = str(row.get(agent_name_col, "")).strip() if agent_name_col else row_agent
                        norm_row_name = " ".join(row_agent_name.lower().split())

                        # Determine file-side payout method and fee amount for comparison and PUT payload
                        _file_payout_str = str(row.get("Payout Method", "")).strip().upper()
                        _carrier_comm_raw = str(row.get("Carrier Commission", "") or row.get(rate_col, "") or row.get(amt_col, "")).strip()
                        if "FIXED FEE" in _file_payout_str:
                            _file_fee_val = parse_amount(_carrier_comm_raw)
                            _file_payout_method_code = 1
                        else:
                            _file_fee_val = parse_rate(_carrier_comm_raw)
                            _file_payout_method_code = 2

                        # File expects OVERRIDE (4) for OVR products / agency-100% — do not treat type-1 COMMISSION as matched
                        _file_ctype = resolve_commission_type(
                            file_commission_split, product_name=prod_name
                        )
                        # Override family = OVR product (file) / stored type 4 (DB)
                        _file_is_override = is_override_product(prod_name) or is_override_product(
                            row.get(prod_col_raw_s2, "") if prod_col_raw_s2 else ""
                        )
                        _group_ctypes.add(4 if _file_is_override else 1)
                        _group_bands.add((file_fm, expected_db_tm))
                        _tier_sl_from_file = (
                            OVR_SL_FROM_FILE_COL in tier_group.columns
                            and any(_clean_helper_text(v) == "Y" for v in tier_group[OVR_SL_FROM_FILE_COL])
                        )

                        match_found = False
                        unmatched_agent = False
                        period_found = False
                        matched_db_value = None
                        db_split_raw = None
                        matched_commission_uuid = None
                        matched_db_rec = None

                        # Score and sort comm_records for candidate matching:
                        # 2 = exact period match (fm and tm match expected)
                        # 1 = open-ended period match (fm matches, db_tm is NULL)
                        # 0 = match on fm only
                        def calc_comm_score(r):
                            # Override and regular commission are separate slab sets:
                            # never pair (or delete) a regular rule for an override row.
                            if is_override_family(rule_effective_type(r)) != _file_is_override:
                                return -1
                            db_fm_val = r.get("month_from")
                            db_tm_val = r.get("month_to")
                            try:
                                fm_i = int(db_fm_val) if db_fm_val is not None else 1
                            except (ValueError, TypeError):
                                fm_i = 1
                            if fm_i != file_fm:
                                return -1
                            if db_tm_val == expected_db_tm:
                                return 2
                            if db_tm_val is None:
                                return 1
                            return 0

                        matching_comm_records = [r for r in comm_records if calc_comm_score(r) >= 0]
                        matching_comm_records.sort(key=calc_comm_score, reverse=True)

                        if matching_comm_records:
                            period_found = True

                        for db_rec in matching_comm_records:
                            db_fm = db_rec.get("month_from")
                            db_tm = db_rec.get("month_to")
                            db_split_raw = db_rec.get("commission_split")
                            db_fee = float(db_rec.get("fee_amount") or 0.0)
                            db_payout_raw = db_rec.get("payout_method")
                            db_payout_code = 1 if (isinstance(db_payout_raw, int) and db_payout_raw == 1) or str(db_payout_raw).strip() in ("1", "FIXED FEE", "FIXED_FEE") else 2

                            # Override VALID rule: a type-4 DB rule for the same months with
                            # Agency 100% (everyone else 0%) and a named Sales Leader + Agent matches,
                            # whatever its fee value, as long as its payout type is the file's type
                            # (file percentage -> DB percentage, file amount -> DB fixed fee).
                            try:
                                _db_stored_type = int(db_rec.get("commission_type") or 1)
                            except (TypeError, ValueError):
                                _db_stored_type = 1
                            if (
                                _file_is_override
                                and _db_stored_type == 4
                                and db_tm == expected_db_tm
                                and db_payout_code == _file_payout_method_code
                                and db_override_accepts_file(db_split_raw, file_commission_split, _tier_sl_from_file)
                            ):
                                match_found = True
                                matched_commission_uuid = db_rec.get("commission_uuid")
                                matched_db_rec = db_rec
                                break

                            if isinstance(db_split_raw, list):
                                # Prioritize splits matching the row's contract role
                                sorted_items = sorted(
                                    db_split_raw,
                                    key=lambda it: 0 if (isinstance(it, dict) and str(it.get("role") or "").strip().lower() == row_role.lower()) else 1
                                )
                                for item in sorted_items:
                                    if not isinstance(item, dict):
                                        continue
                                        
                                    db_code = str(item.get("user_code") or "").strip()
                                    db_name = clean_name_ignoring_id(item.get("user_name"))
                                    norm_db_name = " ".join(db_name.lower().split())
                                    
                                    # Match Agent by user_code, fallback to normalized user_name
                                    is_agent_match = False
                                    if row_agent and db_code and (row_agent == db_code or are_codes_matching(row_agent, db_code)):
                                        is_agent_match = True
                                    elif norm_row_name and norm_db_name and (norm_row_name == norm_db_name or are_names_matching(norm_row_name, norm_db_name)):
                                        is_agent_match = True
                                    elif str(item.get("order", "")) == "1" and (
                                        ("main line benefits" in norm_row_name and "mlb parallel run" in norm_db_name) or
                                        ("mlb parallel run" in norm_row_name and "main line benefits" in norm_db_name)
                                    ):
                                        is_agent_match = True
                                        
                                    if is_agent_match:
                                        db_val = item.get("commission_value")
                                        if db_val is not None:
                                            matched_db_value = float(db_val)
                                            split_matches = (abs(matched_db_value - row_split) <= tolerance or abs(matched_db_value - row_rate) <= tolerance)
                                            fee_matches = (abs(db_fee - _file_fee_val) <= tolerance)
                                            payout_matches = (db_payout_code == _file_payout_method_code)
                                            try:
                                                db_ctype = int(db_rec.get("commission_type") or 1)
                                            except (TypeError, ValueError):
                                                db_ctype = 1
                                            type_matches = db_ctype == _file_ctype
                                            if (
                                                split_matches
                                                and fee_matches
                                                and payout_matches
                                                and type_matches
                                                and db_tm == expected_db_tm
                                            ):
                                                match_found = True
                                            matched_commission_uuid = db_rec.get("commission_uuid")
                                            matched_db_rec = db_rec
                                            if match_found:
                                                break
                                        unmatched_agent = True
                                        
                                if match_found:
                                    break
                                else:
                                    # Record the uuid even if it's a mismatch so we know which one to update
                                    matched_commission_uuid = db_rec.get("commission_uuid")
                                    matched_db_rec = db_rec
                                    
                            # Fallback for Fixed Fee rules where db_split_raw might be empty but fee_amount matches
                            if not match_found and not unmatched_agent and db_fee > 0 and abs(parse_amount(row.get(rate_col, 0)) - db_fee) <= tolerance:
                                matched_commission_uuid = db_rec.get("commission_uuid")
                                matched_db_rec = db_rec
                                if db_tm == expected_db_tm:
                                    match_found = True
                                    break
                            
                            # If we get here, it means we found a rule for the period but failed the agent/rate check
                            if not match_found:
                                matched_commission_uuid = db_rec.get("commission_uuid")
                                matched_db_rec = db_rec

                        if match_found:
                            stage2_statuses[idx] = "MATCHED"
                            unmatched_reasons[idx] = ""
                            db_vals[idx] = json.dumps(db_split_raw) if isinstance(db_split_raw, (list, dict)) else str(db_split_raw)
                            file_vals[idx] = file_split_json
                            _set_audit(idx, matched_db_rec)
                            matched_count += 1
                        else:
                            if not period_found:
                                stage2_statuses[idx] = "NO_COMMISSION_RULE_FOUND_FOR_PERIOD"
                                unmatched_reasons[idx] = f"No commission rule found in DB for period {file_fm} -> {expected_db_tm}."
                                db_vals[idx] = "N/A"
                                file_vals[idx] = file_split_json
                                _set_issue_summary(idx, file_commission_split, [])
                                unmatched_count += 1
                                
                                fee_val = 0.0
                                prem_val = 0.0
                                carrier_comm_raw = str(row.get("Carrier Commission", "") or row.get(rate_col, "") or row.get(amt_col, "")).strip()
                                file_payout_method_str = str(row.get("Payout Method", "")).strip().upper()
                                if "FIXED FEE" in file_payout_method_str:
                                    fee_val = parse_amount(carrier_comm_raw)
                                    payout_method_code = 1
                                else:
                                    fee_val = parse_rate(carrier_comm_raw)
                                    payout_method_code = 2
                                    
                                sanitized_splits = []
                                missing_uuid_in_group = False
                                for split_item in file_commission_split:
                                    s_item = split_item.copy()
                                    c_val = float(s_item.get("commission_value") or 0.0)
                                    s_item["commission_value"] = max(0.0, c_val)
                                    if not s_item.get("user_uuid"):
                                        u_code = str(s_item.get("user_code") or "").strip().lower()
                                        u_name = " ".join(str(s_item.get("user_name") or "").strip().lower().split())
                                        fetched_uuid = db_user_map.get(u_code) if u_code else None
                                        if not fetched_uuid and u_name:
                                            fetched_uuid = db_user_map.get(u_name)
                                        if not fetched_uuid and (u_code or u_name):
                                            fetched_uuid = db_mgr.get_user_uuid(agency_id, user_code=s_item.get("user_code"), user_name=s_item.get("user_name"))
                                            if u_code: db_user_map[u_code] = fetched_uuid
                                            if u_name: db_user_map[u_name] = fetched_uuid
                                        if fetched_uuid:
                                            s_item["user_uuid"] = fetched_uuid
                                    # Check: ALL non-agency splits MUST have user_uuid (regardless of commission_value)
                                    if s_item.get("entity") != "agency" and not s_item.get("user_uuid"):
                                        missing_uuid_in_group = True
                                    sanitized_splits.append(s_item)

                                _ctype = resolve_commission_type(sanitized_splits, product_name=prod_name)
                                missing_payload = {
                                    "policyUuid": found_pol_uuid,
                                    "policyNumber": pol_number,
                                    "writingAgentName": resolve_writing_agent_for_payload(
                                        sanitized_splits,
                                        db_records[0] if db_records else None,
                                        pol_number=pol_number,
                                        pol_uuid=found_pol_uuid,
                                    ),
                                    "ruleName": commission_rule_name_for_type(_ctype),
                                    "version": 1,
                                    "commissionType": _ctype,
                                    "paymentFrequency": 1,
                                    "payoutMethod": payout_method_code,
                                    "monthFrom": file_fm,
                                    "monthTo": expected_db_tm,
                                    "feeAmount": fee_val,
                                    "premiumAmount": prem_val,
                                    "commissionSplit": sanitized_splits
                                }
                                split_issues = commission_payload_issues(missing_payload)
                                if split_issues:
                                    missing_payload["_issue_reason"] = "; ".join(split_issues)
                                    invalid_user_uuid_payloads.append(missing_payload)
                                else:
                                    missing_policy_payloads.append(missing_payload)
                            else:
                                stage2_statuses[idx] = "DB_COMMISSION_VALUE_MISMATCH"
                                if unmatched_agent:
                                    matched_db_value_str = f"{matched_db_value}%" if matched_db_value is not None else "None"
                                    db_fee_curr = float(matched_db_rec.get("fee_amount") or 0.0) if matched_db_rec else 0.0
                                    _ovr_reason = []
                                    if _file_is_override and matched_db_rec:
                                        # Override: fee value is not compared; type, Agency 100% split and months are
                                        _db_pay = matched_db_rec.get("payout_method")
                                        _db_pay_code = 1 if (isinstance(_db_pay, int) and _db_pay == 1) or str(_db_pay).strip() in ("1", "FIXED FEE", "FIXED_FEE") else 2
                                        _lbl = {1: "Fixed Fee", 2: "Percentage"}
                                        if _db_pay_code != _file_payout_method_code:
                                            _ovr_reason.append(
                                                f"override payout type is {_lbl[_db_pay_code]} in DB but {_lbl[_file_payout_method_code]} in file"
                                            )
                                        if not is_override_commission_split(_split_list(matched_db_rec.get("commission_split"))):
                                            _ovr_reason.append("override split is not Agency 100% with everyone else 0%")
                                        if matched_db_rec.get("month_to") != expected_db_tm:
                                            _ovr_reason.append(
                                                f"override months differ (DB to {matched_db_rec.get('month_to')}, file to {expected_db_tm})"
                                            )
                                    if _ovr_reason:
                                        unmatched_reasons[idx] = f"Override for '{row_agent_name}': " + "; ".join(_ovr_reason) + ". "
                                    elif _file_is_override and matched_db_rec:
                                        unmatched_reasons[idx] = (
                                            f"Override for '{row_agent_name}': the Sales Leader or Agent in DB is a different person. "
                                        )
                                    elif matched_db_rec and abs(db_fee_curr - _file_fee_val) > tolerance:
                                        unmatched_reasons[idx] = (
                                            f"Agent '{row_agent_name}' found in DB, but DB fee_amount is {db_fee_curr} (File has {_file_fee_val}). "
                                        )
                                    else:
                                        unmatched_reasons[idx] = (
                                            f"Agent '{row_agent_name}' found in DB, but DB commission_value is {matched_db_value_str} (File has {row_split}%). "
                                        )
                                else:
                                    code_str = f" (Code: {row_agent})" if row_agent else ""
                                    unmatched_reasons[idx] = (
                                        f"No commission rule record found in policy_commission table for agent '{row_agent_name}'{code_str}. "
                                    )
                                db_vals[idx] = json.dumps(db_split_raw) if isinstance(db_split_raw, (list, dict)) else str(db_split_raw)
                                file_vals[idx] = file_split_json
                                _set_audit(idx, matched_db_rec)
                                _set_issue_summary(
                                    idx,
                                    file_commission_split,
                                    db_split_raw if isinstance(db_split_raw, list) else [],
                                    fee_file=_file_fee_val,
                                    fee_db=float(matched_db_rec.get("fee_amount") or 0.0) if matched_db_rec else None,
                                )
                                unmatched_count += 1
                                
                                if matched_commission_uuid and matched_db_rec:
                                    # Merge DB fields into file splits and build diff using role-aware matching
                                    merged_splits = []
                                    diff_lines = []
                                    
                                    db_splits_list = db_split_raw if isinstance(db_split_raw, list) else []

                                    def _role_key_of(item: dict) -> str:
                                        c_val = str(item.get("contract") or item.get("role") or item.get("level") or "").strip().lower()
                                        return classify_contract_role(c_val)

                                    used_db_indices = set()
                                    matched_db_by_fidx = {}
                                    for f_idx, f_split in enumerate(file_commission_split):
                                        f_role = _role_key_of(f_split)
                                        f_code = clean_agent_code(str(f_split.get("user_code") or "").strip().lower())
                                        f_name = _person_name_key(str(f_split.get("user_name") or ""))

                                        best_match_idx = None
                                        # 1. Match by role AND (code or name)
                                        for db_idx, db_item in enumerate(db_splits_list):
                                            if db_idx in used_db_indices:
                                                continue
                                            db_role = _role_key_of(db_item)
                                            if db_role != f_role:
                                                continue
                                            db_code = clean_agent_code(str(db_item.get("user_code") or "").strip().lower())
                                            db_name = _person_name_key(str(db_item.get("user_name") or ""))
                                            if f_role == "agency" and db_role == "agency":
                                                best_match_idx = db_idx
                                                break
                                            if f_code and db_code and (f_code == db_code or are_codes_matching(f_code, db_code)):
                                                best_match_idx = db_idx
                                                break
                                            if f_name and db_name and (f_name == db_name or are_names_matching(f_name, db_name)):
                                                best_match_idx = db_idx
                                                break

                                        # 2. Match by role alone if only one candidate with that role exists
                                        if best_match_idx is None:
                                            role_candidates = [
                                                db_idx for db_idx, db_item in enumerate(db_splits_list)
                                                if db_idx not in used_db_indices and _role_key_of(db_item) == f_role
                                            ]
                                            if len(role_candidates) == 1:
                                                best_match_idx = role_candidates[0]

                                        if best_match_idx is not None:
                                            used_db_indices.add(best_match_idx)
                                            matched_db_by_fidx[f_idx] = db_splits_list[best_match_idx]

                                    for f_idx, f_split in enumerate(file_commission_split):
                                        m_split = f_split.copy()
                                        order_str = str(m_split.get("order", ""))
                                        db_s = matched_db_by_fidx.get(f_idx)
                                        
                                        f_code = str(m_split.get("user_code") or "").strip().lower()
                                        f_name = " ".join(str(m_split.get("user_name") or "").strip().lower().split())

                                        if db_s:
                                            db_code = str(db_s.get("user_code") or "").strip().lower()
                                            db_name = " ".join(clean_name_ignoring_id(db_s.get("user_name")).lower().split())

                                            # Cache DB user_uuid for fast lookup
                                            if db_s.get("user_uuid"):
                                                if db_code:
                                                    db_user_map[db_code] = db_s["user_uuid"]
                                                if db_name:
                                                    db_user_map[db_name] = db_s["user_uuid"]

                                            agent_changed = False
                                            if f_code and db_code:
                                                if not (f_code == db_code or are_codes_matching(f_code, db_code) or are_names_matching(f_name, db_name)):
                                                    agent_changed = True
                                            elif f_name and db_name:
                                                if not (f_name == db_name or are_names_matching(f_name, db_name)):
                                                    agent_changed = True

                                            old_val = float(db_s.get("commission_value") or 0.0)
                                            new_val = float(m_split.get("commission_value") or 0.0)
                                            val_changed = abs(old_val - new_val) > tolerance

                                            db_user_disp = f"{db_s.get('user_name', '')} ({db_s.get('user_code', '')})".strip() or "N/A"
                                            f_user_disp = f"{m_split.get('user_name', '')} ({m_split.get('user_code', '')})".strip() or "N/A"

                                            if agent_changed and val_changed:
                                                diff_lines.append(f"   Order {order_str} | User: {db_user_disp} ➔ {f_user_disp} | Rate: {old_val}% ➔ {new_val}%")
                                            elif agent_changed:
                                                diff_lines.append(f"   Order {order_str} | User: {db_user_disp} ➔ {f_user_disp}")
                                            elif val_changed:
                                                diff_lines.append(f"   Order {order_str} | User: {f_user_disp} | Rate: {old_val}% ➔ {new_val}%")

                                            # Same person: keep the DB identity (uuid / name / code). A different person in
                                            # this role: the file's person is posted. Either way the FILE value is posted.
                                            if not agent_changed:
                                                for key in ["user_uuid", "user_name", "user_code"]:
                                                    if key in db_s:
                                                        m_split[key] = db_s[key]
                                            for key in ["role", "level"]:
                                                if key in db_s and not m_split.get(key):
                                                    m_split[key] = db_s[key]
                                            # File (rolled-up) commission_value is posted, not the DB value
                                            m_split["commission_value"] = new_val
                                        else:
                                            diff_lines.append(f"   Order {order_str} | User: [NEW] ➔ {m_split.get('user_name', 'N/A')} | Rate: [NEW] ➔ {m_split.get('commission_value', 0.0)}%")
                                            if "level" in m_split:
                                                m_split["level"] = normalize_level(m_split["level"])
                                            # A person only in the file is posted with the file value (already on m_split)

                                        if str(m_split.get("order", "")) == "1" or str(m_split.get("role", "")).lower() == "agency":
                                            m_split["user_uuid"] = "62a97a01-2801-4c7c-9cd0-cf38f2e0a4f4"
                                            m_split["user_code"] = ""
                                            m_split["level"] = ""

                                        # Lookup user_uuid if missing (e.g. for new orders)
                                        if not m_split.get("user_uuid"):
                                            u_code = str(m_split.get("user_code") or "").strip().lower()
                                            u_name = " ".join(str(m_split.get("user_name") or "").strip().lower().split())
                                            fetched_uuid = db_user_map.get(u_code) if u_code else None
                                            if not fetched_uuid and u_name:
                                                fetched_uuid = db_user_map.get(u_name)
                                            if not fetched_uuid and (u_code or u_name):
                                                fetched_uuid = db_mgr.get_user_uuid(agency_id, user_code=m_split.get("user_code"), user_name=m_split.get("user_name"))
                                                if u_code:
                                                    db_user_map[u_code] = fetched_uuid
                                                if u_name:
                                                    db_user_map[u_name] = fetched_uuid
                                            if fetched_uuid:
                                                m_split["user_uuid"] = fetched_uuid

                                        c_val = float(m_split.get("commission_value") or 0.0)
                                        m_split["commission_value"] = max(0.0, c_val)

                                        merged_splits.append(m_split)

                                    # People in the DB split but not in the file are DROPPED (not carried over):
                                    # the file split is posted as it is. Each dropped person is listed in the diff.
                                    file_roles = {_role_key_of(f) for f in file_commission_split}
                                    file_people = set()
                                    for f_split in file_commission_split:
                                        f_code = clean_agent_code(str(f_split.get("user_code") or "").strip().lower())
                                        f_name = _person_name_key(str(f_split.get("user_name") or ""))
                                        if f_code:
                                            file_people.add(("code", f_code))
                                        if f_name:
                                            file_people.add(("name", f_name))
                                    for db_idx, db_s in enumerate(db_splits_list):
                                        if db_idx in used_db_indices or _role_key_of(db_s) in file_roles:
                                            continue
                                        db_code = clean_agent_code(str(db_s.get("user_code") or "").strip().lower())
                                        db_name = _person_name_key(str(db_s.get("user_name") or ""))
                                        if (db_code and ("code", db_code) in file_people) or (db_name and ("name", db_name) in file_people):
                                            continue
                                        if db_s.get("user_name") or float(db_s.get("commission_value") or 0.0) != 0.0:
                                            diff_lines.append(
                                                f"   Order {db_s.get('order', '')} | User: {db_s.get('user_name', '')} "
                                                f"({db_s.get('role', '')}) dropped - not in file "
                                                f"(was {float(db_s.get('commission_value') or 0.0)}%)"
                                            )

                                    # Override split-sum correction: when Agency takes 100% in an OVERRIDE, set non-agency to 0% to prevent 199% sums
                                    _candidate_ctype = resolve_commission_type(
                                        merged_splits,
                                        product_name=prod_name,
                                        db_commission_type=matched_db_rec.get("commission_type", 1),
                                    )
                                    _is_override_rule = _file_is_override or _candidate_ctype == 4
                                    if _is_override_rule:
                                        # Override: Agency always 100%, everyone else 0%
                                        agency_item = next((s for s in merged_splits if str(s.get("role", "")).lower() == "agency"), None)
                                        if agency_item:
                                            agency_item["commission_value"] = 100.0
                                            for s in merged_splits:
                                                if s is not agency_item:
                                                    s["commission_value"] = 0.0

                                    merged_splits.sort(key=lambda s: int(s.get("order") or 999))
                                    db_fm_val = int(matched_db_rec.get("month_from") or 1)
                                    db_tm_val = int(matched_db_rec.get("month_to") or 12) if matched_db_rec.get("month_to") is not None else "None"
                                    file_tm_val = expected_db_tm if expected_db_tm is not None else "None"
                                    if db_fm_val != file_fm or matched_db_rec.get("month_to") != expected_db_tm:
                                        diff_lines.append(f"   Period: {db_fm_val}-{db_tm_val} ➔ {file_fm}-{file_tm_val}")

                                    db_fee_val = float(matched_db_rec.get("fee_amount") or 0.0)
                                    # Override: the fee value is not compared or updated (only the payout type)
                                    if not _is_override_rule and abs(db_fee_val - _file_fee_val) > tolerance:
                                        payout_label = "Fixed Fee" if _file_payout_method_code == 1 else "Rate"
                                        diff_lines.append(f"   Fee/Rate ({payout_label}): {db_fee_val} ➔ {_file_fee_val}")

                                    db_payout_raw = matched_db_rec.get("payout_method")
                                    db_payout_code = 1 if (isinstance(db_payout_raw, int) and db_payout_raw == 1) or str(db_payout_raw).strip() in ("1", "FIXED FEE", "FIXED_FEE") else 2
                                    if db_payout_code != _file_payout_method_code:
                                        old_p_label = "Fixed Fee" if db_payout_code == 1 else "Percentage"
                                        new_p_label = "Fixed Fee" if _file_payout_method_code == 1 else "Percentage"
                                        diff_lines.append(f"   Payout Method: {old_p_label} ➔ {new_p_label}")

                                    _ctype = resolve_commission_type(
                                        merged_splits,
                                        product_name=prod_name,
                                        db_commission_type=matched_db_rec.get("commission_type", 1),
                                    )
                                    try:
                                        db_ctype_for_diff = int(matched_db_rec.get("commission_type") or 1)
                                    except (TypeError, ValueError):
                                        db_ctype_for_diff = 1
                                    if db_ctype_for_diff != _ctype:
                                        _type_labels = {1: "COMMISSION", 2: "RENEWAL", 3: "BONUS", 4: "OVERRIDE"}
                                        diff_lines.append(
                                            f"   Commission Type: {_type_labels.get(db_ctype_for_diff, db_ctype_for_diff)} "
                                            f"➔ {_type_labels.get(_ctype, _ctype)}"
                                        )

                                    if not diff_lines:
                                        # If only names were fuzzy matched or rates differed by <0.00%, don't bother PUTting
                                        continue

                                    diff_text = "\n".join(diff_lines)

                                    _default_rule = commission_rule_name_for_type(_ctype)
                                    _rule_name = matched_db_rec.get("rule_name") or _default_rule
                                    # Force OVERRIDE naming when file is agency-100%
                                    if _ctype == 4 and "override" not in str(_rule_name).lower():
                                        _rule_name = _default_rule
                                    _wa_db_rec = (
                                        matched_db_rec
                                        if matched_db_rec.get("writing_agent_user_id")
                                        else (db_records[0] if db_records else None)
                                    )
                                    # Replacement rule: the FILE split values, fee and payout type are posted
                                    # (only uuid / name / code of an unchanged person and premium come from the DB).
                                    # Override: the file payout type is posted, the DB fee value is kept,
                                    # and the split is Agency 100% / everyone else 0% (set above).
                                    mismatch_payload = {
                                        "rule_uuid": matched_commission_uuid,
                                        "policyNumber": pol_number,
                                        "policyUuid": matched_db_rec.get("policy_uuid"),
                                        "writingAgentName": resolve_writing_agent_for_payload(
                                            file_commission_split,
                                            _wa_db_rec,
                                            pol_number=pol_number,
                                            pol_uuid=str(matched_db_rec.get("policy_uuid") or found_pol_uuid or ""),
                                        ),
                                        "ruleName": _rule_name,
                                        "version": matched_db_rec.get("version") or 1,
                                        "commissionType": _ctype,
                                        "paymentFrequency": matched_db_rec.get("payment_frequency", 1),
                                        "payoutMethod": _file_payout_method_code,
                                        "monthFrom": file_fm,
                                        "monthTo": expected_db_tm,
                                        "feeAmount": db_fee_val if _is_override_rule else _file_fee_val,
                                        "premiumAmount": float(matched_db_rec.get("premium_amount") or 0.0),
                                        "commissionSplit": merged_splits,
                                        "_diff_text": diff_text
                                    }
                                    split_issues = commission_payload_issues(mismatch_payload)

                                    # The file split is posted, so it is no longer compared with the DB split total;
                                    # commission_payload_issues() above already requires the posted split to total 100%.
                                    # No person in the posted split may have a negative share (e.g. a carried-over DB -0.31%).
                                    for _s in merged_splits:
                                        _cv = float(_s.get("commission_value") or 0.0)
                                        if _cv < 0:
                                            split_issues.append(
                                                f"Order {_s.get('order', '')} {_s.get('role', '')} "
                                                f"'{_s.get('user_name', '')}' has a negative share ({_cv:g}%)"
                                            )

                                    if split_issues:
                                        mismatch_payload["_issue_reason"] = "; ".join(split_issues)
                                        invalid_user_uuid_payloads.append(mismatch_payload)
                                    else:
                                        mismatch_policy_payloads.append(mismatch_payload)

        # DB slabs inside the file's months that no file slab pairs with (e.g. DB 13-24
        # when the file has 1-11, 12-36) are deleted so the posted file slabs do not overlap.
        # Held when any payload of this group was blocked, so a delete never runs without
        # its replacement, or when the group mixes override and regular rows.
        if _group_comm_records and _group_bands and len(_group_ctypes) == 1:
            if len(invalid_user_uuid_payloads) > _blocked_start:
                pass
            else:
                _replaced_uuids = {
                    str(p.get("rule_uuid") or "").strip()
                    for p in mismatch_policy_payloads[_mismatch_start:]
                }
                _inside = find_db_slabs_inside_file_months(
                    _group_comm_records,
                    _group_bands,
                    file_is_override=(4 in _group_ctypes),
                    skip_uuids=_replaced_uuids,
                )
                _bands_txt = ", ".join(
                    f"{fm}-{tm if tm is not None else 'open'}" for fm, tm in sorted(_group_bands, key=lambda b: b[0])
                )
                for _rec in _inside:
                    _to = _rec.get("month_to")
                    _db_months = f"{_month_to_int(_rec.get('month_from')) or 1}-{_to if _to is not None else 'open'}"
                    mismatch_policy_payloads.append({
                        "rule_uuid": _rec.get("commission_uuid"),
                        "policyNumber": pol_number,
                        "policyUuid": _rec.get("policy_uuid") or _group_pol_uuid,
                        "monthFrom": _rec.get("month_from"),
                        "monthTo": _to,
                        "commissionType": rule_effective_type(_rec),
                        "_delete_only": True,
                        "_diff_text": (
                            f"   DB slab {_db_months} deleted: it sits inside the file months "
                            f"({_bands_txt}), which replace it"
                        ),
                    })
                    slabs_inside_deleted += 1

        if group_idx % 50 == 0 or group_idx == total_groups:
            pct = (group_idx / total_groups) * 100.0
            sys.stdout.write(
                f"\r ⚡ [Stage 2 Chunked DB Verification] {group_idx:,}/{total_groups:,} policy groups verified ({pct:.1f}%) | "
                f"Matched Rows: {matched_count:,} | Unmatched Rows: {unmatched_count:,}"
            )
            sys.stdout.flush()

    print(f"\n   ✅ Stage 2 Database Verification Completed! {matched_count:,} matched rows, {unmatched_count:,} unmatched rows.")
    if slabs_inside_deleted:
        print(
            f"   🧹 {slabs_inside_deleted:,} DB slab(s) sit inside the file months without a matching file slab; "
            f"queued for delete only (no POST)."
        )

    df_out["obtained_product_uuid"] = obtained_product_uuids
    df_out["obtained_policy_uuid"] = obtained_policy_uuids
    df_out["Stage2_Status"] = stage2_statuses
    df_out["Stage2_Unmatched_Reason"] = unmatched_reasons
    df_out["DB_Value"] = db_vals
    df_out["File_Value"] = file_vals
    df_out["created_at"] = created_ats
    df_out["updated_at"] = updated_ats
    df_out["created_by"] = created_bys
    df_out["updated_by"] = updated_bys
    df_out["Difference"] = differences
    df_out["Issue_Party"] = issue_parties
    df_out["Affected_Person"] = affected_persons

    matched_mask = df_out["Stage2_Status"] == "MATCHED"
    ambiguous_mask = df_out["Stage2_Status"] == "AMBIGUOUS_SIMILAR_MATCH"

    df_matched = df_out[matched_mask].copy()
    df_ambiguous = df_out[ambiguous_mask].copy()
    df_unmatched = df_out[~(matched_mask | ambiguous_mask)].copy()

    summary = {
        "stage2_processed": len(df_valid),
        "matched_records": len(df_matched),
        "ambiguous_records": len(df_ambiguous),
        "unmatched_records": len(df_unmatched),
    }

    # Dedup the missing payloads by policyUuid and monthFrom
    unique_missing_payloads = []
    seen_missing = set()
    for mp in missing_policy_payloads:
        dedup_key = (mp["policyUuid"], mp.get("monthFrom"), mp.get("monthTo"))
        if dedup_key not in seen_missing:
            seen_missing.add(dedup_key)
            unique_missing_payloads.append(mp)

    # Dedup the mismatch payloads by rule_uuid
    unique_mismatch_payloads = []
    seen_mismatches = set()
    # Replace payloads first, so a delete-only entry never hides the POST for the same rule
    for mp in sorted(mismatch_policy_payloads, key=lambda x: bool(x.get("_delete_only"))):
        if mp["rule_uuid"] not in seen_mismatches:
            seen_mismatches.add(mp["rule_uuid"])
            unique_mismatch_payloads.append(mp)

    return (
        df_matched,
        df_ambiguous,
        df_unmatched,
        unique_missing_payloads,
        unique_mismatch_payloads,
        invalid_user_uuid_payloads,
        writing_agent_corrections,
    )


def write_updated_policies_list(
    successful_policies: Set[str],
    output_dir: str,
    df_updated: Optional[pd.DataFrame] = None,
    policy_col: Optional[str] = None,
) -> Tuple[str, str]:
    """
    Write a unique list of policies successfully updated via POST recreate.

    Creates:
      - updated_policies_<timestamp>.csv  (Policy No [, Policy UUID])
      - updated_policies_<timestamp>.txt  (one Policy No per line)
    Returns (csv_path, txt_path).
    """
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = os.path.join(output_dir, f"updated_policies_{ts}.csv")
    txt_path = os.path.join(output_dir, f"updated_policies_{ts}.txt")

    uuid_by_policy: Dict[str, str] = {}
    if df_updated is not None and not df_updated.empty and policy_col and policy_col in df_updated.columns:
        uuid_col = next(
            (c for c in df_updated.columns if str(c).strip().lower() in (
                "obtained_policy_uuid", "policy_uuid", "policy uuid", "policyuuid"
            )),
            None,
        )
        if uuid_col:
            for _, row in df_updated.drop_duplicates(subset=[policy_col]).iterrows():
                pol = str(row[policy_col]).strip()
                uuid_val = str(row.get(uuid_col) or "").strip()
                if pol and uuid_val and uuid_val.lower() not in ("nan", "none"):
                    uuid_by_policy[pol] = uuid_val

    sorted_pols = sorted({str(p).strip() for p in successful_policies if str(p).strip()})
    rows = [{"Policy No": p, "Policy UUID": uuid_by_policy.get(p, "")} for p in sorted_pols]
    pd.DataFrame(rows).to_csv(csv_path, index=False)
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(sorted_pols) + ("\n" if sorted_pols else ""))
    return csv_path, txt_path


def write_writing_agent_update_sql(
    corrections: List[dict],
    agency_id: int,
    out_path: str,
) -> str:
    """
    Write UPDATE SQL so policies.writing_agent_user_id matches the file's
    highest-order agent (e.g. Michael Dunavant instead of Jacob Yale).
    Script does NOT execute these updates.
    """
    lines = [
        "-- Writing agent corrections from commission upload file",
        f"-- agency_id = {agency_id}",
        "-- Source of truth: highest-order person on the sheet (Sub Agent / Seq 1).",
        "-- Review with senior eng before running. This script does NOT execute UPDATEs.",
        "",
    ]
    for c in corrections or []:
        pol = str(c.get("policyNumber") or "")
        pol_uuid = str(c.get("policyUuid") or "").strip()
        new_uuid = str(c.get("fileWritingAgentUuid") or "").strip()
        new_name = str(c.get("fileWritingAgentName") or "")
        old_name = str(c.get("dbWritingAgentName") or "")
        old_uid = str(c.get("dbWritingAgentUserId") or "")
        code = str(c.get("fileWritingAgentCode") or "")
        lines.append(f"-- Policy {pol}: '{old_name}' -> '{new_name}' (code {code})")
        if old_uid:
            lines.append(f"-- previous writing_agent_user_id = {old_uid}")
        if not pol_uuid or not new_uuid:
            lines.append("-- SKIPPED: missing policyUuid or fileWritingAgentUuid")
            lines.append("")
            continue
        lines.append("UPDATE policies")
        lines.append(f"SET writing_agent_user_id = '{new_uuid}'")
        lines.append(f"WHERE uuid = '{pol_uuid}'")
        lines.append(f"  AND agency_id = {int(agency_id)};")
        lines.append("")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return os.path.abspath(out_path)


# ---------------------------------------------------------------------------
# STAGE 1.5: DB Commission Snapshot Export
# ---------------------------------------------------------------------------
def fetch_db_commission_snapshot(
    df_stage1_valid: pd.DataFrame,
    db_mgr: "DatabaseManager",
    agency_id: int,
    users_map: Dict[str, str],
) -> pd.DataFrame:
    """
    Stage 1.5: Fetches the current policy_commission rows from the DB for every policy
    that passed Stage 1 validation, and reconstructs them into the same column format
    as the original uploaded file (one row per commission split entry).

    Returns a DataFrame with rows ordered: policy group → split order.
    """
    if df_stage1_valid.empty:
        return pd.DataFrame()

    # Detect input file column names (mirrors stage2 column detection)
    pol_col   = next((c for c in df_stage1_valid.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
    prod_col  = next((c for c in df_stage1_valid.columns if str(c).strip().lower() in ("normalized_product_name", "product name", "product", "product_name", "product_description")), None)
    fm_col    = next((c for c in df_stage1_valid.columns if str(c).strip().lower() in ("from month", "from_month", "month", "month from", "month_from")), None)
    tm_col    = next((c for c in df_stage1_valid.columns if str(c).strip().lower() in ("to month", "to_month", "month to", "month_to")), None)
    seq_col   = next((c for c in df_stage1_valid.columns if str(c).strip().lower() in ("sequence", "seq")), None)
    agent_col = next((c for c in df_stage1_valid.columns if str(c).strip().lower() in ("agent id", "agent_id", "agentid")), None)
    agent_name_col = next((c for c in df_stage1_valid.columns if "agent name" in str(c).strip().lower() or str(c).strip().lower() == "agent_name"), None)
    rate_col  = next((c for c in df_stage1_valid.columns if str(c).strip().lower() in ("rate", "rate %", "commission %")), None)
    amt_col   = next((c for c in df_stage1_valid.columns if str(c).strip().lower() in ("amount", "fee amount", "fee_amount")), None)

    # Resolve canonical output column names (use the actual found column names where available)
    out_pol_col        = pol_col         or "Policy No"
    out_prod_col       = prod_col        or "Product"
    out_fm_col         = fm_col          or "From Month"
    out_tm_col         = tm_col          or "To Month"
    out_seq_col        = seq_col         or "Sequence"
    out_agent_name_col = agent_name_col  or "Agent Name"
    out_agent_id_col   = agent_col       or "Agent ID"
    out_rate_col       = rate_col        or "Rate"
    out_amt_col        = amt_col         or "Amount"

    # Extract unique policy numbers for bulk pre-fetch (same as Stage 2)
    unique_policy_numbers = []
    if pol_col:
        for v in df_stage1_valid[pol_col].dropna().unique():
            p = str(v).strip()
            if p and p.lower() not in ("nan", "none", ""):
                unique_policy_numbers.append(p)

    print(f"   ⚡ Bulk fetching DB commission rows for {len(unique_policy_numbers):,} unique policies...")
    db_map = db_mgr.fetch_all_agency_commission_rules(agency_id, policy_numbers=unique_policy_numbers)
    print(f"   ✅ Bulk loaded {len(db_map):,} policy/product indexes.")

    output_rows = []

    # Iterate over each unique (policy, product, from_month, to_month) tier group from the valid file
    group_cols = [c for c in [pol_col, prod_col, fm_col, tm_col] if c]
    if group_cols:
        grouped = df_stage1_valid.groupby(group_cols, sort=False)
    else:
        grouped = [(None, df_stage1_valid)]

    seen_policy_product_pairs: Set[Tuple[str, str]] = set()

    for _, group in grouped:
        first_row = group.iloc[0]

        pol_number  = str(first_row.get(pol_col, "")).strip()   if pol_col  else ""
        prod_name   = str(first_row.get(prod_col, "")).strip()  if prod_col else ""
        clean_prod  = " ".join(prod_name.lower().split())

        raw_fm = first_row.get(fm_col, 1) if fm_col else 1
        raw_tm = first_row.get(tm_col, 0) if tm_col else 0
        try:
            file_fm = int(float(str(raw_fm).strip().replace("'", "").replace('"', "")))
        except Exception:
            file_fm = 1
        try:
            file_tm = int(float(str(raw_tm).strip().replace("'", "").replace('"', "")))
        except Exception:
            file_tm = 0

        if not pol_number or pol_number.lower() in ("nan", "none", ""):
            continue

        # Look up DB records (composite key first, then policy-number fallback)
        db_records = db_map.get((pol_number, clean_prod), [])
        if not db_records:
            pol_recs = db_map.get(pol_number, [])
            db_records = [
                r for r in pol_recs
                if clean_prod in " ".join(str(r.get("product_name", "")).lower().split())
                or " ".join(str(r.get("product_name", "")).lower().split()) in clean_prod
            ]
            if not db_records and pol_recs:
                db_records = pol_recs

        if not db_records:
            # Policy not found in DB at all — emit a single placeholder row
            output_rows.append({
                out_pol_col:        pol_number,
                out_prod_col:       prod_name,
                out_fm_col:         file_fm,
                out_tm_col:         file_tm if file_tm else "",
                out_seq_col:        "",
                out_agent_name_col: "",
                out_agent_id_col:   "",
                out_rate_col:       "",
                out_amt_col:        "",
                "Carrier Commission": "",
                "Payout Method":     "",
                "Rule Name":         "",
                "DB Commission UUID": "",
                "DB_Status":         "NOT_FOUND_IN_DB",
            })
            continue

        # Filter DB commission records to the tier (month range) matching the file group
        tier_records = [
            r for r in db_records
            if r.get("commission_uuid") is not None
            and (r.get("month_from") is None or int(r.get("month_from") or 1) == file_fm)
        ]

        if not tier_records:
            # Policy exists but has no commission rules at this tier
            output_rows.append({
                out_pol_col:        pol_number,
                out_prod_col:       prod_name,
                out_fm_col:         file_fm,
                out_tm_col:         file_tm if file_tm else "",
                out_seq_col:        "",
                out_agent_name_col: "",
                out_agent_id_col:   "",
                out_rate_col:       "",
                out_amt_col:        "",
                "Carrier Commission": "",
                "Payout Method":     "",
                "Rule Name":         "",
                "DB Commission UUID": "",
                "DB_Status":         "NO_COMMISSION_RULES_IN_DB",
            })
            continue

        # Emit one row per commission rule × per split entry
        for db_rec in tier_records:
            db_commission_uuid = str(db_rec.get("commission_uuid", "") or "")
            rule_name          = str(db_rec.get("rule_name", "") or "")
            db_payout_method   = str(db_rec.get("payout_method", "") or "").upper()
            db_fee_amount      = db_rec.get("fee_amount")   # numeric or None
            db_month_from      = db_rec.get("month_from")   # numeric or None
            db_month_to        = db_rec.get("month_to")     # numeric or None
            db_prem_amount     = db_rec.get("premium_amount")

            # Decode commission_split JSONB (may already be a list/dict or a JSON string)
            raw_split = db_rec.get("commission_split")
            split_entries: List[dict] = []
            if raw_split is not None:
                if isinstance(raw_split, list):
                    split_entries = raw_split
                elif isinstance(raw_split, str) and raw_split.strip():
                    try:
                        parsed = json.loads(raw_split)
                        split_entries = parsed if isinstance(parsed, list) else []
                    except (json.JSONDecodeError, TypeError):
                        split_entries = []

            # Determine carrier commission display value
            if "FIXED" in db_payout_method:
                carrier_comm_display = format_amount(float(db_fee_amount)) if db_fee_amount is not None else ""
            else:
                # For PERCENTAGE, carrier commission = the highest-order (agency) split value
                agency_entry = next((s for s in split_entries if str(s.get("role", "")).lower() in ("agency", "") or s.get("order") == 1), None)
                if agency_entry:
                    carrier_comm_display = format_rate(float(agency_entry.get("commission_value") or 0.0))
                elif db_fee_amount is not None:
                    carrier_comm_display = format_rate(float(db_fee_amount))
                else:
                    carrier_comm_display = ""

            from_m_out = int(db_month_from) if db_month_from is not None else file_fm
            to_m_out   = int(db_month_to)   if db_month_to   is not None else (file_tm if file_tm else "")

            if not split_entries:
                # No split data — emit one row with just the commission rule meta
                rate_out = format_rate(float(db_fee_amount)) if "PERCENT" in db_payout_method and db_fee_amount is not None else ""
                amt_out  = format_amount(float(db_fee_amount)) if "FIXED" in db_payout_method and db_fee_amount is not None else ""
                output_rows.append({
                    out_pol_col:        pol_number,
                    out_prod_col:       prod_name,
                    out_fm_col:         from_m_out,
                    out_tm_col:         to_m_out,
                    out_seq_col:        "",
                    out_agent_name_col: "",
                    out_agent_id_col:   "",
                    out_rate_col:       rate_out,
                    out_amt_col:        amt_out,
                    "Carrier Commission": carrier_comm_display,
                    "Payout Method":     db_payout_method,
                    "Rule Name":         rule_name,
                    "DB Commission UUID": db_commission_uuid,
                    "DB_Status":         "DB_VALUE",
                })
            else:
                # Sort split entries by order (ascending) for consistent row order
                split_entries_sorted = sorted(split_entries, key=lambda s: int(s.get("order") or 0))
                for split in split_entries_sorted:
                    split_order     = int(split.get("order") or 0)
                    split_user_name = str(split.get("user_name", "") or "")
                    split_user_code = str(split.get("user_code", "") or "")
                    split_comm_val  = split.get("commission_value")

                    # Map split value to Rate or Amount depending on payout method
                    if "FIXED" in db_payout_method:
                        rate_out = ""
                        amt_out  = format_amount(float(split_comm_val)) if split_comm_val is not None else ""
                    else:
                        rate_out = format_rate(float(split_comm_val)) if split_comm_val is not None else ""
                        amt_out  = ""

                    output_rows.append({
                        out_pol_col:        pol_number,
                        out_prod_col:       prod_name,
                        out_fm_col:         from_m_out,
                        out_tm_col:         to_m_out,
                        out_seq_col:        split_order,
                        out_agent_name_col: split_user_name,
                        out_agent_id_col:   split_user_code,
                        out_rate_col:       rate_out,
                        out_amt_col:        amt_out,
                        "Carrier Commission": carrier_comm_display,
                        "Payout Method":     db_payout_method,
                        "Rule Name":         rule_name,
                        "DB Commission UUID": db_commission_uuid,
                        "DB_Status":         "DB_VALUE",
                    })

    return pd.DataFrame(output_rows)


def export_db_commission_snapshot(
    df_snapshot: pd.DataFrame,
    versioned_dir: str,
    df_stage1_valid: pd.DataFrame = None,
    filename_prefix: str = "db_commission_snapshot",
) -> str:
    """
    Exports the DB commission snapshot DataFrame to an Excel file.
    Column order mirrors the original uploaded file columns, with DB-only
    informational columns appended at the end.
    Returns the output file path.
    """
    os.makedirs(versioned_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    snapshot_path = os.path.join(versioned_dir, f"{filename_prefix}_{timestamp}.xlsx")

    if df_snapshot.empty:
        df_snapshot.to_excel(snapshot_path, index=False)
        return snapshot_path

    # Build preferred column order: mirror file columns first, then DB-only extras
    db_extra_cols  = ["Carrier Commission", "Payout Method", "Rule Name", "DB Commission UUID", "DB_Status"]
    file_like_cols = [c for c in df_snapshot.columns if c not in db_extra_cols]

    # If we have the original file columns, use them to anchor ordering
    if df_stage1_valid is not None and not df_stage1_valid.empty:
        original_cols = [c for c in df_stage1_valid.columns if not str(c).startswith("_") and c not in db_extra_cols]
        # Keep any snapshot columns that exist in the file order first
        ordered = [c for c in original_cols if c in df_snapshot.columns]
        # Then append any remaining file-like cols not in the original
        ordered += [c for c in file_like_cols if c not in ordered]
    else:
        ordered = file_like_cols

    final_cols = ordered + [c for c in db_extra_cols if c in df_snapshot.columns]
    df_out     = df_snapshot[final_cols].copy()

    # Drop internal helper columns
    drop_cols = [c for c in df_out.columns if str(c).startswith("_")]
    df_out = df_out.drop(columns=drop_cols, errors="ignore")

    df_out.to_excel(snapshot_path, index=False)
    return snapshot_path


# ---------------------------------------------------------------------------
# Output Export & Main CLI Pipeline
# ---------------------------------------------------------------------------
def export_phase1_reports(df_valid: pd.DataFrame, df_invalid: pd.DataFrame, versioned_dir: str) -> Tuple[str, str]:
    """Exports Phase 1 First Validation Reports."""
    os.makedirs(versioned_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    clean_path = os.path.join(versioned_dir, f"file_validated_clean_{timestamp}.xlsx")
    errors_path = os.path.join(versioned_dir, f"file_validated_errors_{timestamp}.xlsx")

    drop_cols = [c for c in df_valid.columns if str(c).startswith("_")]
    df_valid_clean = df_valid.drop(columns=drop_cols, errors="ignore")
    df_invalid_clean = df_invalid.drop(columns=drop_cols, errors="ignore")

    df_valid_clean.to_excel(clean_path, index=False)
    df_invalid_clean.to_excel(errors_path, index=False)

    return clean_path, errors_path


def export_phase2_reports(df_matched: pd.DataFrame, df_unmatched: pd.DataFrame, versioned_dir: str) -> Tuple[str, str]:
    """Exports Phase 2 DB Validated Reports."""
    os.makedirs(versioned_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    matched_path = os.path.join(versioned_dir, f"db_matched_commission_{timestamp}.xlsx")
    unmatched_path = os.path.join(versioned_dir, f"db_unmatched_commission_{timestamp}.xlsx")

    df_matched.to_excel(matched_path, index=False)
    df_unmatched.to_excel(unmatched_path, index=False)

    return matched_path, unmatched_path


def _prepare_policy_commission_api(logger_name: str):
    """Shared SSO + base URL setup for commission DELETE/POST helpers."""
    import logging

    try:
        from dotenv import load_dotenv
    except ImportError:
        def load_dotenv(*args, **kwargs):
            pass

    other_script_dir = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "..", "policy_commission_import_validation_script")
    )
    other_env_path = os.path.join(other_script_dir, ".env")
    if os.path.exists(other_env_path):
        load_dotenv(other_env_path, override=False)
    if other_script_dir not in sys.path:
        sys.path.append(other_script_dir)

    try:
        from import_policy_commission import fetch_sso_token, confirm_payload
    except ImportError as ie:
        print(f"⚠️ Could not import API logic: {ie}")
        return None

    logger = logging.getLogger(logger_name)
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        logger.addHandler(logging.StreamHandler())

    api_base = os.environ.get("API_BASE_URL", "https://prod.api.pieq.ai").rstrip("/")
    sso_url_raw = os.environ.get("AUTH_URL", "https://prod.auth.pieq.ai")
    sso_url = f"{sso_url_raw.rstrip('/')}/realms/pieq-sso/protocol/openid-connect/token"
    client_id = os.environ.get("KEYCLOAK_CLIENT_ID") or os.environ.get("client_id") or "pieq-app"
    username, password = resolve_sso_credentials()

    token = fetch_sso_token(logger, sso_url, client_id, username, password)
    if not token:
        print("⚠️ Could not generate API token. Make sure .env credentials are set.")
        return None

    def _headers(tok: str) -> dict:
        h = {"Content-Type": "application/json", "Authorization": f"Bearer {tok}"}
        if "X_CLIENT_ID" in os.environ:
            h["x-client-id"] = os.environ["X_CLIENT_ID"]
        if "X_SERVICE_NAME" in os.environ:
            h["x-service-name"] = os.environ["X_SERVICE_NAME"]
        return h

    return {
        "logger": logger,
        "api_base": api_base,
        "token": token,
        "sso_url": sso_url,
        "client_id": client_id,
        "username": username,
        "password": password,
        "headers": _headers,
        "fetch_sso_token": fetch_sso_token,
        "confirm_payload": confirm_payload,
    }


def run_delete_commission_uuids(
    commission_uuids: list,
    output_dir: str,
) -> Tuple[Set[str], Set[str], Dict[str, str]]:
    """
    DELETE mismatched policy commission rows via
    DELETE /api/v1/policy/commission/{uuid}.

    HTTP 404 is treated as success (already deleted).
    Returns (ok_uuids, failed_uuids, fail_reasons_by_uuid).
    """
    import urllib.error
    import urllib.request

    ok: Set[str] = set()
    failed: Set[str] = set()
    fail_reasons: Dict[str, str] = {}
    uuids = [str(u).strip() for u in (commission_uuids or []) if str(u).strip()]
    if not uuids:
        return ok, failed, fail_reasons

    api = _prepare_policy_commission_api("AutoDeleteMismatchCommission")
    if not api:
        for uid in uuids:
            failed.add(uid)
            fail_reasons[uid] = "Could not authenticate / prepare API client"
        return ok, failed, fail_reasons

    logger = api["logger"]
    token = api["token"]
    fail_log = os.path.join(
        output_dir, f"delete_api_failures_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )
    print(f"\n🗑️  API DELETE {len(uuids)} mismatch rule(s) via /api/v1/policy/commission/{{uuid}} …")

    for i, uid in enumerate(uuids, start=1):
        url = f"{api['api_base']}/api/v1/policy/commission/{uid}"

        def _delete_once(tok: str) -> Tuple[bool, Optional[str], str]:
            req = urllib.request.Request(url, headers=api["headers"](tok), method="DELETE")
            try:
                with urllib.request.urlopen(req, timeout=30.0) as resp:
                    body = resp.read().decode("utf-8", errors="ignore")
                    if i == 1 or i == len(uuids) or i % 25 == 0:
                        logger.info(f"✅ [{i}/{len(uuids)}] DELETE {uid} → HTTP {resp.status} {body[:120]}")
                    return True, tok, ""
            except urllib.error.HTTPError as e:
                err_body = e.read().decode("utf-8", errors="ignore") if hasattr(e, "read") else str(e)
                if e.code == 404:
                    if i == 1 or i == len(uuids) or i % 25 == 0:
                        logger.info(f"✅ [{i}/{len(uuids)}] DELETE {uid} → already gone (HTTP 404)")
                    return True, tok, ""
                if e.code == 401:
                    new_tok = api["fetch_sso_token"](
                        logger, api["sso_url"], api["client_id"], api["username"], api["password"]
                    )
                    if new_tok:
                        return _delete_once(new_tok)
                logger.error(f"❌ [{i}/{len(uuids)}] DELETE {uid} failed (HTTP {e.code}): {err_body}")
                return False, tok, f"HTTP {e.code}: {err_body}"
            except Exception as err:
                logger.error(f"❌ [{i}/{len(uuids)}] DELETE {uid} failed: {err}")
                return False, tok, str(err)

        success, token, err = _delete_once(token)
        if success:
            ok.add(uid)
            if i == 1 or i == len(uuids) or i % 25 == 0:
                print(f"   DELETE progress {i}/{len(uuids)}", flush=True)
        else:
            failed.add(uid)
            fail_reasons[uid] = err or "DELETE failed"
            print(f"   ❌ DELETE failed [{i}/{len(uuids)}] {uid}: {err}", flush=True)
            with open(fail_log, "a", encoding="utf-8") as f:
                f.write(f"[{datetime.datetime.now().isoformat()}] DELETE {uid}: {err}\n")

    print(f"✅ API DELETE finished: {len(ok)}/{len(uuids)} succeeded.")
    if failed:
        print(f"⚠️ {len(failed)} DELETE(s) failed. Check logs at: {fail_log}")
    return ok, failed, fail_reasons


def run_post_commission_payloads(
    payloads: list,
    output_dir: str,
    auto_confirm: bool = True,
) -> Tuple[Set[str], List[dict]]:
    """
    POST-create policy commission rules (after delete). Never PUT.
    Strips rule_uuid / helper fields.

    Returns (successful_policy_numbers, per_payload_results) where each result has:
      policyNumber, monthFrom, monthTo, status (Updated/Not Updated), reason
    """
    import urllib.error
    import urllib.request

    successful: Set[str] = set()
    results: List[dict] = []
    if not payloads:
        return successful, results

    api = _prepare_policy_commission_api("AutoPostAfterDelete")
    if not api:
        for payload in payloads:
            results.append({
                "policyNumber": str(payload.get("policyNumber") or "Unknown"),
                "monthFrom": payload.get("monthFrom"),
                "monthTo": payload.get("monthTo"),
                "status": "Not Updated",
                "reason": "Could not authenticate / prepare API client",
            })
        return successful, results

    logger = api["logger"]
    token = api["token"]
    confirm_payload = api["confirm_payload"]
    fetch_sso_token = api["fetch_sso_token"]

    post_url = f"{api['api_base']}/api/v1/policy/commission"
    fail_log = os.path.join(
        output_dir, f"post_api_failures_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )
    confirm = auto_confirm
    _headers = api["headers"]
    sso_url = api["sso_url"]
    client_id = api["client_id"]
    username = api["username"]
    password = api["password"]

    for idx, payload in enumerate(payloads):
        payload_to_send = dict(payload)
        diff_text = payload_to_send.pop("_diff_text", "")
        payload_to_send.pop("_issue_reason", None)
        pol_num = str(payload.get("policyNumber") or payload_to_send.get("policyNumber") or "Unknown")
        month_from = payload.get("monthFrom")
        month_to = payload.get("monthTo")
        action = confirm_payload(
            logger, idx + 1, len(payloads), payload_to_send, confirm,
            action_name="POST", diff_text=diff_text,
        )
        if action == "quit":
            results.append({
                "policyNumber": pol_num,
                "monthFrom": month_from,
                "monthTo": month_to,
                "status": "Not Updated",
                "reason": "POST quit by user before this payload",
            })
            for rest in payloads[idx + 1:]:
                results.append({
                    "policyNumber": str(rest.get("policyNumber") or "Unknown"),
                    "monthFrom": rest.get("monthFrom"),
                    "monthTo": rest.get("monthTo"),
                    "status": "Not Updated",
                    "reason": "POST quit by user before this payload",
                })
            break
        if action == "all":
            confirm = True
            action = "yes"
        if action == "skip":
            results.append({
                "policyNumber": pol_num,
                "monthFrom": month_from,
                "monthTo": month_to,
                "status": "Not Updated",
                "reason": "POST skipped by user",
            })
            continue
        if action != "yes":
            results.append({
                "policyNumber": pol_num,
                "monthFrom": month_from,
                "monthTo": month_to,
                "status": "Not Updated",
                "reason": f"POST not confirmed (action={action})",
            })
            continue

        payload_to_send.pop("policyNumber", None)
        payload_to_send.pop("writingAgentName", None)
        payload_to_send.pop("rule_uuid", None)
        body = {k: v for k, v in payload_to_send.items() if not str(k).startswith("_")}

        # Ensure camelCase + snake_case compatibility for each split item so both API and UI bind correctly
        if "commissionSplit" in body and isinstance(body["commissionSplit"], list):
            enriched_splits = []
            for item in body["commissionSplit"]:
                if isinstance(item, dict):
                    clean_item = dict(item)
                    u_uuid = clean_item.get("user_uuid") or clean_item.get("userUuid")
                    if u_uuid:
                        clean_item["user_uuid"] = u_uuid
                        clean_item["userUuid"] = u_uuid
                    clean_item.pop("userId", None)
                    u_name = clean_item.get("user_name") or clean_item.get("userName")
                    if u_name:
                        clean_item["user_name"] = u_name
                        clean_item["userName"] = u_name
                    u_code = clean_item.get("user_code") or clean_item.get("userCode")
                    if u_code:
                        clean_item["user_code"] = u_code
                        clean_item["userCode"] = u_code
                    c_val = clean_item.get("commission_value") if clean_item.get("commission_value") is not None else clean_item.get("commissionValue")
                    if c_val is not None:
                        clean_item["commission_value"] = float(c_val)
                        clean_item["commissionValue"] = float(c_val)
                    enriched_splits.append(clean_item)
                else:
                    enriched_splits.append(item)
            body["commissionSplit"] = enriched_splits

        def _post_once(tok: str) -> Tuple[bool, Optional[str], str]:
            req = urllib.request.Request(
                post_url,
                data=json.dumps([body]).encode("utf-8"),
                headers=_headers(tok),
                method="POST",
            )
            try:
                with urllib.request.urlopen(req, timeout=30.0) as resp:
                    res_json = json.loads(resp.read().decode("utf-8"))
                    logger.info(f"✅ Success POSTing {pol_num}: {res_json}")
                    return True, tok, ""
            except urllib.error.HTTPError as e:
                err_body = e.read().decode("utf-8", errors="ignore") if hasattr(e, "read") else str(e)
                if e.code == 401:
                    new_tok = fetch_sso_token(logger, sso_url, client_id, username, password)
                    if new_tok:
                        return _post_once(new_tok)
                logger.error(f"❌ Failed to POST {pol_num} (HTTP {e.code}): {err_body}")
                with open(fail_log, "a", encoding="utf-8") as f:
                    f.write(f"[{datetime.datetime.now().isoformat()}] POST HTTP {e.code} for {pol_num}: {err_body}\n")
                return False, tok, f"POST HTTP {e.code}: {err_body}"
            except Exception as err:
                logger.error(f"❌ Failed to POST {pol_num}: {err}")
                with open(fail_log, "a", encoding="utf-8") as f:
                    f.write(f"[{datetime.datetime.now().isoformat()}] POST Exception for {pol_num}: {err}\n")
                return False, tok, f"POST exception: {err}"

        ok, token, err = _post_once(token)
        if ok:
            successful.add(pol_num)
            results.append({
                "policyNumber": pol_num,
                "monthFrom": month_from,
                "monthTo": month_to,
                "status": "Updated",
                "reason": "Successfully updated via API POST",
            })
        else:
            results.append({
                "policyNumber": pol_num,
                "monthFrom": month_from,
                "monthTo": month_to,
                "status": "Not Updated",
                "reason": err or "POST failed",
            })

    if os.path.exists(fail_log):
        print(f"\n⚠️ Some POST calls failed. Check logs at: {fail_log}")
    print(f"✅ POST recreate finished: {len(successful)}/{len(payloads)} succeeded.")
    return successful, results


def write_policy_update_status_excel(
    output_dir: str,
    post_payloads: Optional[list] = None,
    post_results: Optional[List[dict]] = None,
    mismatch_payloads: Optional[list] = None,
    deleted_ok: Optional[Set[str]] = None,
    deleted_failed: Optional[Set[str]] = None,
    delete_fail_reasons: Optional[Dict[str, str]] = None,
    post_skipped_reason: Optional[str] = None,
) -> str:
    """
    Write Excel: Policy No | Status (Updated/Not Updated) | Reason
    after API DELETE + POST. One row per policy.
    """
    deleted_ok = deleted_ok or set()
    deleted_failed = deleted_failed or set()
    delete_fail_reasons = delete_fail_reasons or {}
    post_results = post_results or []
    post_payloads = post_payloads or []
    mismatch_payloads = mismatch_payloads or []

    uuid_to_policy: Dict[str, str] = {}
    for p in mismatch_payloads:
        uid = str((p or {}).get("rule_uuid") or "").strip()
        pol = str((p or {}).get("policyNumber") or "").strip()
        if uid and pol:
            uuid_to_policy[uid] = pol

    by_policy: Dict[str, dict] = {}

    def _ensure(pol: str):
        key = str(pol or "").strip() or "Unknown"
        if key not in by_policy:
            by_policy[key] = {
                "Policy No": key,
                "Status": "Not Updated",
                "Reason": "",
                "_ok": 0,
                "_fail": 0,
                "_reasons": [],
            }
        return by_policy[key]

    for uid in deleted_failed:
        pol = uuid_to_policy.get(uid, uid)
        row = _ensure(pol)
        row["_fail"] += 1
        row["_reasons"].append(
            f"API DELETE failed for {uid}: {delete_fail_reasons.get(uid, 'unknown error')}"
        )

    for uid in deleted_ok:
        pol = uuid_to_policy.get(uid)
        if pol:
            _ensure(pol)

    if post_skipped_reason:
        for p in post_payloads:
            pol = str((p or {}).get("policyNumber") or "Unknown")
            row = _ensure(pol)
            row["_fail"] += 1
            row["_reasons"].append(post_skipped_reason)
    else:
        for r in post_results:
            pol = str(r.get("policyNumber") or "Unknown")
            row = _ensure(pol)
            if r.get("status") == "Updated":
                row["_ok"] += 1
            else:
                row["_fail"] += 1
                reason = str(r.get("reason") or "Not updated").strip()
                mf, mt = r.get("monthFrom"), r.get("monthTo")
                if mf is not None or mt is not None:
                    reason = f"[months {mf}->{mt}] {reason}"
                row["_reasons"].append(reason)

        # Policies intended for POST but missing from results (shouldn't happen often)
        seen_in_results = {str(r.get("policyNumber") or "") for r in post_results}
        for p in post_payloads:
            pol = str((p or {}).get("policyNumber") or "Unknown")
            if pol not in seen_in_results:
                row = _ensure(pol)
                row["_fail"] += 1
                row["_reasons"].append("No POST result recorded")

    rows = []
    for pol in sorted(by_policy.keys(), key=lambda x: str(x)):
        info = by_policy[pol]
        if info["_ok"] > 0 and info["_fail"] == 0:
            status = "Updated"
            reason = "Successfully updated via API"
        elif info["_ok"] > 0 and info["_fail"] > 0:
            status = "Not Updated"
            reason = (
                f"Partial: {info['_ok']} slab(s) updated, {info['_fail']} failed. "
                + "; ".join(info["_reasons"][:5])
            )
        else:
            status = "Not Updated"
            reason = "; ".join(info["_reasons"][:8]) if info["_reasons"] else "Not updated"
        rows.append({"Policy No": info["Policy No"], "Status": status, "Reason": reason})

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(output_dir, f"policy_update_status_{ts}.xlsx")
    os.makedirs(output_dir, exist_ok=True)
    pd.DataFrame(rows, columns=["Policy No", "Status", "Reason"]).to_excel(out_path, index=False)
    return out_path


# ---------------------------------------------------------------------------
# Readable validation summary (validation_summary_<run>.md)
# ---------------------------------------------------------------------------
MD_MAX_ITEMS_PER_SECTION = 300

STAGE1_STATUS_TEXT = {
    "INVALID_POLICY_ROW_COUNT": "Each month band needs 3 or 4 people (Agency, Sales Leader, Agent, and an optional Sub Agent). This band has a different number.",
    "INVALID_OVERRIDE_LEVELS": "An override needs an Agency, a Sales Leader and an Agent. One of them is missing and could not be taken from the policy's regular rule.",
    "NEGATIVE_SPLIT_SHARE": "One person's share works out below 0%, so a higher level is paid less than the level under it.",
    "INVALID_SPLIT_SUM": "The shares in this band do not add up to 100%.",
    "MISSING_POLICY_NUMBER": "The row has no policy number.",
    "UNMATCHED_MONTH_RANGE": "The months do not fit any tier in the rate sheet.",
    "UNMATCHED_PRODUCT_IN_RATE_SHEET": "The product is not in the rate sheet.",
    "VALUE_MISMATCH": "The commission in the file does not match the rate sheet.",
}


def _md_months(fm, tm) -> str:
    def _num(v):
        s = str(v if v is not None else "").strip()
        if s.endswith(".0"):
            s = s[:-2]
        return "" if s.lower() in ("", "nan", "none") else s
    a, b = _num(fm), _num(tm)
    if not a:
        return "Months: not given"
    if not b or b == "0":
        return f"Months {a} onwards"
    return f"Months {a}–{b}"


def _md_escape(text) -> str:
    return str(text if text is not None else "").replace("|", "/").replace("\n", " ").strip()


def explain_reason(reason: str) -> str:
    """Turn a Stage 2 reason into one plain sentence."""
    r = str(reason or "").strip()
    if not r or r.lower() == "nan":
        return ""
    m = re.search(r"Agent '(.+?)' found in DB, but DB commission_value is ([\d.]+|None)%? \(File has ([\d.]+)%\)", r)
    if m:
        name = " ".join(m.group(1).split()).title()
        return f"{name}: the file gives {m.group(3)}%, the system has {m.group(2)}%."
    m = re.search(r"Agent '(.+?)' found in DB, but DB fee_amount is ([\d.]+) \(File has ([\d.]+)\)", r)
    if m:
        return f"The carrier commission is ${m.group(3)} in the file but ${m.group(2)} in the system."
    m = re.search(r"No commission rule record found in policy_commission table for agent '(.+?)'", r)
    if m:
        name = " ".join(m.group(1).split()).title()
        return f"{name} is in the file but is not on the system's rule for these months."
    m = re.search(r"No commission rule found in DB for period (\S+) -> (\S+)\.", r)
    if m:
        return "The policy has commission rules, but none for these months. A new rule will be added from the file."
    if "No commission rules in policy_commission table" in r:
        return "The policy is in the system but has no commission rules yet. The file's rules will be added."
    if "not found in policies table" in r:
        return "This policy number is not in the system (also tried with the trailing digits removed)."
    m = re.search(r"Policy number '(.+?)' matched (\d+) policies after removing the suffix \('(.+?)'\); (.+?)\. Not updated\.", r)
    if m:
        return (
            f"The number is not in the system as written. With the last digit(s) removed ({m.group(3)}) "
            f"it matches {m.group(2)} different policies, and the product could not tell them apart: "
            f"{m.group(4)}. Nothing was changed."
        )
    m = re.search(r"Policy number '(.+?)\.\.\.' matched (\d+) policies in DB; (.+?)\. Not updated\.", r)
    if m:
        return (
            f"The first 11 characters ({m.group(1)}) match {m.group(2)} different policies, and the product "
            f"could not tell them apart: {m.group(3)}. Nothing was changed."
        )
    if "matched multiple policies" in r or ("matched" in r and "policies in DB" in r):
        return f"The policy number matches more than one policy in the system, so nothing was changed. ({r})"
    return r


def _md_split_table(db_val, file_val) -> Tuple[List[str], bool]:
    """Markdown table rows comparing file vs DB people; returns (lines, any_difference)."""
    try:
        db_list = json.loads(db_val) if isinstance(db_val, str) and db_val.strip().startswith("[") else db_val
        file_list = json.loads(file_val) if isinstance(file_val, str) and file_val.strip().startswith("[") else file_val
    except (json.JSONDecodeError, TypeError):
        return [], False
    if not isinstance(db_list, list) or not isinstance(file_list, list):
        return [], False
    db_by = {it.get("order"): it for it in db_list if isinstance(it, dict)}
    fi_by = {it.get("order"): it for it in file_list if isinstance(it, dict)}
    lines = [
        "| Role | In the file | In the system | File % | System % | Same? |",
        "|------|-------------|---------------|--------|----------|-------|",
    ]
    any_diff = False
    for o in sorted(set(db_by) | set(fi_by), key=lambda x: (x is None, x)):
        d, f = db_by.get(o, {}), fi_by.get(o, {})
        role = f.get("role") or d.get("role") or ""
        d_cv, f_cv = d.get("commission_value"), f.get("commission_value")
        d_name = " ".join(str(d.get("user_name") or "").split())
        f_name = " ".join(str(f.get("user_name") or "").split())
        if "sub agent" in role.lower() and not f_name and not d_name:
            continue  # empty Sub Agent slot on both sides
        try:
            same_val = d_cv is not None and f_cv is not None and abs(float(d_cv) - float(f_cv)) <= 0.05
        except (TypeError, ValueError):
            same_val = False
        same_person = (
            (not d_name and not f_name)
            or are_names_matching(f_name.lower(), d_name.lower())
            or are_codes_matching(str(f.get("user_code") or ""), str(d.get("user_code") or ""))
            or ("main line benefits" in f_name.lower() and "mlb parallel run" in d_name.lower())
        )
        ok = same_val and same_person
        any_diff = any_diff or not ok
        fv = f"{f_cv}%" if f_cv is not None else "—"
        dv = f"{d_cv}%" if d_cv is not None else "—"
        mark = "✅" if ok else "❌"
        if not ok:
            fv, dv = f"**{fv}**", f"**{dv}**"
        lines.append(
            f"| {_md_escape(role)} | {_md_escape(f_name.title()) or '—'} | {_md_escape(d_name.title()) or '—'} "
            f"| {fv} | {dv} | {mark} |"
        )
    return lines, any_diff


def _md_band_groups(df: pd.DataFrame, reason_col: str):
    """Group report rows by (policy, product, from, to); yields (key, rows_df)."""
    if df is None or df.empty:
        return []
    pol = _find_col(df, ("policy no", "policy number", "policy", "policyno", "policy_number"))
    prod = _find_col(df, ("product", "product name", "normalized_product_name"))
    fm = _find_col(df, ("from month", "from_month", "month_from"))
    tm = _find_col(df, ("to month", "to_month", "month_to"))
    keys = [c for c in (pol, prod, fm, tm) if c]
    if not keys:
        return []
    out = []
    for key, g in df.groupby(keys, sort=True, dropna=False):
        key = key if isinstance(key, tuple) else (key,)
        vals = dict(zip(keys, key))
        out.append((
            str(vals.get(pol, "")),
            str(vals.get(prod, "")),
            _md_months(vals.get(fm), vals.get(tm)),
            g,
        ))
    return out


def write_readable_summary_md(
    md_path: str,
    *,
    source_file: str,
    counts: dict,
    df_stage1_invalid: pd.DataFrame,
    df_stage2_unmatched: pd.DataFrame,
    df_stage2_ambiguous: pd.DataFrame,
    missing_payloads: list,
    mismatch_payloads: list,
    blocked_payloads: list,
    snapshot_path: Optional[str] = None,
) -> None:
    """Write the validation summary in plain language, one entry per policy + month band."""
    c = counts
    n_post_new = len(missing_payloads or [])
    n_replace = sum(1 for p in (mismatch_payloads or []) if not p.get("_delete_only"))
    n_delete_only = sum(1 for p in (mismatch_payloads or []) if p.get("_delete_only"))
    n_blocked = len(blocked_payloads or [])
    W = []
    W.append("# Commission validation summary\n")
    W.append(f"**File:** {os.path.basename(str(source_file or ''))}  ")
    W.append(f"**Run at:** {datetime.datetime.now().strftime('%d %b %Y, %H:%M')}\n")

    W.append("## At a glance\n")
    W.append("| | Rows | Policies |")
    W.append("|---|---:|---:|")
    W.append(f"| Rows in the file | {c['tot_records']:,} | {c.get('tot_pols', 0):,} |")
    if c.get("s1_dup"):
        W.append(f"| Exact duplicate rows (removed) | {c['s1_dup']:,} | |")
    if c.get("s1_zero"):
        W.append(f"| Zero commission (ignored) | {c['s1_zero']:,} | |")
    W.append(f"| ❌ Failed the file checks (not compared) | {c['s1_err']:,} | {c.get('s1_err_pols', 0):,} |")
    W.append(f"| ✅ Same as the system | {c['s2_match']:,} | {c.get('s2_match_pols', 0):,} |")
    W.append(f"| ⚠️ Different from the system, or missing there | {c['s2_unmatch']:,} | {c.get('s2_unmatch_pols', 0):,} |")
    if c.get("s2_ambig"):
        W.append(f"| ⚠️ Could not tell which policy (not compared) | {c['s2_ambig']:,} | {c.get('s2_ambig_pols', 0):,} |")
    W.append("")

    W.append("## What will change if you confirm the update\n")
    if not (n_post_new or n_replace or n_delete_only):
        W.append("**Nothing.** There is nothing to delete or post: the file already matches the system, "
                 "or the differences were set aside below.\n")
    else:
        W.append("Nothing has been changed yet. These changes happen only after you confirm.\n")
        if n_post_new:
            W.append(f"- **{n_post_new:,}** new rule(s) will be added (the system has none for those months).")
        if n_replace:
            W.append(f"- **{n_replace:,}** rule(s) differ from the file: the system rule is deleted and posted again.")
        if n_delete_only:
            W.append(f"- **{n_delete_only:,}** system rule(s) sit inside months the file now covers with other slabs; they are deleted only.")
        W.append("")
    if n_blocked:
        W.append(f"⛔ **{n_blocked:,}** rule(s) are held back and will not be posted, because they failed the final checks. "
                 "See *Held back* below.\n")

    def _section(title: str, intro: str, groups, body_fn):
        if not groups:
            return
        W.append(f"## {title} ({len(groups):,})\n")
        W.append(intro + "\n")
        for i, (pol, prod, months, g) in enumerate(groups):
            if i >= MD_MAX_ITEMS_PER_SECTION:
                W.append(f"_…and {len(groups) - MD_MAX_ITEMS_PER_SECTION:,} more; see the Excel report._\n")
                break
            W.append(f"### Policy {pol} · {months}")
            W.append(f"*{_md_escape(prod)}*\n")
            body_fn(g)
            W.append("")

    # 1) Differences / missing in the system
    def _s2_body(g):
        reasons = []
        for r in g.get("Stage2_Unmatched_Reason", pd.Series(dtype=str)).tolist():
            t = explain_reason(r)
            if t and t not in reasons:
                reasons.append(t)
        for t in reasons:
            W.append(f"- {t}")
        first = g.iloc[0]
        table, _ = _md_split_table(first.get("DB_Value"), first.get("File_Value"))
        if table:
            W.append("")
            W.extend(table)

    _section(
        "Different from the system",
        "The file and the system disagree for these policies and months. ❌ marks the lines that differ.",
        _md_band_groups(df_stage2_unmatched, "Stage2_Unmatched_Reason"),
        _s2_body,
    )

    # 2) Ambiguous
    def _amb_body(g):
        t = explain_reason(g.iloc[0].get("Stage2_Unmatched_Reason", ""))
        if t:
            W.append(f"- {t}")

    amb_groups = _md_band_groups(df_stage2_ambiguous, "Stage2_Unmatched_Reason")
    # One entry per policy is enough for ambiguous: the reason is the same for every band
    seen_pol, amb_one = set(), []
    for item in amb_groups:
        if item[0] not in seen_pol:
            seen_pol.add(item[0])
            amb_one.append((item[0], item[1], "all months", item[3]))
    _section(
        "Could not tell which policy",
        "These policies were not compared or changed. Check the policy number or product in the file.",
        amb_one,
        _amb_body,
    )

    # 3) Stage 1 errors
    def _s1_body(g):
        statuses = [s for s in g.get("Stage1_Status", pd.Series(dtype=str)).dropna().unique().tolist()]
        for s in statuses:
            W.append(f"- {STAGE1_STATUS_TEXT.get(str(s), str(s))}")
        details = []
        for r in g.get("Stage1_Error_Reason", pd.Series(dtype=str)).dropna().tolist():
            r = " ".join(str(r).split())
            if r and r.lower() != "nan" and r not in details:
                details.append(r)
        for d in details[:3]:
            W.append(f"  - Detail: {_md_escape(d)}")
        name_col = _find_col(g, ("agent name", "agent_name"))
        contract_col = _find_col(g, ("contract", "level", "role"))
        if name_col:
            people = []
            for _, row in g.iterrows():
                who = " ".join(str(row.get(name_col, "")).split()).title()
                lvl = str(row.get(contract_col, "") if contract_col else "").strip()
                lvl = "" if lvl.lower() == "nan" else lvl
                people.append(f"{who} ({lvl or 'no contract level'})")
            W.append(f"  - People in this band: {', '.join(people)}")

    _section(
        "Failed the file checks",
        "These rows were not compared with the system. Fix them in the file and run again.",
        _md_band_groups(df_stage1_invalid, "Stage1_Error_Reason"),
        _s1_body,
    )

    # 4) Held back
    if n_blocked:
        W.append(f"## Held back ({n_blocked:,})\n")
        W.append("These rules were built but will not be posted, because they failed the final checks.\n")
        W.append("| Policy | Months | Why |")
        W.append("|---|---|---|")
        for i, p in enumerate(blocked_payloads):
            if i >= MD_MAX_ITEMS_PER_SECTION:
                W.append(f"\n_…and {n_blocked - MD_MAX_ITEMS_PER_SECTION:,} more; see rejected_missing_user_uuid_payloads_*.json._")
                break
            why = str(p.get("_issue_reason") or "").replace("Order ", "Line ")
            why = why.replace("level is null", "has no contract level").replace("Split sum mismatch DB left != File right", "the new shares do not add up to the system's total")
            W.append(f"| {p.get('policyNumber', '')} | {_md_months(p.get('monthFrom'), p.get('monthTo'))} | {_md_escape(why)} |")
        W.append("")

    if snapshot_path:
        W.append(f"---\nSystem rules at the time of this run: [{os.path.basename(snapshot_path)}]({snapshot_path})\n")

    with open(md_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(W) + "\n")


def main():
    parser = argparse.ArgumentParser(description="2-Stage Policy Commission Import Validation & DB Verification Script")
    parser.add_argument(
        "file_path",
        nargs="?",
        default=None,
        help="Optional positional path (same as --file / UPLOAD_FILE_PATH in .env)",
    )
    parser.add_argument("--file", type=str, help="Path to input commission upload file or directory (CSV/Excel)")
    parser.add_argument("--rate-sheet", type=str, help="Path to master ProductWise Commission Rate.xlsx file")
    parser.add_argument("--norm-file", type=str, help="Path to product normalization table")
    parser.add_argument("--agent-roster", type=str, help="Path to Master Agent Roster file")
    parser.add_argument("--policy-mapping", type=str, help="Path to Recovered Policies Excel mapping file")
    parser.add_argument("--env-file", type=str, default=".env", help="Path to .env configuration file")
    parser.add_argument("--output-dir", type=str, default="pipeline_output/commission_validation_db", help="Directory to save output reports")
    parser.add_argument("--agency-id", type=int, default=5, help="Agency ID for database query filtering (default: 5)")
    parser.add_argument("--dry-run", action="store_true", help="Simulate DB queries for testing without active DB connection")
    parser.add_argument("--tolerance", type=float, default=0.02, help="Tolerance for commission value matching (default: 0.02 / +-0.02%%)")
    parser.add_argument(
        "--yes", "-y",
        action="store_true",
        help="Reserved for non-interactive helpers; delete gate still asks y/n, then POST auto-runs",
    )
    parser.add_argument(
        "--skip-normalize",
        action="store_true",
        help="Skip year-band From-To normalization and override 3-row expansion before Stage 1",
    )
    parser.add_argument(
        "--skip-delete-gate",
        action="store_true",
        help="Skip interactive delete/POST confirm (use with --force-post-mismatches or --post-missing-only)",
    )
    parser.add_argument(
        "--force-post-mismatches",
        action="store_true",
        help="API-DELETE mismatch uuids via /api/v1/policy/commission/{uuid}, then AUTO-POST missing+mismatch rules",
    )
    parser.add_argument(
        "--post-missing-only",
        action="store_true",
        help="Exclude mismatch rules from AUTO-POST; only POST missing-in-DB rules (no API delete)",
    )

    args = parser.parse_args()

    env_loaded = load_env_file(args.env_file, override=True)
    print("==========================================================")
    print("   PIEQ 2-STAGE POLICY COMMISSION VALIDATION & DB ENGINE  ")
    print("==========================================================")
    print(f"⚙️ Loaded {env_loaded} environment variable(s) from '{args.env_file}'.")

    agency_id = int(os.getenv("AGENCY_ID", args.agency_id))
    # Prefer explicit flags, then .env UPLOAD_FILE_PATH (positional is optional override)
    input_file = args.file or args.file_path or os.getenv("UPLOAD_FILE_PATH")

    if not input_file or not os.path.exists(input_file):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        possible = [f for f in os.listdir(script_dir) if f.endswith((".csv", ".xlsx")) and "commission" in f.lower()]
        if possible:
            input_file = os.path.join(script_dir, possible[0])

    if not input_file or not os.path.exists(input_file):
        print("❌ Error: Input file not specified or file not found. Provide via --file or UPLOAD_FILE_PATH in .env.")
        sys.exit(1)

    if os.path.isdir(input_file):
        print(f"📖 Input Upload Folder: '{input_file}' (all CSV and Excel files)")
    else:
        print(f"📖 Input Upload File: '{input_file}'")
    print(f"🏢 Target Agency ID: {agency_id}")

    print("\n----------------------------------------------------------")
    print("PRE-FLIGHT DATABASE CONNECTION CHECK")
    print("----------------------------------------------------------")
    db_mgr = DatabaseManager()
    db_connected = False
    if not args.dry_run:
        print("🔌 Testing Database Connection before processing...")
        db_connected = db_mgr.connect()
        if db_connected:
            print("   ✅ Pre-flight Database Connection: SUCCESSFUL!")
        else:
            print("   ⚠️ Pre-flight Database Connection: FAILED (Stage 2 will run in DRY-RUN mode).")
    else:
        print("   ℹ️ Running in DRY-RUN mode (Skipping Database connection test).")

    df_rules = None
    rate_sheet_path = args.rate_sheet or os.getenv("RATE_SHEET_PATH")
    if rate_sheet_path and os.path.exists(rate_sheet_path):
        print(f"\n📖 Loading Master Rate Sheet from '{rate_sheet_path}'...")
        df_rules = unpivot_product_commission_rates(rate_sheet_path)
        print(f"   ✅ Extracted {len(df_rules):,} active rate rules.")

    norm_dict = {}
    norm_file_path = args.norm_file or os.getenv("NORM_FILE_PATH")
    if norm_file_path and os.path.exists(norm_file_path):
        print(f"📖 Loading Product Normalization Map from '{norm_file_path}'...")
        norm_dict = load_product_normalization(norm_file_path)
        print(f"   ✅ Extracted {len(norm_dict):,} product mappings.")

    agent_roster_dict = {}
    agent_roster_path = args.agent_roster or os.getenv("AGENT_ROSTER_PATH")
    if agent_roster_path and os.path.exists(agent_roster_path):
        print(f"📖 Loading Agent Roster from '{agent_roster_path}'...")
        agent_roster_dict = load_agent_roster(agent_roster_path)
        print(f"   ✅ Extracted {len(agent_roster_dict):,} agent mappings.")

    print("\n----------------------------------------------------------")
    print("STAGE 1: DIFFERENTIAL ROLL-UP, SPLIT NORMALIZATION & RATE CHECK")
    print("----------------------------------------------------------")
    df_comm = load_file(input_file)
    print(f"📄 Loaded {len(df_comm):,} total input rows.")
    
    # Preprocess Policy Numbers to remove single quotes
    pol_col = next((c for c in df_comm.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
    if pol_col:
        print(f"🧹 Preprocessing: Stripping single quotes from '{pol_col}'...")
        df_comm[pol_col] = df_comm[pol_col].astype(str).str.replace("'", "", regex=False)
        
    df_comm["_is_mapped"] = False
    
    # Optional Policy Mapping Replacement
    policy_mapping_path = args.policy_mapping or os.getenv("POLICY_MAPPING_PATH")
    if policy_mapping_path and os.path.exists(policy_mapping_path) and pol_col:
        try:
            print(f"\n📂 Found Policy Mapping file: '{policy_mapping_path}'")
            user_input = "y" if args.yes else input("   Do you want to apply Processed Policy Numbers? (y/n): ").strip().lower()
            if user_input in ('y', 'yes'):
                print("   🔄 Loading mapping file and applying replacements...")
                df_map = pd.read_excel(policy_mapping_path) if policy_mapping_path.endswith((".xlsx", ".xls")) else pd.read_csv(policy_mapping_path)
                
                map_pol_col = next((c for c in df_map.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno")), None)
                map_proc_col = next((c for c in df_map.columns if str(c).strip().lower() in ("processed policy number", "processed_policy_number")), None)
                
                if map_pol_col and map_proc_col:
                    df_map[map_pol_col] = df_map[map_pol_col].astype(str).str.replace("'", "", regex=False).str.strip()
                    df_map[map_proc_col] = df_map[map_proc_col].astype(str).str.replace("'", "", regex=False).str.strip()
                    
                    df_map_valid = df_map[df_map[map_proc_col] != "nan"]
                    df_map_valid = df_map_valid[df_map_valid[map_proc_col] != ""]
                    
                    policy_dict = dict(zip(df_map_valid[map_pol_col], df_map_valid[map_proc_col]))
                    
                    df_comm["_mapped_pol"] = df_comm[pol_col].map(policy_dict)
                    mapped_mask = df_comm["_mapped_pol"].notna()
                    replaced_count = mapped_mask.sum()
                    
                    if replaced_count > 0:
                        df_comm.loc[mapped_mask, pol_col] = df_comm.loc[mapped_mask, "_mapped_pol"]
                        df_comm.loc[mapped_mask, "_is_mapped"] = True
                        print(f"   ✅ Successfully replaced {replaced_count:,} policy numbers using the mapping file.")
                    else:
                        print("   ⚠️ No matching policy numbers found in the mapping file to replace.")
                        
                    df_comm = df_comm.drop(columns=["_mapped_pol"], errors="ignore")
                else:
                    print(f"   ⚠️ Could not find 'Policy No' or 'Processed Policy Number' columns in mapping file. Skipping mapping.")
            else:
                print("   ⏩ Skipping Policy Mapping as per user input.")
        except Exception as e:
            print(f"   ❌ Error applying Policy Mapping: {e}")
            
    # Preprocessing: Upfront Deduplication
    ignore_cols = ["transaction_id", "source_batch_id", "uploaded_at", "id"]
    subset_cols = [c for c in df_comm.columns if str(c).strip().lower() not in ignore_cols]
    
    is_dup = df_comm.duplicated(subset=subset_cols, keep="first")
    df_duplicates = df_comm[is_dup].copy()
    df_comm_dedup = df_comm[~is_dup].copy().reset_index(drop=True)
    
    if len(df_duplicates) > 0:
        print(f"🧹 Preprocessing: Removed {len(df_duplicates):,} exact duplicate rows.")

    # Overrides on hold: set OVR rows aside (not validated, never deleted or posted)
    df_override_on_hold = pd.DataFrame()
    if not OVERRIDE_UPDATES_ENABLED:
        _prod_col_hold = next((c for c in df_comm_dedup.columns if str(c).strip().lower() in ("product", "product name")), None)
        if _prod_col_hold:
            _ovr_mask = df_comm_dedup[_prod_col_hold].apply(is_override_product)
            df_override_on_hold = df_comm_dedup[_ovr_mask].copy()
            df_comm_dedup = df_comm_dedup[~_ovr_mask].copy().reset_index(drop=True)
            if len(df_override_on_hold):
                _hold_pols = df_override_on_hold[pol_col].nunique() if pol_col else 0
                print(
                    f"⏸️  Overrides on hold: {len(df_override_on_hold):,} OVR row(s) across {_hold_pols:,} policy(ies) set aside "
                    f"(no override check, no DELETE, no POST)."
                )

    # Year-band From-To alignment + override Agency 100% -> 3 rows (before Stage 1 roll-up)
    if not args.skip_normalize:
        print("\n----------------------------------------------------------")
        print("PRE-STAGE 1: YEAR-BAND NORMALIZE + OVERRIDE 3-ROW EXPAND")
        print("----------------------------------------------------------")
        df_comm_dedup, overlap_summary = fix_inclusive_month_band_overlaps(df_comm_dedup)
        n_overlap = int(overlap_summary.get("rows_fixed", 0) or 0)
        if n_overlap:
            print(
                f"🔗 Inclusive month-band overlap fix: {n_overlap:,} row(s) "
                f"(From bumped to previous To+1 so bands do not share an endpoint)"
            )
            for ex in overlap_summary.get("examples") or []:
                print(
                    f"   • Policy {ex.get('policy')}: "
                    f"{ex.get('old_from')}->{ex.get('to')} became "
                    f"{ex.get('new_from')}->{ex.get('to')}"
                )
        else:
            print("🔗 Inclusive month-band overlap fix: 0 row(s) (no endpoint overlaps)")
        before_rows = len(df_comm_dedup)
        df_comm_dedup, band_summary = normalize_year_bands(df_comm_dedup)
        print(
            f"📐 Year-band normalize: {band_summary.get('groups_touched', 0):,} group(s) re-sliced "
            f"({band_summary.get('rows_in', before_rows):,} -> {band_summary.get('rows_out', len(df_comm_dedup)):,} rows)"
        )
        db_for_override = db_mgr if db_connected else None
        df_comm_dedup, ovr_summary = ensure_override_three_rows(
            df_comm_dedup, db_mgr=db_for_override, agency_id=agency_id
        )
        print(
            f"🧩 Override 3-row expand: {ovr_summary.get('groups_expanded', 0):,} group(s), "
            f"+{ovr_summary.get('rows_appended', 0):,} row(s) "
            f"(DB lookups={ovr_summary.get('db_lookups', 0)}, misses={ovr_summary.get('db_misses', 0)})"
        )
    else:
        print("⏩ Skipping year-band normalize / override expand (--skip-normalize).")

    df_normalized_export = df_comm_dedup.copy()

    df_stage1_valid, df_stage1_invalid, df_stage1_zero, s1_summary = stage1_pre_validate_dataset(
        df_comm=df_comm_dedup,
        df_rules=df_rules,
        norm_dict=norm_dict,
        agent_roster_dict=agent_roster_dict,
        tolerance=args.tolerance,
    )

    print(f"✅ Stage 1 Pre-Validation Passed : {s1_summary['stage1_valid_records']:,} rows")
    print(f"❌ Stage 1 Pre-Validation Failed : {s1_summary['stage1_invalid_records']:,} rows (Bypassing DB queries)")
    if s1_summary['stage1_zero_ignored'] > 0:
        print(f"🟡 Stage 1 Zero Comm Ignored   : {s1_summary['stage1_zero_ignored']:,} rows")

    # Create a single dynamic version folder for this run
    run_version = f"v{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
    versioned_dir = os.path.join(args.output_dir, run_version)
    input_dir = os.path.join(versioned_dir, "input")
    output_dir = os.path.join(versioned_dir, "output")
    
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    archive_upload_input(input_file, input_dir)

    # Export pre-Stage-1 normalized sheet (year-bands + override rows)
    ts_norm = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    normalized_csv = os.path.join(output_dir, f"normalized_yearband_override_{ts_norm}.csv")
    normalized_xlsx = os.path.join(output_dir, f"normalized_yearband_override_{ts_norm}.xlsx")
    try:
        df_normalized_export.to_csv(normalized_csv, index=False)
        df_normalized_export.to_excel(normalized_xlsx, index=False)
        print(f" 📁 Normalized CSV     : {normalized_csv}")
        print(f" 📁 Normalized Excel   : {normalized_xlsx}")
    except Exception as e:
        print(f"⚠️ Could not export normalized sheet: {e}")

    # Phase 1 Exports
    phase1_clean_path, phase1_error_path = export_phase1_reports(df_stage1_valid, df_stage1_invalid, output_dir)
    print(f" 📁 Phase 1 Clean File  : {phase1_clean_path}")
    print(f" 📁 Phase 1 Error File  : {phase1_error_path}")
    
    dup_path = ""
    if len(df_duplicates) > 0:
        dup_path = os.path.join(output_dir, f"file_validated_duplicates_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        df_duplicates.to_excel(dup_path, index=False)
        print(f" 📁 Phase 1 Duplicates  : {dup_path}")
        
    if len(df_override_on_hold) > 0:
        hold_path = os.path.join(output_dir, f"overrides_on_hold_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        df_override_on_hold.to_excel(hold_path, index=False)
        print(f" 📁 Overrides on hold   : {hold_path}")

    zero_path = ""
    if len(df_stage1_zero) > 0:
        zero_path = os.path.join(output_dir, f"file_validated_zero_comm_ignored_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        df_stage1_zero.to_excel(zero_path, index=False)
        print(f" 📁 Phase 1 Zero Comm   : {zero_path}")

    print("\n----------------------------------------------------------")
    print(f"STAGE 2: READ-ONLY DATABASE VERIFICATION (AGENCY_ID = {agency_id})")
    print("----------------------------------------------------------")

    dry_run_mode = args.dry_run or (not db_connected)
    if dry_run_mode and not args.dry_run:
        print("⚠️ Notice: Database unestablished. Executing Stage 2 in DRY-RUN simulation mode.")

    # ----------------------------------------------------------
    # STAGE 1.5: DB COMMISSION SNAPSHOT EXPORT
    # ----------------------------------------------------------
    print("\n----------------------------------------------------------")
    print("STAGE 1.5: FETCHING CURRENT DB COMMISSION VALUES FOR VALID POLICIES")
    print("----------------------------------------------------------")
    snapshot_path = ""
    if db_connected and not dry_run_mode:
        print(f"🔍 Fetching DB commission snapshot for {df_stage1_valid[next((c for c in df_stage1_valid.columns if str(c).strip().lower() in ('policy no', 'policy number', 'policy', 'policyno', 'policy_number')), df_stage1_valid.columns[0])].nunique() if not df_stage1_valid.empty else 0:,} unique policies...")
        users_map_snapshot = db_mgr.fetch_all_users_map(agency_id)
        df_db_snapshot = fetch_db_commission_snapshot(
            df_stage1_valid=df_stage1_valid,
            db_mgr=db_mgr,
            agency_id=agency_id,
            users_map=users_map_snapshot,
        )
        snapshot_path = export_db_commission_snapshot(
            df_snapshot=df_db_snapshot,
            versioned_dir=output_dir,
            df_stage1_valid=df_stage1_valid,
        )
        snapshot_pol_col = next((c for c in df_db_snapshot.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
        snapshot_unique_pols = df_db_snapshot[snapshot_pol_col].nunique() if snapshot_pol_col and not df_db_snapshot.empty else 0
        print(f"   ✅ DB Snapshot: {len(df_db_snapshot):,} split rows across {snapshot_unique_pols:,} policies.")
        print(f" 📁 DB Commission Snapshot : {snapshot_path}")
    else:
        print("   ⚠️ Skipping DB snapshot (no active DB connection or dry-run mode).")

    (
        df_stage2_matched,
        df_stage2_ambiguous,
        df_stage2_unmatched,
        unique_missing_payloads,
        unique_mismatch_payloads,
        invalid_user_uuid_payloads,
        writing_agent_corrections,
    ) = stage2_db_verify_dataset(
        df_valid=df_stage1_valid,
        db_mgr=db_mgr,
        agency_id=agency_id,
        dry_run=dry_run_mode,
        tolerance=args.tolerance,
    )
    db_mgr.close()

    if writing_agent_corrections:
        wa_sql_path = os.path.join(output_dir, "update_writing_agent_from_file.sql")
        wa_json_path = os.path.join(
            output_dir,
            f"writing_agent_corrections_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
        )
        abs_wa_sql = write_writing_agent_update_sql(
            writing_agent_corrections, agency_id=agency_id, out_path=wa_sql_path
        )
        with open(wa_json_path, "w", encoding="utf-8") as f:
            json.dump(writing_agent_corrections, f, indent=2)
        print(
            f" ✍️ Writing-agent corrections: {len(writing_agent_corrections):,} policy(ies) "
            f"(file highest-order agent preferred over DB writing_agent_user_id)"
        )
        for c in writing_agent_corrections[:15]:
            print(
                f"   • {c.get('policyNumber')}: "
                f"DB '{c.get('dbWritingAgentName')}' → file '{c.get('fileWritingAgentName')}'"
            )
        print(f" 📁 UPDATE SQL (share with senior eng; not executed): {abs_wa_sql}")
        print(f" 📁 Corrections JSON: {wa_json_path}")

    if invalid_user_uuid_payloads:
        missing_uuid_file = os.path.join(output_dir, f"rejected_missing_user_uuid_payloads_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        with open(missing_uuid_file, "w") as f:
            json.dump(invalid_user_uuid_payloads, f, indent=2)
        print(f" ⚠️ Excluded {len(invalid_user_uuid_payloads):,} payload(s) that failed split validation. Logged to: {missing_uuid_file}")

    df_all_unmatched = pd.concat([df_stage1_invalid, df_stage2_unmatched], ignore_index=True)

    # Phase 2 Exports
    phase2_matched_path, phase2_unmatched_path = export_phase2_reports(df_stage2_matched, df_stage2_unmatched, output_dir)
    
    ambiguous_path = ""
    if len(df_stage2_ambiguous) > 0:
        ambiguous_path = os.path.join(output_dir, f"file_validated_ambiguous_matches_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        df_stage2_ambiguous.to_excel(ambiguous_path, index=False)
        print(f" 📁 Phase 2 Ambiguous  : {ambiguous_path}")

    tot_records = len(df_comm)
    s1_dup = len(df_duplicates)
    tot_unique = len(df_comm_dedup)
    
    s1_zero = len(df_stage1_zero)
    s1_clean = len(df_stage1_valid)
    s1_err = len(df_stage1_invalid)
    
    s2_match = len(df_stage2_matched)
    s2_ambig = len(df_stage2_ambiguous)
    s2_unmatch = len(df_stage2_unmatched)
    tot_unmatch = len(df_all_unmatched)
    
    pol_col = next((c for c in df_comm.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
    
    def count_unique_pols(df):
        if len(df) == 0 or not pol_col or pol_col not in df.columns: return 0
        return df[pol_col].nunique()
        
    s1_err_pols = count_unique_pols(df_stage1_invalid)
    s2_unmatch_pols = count_unique_pols(df_stage2_unmatched)
    s2_ambig_pols = count_unique_pols(df_stage2_ambiguous)
    tot_unmatch_pols = count_unique_pols(df_all_unmatched)
    
    s1_err_str = f"{s1_err:,} ({s1_err_pols:,} pols)" if s1_err > 0 else "0"
    s2_unmatch_str = f"{s2_unmatch:,} ({s2_unmatch_pols:,} pols)" if s2_unmatch > 0 else "0"
    s2_ambig_str = f"{s2_ambig:,} ({s2_ambig_pols:,} pols)" if s2_ambig > 0 else "0"
    tot_unmatch_str = f"{tot_unmatch:,} ({tot_unmatch_pols:,} pols)" if tot_unmatch > 0 else "0"

    pct_s1_dup = round((s1_dup / tot_records) * 100, 2) if tot_records > 0 else 0
    pct_tot_unique = round((tot_unique / tot_records) * 100, 2) if tot_records > 0 else 0
    
    pct_s1_zero = round((s1_zero / tot_unique) * 100, 2) if tot_unique > 0 else 0
    pct_s1_clean = round((s1_clean / tot_unique) * 100, 2) if tot_unique > 0 else 0
    pct_s1_err = round((s1_err / tot_unique) * 100, 2) if tot_unique > 0 else 0
    
    pct_s2_match = round((s2_match / s1_clean) * 100, 2) if s1_clean > 0 else 0
    pct_s2_ambig = round((s2_ambig / s1_clean) * 100, 2) if s1_clean > 0 else 0
    pct_s2_unmatch = round((s2_unmatch / s1_clean) * 100, 2) if s1_clean > 0 else 0
    
    pct_tot_unmatch = round((tot_unmatch / tot_unique) * 100, 2) if tot_unique > 0 else 0

    print("\n=========================================================================================")
    print("                              FINAL EXECUTION SUMMARY TABLE                              ")
    print("=========================================================================================")
    print("+------------------------------------+-----------------------+------------+------------+")
    print("| Metric / Pipeline Stage            | Record Count          | Out of     | Percentage |")
    print("+------------------------------------+-----------------------+------------+------------+")
    print(f"| Total Input Records Uploaded       | {tot_records:<21,} | {tot_records:<10,} | 100.00%    |")
    if s1_dup > 0:
        print(f"| 🟡 Exact Duplicates (Removed)      | {s1_dup:<21,} | {tot_records:<10,} | {pct_s1_dup:>6.2f}%   |")
    print(f"| 🔵 Unique Records Processed        | {tot_unique:<21,} | {tot_records:<10,} | {pct_tot_unique:>6.2f}%   |")
    print("+------------------------------------+-----------------------+------------+------------+")
    if s1_zero > 0:
        print(f"| 🟡 Phase 1 Zero Comm (Ignored)     | {s1_zero:<21,} | {tot_unique:<10,} | {pct_s1_zero:>6.2f}%   |")
    print(f"| 🟢 Phase 1 Pre-Validated Clean     | {s1_clean:<21,} | {tot_unique:<10,} | {pct_s1_clean:>6.2f}%   |")
    print(f"| 🔴 Phase 1 Pre-Validation Errors   | {s1_err_str:<21} | {tot_unique:<10,} | {pct_s1_err:>6.2f}%   |")
    print("+------------------------------------+-----------------------+------------+------------+")
    print(f"| 🟢 Phase 2 DB Matched Records      | {s2_match:<21,} | {s1_clean:<10,} | {pct_s2_match:>6.2f}%   |")
    if s2_ambig > 0:
        print(f"| 🟡 Phase 2 Ambiguous (Skipped)     | {s2_ambig_str:<21} | {s1_clean:<10,} | {pct_s2_ambig:>6.2f}%   |")
    print(f"| 🔴 Phase 2 DB Unmatched Records    | {s2_unmatch_str:<21} | {s1_clean:<10,} | {pct_s2_unmatch:>6.2f}%   |")
    print("+------------------------------------+-----------------------+------------+------------+")
    print(f"| ❌ TOTAL FAILED / UNMATCHED (P1+P2)| {tot_unmatch_str:<21} | {tot_unique:<10,} | {pct_tot_unmatch:>6.2f}%   |")
    print("+------------------------------------+-----------------------+------------+------------+")
    
    if s1_err > 0 or s2_unmatch > 0 or s2_ambig > 0:
        print("\n-----------------------------------------------------------------------------------------")
        print("                              ERROR & MISMATCH BREAKDOWN                                 ")
        print("-----------------------------------------------------------------------------------------")
        
        def group_error_reason(reason: str) -> str:
            r = str(reason).lower()
            if "not found in policies table" in r:
                return "Policy number not found in policies table"
            if "policy group not found in db at all" in r or "no commission rule record found in policy_commission table" in r or "not found in db for" in r or "not found in db policy commission splits" in r or "not found in db commiss" in r:
                return "No commission rule record found in policy_commission table"
            if "found in db, but db commission_value is" in r:
                return "Commission mismatch with DB"
            if "similar match for" in r and "returned multiple different policies" in r:
                return "Ambiguous match (multiple policies)"
            if "split rows instead of expected" in r:
                return "Invalid split row count"
            if "not found in master rate sheet" in r:
                return "Product missing in rate sheet"
            if "out of rate tier ranges" in r:
                return "Month out of rate tier range"
            if "group split share sum is" in r:
                return "Split share sum is not 100%"
            if "carrier commission mismatch" in r:
                return "Carrier commission mismatch"
            return str(reason)

        if s1_err > 0 and "Stage1_Error_Reason" in df_stage1_invalid.columns:
            print("🔴 Phase 1 Pre-Validation Errors:")
            s1_counts = df_stage1_invalid["Stage1_Error_Reason"].apply(group_error_reason).value_counts()
            for reason, count in s1_counts.items():
                print(f"   - {str(reason)[:70]:<70} : {count} records")
                
        if s2_unmatch > 0 and "Stage2_Unmatched_Reason" in df_stage2_unmatched.columns:
            print("\n🔴 Phase 2 DB Unmatched:")
            s2_counts = df_stage2_unmatched["Stage2_Unmatched_Reason"].apply(group_error_reason).value_counts()
            for reason, count in s2_counts.items():
                print(f"   - {str(reason)[:70]:<70} : {count} records")
                
        if s2_ambig > 0 and "Stage2_Unmatched_Reason" in df_stage2_ambiguous.columns:
            print("\n🟡 Phase 2 Ambiguous:")
            s2a_counts = df_stage2_ambiguous["Stage2_Unmatched_Reason"].apply(group_error_reason).value_counts()
            for reason, count in s2a_counts.items():
                print(f"   - {str(reason)[:70]:<70} : {count} records")

    print("\n-----------------------------------------------------------------------------------------")
    print("OUTPUT FILE REPORTS GENERATED:")
    print(f" 📁 Phase 1 Clean File  : {phase1_clean_path}")
    print(f" 📁 Phase 1 Error File  : {phase1_error_path}")
    if s1_dup > 0:
        print(f" 📁 Phase 1 Duplicates  : {dup_path}")
    if s1_zero > 0:
        print(f" 📁 Phase 1 Zero Comm   : {zero_path}")
    print(f" 📁 Phase 2 Matched File: {phase2_matched_path}")
    if s2_ambig > 0:
        print(f" 📁 Phase 2 Ambiguous   : {ambiguous_path}")
    print(f" 📁 Phase 2 Unmatched   : {phase2_unmatched_path}")
    if snapshot_path:
        print(f" 📁 DB Snapshot File    : {snapshot_path}")
    print("=========================================================================================")
    # Write the readable markdown summary (one entry per policy + month band)
    md_summary_path = os.path.join(output_dir, f"validation_summary_{run_version}.md")
    try:
        write_readable_summary_md(
            md_summary_path,
            source_file=input_file or "",
            counts={
                "tot_records": tot_records,
                "tot_pols": count_unique_pols(df_comm),
                "s1_dup": s1_dup,
                "s1_zero": s1_zero,
                "s1_err": s1_err,
                "s1_err_pols": count_unique_pols(df_stage1_invalid),
                "s2_match": s2_match,
                "s2_match_pols": count_unique_pols(df_stage2_matched),
                "s2_unmatch": s2_unmatch,
                "s2_unmatch_pols": count_unique_pols(df_stage2_unmatched),
                "s2_ambig": s2_ambig,
                "s2_ambig_pols": count_unique_pols(df_stage2_ambiguous),
            },
            df_stage1_invalid=df_stage1_invalid,
            df_stage2_unmatched=df_stage2_unmatched,
            df_stage2_ambiguous=df_stage2_ambiguous,
            missing_payloads=unique_missing_payloads,
            mismatch_payloads=unique_mismatch_payloads,
            blocked_payloads=invalid_user_uuid_payloads,
            snapshot_path=snapshot_path,
        )
    except Exception as e:
        print(f"⚠️ Failed to write markdown summary: {e}")
    print(f"\n📝 Validation summary markdown saved to: {md_summary_path}")
    successful_policies = set()

    # ------------------------------------------------------------------
    # FLOW: API DELETE mismatches -> POST recreate (no PUT)
    #   1) Collect mismatched policy_commission.uuid values
    #   2) Optionally exclude mismatches from POST (missing-only)
    #   3) On confirm / --force-post-mismatches: DELETE each uuid via
    #      DELETE /api/v1/policy/commission/{uuid}, then POST recreate
    # Matched slabs are not deleted and are not posted.
    # ------------------------------------------------------------------
    def _dedupe_post_payloads(payloads):
        deduped = []
        seen_keys = set()
        for p in payloads:
            if p.get("_delete_only"):
                # DB slab inside the file months: deleted, never posted
                continue
            key = (
                str(p.get("policyUuid") or ""),
                str(p.get("monthFrom")),
                str(p.get("monthTo")),
                str(p.get("commissionType")),
            )
            if key in seen_keys:
                continue
            seen_keys.add(key)
            deduped.append(p)
        return deduped

    from generate_delete_commission_sql import (
        mismatch_commission_uuids,
        write_mismatch_delete_sql,
    )

    mismatch_uuids = mismatch_commission_uuids(unique_mismatch_payloads)
    delete_line_count = len(mismatch_uuids)
    missing_only = bool(getattr(args, "post_missing_only", False))
    # --force-post-mismatches means include mismatches; it wins over --post-missing-only
    if args.force_post_mismatches:
        missing_only = False

    all_candidate_payloads = _dedupe_post_payloads(
        list(unique_missing_payloads or []) + list(unique_mismatch_payloads or [])
    )
    post_payloads = all_candidate_payloads
    do_api_delete = False

    # One machine-readable line for the UI (ui/runner.parse_gate_hints): lets it hide the
    # DELETE/POST confirmation when there is nothing to delete or post.
    _gate_mismatch_post = sum(1 for p in (unique_mismatch_payloads or []) if not p.get("_delete_only"))
    print(
        f"GATE_SUMMARY missing_post={len(unique_missing_payloads or [])} "
        f"mismatch_post={_gate_mismatch_post} delete={delete_line_count} "
        f"blocked={len(invalid_user_uuid_payloads or [])}"
    )

    # Always generate the mismatch DELETE SQL (UUID list) when there is anything to delete.
    delete_sql_path = None
    if mismatch_uuids:
        delete_md_path = os.path.join(output_dir, "delete_commission_agency5.sql")
        docs_md_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "docs", "delete_commission_agency5.sql"
        )
        try:
            delete_sql_path = write_mismatch_delete_sql(
                out_path=delete_md_path,
                agency_id=agency_id,
                commission_uuids=mismatch_uuids,
            )
            try:
                write_mismatch_delete_sql(
                    out_path=docs_md_path,
                    agency_id=agency_id,
                    commission_uuids=mismatch_uuids,
                )
            except Exception:
                pass
        except Exception as de:
            print(f"⚠️ Could not write delete SQL: {de}")

    if not all_candidate_payloads:
        print("\nℹ️ No missing/mismatch payloads to recreate via POST.")
        if delete_line_count:
            print(f"   (Mismatch delete SQL still lists {delete_line_count} line(s) that would be deleted.)")
            if delete_sql_path:
                print(f"   SQL: {delete_sql_path}")
    elif args.skip_delete_gate and not args.force_post_mismatches and not missing_only:
        print("\n⏩ --skip-delete-gate set without --force-post-mismatches: skipping delete gate and POST recreate.")
        print(f"   Lines that will be deleted (if you later API-DELETE mismatches): {delete_line_count}")
        if delete_sql_path:
            print(f"   Delete SQL generated: {delete_sql_path}")
        print("   Use --force-post-mismatches to API-DELETE mismatches then POST,")
        print("   or --post-missing-only to POST missing rules without delete.")
    else:
        delete_approved = bool(args.force_post_mismatches) or (
            bool(args.skip_delete_gate) and missing_only
        )
        if args.force_post_mismatches:
            do_api_delete = bool(mismatch_uuids)
            print(f"\n🗑️  --force-post-mismatches: lines that will be deleted via API: {delete_line_count}")
            if delete_sql_path:
                print(f"   Delete SQL generated: {delete_sql_path}")

        if not args.force_post_mismatches and not (args.skip_delete_gate and missing_only):
            print("\n" + "=" * 80)
            print("FLOW STEP 1/2 — API DELETE GATE")
            print("=" * 80)
            invalid_mismatch_rows = [
                p for p in (invalid_user_uuid_payloads or []) if str(p.get("rule_uuid") or "").strip()
            ]
            print(f"Rules available to create: {len(all_candidate_payloads)}")
            print(f"  - missing in DB (no delete): {len(unique_missing_payloads or [])}")
            print(f"  - mismatch rows to delete:   {delete_line_count}")
            _delete_only_count = sum(1 for p in (unique_mismatch_payloads or []) if p.get("_delete_only"))
            if _delete_only_count:
                print(
                    f"    (of which {_delete_only_count} are DB slabs inside the file months: "
                    f"delete only, the file slabs replace them)"
                )
            print(f"  - lines that will be deleted: {delete_line_count}")
            print("  - endpoint: DELETE /api/v1/policy/commission/{uuid}")
            print("  - matched slabs are omitted from DELETE and are not posted")
            if invalid_mismatch_rows:
                print(
                    f"  - mismatch rows left in place (split checks failed, so they are not deleted): "
                    f"{len(invalid_mismatch_rows)}"
                )
            if delete_sql_path:
                print(f"Delete SQL generated ({delete_line_count} uuid line(s)): {delete_sql_path}")
            else:
                print("⚠️ Delete SQL was not generated.")

            if mismatch_uuids and not missing_only:
                exclude_prompt = (
                    "\nExclude mismatch rules from this update?\n"
                    "  y = POST missing-in-DB only (no API delete)\n"
                    "  n = API-DELETE mismatches, then POST missing + mismatches\n"
                    "Type 'y' or 'n': "
                )
                exclude_ans = input(exclude_prompt).strip().lower()
                missing_only = exclude_ans in ("y", "yes")

            if missing_only:
                post_payloads = _dedupe_post_payloads(list(unique_missing_payloads or []))
                do_api_delete = False
                print(
                    f"\n➡️ Mismatch API-DELETE/POST eliminated for this run. "
                    f"Will POST missing only: {len(post_payloads)} payload(s)."
                )
                if not post_payloads:
                    print("ℹ️ No missing payloads left to POST after excluding mismatches.")
                    delete_approved = False
                else:
                    confirm_prompt = (
                        "\nFLOW STEP 2/2 — POST missing rules only (mismatches skipped).\n"
                        "Type 'y' to AUTO-POST missing rules now (n to stop): "
                    )
                    confirm_del = input(confirm_prompt).strip().lower()
                    delete_approved = confirm_del in ("y", "yes")
            elif mismatch_uuids:
                post_payloads = all_candidate_payloads
                confirm_prompt = (
                    f"\nFLOW STEP 2/2 — API-DELETE {delete_line_count} line(s) "
                    "via /api/v1/policy/commission/{{uuid}}, then POST recreate?\n"
                    "Type 'y' to DELETE + POST now (n to stop): "
                )
                confirm_del = input(confirm_prompt).strip().lower()
                delete_approved = confirm_del in ("y", "yes")
                do_api_delete = delete_approved
            else:
                post_payloads = all_candidate_payloads
                do_api_delete = False
                confirm_prompt = (
                    "\nFLOW STEP 2/2 — No mismatch rows to delete.\n"
                    "Type 'y' to AUTO-POST missing rules now (n to stop): "
                )
                confirm_del = input(confirm_prompt).strip().lower()
                delete_approved = confirm_del in ("y", "yes")
        elif missing_only:
            post_payloads = _dedupe_post_payloads(list(unique_missing_payloads or []))
            do_api_delete = False
            print(
                f"\n➡️ --post-missing-only: POST missing only "
                f"({len(post_payloads)} payload(s); mismatches excluded)."
            )

        deleted_ok: Set[str] = set()
        deleted_failed: Set[str] = set()
        delete_fail_reasons: Dict[str, str] = {}
        post_results: List[dict] = []
        post_skipped_reason: Optional[str] = None

        if delete_approved and do_api_delete and mismatch_uuids:
            print(f"\n🗑️  Deleting {delete_line_count} mismatch line(s) via API…")
            if delete_sql_path:
                print(f"   SQL reference: {delete_sql_path}")
            deleted_ok, deleted_failed, delete_fail_reasons = run_delete_commission_uuids(
                commission_uuids=mismatch_uuids,
                output_dir=output_dir,
            )
            if deleted_failed:
                print(
                    f"⚠️ Stopping before POST: {len(deleted_failed)} mismatch DELETE(s) failed "
                    f"({len(deleted_ok)} succeeded)."
                )
                delete_approved = False
                post_skipped_reason = (
                    f"POST skipped because API DELETE failed for {len(deleted_failed)} uuid(s)"
                )

        if delete_approved and post_payloads:
            scope = "missing only" if missing_only else "missing + mismatch"
            print("\n" + "=" * 80)
            print(f"FLOW — AUTO-POST recreate ({scope}): {len(post_payloads)} payload(s) (no PUT)")
            print("=" * 80)
            successful_policies, post_results = run_post_commission_payloads(
                payloads=post_payloads,
                output_dir=output_dir,
                auto_confirm=True,
            )
        elif delete_approved and not post_payloads:
            print("\nℹ️ Nothing to POST.")
        else:
            print("Stopped before API DELETE/POST. Re-run:")
            print("  python validate_policy_commission_db.py -y --force-post-mismatches --skip-delete-gate")
            print("Or to POST missing only (no delete):")
            print("  python validate_policy_commission_db.py -y --post-missing-only --skip-delete-gate")

        # Always emit Policy No / Status / Reason after a delete and/or post attempt
        if post_results or post_skipped_reason or deleted_ok or deleted_failed:
            try:
                status_path = write_policy_update_status_excel(
                    output_dir=output_dir,
                    post_payloads=post_payloads,
                    post_results=post_results,
                    mismatch_payloads=unique_mismatch_payloads,
                    deleted_ok=deleted_ok,
                    deleted_failed=deleted_failed,
                    delete_fail_reasons=delete_fail_reasons,
                    post_skipped_reason=post_skipped_reason,
                )
                print(f"\n📊 Policy update status Excel: {status_path}")
            except Exception as ste:
                print(f"⚠️ Could not write policy update status Excel: {ste}")

    if successful_policies:
        pol_col_s2 = next((c for c in df_stage2_unmatched.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
        df_updated = pd.DataFrame()
        if pol_col_s2:
            df_updated = df_stage2_unmatched[df_stage2_unmatched[pol_col_s2].astype(str).isin([str(p) for p in successful_policies])].copy()
            if not df_updated.empty:
                df_updated["Stage2_Status"] = "MATCHED (AUTO_UPDATED)"
                df_updated["Stage2_Unmatched_Reason"] = "Successfully updated via API"
                out_path = os.path.join(output_dir, f"db_verification_auto_updated_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                df_updated.to_excel(out_path, index=False)
                print(f"\n✅ Created Auto-Updated DB Match File with {len(df_updated)} rows: {out_path}")

        try:
            csv_path, txt_path = write_updated_policies_list(
                successful_policies=successful_policies,
                output_dir=output_dir,
                df_updated=df_updated if not df_updated.empty else None,
                policy_col=pol_col_s2,
            )
            print(
                f"✅ Updated policies list ({len(successful_policies)} unique): "
                f"{csv_path} | {txt_path}"
            )
        except Exception as upe:
            print(f"⚠️ Could not write updated policies list: {upe}")

        # Fetch latest post-update DB commission snapshot directly from database
        print("\n" + "="*80)
        print("🔄 FETCHING LATEST POST-UPDATE POLICY COMMISSIONS FROM DATABASE...")
        print("="*80)
        try:
            db_mgr_post = DatabaseManager()
            if db_mgr_post.connect():
                users_map_post = db_mgr_post.fetch_all_users_map(agency_id)
                df_post_snapshot = fetch_db_commission_snapshot(
                    df_stage1_valid=df_stage1_valid,
                    db_mgr=db_mgr_post,
                    agency_id=agency_id,
                    users_map=users_map_post,
                )
                if not df_post_snapshot.empty:
                    post_snapshot_path = export_db_commission_snapshot(
                        df_snapshot=df_post_snapshot,
                        versioned_dir=output_dir,
                        df_stage1_valid=df_stage1_valid,
                        filename_prefix="db_commission_snapshot_post_update",
                    )
                    post_pol_col = next((c for c in df_post_snapshot.columns if str(c).strip().lower() in ("policy no", "policy number", "policy", "policyno", "policy_number")), None)
                    post_unique_pols = df_post_snapshot[post_pol_col].nunique() if post_pol_col else 0
                    print(f"   ✅ Post-Update DB Snapshot: {len(df_post_snapshot):,} split rows across {post_unique_pols:,} policies.")
                    print(f" 📁 Post-Update DB Commission Snapshot File: {post_snapshot_path}")
                db_mgr_post.close()
            else:
                print("   ⚠️ Could not connect to DB for post-update snapshot.")
        except Exception as pe:
            print(f"   ⚠️ Error fetching post-update DB snapshot: {pe}")

if __name__ == "__main__":
    main()
