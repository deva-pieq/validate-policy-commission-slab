#!/usr/bin/env python3
"""
Raw DELETE SQL for mismatched policy_commission rows (agency_id=5).

The validation run writes this file. It deletes by policy_commission.uuid for
slabs that do not match the sheet. Policies whose slabs all match are omitted.
Missing slabs have no row to delete.

Does NOT run DELETE.
Output is raw SQL only (.sql): BEGIN / DELETE / ROLLBACK.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from typing import List, Optional, Set

_UUID_RE = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"
)

from validate_policy_commission_db import (
    DatabaseManager,
    load_env_file,
    load_file,
    policy_suffix_candidates,
)


def clean_policy_number(val) -> str:
    return str(val or "").strip().strip("'\"").strip()


def extract_policy_numbers(df) -> List[str]:
    pol_col = next(
        (
            c
            for c in df.columns
            if str(c).strip().lower()
            in ("policy no", "policy number", "policy", "policyno", "policy_number")
        ),
        None,
    )
    if not pol_col:
        raise KeyError("Could not find Policy No column in sheet.")
    seen: Set[str] = set()
    ordered: List[str] = []
    for raw in df[pol_col].tolist():
        p = clean_policy_number(raw)
        if not p or p.lower() == "nan":
            continue
        if p not in seen:
            seen.add(p)
            ordered.append(p)
    return ordered


def sql_text_array(values: List[str], indent: str = "    ") -> str:
    if not values:
        return "ARRAY[]::text[]"
    lines = []
    for v in values:
        esc = v.replace("'", "''")
        lines.append(f"{indent}'{esc}'")
    return "ARRAY[\n" + ",\n".join(lines) + "\n]::text[]"


def sql_values_list(values: List[str], indent: str = "  ") -> str:
    """Format VALUES ('a'),('b'),... for INSERT into temp table."""
    if not values:
        return f"{indent}-- (empty)"
    lines = []
    for v in values:
        esc = v.replace("'", "''")
        lines.append(f"{indent}('{esc}')")
    return ",\n".join(lines)


def build_match_numbers(policy_numbers: List[str]) -> List[str]:
    out: List[str] = []
    seen: Set[str] = set()
    for p in policy_numbers:
        for cand in [p] + policy_suffix_candidates(p):
            if cand and cand not in seen:
                seen.add(cand)
                out.append(cand)
    return out


def mismatch_commission_uuids(mismatch_payloads: List[dict]) -> List[str]:
    """Commission row ids that failed the sheet compare and will be recreated."""
    ordered: List[str] = []
    seen: Set[str] = set()
    for payload in mismatch_payloads or []:
        uid = str((payload or {}).get("rule_uuid") or "").strip()
        if not _UUID_RE.match(uid) or uid.lower() in seen:
            continue
        seen.add(uid.lower())
        ordered.append(uid)
    return ordered


def sql_uuid_array(values: List[str], indent: str = "    ") -> str:
    if not values:
        return "ARRAY[]::uuid[]"
    lines = [f"{indent}'{v}'" for v in values]
    return "ARRAY[\n" + ",\n".join(lines) + "\n]::uuid[]"


def write_mismatch_delete_sql(
    out_path: str,
    agency_id: int,
    commission_uuids: List[str],
) -> str:
    """DELETE only the given policy_commission.uuid rows. Matched slabs are not listed."""
    safe_ids = [v for v in commission_uuids if _UUID_RE.match(str(v).strip())]
    array_sql = sql_uuid_array(safe_ids)
    body = f"""BEGIN;

DELETE FROM policy_commission
WHERE agency_id = {int(agency_id)}
  AND uuid = ANY(
{array_sql}
  );

ROLLBACK;
"""
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(body)
    return os.path.abspath(out_path)


def write_raw_delete_sql_md(
    out_path: str,
    agency_id: int,
    match_numbers: List[str],
    source_label: str = "",
    sheet_policy_count: int = 0,
) -> str:
    """Write raw SQL only: BEGIN / DELETE / ROLLBACK. No comments or markdown."""
    array_sql = sql_text_array(match_numbers)
    body = f"""BEGIN;

DELETE FROM policy_commission pc
USING policies p
WHERE pc.policy_uuid = p.uuid
  AND pc.agency_id = {agency_id}
  AND p.agency_id = {agency_id}
  AND p.number = ANY(
{array_sql}
  );
COMMIT;
ROLLBACK;
"""
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    # utf-8 no BOM, LF newlines — avoids paste/client parse errors from .md / CRLF / BOM
    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(body)
    return os.path.abspath(out_path)


def write_delete_md_for_file(
    input_path: str,
    agency_id: int = 5,
    out_path: str = "docs/delete_commission_agency5.sql",
) -> str:
    """Delete every commission row for policy numbers on the sheet.

    The validation gate does not call this. Use write_mismatch_delete_sql.
    """
    df = load_file(input_path)
    sheet_policies = extract_policy_numbers(df)
    match_numbers = build_match_numbers(sheet_policies)
    return write_raw_delete_sql_md(
        out_path=out_path,
        agency_id=agency_id,
        match_numbers=match_numbers,
        source_label=os.path.abspath(input_path),
        sheet_policy_count=len(sheet_policies),
    )


def run_db_preview(db_mgr: DatabaseManager, agency_id: int, match_numbers: List[str]) -> Optional[dict]:
    if not db_mgr.conn:
        return None
    try:
        from psycopg2.extras import RealDictCursor

        with db_mgr.conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(
                """
                SELECT COUNT(*) AS cnt,
                       COUNT(DISTINCT p.number) AS distinct_pols
                FROM policy_commission pc
                JOIN policies p ON p.uuid = pc.policy_uuid AND p.agency_id = pc.agency_id
                WHERE pc.agency_id = %s
                  AND p.agency_id = %s
                  AND p.number = ANY(%s);
                """,
                (agency_id, agency_id, match_numbers),
            )
            row = cur.fetchone()
            return {
                "commission_count": int(row["cnt"]),
                "distinct_pol_with_comm": int(row["distinct_pols"]),
            }
    except Exception as e:
        print(f"DB preview failed: {e}")
        return None


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Delete SQL is written by the validator for mismatched commission rows only."
    )
    parser.add_argument(
        "file_path",
        nargs="?",
        default=None,
        help="Unused. Policy numbers on the sheet are not a delete list.",
    )
    parser.add_argument("--file", type=str, default=None, help="Unused. Kept so old commands fail clearly.")
    parser.add_argument("--env-file", type=str, default=".env")
    parser.add_argument("--agency-id", type=int, default=5)
    parser.add_argument(
        "--out",
        type=str,
        default="docs/delete_commission_agency5.sql",
        help="Output path written by validate_policy_commission_db.py",
    )
    parser.add_argument(
        "--skip-db-preview",
        action="store_true",
        help="Unused. This script no longer previews a policy-wide delete.",
    )
    parser.parse_args()
    print("This script no longer deletes every policy on the sheet.")
    print("A policy whose slabs all match is left out. A mixed policy deletes only the mismatched commission row.")
    print("Run the validator. It writes delete_commission_agency5.sql from those mismatched row ids:")
    print("  python validate_policy_commission_db.py")
    return 2


if __name__ == "__main__":
    sys.exit(main())
