#!/usr/bin/env python3
"""
Raw DELETE SQL for policy_commission (agency_id=5).

Usage (file from --file, positional, or UPLOAD_FILE_PATH in .env):
  python generate_delete_commission_sql.py --file Files
  python generate_delete_commission_sql.py --file "Files/foo.csv"
  python generate_delete_commission_sql.py   # uses UPLOAD_FILE_PATH from .env

Does NOT run DELETE. Optional read-only COUNT preview only.
Output is raw SQL only (.sql): BEGIN / DELETE / ROLLBACK.
"""

from __future__ import annotations

import argparse
import os
import sys
from typing import List, Optional, Set

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
        description="Write raw DELETE SQL markdown for policy_commission"
    )
    parser.add_argument(
        "file_path",
        nargs="?",
        default=None,
        help="Optional positional path (same as --file / UPLOAD_FILE_PATH)",
    )
    parser.add_argument("--file", type=str, default=None, help="Sheet CSV/Excel path or directory")
    parser.add_argument("--env-file", type=str, default=".env")
    parser.add_argument("--agency-id", type=int, default=5)
    parser.add_argument(
        "--out",
        type=str,
        default="docs/delete_commission_agency5.sql",
        help="Output .sql path",
    )
    parser.add_argument(
        "--skip-db-preview",
        action="store_true",
        help="Do not connect for COUNT preview",
    )
    args = parser.parse_args()

    load_env_file(args.env_file, override=True)
    agency_id = int(os.getenv("AGENCY_ID", args.agency_id))

    input_path = args.file or args.file_path or os.getenv("UPLOAD_FILE_PATH")
    if not input_path or not os.path.exists(input_path):
        print("Provide file via --file or UPLOAD_FILE_PATH in .env")
        print(f"Input not found: {input_path!r}")
        return 1

    print(f"Source: {input_path}")
    print(f"Agency ID: {agency_id}")
    print("DELETE is NOT executed by this script.")

    df = load_file(input_path)
    sheet_policies = extract_policy_numbers(df)
    match_numbers = build_match_numbers(sheet_policies)
    print(f"Sheet policies: {len(sheet_policies):,} | Match candidates: {len(match_numbers):,}")

    if not args.skip_db_preview:
        db_mgr = DatabaseManager()
        if db_mgr.connect():
            preview = run_db_preview(db_mgr, agency_id, match_numbers)
            if preview:
                print(f"Commission rows to delete: {preview['commission_count']:,}")
                print(f"Distinct policies with commission: {preview['distinct_pol_with_comm']:,}")
            db_mgr.close()

    abs_out = write_raw_delete_sql_md(
        out_path=args.out,
        agency_id=agency_id,
        match_numbers=match_numbers,
        source_label=os.path.abspath(input_path),
        sheet_policy_count=len(sheet_policies),
    )
    print(f"Wrote: {abs_out}")
    print("Share .sql with senior eng. After DELETE + remaining=0, re-run validator with --force-post-mismatches.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
