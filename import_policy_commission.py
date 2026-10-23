#!/usr/bin/env python3
"""
Policy Commission Data Import Script with Interactive Payload Confirmation & Logging
===================================================================================
Automates uploading policy commission files and creating policy commission setup rules
via the PieQ Backend APIs with step-by-step logging, request timeouts, retries,
resuming failed imports, Keycloak SSO token auto-auth, and interactive payload confirmation.

Usage:
  python import_policy_commission.py --file /Users/nirmalrajaak/Downloads/Commission_Upload_today.csv
  python import_policy_commission.py --resume <IMPORT_UUID> --yes
  python import_policy_commission.py --api-base-url https://prod.api.pieq.ai

Env Config Options (.env):
  - API_BASE_URL (e.g. https://prod.api.pieq.ai)
  - AUTH_TOKEN / BEARER_TOKEN (Bearer Auth Token)
  - AUTH_URL / SSO_TOKEN_URL (e.g. https://prod.auth.pieq.ai)
  - KEYCLOAK_CLIENT_ID / SSO_CLIENT_ID (default: pieq-app)
  - USERNAME / SSO_USERNAME
  - PASSWORD / SSO_PASSWORD
  - X_CLIENT_ID / X_SERVICE_NAME (Optional custom headers)
  - BATCH_SIZE (default: 5)
  - REQUEST_TIMEOUT (default: 30 seconds)
  - UPLOAD_FILE_PATH
  - AUTO_CONFIRM (1 to auto-confirm all payloads)
"""

import os
import sys
import re
import csv
import json
import uuid
import time
import socket
import logging
import getpass
import argparse
import urllib.request
import urllib.parse
import urllib.error
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

# Set default global socket timeout to prevent indefinite hangs
socket.setdefaulttimeout(30.0)


# ---------------------------------------------------------------------------
# Logging Setup
# ---------------------------------------------------------------------------
def setup_logger(verbose: bool = False):
    """Initializes standard logger with clean timestamp formatting."""
    logger = logging.getLogger("PolicyCommissionImport")
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter(
        fmt="[%(asctime)s] [%(levelname)s] %(message)s", datefmt="%H:%M:%S"
    )
    handler.setFormatter(formatter)
    if not logger.handlers:
        logger.addHandler(handler)
    return logger


# ---------------------------------------------------------------------------
# Helper: Load .env File
# ---------------------------------------------------------------------------
def load_env_file(logger, env_file_path=".env"):
    """Loads key-value pairs from a .env file into os.environ if present."""
    if not os.path.exists(env_file_path):
        logger.debug(f"No env file found at {env_file_path}")
        return
    try:
        with open(env_file_path, "r", encoding="utf-8") as f:
            count = 0
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip().strip("'").strip('"')
                    if key and key not in os.environ:
                        os.environ[key] = value
                        count += 1
        logger.info(f"Loaded {count} environment variable(s) from {env_file_path}")
    except Exception as e:
        logger.warning(f"Could not read env file {env_file_path}: {e}")


# ---------------------------------------------------------------------------
# Keycloak SSO Token Auto-Auth
# ---------------------------------------------------------------------------
def fetch_sso_token(
    logger,
    token_url: str,
    client_id: str,
    username: str,
    password: str,
) -> str:
    """Fetches a fresh access token from Keycloak SSO using resource owner password grant."""
    logger.info(f"Requesting fresh SSO access token for user '{username}' from {token_url}...")
    payload_data = urllib.parse.urlencode(
        {
            "grant_type": "password",
            "client_id": client_id,
            "username": username,
            "password": password,
        }
    ).encode("utf-8")

    req = urllib.request.Request(
        token_url,
        data=payload_data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=15.0) as resp:
            res_bytes = resp.read()
            res_json = json.loads(res_bytes.decode("utf-8"))
            token = res_json.get("access_token")
            if token:
                logger.info("Successfully fetched fresh SSO access token!")
                return token
            else:
                logger.error(f"SSO response missing access_token: {res_json}")
                return None
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="ignore")
        logger.error(f"Failed to fetch SSO token (HTTP {e.code}): {err_body}")
        return None
    except Exception as e:
        logger.error(f"Exception fetching SSO token: {e}")
        return None


# ---------------------------------------------------------------------------
# Field Normalization & Parsing Helpers (Matching Frontend Logic)
# ---------------------------------------------------------------------------
def normalize_field_key(key: str) -> str:
    """Normalizes field names by stripping, lowercasing, and turning underscores/dashes to spaces."""
    return re.sub(r"[_\s-]+", " ", str(key).strip().lower())


def read_field(data: dict, aliases: list) -> str:
    """Reads a field value from a row dictionary given a list of candidate alias names."""
    if not data:
        return ""
    norm_map = {normalize_field_key(k): v for k, v in data.items() if v is not None}
    for alias in aliases:
        norm_alias = normalize_field_key(alias)
        if norm_alias in norm_map:
            val = norm_map[norm_alias]
            if val is not None:
                val_str = str(val).strip()
                if val_str and val_str.lower() != "none":
                    return val_str
    return ""


def find_field_value(data: dict, aliases: list) -> dict:
    """Checks for the presence of a field in a dictionary, preserving nulls."""
    if not data:
        return {"found": False, "value": None}
    norm_map = {normalize_field_key(k): v for k, v in data.items()}
    for alias in aliases:
        norm_alias = normalize_field_key(alias)
        if norm_alias in norm_map:
            return {"found": True, "value": norm_map[norm_alias]}
    return {"found": False, "value": None}


def parse_float(val) -> float:
    """Parses numeric float values, stripping currency symbols, percentages, and commas."""
    if val is None:
        return 0.0
    val_str = str(val).strip()
    if not val_str:
        return 0.0
    cleaned = re.sub(r"[^0-9.-]", "", val_str)
    if not cleaned or cleaned == "-":
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def parse_commission_type(val) -> int:
    """
    Converts Commission Type string or number to backend enum int:
    1 = REGULAR / Commission
    2 = RENEWAL
    3 = BONUS
    4 = OVERRIDE
    """
    if val is None or val == "":
        return 1
    try:
        val_int = int(val)
        if val_int in (1, 2, 3, 4):
            return val_int
    except (ValueError, TypeError):
        pass
    norm = re.sub(r"[_\s-]+", " ", str(val).strip().lower())
    if norm in ("regular", "commission", "reg"):
        return 1
    if norm in ("renewal", "ren"):
        return 2
    if norm in ("bonus", "bon"):
        return 3
    if norm in ("override", "over"):
        return 4
    return 1


def parse_payment_frequency(val) -> int:
    """
    Converts Payment Frequency string or number to backend enum int:
    1 = MONTHLY
    2 = QUARTERLY
    3 = HALF_YEARLY
    4 = ANNUAL
    """
    if val is None or val == "":
        return 1
    try:
        val_int = int(val)
        if val_int in (1, 2, 3, 4):
            return val_int
    except (ValueError, TypeError):
        pass
    norm = re.sub(r"[_\s-]+", " ", str(val).strip().lower())
    if norm == "monthly":
        return 1
    if norm == "quarterly":
        return 2
    if norm in ("half yearly", "halfyearly", "semi annual", "semiannual"):
        return 3
    if norm in ("annual", "annually"):
        return 4
    return 1


def parse_payout_method(val) -> int:
    """
    Converts Payout Method string or number to backend enum int:
    1 = FIXED_FEE
    2 = PERCENTAGE
    """
    if val is None or val == "":
        return 1
    try:
        val_int = int(val)
        if val_int in (1, 2):
            return val_int
    except (ValueError, TypeError):
        pass
    norm = re.sub(r"[_\s-]+", " ", str(val).strip().lower())
    if "%" in norm or "percent" in norm or "percentage" in norm:
        return 2
    if "$" in norm or "fixed" in norm or "fee" in norm or "dollar" in norm or "flat" in norm:
        return 1
    return 1


def extract_carrier_and_product(data: dict) -> tuple:
    """Extracts carrier and product name from row data."""
    carrier = read_field(data, ["Carrier", "carrier"])
    product = read_field(
        data,
        [
            "Product Name",
            "product name",
            "product_name",
            "productName",
            "Product",
            "product",
        ],
    )
    return carrier, product


def classify_contract_role(contract_val: str) -> str:
    """Map Contract column to agency | sales_leader | agent | sub_agent."""
    c = str(contract_val or "").strip().lower()
    if c in ("agency", ""):
        return "agency"
    if "sales" in c or "leader" in c:
        return "sales_leader"
    if "sub" in c or re.match(r"^sa\d+$", c):
        return "sub_agent"
    return "agent"


# ---------------------------------------------------------------------------
# Core Payload Builder: Grouping & Transform
# ---------------------------------------------------------------------------
def build_policy_commission_import_payload(group_records: list) -> dict:
    """
    Builds single PolicyCommissionImportPayload from a group of records matching
    (Policy Number + Month From + Month To).
    """
    primary_record = group_records[0]
    data = primary_record.get("data") or {}

    record_uuids = [r["uuid"] for r in group_records if r.get("uuid")]

    policy_number = ""
    product_uuid = None

    for record in group_records:
        r_data = record.get("data") or {}
        r_root = record
        if not policy_number:
            policy_number = (
                read_field(
                    r_data,
                    [
                        "policy_number",
                        "policyNumber",
                        "Policy Number",
                        "policy_uuid",
                        "policyUuid",
                        "Policy",
                        "policy",
                    ],
                )
                or str(
                    r_root.get("policy_number")
                    or r_root.get("policyNumber")
                    or r_root.get("policy_uuid")
                    or r_root.get("policyUuid")
                    or ""
                )
            )
        if not product_uuid:
            p_uuid = read_field(r_data, ["product_uuid", "productUuid"]) or r_root.get(
                "product_uuid"
            )
            if p_uuid:
                product_uuid = str(p_uuid)

    carrier, product = extract_carrier_and_product(data)
    rule_name = f"{carrier or 'Carrier'} - {product or 'Product'} - COMMISSION"

    raw_comm_type = read_field(
        data, ["Commission Type", "commission_type", "commissionType", "Type", "type"]
    )
    commission_type = parse_commission_type(raw_comm_type)

    raw_payment_freq = read_field(
        data,
        [
            "Payment Frequency",
            "payment_frequency",
            "paymentFrequency",
            "Frequency",
            "frequency",
        ],
    )
    payment_frequency = parse_payment_frequency(raw_payment_freq)

    raw_from_month = read_field(
        data,
        [
            "From Month",
            "from_month",
            "fromMonth",
            "from month",
            "Month From",
            "month_from",
            "monthFrom",
            "month from",
        ],
    )
    month_from = int(raw_from_month) if raw_from_month and raw_from_month.isdigit() else 1

    raw_to_month_field = find_field_value(
        data,
        [
            "To Month",
            "to_month",
            "toMonth",
            "to month",
            "Month To",
            "month_to",
            "monthTo",
            "month to",
        ],
    )
    month_to = 12
    if raw_to_month_field["found"]:
        val = raw_to_month_field["value"]
        if val is None or str(val).strip() == "" or str(val).strip().lower() == "null":
            month_to = None
        else:
            try:
                month_to = int(str(val).strip())
            except ValueError:
                month_to = None

    raw_payout_method = read_field(
        data, ["Payout Method", "payout_method", "payoutMethod", "payout method"]
    )
    payout_method = parse_payout_method(raw_payout_method)

    raw_fee_str = read_field(
        data,
        [
            "Carrier Commission",
            "carrier_commission",
            "CarrierCommission",
            "Fee Amount",
            "fee_amount",
            "feeAmount",
            "fee amount",
            "Fee",
            "fee",
        ],
    )
    fee_amount = parse_float(raw_fee_str)

    premium_str = read_field(data, ["Premium", "premium"])
    premium_amount = parse_float(premium_str)

    raw_split_items = []
    for record in group_records:
        r_data = record.get("data") or {}
        sequence_str = read_field(r_data, ["Sequence", "sequence"])
        try:
            sequence = int(sequence_str) if sequence_str else 0
        except ValueError:
            sequence = 0

        agent_id = read_field(
            r_data, ["Agent ID", "agent_id", "agentId", "agent id"]
        ).strip()
        agent_name = (
            read_field(
                r_data, ["Agent Name", "agent_name", "agentName", "agent name"]
            ).strip()
            or agent_id
        )

        contract = read_field(
            r_data, ["Contract", "contract", "Role", "role", "Level", "level"]
        )

        rate_str = read_field(
            r_data, ["Rate %", "Rate $", "Rate", "rate", "Commission %"]
        )
        rate_val = parse_float(rate_str)

        raw_split_items.append(
            {
                "user_name": agent_name,
                "user_code": agent_id if agent_id else None,
                "rateValue": rate_val,
                "sequence": sequence,
                "contract": contract,
            }
        )

    has_contract_info = any(bool(item.get("contract")) for item in raw_split_items)

    def _role_rank(item: dict) -> int:
        role = classify_contract_role(item.get("contract", ""))
        return {"agency": 0, "sales_leader": 1, "agent": 2, "sub_agent": 3}.get(role, 2)

    if has_contract_info:
        raw_split_items.sort(key=lambda x: (_role_rank(x), -x["sequence"]))
    else:
        # Fallback: Sort DESCENDING by Sequence: highest sequence is Agency (order 1)
        raw_split_items.sort(key=lambda x: x["sequence"], reverse=True)

    commission_split = []
    for index, item in enumerate(raw_split_items):
        order = index + 1
        if has_contract_info:
            role_key = classify_contract_role(item.get("contract", ""))
            if role_key == "agency":
                role = "Agency"
                entity = "agency"
            elif role_key == "sales_leader":
                role = "Sales Leader"
                entity = "agent"
            elif role_key == "sub_agent":
                role = "Sub Agent"
                entity = "agent"
            else:
                role = "Agent"
                entity = "agent"
        else:
            if order == 1:
                role = "Agency"
                entity = "agency"
            elif order == 2:
                role = "Sales Leader"
                entity = "agent"
            elif order == 3:
                role = "Agent"
                entity = "agent"
            else:
                role = "Sub Agent"
                entity = "agent"

        level = "" if role == "Agency" else f"Level {index}"

        commission_split.append(
            {
                "order": order,
                "role": role,
                "user_name": item["user_name"],
                "user_uuid": None,
                "user_code": item["user_code"],
                "commission_value": item["rateValue"],
                "entity": entity,
                "level": level,
            }
        )

    payload = {
        "record_uuids": record_uuids,
        "policyNumber": str(policy_number),
        "ruleName": rule_name,
        "version": 1,
        "commissionType": commission_type,
        "paymentFrequency": payment_frequency,
        "monthFrom": month_from,
        "monthTo": month_to,
        "payoutMethod": payout_method,
        "feeAmount": fee_amount,
        "premiumAmount": premium_amount,
        "commissionSplit": commission_split,
    }

    if carrier:
        payload["carrier"] = carrier
    if product:
        payload["product"] = product
    if product_uuid:
        payload["productUuid"] = str(product_uuid)

    return payload


def group_records_to_payloads(logger, records: list) -> list:
    """Groups pending records by policyNumber + monthFrom + monthToKey and generates payload dicts."""
    completed_records = [r for r in records if r.get("status") == 1]
    pending_records = [r for r in records if r.get("status") != 1]

    if completed_records:
        logger.info(f"Skipping {len(completed_records)} record(s) already completed (status=1).")

    if not pending_records:
        logger.info("No pending/unprocessed records found.")
        return []

    logger.info(f"Grouping {len(pending_records)} pending record(s)...")
    records_by_group = defaultdict(list)

    for record in pending_records:
        r_data = record.get("data") or {}
        r_root = record
        policy_number = str(
            read_field(
                r_data,
                [
                    "policy_number",
                    "policyNumber",
                    "Policy Number",
                    "policy_uuid",
                    "policyUuid",
                    "Policy",
                    "policy",
                ],
            )
            or r_root.get("policy_number")
            or r_root.get("policyNumber")
            or r_root.get("policy_uuid")
            or r_root.get("policyUuid")
            or ""
        )

        raw_from_month = read_field(
            r_data,
            [
                "From Month",
                "from_month",
                "fromMonth",
                "from month",
                "Month From",
                "month_from",
                "monthFrom",
                "month from",
            ],
        )
        month_from = int(raw_from_month) if raw_from_month and raw_from_month.isdigit() else 1

        raw_to_month_field = find_field_value(
            r_data,
            [
                "To Month",
                "to_month",
                "toMonth",
                "to month",
                "Month To",
                "month_to",
                "monthTo",
                "month to",
            ],
        )
        month_to_key = "12"
        if raw_to_month_field["found"]:
            val = raw_to_month_field["value"]
            if val is None or str(val).strip() == "" or str(val).strip().lower() == "null":
                month_to_key = "null"
            else:
                try:
                    parsed_to = int(str(val).strip())
                    month_to_key = str(parsed_to)
                except ValueError:
                    month_to_key = "null"

        group_key = f"{policy_number}|{month_from}|{month_to_key}"
        records_by_group[group_key].append(record)

    payloads = []
    for group_list in records_by_group.values():
        payloads.append(build_policy_commission_import_payload(group_list))

    logger.info(f"Grouped into {len(payloads)} unique pending policy commission setup rule(s).")
    return payloads


AGENCY_USER_UUID = "62a97a01-2801-4c7c-9cd0-cf38f2e0a4f4"


def _person_name_key(value) -> str:
    """Lowercase name with collapsed spaces, without a trailing '(Code: ...)' or '(ID)' label."""
    text = str(value or "")
    text = re.sub(r"\s*\([^)]*\)\s*$", "", text)
    key = " ".join(text.casefold().split())
    # This file uses both forms for the same writing agent.
    if key == "greg teipel":
        return "gregory teipel"
    # Same agency row, stored under either display name.
    if key == "mlb parallel run":
        return "main line benefits"
    return key


def _blank_split_value(value) -> bool:
    if value is None:
        return True
    return str(value).strip().lower() in ("", "none", "null", "nan")


def _commission_value_blank(value) -> bool:
    if _blank_split_value(value):
        return True
    try:
        return float(value) == 0.0
    except (TypeError, ValueError):
        return False


def _is_empty_sub_agent(split: dict) -> bool:
    """A sub-agent row with no uuid, name, or commission amount is not a real split."""
    role = str(split.get("role") or "").strip().casefold()
    if role != "sub agent":
        return False
    return (
        _blank_split_value(split.get("user_uuid"))
        and _blank_split_value(split.get("user_name"))
        and _commission_value_blank(split.get("commission_value"))
    )


def commission_payload_issues(payload: dict) -> list:
    """Return split-rule failures. An empty list means the payload may be posted."""
    issues = []
    splits = payload.get("commissionSplit") or []
    if not isinstance(splits, list):
        return ["commissionSplit must be a list"]

    if len(splits) not in (3, 4):
        issues.append(f"Expected 3 or 4 splits, found {len(splits)}")

    for split in splits:
        if not isinstance(split, dict):
            issues.append("A commission split entry is not an object")
            continue
        if _is_empty_sub_agent(split):
            continue
        try:
            order = int(split.get("order"))
        except (TypeError, ValueError):
            order = split.get("order")
        if _blank_split_value(split.get("user_uuid")):
            issues.append(f"Order {order} user_uuid is null")
        if order != 1 and _blank_split_value(split.get("level")):
            issues.append(f"Order {order} level is null")

    order1 = next((split for split in splits if isinstance(split, dict) and str(split.get("order")) == "1"), None)
    if order1 is None:
        issues.append("Order 1 agency split is missing")
    else:
        if str(order1.get("role") or "").strip().casefold() != "agency":
            issues.append(f"Order 1 role must be Agency, found '{order1.get('role')}'")
        if str(order1.get("entity") or "").strip().casefold() != "agency":
            issues.append(f"Order 1 entity must be agency, found '{order1.get('entity')}'")
        if _person_name_key(order1.get("user_name")) != "main line benefits":
            issues.append(f"Order 1 must be Main Line Benefits, found '{order1.get('user_name')}'")
        if str(order1.get("user_uuid") or "").strip().lower() != AGENCY_USER_UUID:
            issues.append(f"Order 1 user_uuid must be {AGENCY_USER_UUID}")

    numeric_splits = [
        split for split in splits
        if isinstance(split, dict) and not _is_empty_sub_agent(split)
    ]
    if numeric_splits:
        def order_num(split):
            try:
                return int(split.get("order"))
            except (TypeError, ValueError):
                return -1

        highest = max(numeric_splits, key=order_num)
        writing_name = _person_name_key(payload.get("writingAgentName"))
        highest_name = _person_name_key(highest.get("user_name"))
        if not writing_name:
            issues.append("Policy writing agent name is missing, so the highest order cannot be checked")
        elif highest_name != writing_name:
            issues.append(
                f"Highest order {highest.get('order')} '{highest.get('user_name')}' does not match "
                f"writing agent '{payload.get('writingAgentName')}'"
            )

    total = 0.0
    for split in numeric_splits:
        try:
            total += float(split.get("commission_value") or 0)
        except (TypeError, ValueError):
            issues.append(f"Order {split.get('order')} commission_value is not a number")
    if round(total, 2) != 100.00:
        issues.append(f"Commission split sum is {round(total, 2):.2f}, expected 100")

    return issues


# ---------------------------------------------------------------------------
# Interactive Payload Confirmation UI
# ---------------------------------------------------------------------------
def confirm_payload(logger, idx: int, total: int, payload: dict, auto_confirm: bool, action_name: str = "POST", diff_text: str = "") -> str:
    """
    Displays payload structure and prompts the user for confirmation.
    Returns: 'yes', 'skip', 'all', or 'quit'
    """
    pol_num = payload.get("policyNumber", "Unknown")
    rule_name = payload.get("ruleName", "N/A")
    splits = payload.get("commissionSplit", [])

    print("\n" + "=" * 78)
    print(f"  [PAYLOAD CONFIRMATION {idx}/{total}] Policy #{pol_num}")
    print("=" * 78)
    print(f" Policy writing agent: {payload.get('writingAgentName') or 'N/A'}")
    print(f" Rule Name:      {rule_name}")
    print(f" CommissionType: {payload.get('commissionType')} (1=REGULAR, 2=RENEWAL, 3=BONUS, 4=OVERRIDE)")
    print(f" PaymentFreq:    {payload.get('paymentFrequency')} (1=MONTHLY, 2=QUARTERLY, 3=HALF_YEARLY, 4=ANNUAL)")
    print(f" PayoutMethod:   {payload.get('payoutMethod')} (1=FIXED_FEE, 2=PERCENTAGE)")
    print(f" Month Range:    {payload.get('monthFrom')} to {payload.get('monthTo')}")
    print(f" Fee Amount:     ${payload.get('feeAmount', 0.0)}")
    print(f" Premium Amount: ${payload.get('premiumAmount', 0.0)}")
    print(f" Record UUIDs ({len(payload.get('record_uuids', []))}): {payload.get('record_uuids', [])[:3]}...")
    print("\n Commission Hierarchy Splits:")
    for s in splits:
        user_info = f"{s.get('user_name', '')} (Code: {s.get('user_code') or 'N/A'})"
        level_str = f" [{s['level']}]" if s.get("level") else ""
        print(f"   Order {s.get('order')} | Role: {s.get('role', 'nan'):<12}{level_str:<10} | Entity: {s.get('entity'):<7} | Rate: {s.get('commission_value')}% | User: {user_info}")

    split_issues = commission_payload_issues(payload)
    print("\n Split validation:")
    if split_issues:
        print("   FAIL")
        for issue in split_issues:
            print(f"   - {issue}")
    else:
        split_total = round(sum(float(s.get("commission_value") or 0) for s in splits), 2)
        print(f"   PASS | {len(splits)} splits | sum {split_total:.2f} | order 1 is Main Line Benefits")

    if diff_text:
        print("\n" + "-" * 78)
        print(" 🔍 WHAT WILL BE CHANGED (DIFF FROM DB):")
        print(diff_text)
        print("-" * 78)

    print("\n Full Payload JSON Structure:")
    print(json.dumps(payload, indent=2))
    print("=" * 78)

    if split_issues:
        logger.error(f"Blocked {action_name} for Policy #{pol_num}: " + "; ".join(split_issues))
        print(f"\n{action_name} blocked. Fix the split validation failures above before this policy can be imported.")
        return "skip"

    if auto_confirm:
        logger.info(f"Auto-confirming payload for Policy #{pol_num} (--yes flag active).")
        return "yes"

    try:
        while True:
            choice = (
                input(f"\nProceed to {action_name} import for this policy? [y=Yes / n=Skip / a=Approve All / q=Quit] ")
                .strip()
                .lower()
            )
            if choice in ("y", "yes", ""):
                return "yes"
            elif choice in ("n", "no", "s", "skip"):
                return "skip"
            elif choice in ("a", "all", "approve"):
                return "all"
            elif choice in ("q", "quit", "exit"):
                return "quit"
            else:
                print("Invalid choice. Please enter 'y' (yes), 'n' (skip), 'a' (approve all), or 'q' (quit).")
    except (EOFError, KeyboardInterrupt):
        print("\nInput cancelled. Aborting.")
        return "quit"


# ---------------------------------------------------------------------------
# API Client (Stdlib HTTP / urllib with Timeout & Retry Support)
# ---------------------------------------------------------------------------
class PieqApiClient:

    def __init__(self, logger, base_url: str, token: str = None, headers: dict = None, timeout: float = 30.0, max_retries: int = 2):
        self.logger = logger
        self.base_url = base_url.rstrip("/")
        self.token = token
        self.default_headers = headers or {}
        self.timeout = timeout
        self.max_retries = max_retries

    def _get_headers(self, content_type="application/json") -> dict:
        headers = dict(self.default_headers)
        if content_type:
            headers["Content-Type"] = content_type
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"

        # Pass custom identity headers from env if available
        if "X_CLIENT_ID" in os.environ:
            headers["x-client-id"] = os.environ["X_CLIENT_ID"]
        if "X_SERVICE_NAME" in os.environ:
            headers["x-service-name"] = os.environ["X_SERVICE_NAME"]

        return headers

    def upload_file(self, file_path: str, entity_type: str = "Policy Commission") -> dict:
        """Uploads CSV file via multipart/form-data POST /api/v1/data-imports/uploads."""
        url = f"{self.base_url}/api/v1/data-imports/uploads"
        self.logger.info(f"Posting file to {url} (entity_type={entity_type}, timeout={self.timeout}s)...")

        boundary = f"----WebKitFormBoundary{uuid.uuid4().hex}"
        filename = os.path.basename(file_path)

        with open(file_path, "rb") as f:
            file_data = f.read()

        body = bytearray()
        body.extend(f"--{boundary}\r\n".encode("utf-8"))
        body.extend(
            f'Content-Disposition: form-data; name="entity_type"\r\n\r\n{entity_type}\r\n'.encode(
                "utf-8"
            )
        )

        body.extend(f"--{boundary}\r\n".encode("utf-8"))
        body.extend(
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'.encode(
                "utf-8"
            )
        )
        body.extend(b"Content-Type: text/csv\r\n\r\n")
        body.extend(file_data)
        body.extend(b"\r\n")
        body.extend(f"--{boundary}--\r\n".encode("utf-8"))

        headers = self._get_headers(
            content_type=f"multipart/form-data; boundary={boundary}"
        )
        req = urllib.request.Request(
            url, data=bytes(body), headers=headers, method="POST"
        )

        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                res_bytes = resp.read()
                res_json = json.loads(res_bytes.decode("utf-8"))
                if isinstance(res_json, dict) and "data" in res_json:
                    return res_json["data"]
                return res_json
        except urllib.error.HTTPError as e:
            err_body = e.read().decode("utf-8", errors="ignore")
            self.logger.error(f"File upload failed (HTTP {e.code}): {err_body}")
            raise Exception(f"File upload failed (HTTP {e.code}): {err_body}")
        except (socket.timeout, urllib.error.URLError) as e:
            self.logger.error(f"File upload connection failed / timed out: {e}")
            raise Exception(f"File upload network error: {e}")

    def fetch_records(self, import_uuid: str) -> list:
        """Fetches import records GET /api/v1/data-imports/uploads/{import_uuid}."""
        url = f"{self.base_url}/api/v1/data-imports/uploads/{import_uuid}"
        self.logger.info(f"Querying import records for UUID {import_uuid} from {url}...")
        req = urllib.request.Request(
            url, headers=self._get_headers(), method="GET"
        )

        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                res_bytes = resp.read()
                res_json = json.loads(res_bytes.decode("utf-8"))
                if isinstance(res_json, dict) and "data" in res_json:
                    return res_json["data"]
                return res_json if isinstance(res_json, list) else [res_json]
        except urllib.error.HTTPError as e:
            err_body = e.read().decode("utf-8", errors="ignore")
            self.logger.error(f"Failed to fetch import records (HTTP {e.code}): {err_body}")
            raise Exception(
                f"Failed to fetch import records (HTTP {e.code}): {err_body}"
            )
        except (socket.timeout, urllib.error.URLError) as e:
            self.logger.error(f"Fetch import records timed out / network error: {e}")
            raise Exception(f"Fetch import records network error: {e}")

    def create_policy_commission(self, payload: dict) -> dict:
        """Posts single policy commission payload to POST /api/v1/policy/commission/data-import with retries and timeout."""
        url = f"{self.base_url}/api/v1/policy/commission/data-import"
        pol_num = payload.get("policyNumber", "Unknown")
        json_data = json.dumps(payload).encode("utf-8")

        last_error = None
        for attempt in range(1, self.max_retries + 2):
            req = urllib.request.Request(
                url, data=json_data, headers=self._get_headers(), method="POST"
            )
            try:
                self.logger.debug(f"Attempt {attempt}: Posting payload for Policy #{pol_num} to {url} (timeout={self.timeout}s)...")
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    res_bytes = resp.read()
                    res_json = json.loads(res_bytes.decode("utf-8"))
                    return res_json
            except urllib.error.HTTPError as e:
                err_body = e.read().decode("utf-8", errors="ignore")
                last_error = f"HTTP {e.code}: {err_body}"
                self.logger.warning(f"Attempt {attempt} failed for Policy #{pol_num}: {last_error}")
                if e.code < 500:
                    # Do not retry client 4xx errors
                    break
            except (socket.timeout, urllib.error.URLError) as e:
                last_error = f"Network timeout/error ({e})"
                self.logger.warning(f"Attempt {attempt} timed out for Policy #{pol_num}: {last_error}")
            except Exception as e:
                last_error = str(e)
                self.logger.warning(f"Attempt {attempt} exception for Policy #{pol_num}: {last_error}")

            if attempt <= self.max_retries:
                time.sleep(1.0 * attempt)

        raise Exception(last_error or "Unknown HTTP failure")


# ---------------------------------------------------------------------------
# Offline CSV Reader (for Dry-Run / Offline Mode)
# ---------------------------------------------------------------------------
def load_records_from_csv(logger, file_path: str) -> list:
    """Reads a CSV file directly into mock DataImportRecord objects for offline processing / dry-run."""
    logger.info(f"Reading CSV file locally: {file_path}")
    records = []
    with open(file_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader):
            clean_row = {k.strip(): (v.strip() if v else "") for k, v in row.items() if k}
            records.append(
                {
                    "uuid": f"mock-record-{i+1}",
                    "dataImportUuid": "mock-import-uuid",
                    "data": clean_row,
                    "status": 0,
                    "created_at": time.strftime("%Y-%m-%dT%H:%M:%SZ"),
                    "updated_at": time.strftime("%Y-%m-%dT%H:%M:%SZ"),
                }
            )
    logger.info(f"Successfully read {len(records)} row records from CSV.")
    return records


# ---------------------------------------------------------------------------
# Main Execution Pipeline
# ---------------------------------------------------------------------------
def main():
    logger = setup_logger(verbose=False)

    # 1. ALWAYS load .env file FIRST before constructing argparse options
    env_file = ".env"
    for i, a in enumerate(sys.argv):
        if a in ("-e", "--env-file") and i + 1 < len(sys.argv):
            env_file = sys.argv[i + 1]

    load_env_file(logger, env_file)

    parser = argparse.ArgumentParser(
        description="Run Policy Commission Data Import from terminal with env configs and confirmation."
    )
    parser.add_argument(
        "-f",
        "--file",
        default=os.environ.get("UPLOAD_FILE_PATH"),
        help="Path to commission upload CSV file",
    )
    parser.add_argument(
        "-r",
        "--resume",
        default=os.environ.get("RESUME_IMPORT_UUID"),
        help="Resume processing for an existing import UUID (skips uploading step)",
    )
    parser.add_argument(
        "-e",
        "--env-file",
        default=".env",
        help="Path to .env configuration file (default: .env)",
    )
    parser.add_argument(
        "-u",
        "--api-base-url",
        default=None,
        help="Target API base URL (overrides .env API_BASE_URL)",
    )
    parser.add_argument(
        "-t",
        "--token",
        default=None,
        help="Bearer Auth Token (overrides .env AUTH_TOKEN)",
    )
    parser.add_argument(
        "-U",
        "--username",
        default=None,
        help="SSO Username for Keycloak authentication",
    )
    parser.add_argument(
        "-P",
        "--password",
        default=None,
        help="SSO Password for Keycloak authentication",
    )
    parser.add_argument(
        "--sso-url",
        default=None,
        help="Keycloak SSO token URL (overrides .env AUTH_URL / SSO_TOKEN_URL)",
    )
    parser.add_argument(
        "--sso-client-id",
        default=None,
        help="Keycloak SSO Client ID",
    )
    parser.add_argument(
        "-b",
        "--batch-size",
        type=int,
        default=None,
        help="Parallel worker batch size (default: 5)",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=None,
        help="Request timeout in seconds per API call (default: 30.0)",
    )
    parser.add_argument(
        "-y",
        "--yes",
        action="store_true",
        default=os.environ.get("AUTO_CONFIRM") == "1",
        help="Automatically confirm all policy payloads without interactive prompt",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Parse file and generate payloads without calling remote API endpoints",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Print verbose detailed outputs"
    )

    args = parser.parse_args()

    # Re-evaluate logger verbosity
    if args.verbose:
        logger.setLevel(logging.DEBUG)

    # 2. Strict prioritization: CLI argument > .env environment variable > Default fallback
    api_base_url = (
        args.api_base_url
        or os.environ.get("API_BASE_URL")
        or os.environ.get("API_URL")
        or os.environ.get("BASE_URL")
        or "https://prod.api.pieq.ai"
    )

    token = (
        args.token
        or os.environ.get("AUTH_TOKEN")
        or os.environ.get("BEARER_TOKEN")
    )

    file_path = args.file or "/Users/nirmalrajaak/Downloads/Commission_Upload_today.csv"
    resume_uuid = args.resume or os.environ.get("RESUME_IMPORT_UUID")
    auto_confirm = args.yes or (os.environ.get("AUTO_CONFIRM") == "1")

    batch_size = args.batch_size or int(os.environ.get("BATCH_SIZE", "5"))
    request_timeout = args.timeout or float(os.environ.get("REQUEST_TIMEOUT", "30.0"))

    # Set socket default timeout according to final timeout
    socket.setdefaulttimeout(request_timeout)

    sso_url_raw = (
        args.sso_url
        or os.environ.get("AUTH_URL")
        or os.environ.get("SSO_TOKEN_URL")
        or os.environ.get("SSO_URL")
    )
    if sso_url_raw:
        if not sso_url_raw.endswith("/token"):
            sso_url = f"{sso_url_raw.rstrip('/')}/realms/pieq-sso/protocol/openid-connect/token"
        else:
            sso_url = sso_url_raw
    else:
        sso_url = "https://prod.auth.pieq.ai/realms/pieq-sso/protocol/openid-connect/token"

    sso_client_id = (
        args.sso_client_id
        or os.environ.get("KEYCLOAK_CLIENT_ID")
        or os.environ.get("SSO_CLIENT_ID")
        or "pieq-app"
    )
    sso_username = args.username or os.environ.get("USERNAME") or os.environ.get("SSO_USERNAME")
    sso_password = args.password or os.environ.get("PASSWORD") or os.environ.get("SSO_PASSWORD")

    # If no token provided or token expired/requested, attempt SSO login if password available
    if not token and sso_username and sso_password:
        token = fetch_sso_token(logger, sso_url, sso_client_id, sso_username, sso_password)

    logger.info("=========================================================")
    logger.info("        POLICY COMMISSION DATA IMPORT RUNNER             ")
    logger.info("=========================================================")
    logger.info(f"Target API Base URL:  {api_base_url}")
    logger.info(f"SSO Auth URL:         {sso_url}")
    logger.info(f"Upload File:          {file_path if not resume_uuid else '(N/A - Resuming)'}")
    logger.info(f"Resume Import UUID:   {resume_uuid or 'No'}")
    logger.info(f"Dry Run Mode:         {args.dry_run}")
    logger.info(f"Auto-Confirm (--yes): {auto_confirm}")
    logger.info(f"Batch Size:           {batch_size}")
    logger.info(f"Request Timeout:      {request_timeout}s")
    logger.info(f"Auth Token Provided:  {'Yes' if token else 'No'}")
    logger.info("=========================================================\n")

    if not resume_uuid and not os.path.exists(file_path):
        logger.error(f"Input file not found: {file_path}")
        sys.exit(1)

    # If Dry Run Mode
    if args.dry_run:
        logger.info("[DRY-RUN MODE] Parsing CSV file locally...")
        records = load_records_from_csv(logger, file_path)
        payloads = group_records_to_payloads(logger, records)

        print(f"\n[DRY-RUN] Generated {len(payloads)} policy commission rule setup(s).\n")
        auto_confirm_dr = auto_confirm

        for i, payload in enumerate(payloads, 1):
            action = confirm_payload(logger, i, len(payloads), payload, auto_confirm=auto_confirm_dr)
            if action == "all":
                auto_confirm_dr = True
            elif action == "quit":
                logger.info("Dry-run preview aborted by user.")
                break
            elif action == "skip":
                logger.info(f"Skipped preview for Policy #{payload.get('policyNumber')}")

        logger.info("\n[DRY-RUN] Dry run completed. No remote API requests were sent.")
        return

    # Live API Execution Workflow
    client = PieqApiClient(
        logger=logger,
        base_url=api_base_url,
        token=token,
        timeout=request_timeout,
    )

    import_uuid = resume_uuid

    if not import_uuid:
        logger.info("[STEP 1/3] Uploading file to backend API...")
        try:
            upload_res = client.upload_file(file_path=file_path, entity_type="Policy Commission")
            import_uuid = upload_res.get("uuid")
            record_count = upload_res.get("record_count", 0)
            logger.info(f"File uploaded successfully! Import UUID: {import_uuid} (Records: {record_count})")
        except Exception as e:
            logger.error(f"File upload failed: {e}")
            sys.exit(1)
    else:
        logger.info(f"[STEP 1/3 SKIPPED] Resuming execution for existing Import UUID: {import_uuid}")

    logger.info("\n[STEP 2/3] Fetching imported record details from server...")
    records = None
    try:
        records = client.fetch_records(import_uuid)
        logger.info(f"Fetched {len(records)} total record(s) for import {import_uuid}.")
    except Exception as e:
        err_str = str(e)
        if "401" in err_str or "Invalid token" in err_str or "Unauthorized" in err_str:
            logger.warning("Token expired or unauthorized (HTTP 401). Attempting fresh SSO login...")
            fresh_username = sso_username or input("Enter SSO Username (e.g. nirmal.k+5@pieq.ai): ").strip()
            fresh_password = sso_password or getpass.getpass("Enter SSO Password: ").strip()
            if fresh_username and fresh_password:
                fresh_token = fetch_sso_token(logger, sso_url, sso_client_id, fresh_username, fresh_password)
                if fresh_token:
                    client.token = fresh_token
                    try:
                        records = client.fetch_records(import_uuid)
                        logger.info(f"Retry successful! Fetched {len(records)} total record(s).")
                    except Exception as retry_err:
                        logger.error(f"Failed to fetch import records after token refresh: {retry_err}")
                        sys.exit(1)
                else:
                    logger.error("Could not obtain fresh SSO token. Please provide valid credentials.")
                    sys.exit(1)
            else:
                logger.error("No SSO credentials provided.")
                sys.exit(1)
        else:
            logger.error(f"Failed to fetch import records: {e}")
            sys.exit(1)

    logger.info("\n[STEP 3/3] Grouping records & Filtering pending/unprocessed payloads...")
    payloads = group_records_to_payloads(logger, records)
    total_payloads = len(payloads)

    if total_payloads == 0:
        logger.info("All records for this import have already been successfully processed (status=1). Nothing to submit.")
        return

    logger.info(f"Grouped into {total_payloads} pending policy commission setup rule(s). Starting confirmation flow...")

    approved_payloads = []
    skipped_count = 0
    auto_confirm_run = auto_confirm

    for i, payload in enumerate(payloads, 1):
        action = confirm_payload(logger, i, total_payloads, payload, auto_confirm=auto_confirm_run)
        if action == "all":
            auto_confirm_run = True
            approved_payloads.append(payload)
        elif action == "yes":
            approved_payloads.append(payload)
        elif action == "skip":
            skipped_count += 1
            logger.warning(f"Skipping import for Policy #{payload.get('policyNumber')}")
        elif action == "quit":
            logger.info("User requested to abort execution. Submitting only approved payloads so far...")
            break

    if not approved_payloads:
        logger.info("No policy payloads approved for submission. Exiting.")
        return

    logger.info(f"\nSubmitting {len(approved_payloads)} approved policy commission setup(s) to API (Batch Size: {batch_size}, Timeout: {request_timeout}s)...")

    success_count = 0
    failure_count = 0
    errors_list = []
    completed_count = 0
    total_approved = len(approved_payloads)

    def submit_payload(idx, payload):
        pol_num = payload["policyNumber"]
        try:
            logger.info(f"Sending POST payload for Batch #{idx:2d}/{total_approved} | Policy #{pol_num}...")
            res = client.create_policy_commission(payload)
            return (idx, True, pol_num, res)
        except Exception as err:
            return (idx, False, pol_num, str(err))

    with ThreadPoolExecutor(max_workers=batch_size) as executor:
        futures = [
            executor.submit(submit_payload, i, payload)
            for i, payload in enumerate(approved_payloads, 1)
        ]

        for future in as_completed(futures):
            completed_count += 1
            idx, ok, pol_num, info = future.result()
            if ok:
                success_count += 1
                logger.info(f" -> [SUCCESS {completed_count}/{total_approved}] Batch #{idx:2d} | Policy #{pol_num} commission rule created successfully.")
            else:
                failure_count += 1
                errors_list.append((pol_num, info))
                logger.error(f" -> [FAILED {completed_count}/{total_approved}]  Batch #{idx:2d} | Policy #{pol_num} - Error: {info}")

    logger.info("\n=========================================================")
    logger.info("                    IMPORT SUMMARY                       ")
    logger.info("=========================================================")
    logger.info(f"Total Rules Formatted:  {total_payloads}")
    logger.info(f"Approved for Import:    {len(approved_payloads)}")
    logger.info(f"Skipped by User:        {skipped_count}")
    logger.info(f"Successfully Created:   {success_count}")
    logger.info(f"Failed / Timed Out:     {failure_count}")

    if errors_list:
        logger.error("\nDetailed Failure List:")
        for pol_num, err in errors_list:
            logger.error(f" - Policy #{pol_num}: {err}")
    print("\n")
    logger.info("=========================================================")


if __name__ == "__main__":
    main()
