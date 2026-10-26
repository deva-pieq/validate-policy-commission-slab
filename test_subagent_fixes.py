#!/usr/bin/env python3
"""
Automated Regression Test Suite for APL Commission Validation & Import Engine
Targeted Fixes for Sub-Agent Commission Sequence & Value Preservation
"""

import sys
import copy
import unittest
import pandas as pd

# Ensure UTF-8 output encoding on Windows terminals
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from validate_policy_commission_db import (
    classify_contract_role,
    _parse_split_people,
    ensure_override_three_rows,
    build_file_commission_split,
    fix_inclusive_month_band_overlaps,
    normalize_year_bands,
    normalize_level,
    stage2_db_verify_dataset,
)

from import_policy_commission import (
    build_policy_commission_import_payload,
    classify_contract_role as import_classify_contract_role,
)


class MockDBManager:
    """Mock DatabaseManager for isolated unit testing of Stage 2 logic."""
    def __init__(self, db_rules_map=None, users_map=None):
        self.conn = True
        self.db_rules_map = db_rules_map or {}
        self.users_map = users_map or {}

    def fetch_all_agency_commission_rules(self, agency_id, policy_numbers=None):
        return self.db_rules_map

    def fetch_all_users_map(self, agency_id):
        return self.users_map

    def format_writing_agent(self, user_id):
        return "Adam Agent"

    def batch_fetch_policy_commission_rules(self, agency_id, pol_list, _):
        return {p: self.db_rules_map.get(p, []) for p in pol_list if p in self.db_rules_map}


class TestSubAgentCommissionFixes(unittest.TestCase):
    """Test suite verifying all sub-agent sequence, role, and DB preservation fixes."""

    # -------------------------------------------------------------------------
    # 1. Participant Hierarchy Tests
    # -------------------------------------------------------------------------
    def test_standard_three_person_hierarchy(self):
        """Standard 3-person: Agency -> Sales Leader -> Agent."""
        df = pd.DataFrame([
            {"Agent ID": "AGY01", "Agent Name": "Main Line Benefits", "Contract": "Agency", "Sequence": 10, "Rate %": 100.0},
            {"Agent ID": "SL01", "Agent Name": "Sarah Leader", "Contract": "SALES LEADER", "Sequence": 7, "Rate %": 10.0},
            {"Agent ID": "AGT01", "Agent Name": "Adam Agent", "Contract": "LVL 1", "Sequence": 4, "Rate %": 85.0},
        ])
        split = build_file_commission_split(
            df,
            seq_col="Sequence",
            agent_col="Agent ID",
            agent_name_col="Agent Name",
            contract_col="Contract",
            split_col="Rate %",
        )
        self.assertEqual(len(split), 3)
        self.assertEqual([s["role"] for s in split], ["Agency", "Sales Leader", "Agent"])
        self.assertEqual([s["order"] for s in split], [1, 2, 3])
        self.assertEqual(split[0]["entity"], "agency")
        self.assertEqual(split[1]["entity"], "agent")
        self.assertEqual(split[2]["entity"], "agent")
        # Writing agent should be order 3 (Agent)
        self.assertEqual(split[2]["user_code"], "AGT01")

    def test_standard_four_person_hierarchy(self):
        """Standard 4-person: Agency -> Sales Leader -> Agent -> Sub-Agent."""
        df = pd.DataFrame([
            {"Agent ID": "AGY01", "Agent Name": "Main Line Benefits", "Contract": "Agency", "Sequence": 10, "Rate %": 100.0},
            {"Agent ID": "SL01", "Agent Name": "Sarah Leader", "Contract": "SALES LEADER", "Sequence": 7, "Rate %": 10.0},
            {"Agent ID": "AGT01", "Agent Name": "Adam Agent", "Contract": "LVL 1", "Sequence": 4, "Rate %": 85.0},
            {"Agent ID": "SUB01", "Agent Name": "Sam SubAgent", "Contract": "SA1", "Sequence": 1, "Rate %": 80.0},
        ])
        split = build_file_commission_split(
            df,
            seq_col="Sequence",
            agent_col="Agent ID",
            agent_name_col="Agent Name",
            contract_col="Contract",
            split_col="Rate %",
        )
        self.assertEqual(len(split), 4)
        self.assertEqual([s["role"] for s in split], ["Agency", "Sales Leader", "Agent", "Sub Agent"])
        self.assertEqual([s["order"] for s in split], [1, 2, 3, 4])
        # Writing agent must be Sub Agent (order 4)
        self.assertEqual(split[3]["role"], "Sub Agent")
        self.assertEqual(split[3]["user_code"], "SUB01")
        self.assertEqual(split[3]["order"], 4)

    def test_subagent_present_without_sales_leader(self):
        """Case 3: Sub-Agent present without Sales Leader (Agency -> Agent -> Sub-Agent)."""
        df = pd.DataFrame([
            {"Agent ID": "AGY01", "Agent Name": "Main Line Benefits", "Contract": "Agency", "Sequence": 10, "Rate %": 100.0},
            {"Agent ID": "AGT01", "Agent Name": "Adam Agent", "Contract": "LVL 1", "Sequence": 4, "Rate %": 85.0},
            {"Agent ID": "SUB01", "Agent Name": "Sam SubAgent", "Contract": "SUBAGT1", "Sequence": 1, "Rate %": 80.0},
        ])
        split = build_file_commission_split(
            df,
            seq_col="Sequence",
            agent_col="Agent ID",
            agent_name_col="Agent Name",
            contract_col="Contract",
            split_col="Rate %",
        )
        self.assertEqual(len(split), 3)
        # Roles must reflect actual identity: Agency, Agent, Sub Agent (NOT Sales Leader!)
        self.assertEqual([s["role"] for s in split], ["Agency", "Agent", "Sub Agent"])
        self.assertEqual([s["order"] for s in split], [1, 2, 3])
        self.assertEqual(split[1]["role"], "Agent")
        self.assertEqual(split[2]["role"], "Sub Agent")
        self.assertEqual(split[2]["user_code"], "SUB01")

    def test_same_person_occupying_two_roles(self):
        """Same person as Sales Leader and Agent without role collapsing."""
        df = pd.DataFrame([
            {"Agent ID": "AGY01", "Agent Name": "MLB", "Contract": "Agency", "Sequence": 10, "Rate %": 100.0},
            {"Agent ID": "JASON01", "Agent Name": "Jason Hoffmann", "Contract": "SALES LEADER", "Sequence": 7, "Rate %": 5.0},
            {"Agent ID": "JASON01", "Agent Name": "Jason Hoffmann", "Contract": "LVL 1", "Sequence": 4, "Rate %": 90.0},
            {"Agent ID": "SUB01", "Agent Name": "Writing SubAgent", "Contract": "SA1", "Sequence": 1, "Rate %": 80.0},
        ])
        split = build_file_commission_split(
            df,
            seq_col="Sequence",
            agent_col="Agent ID",
            agent_name_col="Agent Name",
            contract_col="Contract",
            split_col="Rate %",
        )
        self.assertEqual(len(split), 4)
        roles = [s["role"] for s in split]
        self.assertEqual(roles, ["Agency", "Sales Leader", "Agent", "Sub Agent"])
        self.assertEqual(split[1]["user_code"], "JASON01")
        self.assertEqual(split[2]["user_code"], "JASON01")
        self.assertEqual(split[1]["role"], "Sales Leader")
        self.assertEqual(split[2]["role"], "Agent")

    def test_non_contiguous_and_duplicate_sequences(self):
        """Input sequence numbers with gaps and duplicates must map to contiguous 1..n."""
        df = pd.DataFrame([
            {"Agent ID": "AGY01", "Agent Name": "Agency", "Contract": "Agency", "Sequence": 99, "Rate %": 100.0},
            {"Agent ID": "SL01", "Agent Name": "Leader", "Contract": "SALES LEADER", "Sequence": 50, "Rate %": 10.0},
            {"Agent ID": "AGT01", "Agent Name": "Agent", "Contract": "GA", "Sequence": 50, "Rate %": 85.0},
            {"Agent ID": "SUB01", "Agent Name": "SubAgent", "Contract": "SA1", "Sequence": 1, "Rate %": 80.0},
        ])
        split = build_file_commission_split(
            df,
            seq_col="Sequence",
            agent_col="Agent ID",
            agent_name_col="Agent Name",
            contract_col="Contract",
            split_col="Rate %",
        )
        orders = [s["order"] for s in split]
        self.assertEqual(orders, [1, 2, 3, 4])
        self.assertEqual(split[0]["role"], "Agency")
        self.assertEqual(split[3]["role"], "Sub Agent")

    # -------------------------------------------------------------------------
    # 2. Month-Band Behavior Regression
    # -------------------------------------------------------------------------
    def test_a03725273_subagent_month_bands(self):
        """A03725273: sub-agent 13-0 across stepped renewal bands must not overlap."""
        df = pd.DataFrame([
            {"Policy Number": "A03725273", "Agent ID": "SUB01", "From Month": 1, "To Month": 12, "Contract": "SA1", "Rate %": "80.0%", "Sequence": 1},
            {"Policy Number": "A03725273", "Agent ID": "SUB01", "From Month": 13, "To Month": 0, "Contract": "SA1", "Rate %": "10.0%", "Sequence": 1},
        ])
        fixed_df, stats = fix_inclusive_month_band_overlaps(df)
        self.assertEqual(len(fixed_df), 2)
        from_months = list(fixed_df["From Month"])
        self.assertEqual(from_months, [1, 13])

    def test_a03814850_inclusive_endpoint_correction(self):
        """A03814850: inclusive endpoint correction across multi-year bands."""
        df = pd.DataFrame([
            {"Policy Number": "A03814850", "Agent ID": "AGT01", "From Month": 1, "To Month": 12, "Contract": "LVL 1", "Rate %": "80.0%"},
            {"Policy Number": "A03814850", "Agent ID": "AGT01", "From Month": 13, "To Month": 24, "Contract": "LVL 1", "Rate %": "10.0%"},
            {"Policy Number": "A03814850", "Agent ID": "AGT01", "From Month": 25, "To Month": 36, "Contract": "LVL 1", "Rate %": "5.0%"},
            {"Policy Number": "A03814850", "Agent ID": "AGT01", "From Month": 37, "To Month": 0, "Contract": "LVL 1", "Rate %": "2.0%"},
        ])
        fixed_df, stats = fix_inclusive_month_band_overlaps(df)
        self.assertEqual(len(fixed_df), 4)
        self.assertEqual(list(fixed_df["From Month"]), [1, 13, 25, 37])

    def test_ahc6743744_symmetrical_four_person_baseline(self):
        """AHC6743744: symmetrical 4-person baseline preserves all participants."""
        df = pd.DataFrame([
            {"Agent ID": "62a97a01", "Agent Name": "Agency", "Contract": "Agency", "Sequence": 4, "Rate %": 100.0},
            {"Agent ID": "SL_CODE", "Agent Name": "Leader", "Contract": "SALES LEADER", "Sequence": 3, "Rate %": 7.0},
            {"Agent ID": "AGT_CODE", "Agent Name": "Writing Agent", "Contract": "LVL 1", "Sequence": 2, "Rate %": 92.0},
            {"Agent ID": "SUB_CODE", "Agent Name": "Sub Agent", "Contract": "SA1", "Sequence": 1, "Rate %": 0.0},
        ])
        split = build_file_commission_split(
            df,
            seq_col="Sequence",
            agent_col="Agent ID",
            agent_name_col="Agent Name",
            contract_col="Contract",
            split_col="Rate %",
        )
        self.assertEqual(len(split), 4)
        self.assertEqual([s["role"] for s in split], ["Agency", "Sales Leader", "Agent", "Sub Agent"])
        self.assertEqual([s["order"] for s in split], [1, 2, 3, 4])

    # -------------------------------------------------------------------------
    # 3. Existing-Policy Updates & Rule C Value Preservation
    # -------------------------------------------------------------------------
    def test_rule_c_preserves_db_values_in_mismatch_payload(self):
        """Existing-policy updates must preserve DB commission rates, feeAmount, and payoutMethod."""
        pol_num = "TEST_POL_01"
        prod_name = "Dental Select"
        clean_prod = "dental select"

        db_splits = [
            {"order": 1, "role": "Agency", "user_uuid": "62a97a01-2801-4c7c-9cd0-cf38f2e0a4f4", "commission_value": 0.0, "entity": "agency"},
            {"order": 2, "role": "Sales Leader", "user_uuid": "UUID_SL", "commission_value": 7.0, "entity": "agent", "user_name": "Sarah Leader", "user_code": "SL01"},
            {"order": 3, "role": "Agent", "user_uuid": "UUID_AGT", "commission_value": 93.0, "entity": "agent", "user_name": "Adam Agent", "user_code": "AGT01"},
        ]

        db_rec = {
            "policy_number": pol_num,
            "product_name": prod_name,
            "policy_uuid": "POL_UUID_01",
            "product_uuid": "PROD_UUID_01",
            "commission_uuid": "COMM_UUID_01",
            "rule_name": f"{prod_name} - COMMISSION",
            "version": 1,
            "commission_type": 1,
            "payment_frequency": 1,
            "month_from": 1,
            "month_to": None,
            "fee_amount": 15.50,
            "payout_method": 2,  # PERCENTAGE
            "premium_amount": 100.0,
            "commission_split": db_splits,
            "writing_agent_user_id": "UUID_AGT",
        }

        mock_db = MockDBManager(
            db_rules_map={
                (pol_num, clean_prod): [db_rec],
                pol_num: [db_rec],
            },
            users_map={
                "sl01": "UUID_SL",
                "agt01": "UUID_AGT",
                "sub01": "UUID_SUB",
                "sarah leader": "UUID_SL",
                "adam agent": "UUID_AGT",
                "sam subagent": "UUID_SUB",
            }
        )

        # File rows: introduces a new Sub Agent (Order 4) and has different split shares/fee in file
        df_valid = pd.DataFrame([
            {"Policy Number": pol_num, "Normalized_Product_Name": prod_name, "From Month": 1, "To Month": 0, "Sequence": 10, "Agent ID": "AGY01", "Agent Name": "Main Line Benefits", "Contract": "Agency", "Rate %": 0.0, "Normalized_Split_Share_%": 0.0, "Payout Method": "Percentage", "Carrier Commission": "100%"},
            {"Policy Number": pol_num, "Normalized_Product_Name": prod_name, "From Month": 1, "To Month": 0, "Sequence": 7, "Agent ID": "SL01", "Agent Name": "Sarah Leader", "Contract": "SALES LEADER", "Rate %": 10.0, "Normalized_Split_Share_%": 10.0, "Payout Method": "Percentage", "Carrier Commission": "100%"},
            {"Policy Number": pol_num, "Normalized_Product_Name": prod_name, "From Month": 1, "To Month": 0, "Sequence": 4, "Agent ID": "AGT01", "Agent Name": "Adam Agent", "Contract": "LVL 1", "Rate %": 85.0, "Normalized_Split_Share_%": 85.0, "Payout Method": "Percentage", "Carrier Commission": "100%"},
            {"Policy Number": pol_num, "Normalized_Product_Name": prod_name, "From Month": 1, "To Month": 0, "Sequence": 1, "Agent ID": "SUB01", "Agent Name": "Sam SubAgent", "Contract": "SA1", "Rate %": 5.0, "Normalized_Split_Share_%": 5.0, "Payout Method": "Percentage", "Carrier Commission": "100%"},
        ])

        results = stage2_db_verify_dataset(df_valid, mock_db, agency_id=5, dry_run=False)
        # Returns: (df_matched, df_ambiguous, df_unmatched, missing_payloads, mismatch_payloads, invalid_payloads, wa_corrections)
        mismatch_payloads = results[4]
        self.assertEqual(len(mismatch_payloads), 1)
        m_payload = mismatch_payloads[0]

        # Rule C: feeAmount and payoutMethod preserved from DB
        self.assertEqual(m_payload["feeAmount"], 15.50)
        self.assertEqual(m_payload["payoutMethod"], 2)

        # Rule C: Existing participants preserve their authoritative DB commission values and UUIDs
        m_splits = {s["order"]: s for s in m_payload["commissionSplit"]}
        self.assertEqual(m_splits[1]["commission_value"], 0.0)
        self.assertEqual(m_splits[2]["commission_value"], 7.0)   # DB 7.0 preserved, NOT File 10.0
        self.assertEqual(m_splits[2]["user_uuid"], "UUID_SL")

        self.assertEqual(m_splits[3]["commission_value"], 93.0)  # DB 93.0 preserved, NOT File 85.0
        self.assertEqual(m_splits[3]["user_uuid"], "UUID_AGT")

        # Newly introduced Sub-Agent gets 0.0% so total remains 100% (7% + 93% + 0% = 100%)
        self.assertEqual(m_splits[4]["role"], "Sub Agent")
        self.assertEqual(m_splits[4]["commission_value"], 0.0)
        self.assertEqual(m_splits[4]["user_code"], "SUB01")

    # -------------------------------------------------------------------------
    # 4. Override Reconstruction & 199% Defect Fix
    # -------------------------------------------------------------------------
    def test_parse_split_people_preserves_subagent(self):
        """_parse_split_people must distinctly extract sub_agent without collapsing into agent."""
        db_splits = [
            {"order": 1, "role": "Agency", "entity": "agency", "user_uuid": "U1"},
            {"order": 2, "role": "Sales Leader", "entity": "agent", "user_name": "Leader", "user_code": "SL01"},
            {"order": 3, "role": "Agent", "entity": "agent", "user_name": "Agent", "user_code": "AGT01"},
            {"order": 4, "role": "Sub Agent", "entity": "agent", "user_name": "SubAgent", "user_code": "SUB01"},
        ]
        people = _parse_split_people(db_splits)
        self.assertIn("sub_agent", people)
        self.assertEqual(people["sub_agent"]["name"], "SubAgent")
        self.assertEqual(people["sub_agent"]["code"], "SUB01")

        self.assertIn("agent", people)
        self.assertEqual(people["agent"]["name"], "Agent")
        self.assertEqual(people["agent"]["code"], "AGT01")

        self.assertIn("sales_leader", people)
        self.assertEqual(people["sales_leader"]["name"], "Leader")

    def test_ensure_override_preserves_subagent_four_rows(self):
        """ensure_override_three_rows must produce 4 rows if sub-agent exists in DB info."""
        pol_num = "POL_OVR_01"
        prod_name = "Health OVR"

        df = pd.DataFrame([
            {"Policy Number": pol_num, "Product Name": prod_name, "From Month": 1, "To Month": 12, "Rate %": "100.0%", "Contract": "Agency", "Sequence": 10, "Agent ID": "AGY01", "Agent Name": "Main Line Benefits"}
        ])

        db_splits = [
            {"order": 1, "role": "Agency", "user_uuid": "U1"},
            {"order": 2, "role": "Sales Leader", "user_name": "Leader", "user_uuid": "U2", "user_code": "SL01"},
            {"order": 3, "role": "Agent", "user_name": "Agent", "user_uuid": "U3", "user_code": "AGT01"},
            {"order": 4, "role": "Sub Agent", "user_name": "SubAgent", "user_uuid": "U4", "user_code": "SUB01"},
        ]

        mock_db = MockDBManager(
            db_rules_map={
                pol_num: [{"commission_split": db_splits}]
            }
        )

        expanded_df, stats = ensure_override_three_rows(df, db_mgr=mock_db)
        self.assertEqual(len(expanded_df), 4)
        roles = [classify_contract_role(r) for r in expanded_df["Contract"]]
        self.assertIn("sub_agent", roles)
        self.assertIn("agent", roles)
        self.assertIn("sales_leader", roles)
        self.assertIn("agency", roles)

    def test_override_agency_100_percent_does_not_exceed_100(self):
        """199% override defect fix: Agency 100% must zero out non-agency shares so total <= 100.0%."""
        pol_num = "TEST_OVR_199"
        prod_name = "Accident OVR"
        clean_prod = "accident ovr"

        db_splits = [
            {"order": 1, "role": "Agency", "user_uuid": "62a97a01-2801-4c7c-9cd0-cf38f2e0a4f4", "commission_value": 100.0, "entity": "agency"},
            {"order": 2, "role": "Sales Leader", "user_uuid": "UUID_SL", "commission_value": 7.0, "entity": "agent", "user_name": "Sarah Leader", "user_code": "SL01"},
            {"order": 3, "role": "Agent", "user_uuid": "UUID_AGT", "commission_value": 92.0, "entity": "agent", "user_name": "Adam Agent", "user_code": "AGT01"},
        ]

        db_rec = {
            "policy_number": pol_num,
            "product_name": prod_name,
            "policy_uuid": "POL_UUID_OVR",
            "product_uuid": "PROD_UUID_OVR",
            "commission_uuid": "COMM_UUID_OVR",
            "rule_name": f"{prod_name} - OVERRIDE",
            "version": 1,
            "commission_type": 4,  # OVERRIDE
            "payment_frequency": 1,
            "month_from": 1,
            "month_to": None,
            "fee_amount": 0.0,
            "payout_method": 2,  # PERCENTAGE
            "premium_amount": 50.0,
            "commission_split": db_splits,
            "writing_agent_user_id": "UUID_AGT",
        }

        mock_db = MockDBManager(
            db_rules_map={
                (pol_num, clean_prod): [db_rec],
                pol_num: [db_rec],
            },
            users_map={
                "sl01": "UUID_SL",
                "agt01": "UUID_AGT",
                "sarah leader": "UUID_SL",
                "adam agent": "UUID_AGT",
            }
        )

        # File has Agency 100% and mismatched non-agency rows
        df_valid = pd.DataFrame([
            {"Policy Number": pol_num, "Normalized_Product_Name": prod_name, "From Month": 1, "To Month": 0, "Sequence": 10, "Agent ID": "AGY01", "Agent Name": "Main Line Benefits", "Contract": "Agency", "Rate %": 100.0, "Normalized_Split_Share_%": 100.0, "Payout Method": "Percentage", "Carrier Commission": "100%"},
            {"Policy Number": pol_num, "Normalized_Product_Name": prod_name, "From Month": 1, "To Month": 0, "Sequence": 7, "Agent ID": "SL01", "Agent Name": "Sarah Leader", "Contract": "SALES LEADER", "Rate %": 0.0, "Normalized_Split_Share_%": 0.0, "Payout Method": "Percentage", "Carrier Commission": "100%"},
            {"Policy Number": pol_num, "Normalized_Product_Name": prod_name, "From Month": 1, "To Month": 0, "Sequence": 4, "Agent ID": "AGT01", "Agent Name": "Adam Agent", "Contract": "LVL 1", "Rate %": 0.0, "Normalized_Split_Share_%": 0.0, "Payout Method": "Percentage", "Carrier Commission": "100%"},
        ])

        results = stage2_db_verify_dataset(df_valid, mock_db, agency_id=5, dry_run=False)
        mismatch_payloads = results[4]
        self.assertEqual(len(mismatch_payloads), 1)
        m_payload = mismatch_payloads[0]

        # Calculate total split sum in mismatch payload
        total_comm = sum(s["commission_value"] for s in m_payload["commissionSplit"])
        self.assertAlmostEqual(total_comm, 100.00, places=2)

        # Verify non-agency shares were zeroed out instead of adding 7% + 92% = 199%
        splits = {s["order"]: s for s in m_payload["commissionSplit"]}
        self.assertEqual(splits[1]["commission_value"], 100.0)
        self.assertEqual(splits[2]["commission_value"], 0.0)
        self.assertEqual(splits[3]["commission_value"], 0.0)

    # -------------------------------------------------------------------------
    # 5. Import Payload Builder Tests
    # -------------------------------------------------------------------------
    def test_import_payload_role_classification(self):
        """Import payload builder must assign roles by contract, not list index."""
        # 3 participants with Sub Agent and NO Sales Leader
        records = [
            {"uuid": "r1", "policy_number": "POL_IMP_01", "data": {"Agent ID": "AGY01", "Agent Name": "Agency", "Contract": "Agency", "Sequence": "10", "Rate %": "100.0%"}},
            {"uuid": "r2", "policy_number": "POL_IMP_01", "data": {"Agent ID": "AGT01", "Agent Name": "Adam Agent", "Contract": "LVL 1", "Sequence": "5", "Rate %": "85.0%"}},
            {"uuid": "r3", "policy_number": "POL_IMP_01", "data": {"Agent ID": "SUB01", "Agent Name": "Sam SubAgent", "Contract": "SA1", "Sequence": "1", "Rate %": "80.0%"}},
        ]
        payload = build_policy_commission_import_payload(records)
        splits = payload["commissionSplit"]
        self.assertEqual(len(splits), 3)

        # Index 0 is Agency (Order 1)
        self.assertEqual(splits[0]["role"], "Agency")
        self.assertEqual(splits[0]["order"], 1)

        # Index 1 MUST be Agent (Order 2), NOT mislabeled as Sales Leader!
        self.assertEqual(splits[1]["role"], "Agent")
        self.assertEqual(splits[1]["order"], 2)
        self.assertEqual(splits[1]["user_code"], "AGT01")

        # Index 2 MUST be Sub Agent (Order 3), NOT mislabeled as Agent!
        self.assertEqual(splits[2]["role"], "Sub Agent")
        self.assertEqual(splits[2]["order"], 3)
        self.assertEqual(splits[2]["user_code"], "SUB01")

    def test_import_payload_standard_four_person(self):
        """Import payload builder for standard 4-person hierarchy."""
        records = [
            {"uuid": "r1", "policy_number": "POL_IMP_02", "data": {"Agent ID": "AGY01", "Agent Name": "Agency", "Contract": "Agency", "Sequence": "10", "Rate %": "100.0%"}},
            {"uuid": "r2", "policy_number": "POL_IMP_02", "data": {"Agent ID": "SL01", "Agent Name": "Leader", "Contract": "SALES LEADER", "Sequence": "7", "Rate %": "10.0%"}},
            {"uuid": "r3", "policy_number": "POL_IMP_02", "data": {"Agent ID": "AGT01", "Agent Name": "Agent", "Contract": "LVL 1", "Sequence": "4", "Rate %": "85.0%"}},
            {"uuid": "r4", "policy_number": "POL_IMP_02", "data": {"Agent ID": "SUB01", "Agent Name": "SubAgent", "Contract": "SA1", "Sequence": "1", "Rate %": "80.0%"}},
        ]
        payload = build_policy_commission_import_payload(records)
        splits = payload["commissionSplit"]
        self.assertEqual(len(splits), 4)
        self.assertEqual([s["role"] for s in splits], ["Agency", "Sales Leader", "Agent", "Sub Agent"])
        self.assertEqual([s["order"] for s in splits], [1, 2, 3, 4])
    def test_dummy_subagent_slot_not_preserved(self):
        """Empty placeholder Sub Agent slot (no UUID, 0% value) from legacy DB rule must NOT be added to 3-participant file policy."""
        df_valid = pd.DataFrame([
            {"Policy No": "POL_DUMMY_3", "Normalized_Product_Name": "ProdA", "Normalized_Split_Share_%": 20.0, "Sequence": 10, "Contract": "Agency", "Agent ID": "AGY", "Agent Name": "Main Line Benefits", "From Month": 1, "To Month": 12, "Payout Method": "RATE", "Carrier Commission": 100},
            {"Policy No": "POL_DUMMY_3", "Normalized_Product_Name": "ProdA", "Normalized_Split_Share_%": 10.0, "Sequence": 7, "Contract": "Sales Leader", "Agent ID": "8307933", "Agent Name": "Greg Teipel", "From Month": 1, "To Month": 12, "Payout Method": "RATE", "Carrier Commission": 100},
            {"Policy No": "POL_DUMMY_3", "Normalized_Product_Name": "ProdA", "Normalized_Split_Share_%": 70.0, "Sequence": 4, "Contract": "Level 5", "Agent ID": "16018575", "Agent Name": "Kevin Berthelsen", "From Month": 1, "To Month": 12, "Payout Method": "RATE", "Carrier Commission": 100},
        ])
        # DB has 4 slots, but slot 4 is an empty dummy placeholder
        db_split = [
            {"order": 1, "role": "Agency", "user_uuid": "U1", "user_name": "Main Line Benefits", "user_code": "", "commission_value": 20.0},
            {"order": 2, "role": "Sales Leader", "user_uuid": "U2", "user_name": "Gregory Teipel", "user_code": "8307933", "commission_value": 10.0},
            {"order": 3, "role": "Agent", "user_uuid": "U3", "user_name": "Kevin Berthelsen", "user_code": "16018575", "commission_value": 70.0},
            {"order": 4, "role": "Sub Agent", "user_uuid": None, "user_name": "", "user_code": "", "commission_value": 0.0},
        ]
        db_rules = {
            "POL_DUMMY_3": [{
                "policy_number": "POL_DUMMY_3", "policy_uuid": "P_DUMMY", "product_uuid": "PR_1", "product_name": "ProdA",
                "commission_uuid": "COMM_DUMMY", "month_from": 1, "month_to": 12, "fee_amount": 100.0, "payout_method": 2,
                "commission_type": 1, "commission_split": db_split
            }]
        }
        mock_db = MockDBManager(db_rules_map=db_rules, users_map={"8307933": "U2", "16018575": "U3"})
        results = stage2_db_verify_dataset(df_valid, mock_db, agency_id=5, dry_run=False)
        df_matched = results[0]
        # Should match cleanly
        self.assertEqual(len(df_matched), 3)


    def test_blank_contract_continuation_keeps_three_splits(self):
        """Later bands with a blank Contract must keep Agent, Sales Leader, and Agency."""
        df = pd.DataFrame([
            {"Policy No": "685711160", "Product": "NCD MetLife -24%", "From Month": 1, "To Month": 12, "Contract": "LVL5", "Sequence": 1, "Agent ID": "'17425766", "Rate": "17.76%"},
            {"Policy No": "685711160", "Product": "NCD MetLife -24%", "From Month": 13, "To Month": 60, "Contract": None, "Sequence": 1, "Agent ID": "'17425766", "Rate": "15.84%"},
            {"Policy No": "685711160", "Product": "NCD MetLife -24%", "From Month": 1, "To Month": 12, "Contract": "Sales Leader", "Sequence": 2, "Agent ID": "'17425766SL", "Rate": "19.72%"},
            {"Policy No": "685711160", "Product": "NCD MetLife -24%", "From Month": 13, "To Month": 60, "Contract": None, "Sequence": 2, "Agent ID": "'17425766SL", "Rate": "17.52%"},
            {"Policy No": "685711160", "Product": "NCD MetLife -24%", "From Month": 1, "To Month": 12, "Contract": "Agency", "Sequence": 3, "Agent ID": "'20411454", "Rate": "24.00%"},
            {"Policy No": "685711160", "Product": "NCD MetLife -24%", "From Month": 13, "To Month": 60, "Contract": None, "Sequence": 3, "Agent ID": "'20411454", "Rate": "24.00%"},
        ])
        out, _ = normalize_year_bands(df)
        band = out[out["From Month"].astype(int) == 13]
        self.assertEqual(len(band), 3)
        roles = sorted(classify_contract_role(c) for c in band["Contract"])
        self.assertEqual(roles, ["agency", "agent", "sales_leader"])

    def test_all_blank_contracts_keep_three_people(self):
        """A sheet with no Contract at all still keeps Agent, Sales Leader, and Agency."""
        df = pd.DataFrame([
            {"Policy No": "970060", "Product": "Sunlife", "From Month": 1, "To Month": 12, "Contract": None, "Sequence": 1, "Agent ID": "'8244481", "Agent Name": "PAMELA  DEATON", "Rate": "10.92%"},
            {"Policy No": "970060", "Product": "Sunlife", "From Month": 13, "To Month": 0, "Contract": None, "Sequence": 1, "Agent ID": "'8244481", "Agent Name": "PAMELA  DEATON", "Rate": "10.92%"},
            {"Policy No": "970060", "Product": "Sunlife", "From Month": 1, "To Month": 12, "Contract": None, "Sequence": 2, "Agent ID": "'7146581SL", "Agent Name": "ANTHONY  HUNSBERGER", "Rate": "11.83%"},
            {"Policy No": "970060", "Product": "Sunlife", "From Month": 13, "To Month": 0, "Contract": None, "Sequence": 2, "Agent ID": "'7146581SL", "Agent Name": "ANTHONY  HUNSBERGER", "Rate": "11.31%"},
            {"Policy No": "970060", "Product": "Sunlife", "From Month": 1, "To Month": 12, "Contract": None, "Sequence": 3, "Agent ID": "'20411454", "Agent Name": "MAIN LINE  BENEFITS", "Rate": "13.00%"},
            {"Policy No": "970060", "Product": "Sunlife", "From Month": 13, "To Month": 0, "Contract": None, "Sequence": 3, "Agent ID": "'20411454", "Agent Name": "MAIN LINE  BENEFITS", "Rate": "13.00%"},
        ])
        out, summary = normalize_year_bands(df)
        self.assertEqual(summary["rows_out"], 6)
        roles = sorted(classify_contract_role(c) for c in out["Contract"])
        self.assertEqual(roles, ["agency", "agency", "agent", "agent", "sales_leader", "sales_leader"])

    def test_two_agents_with_same_role_are_both_kept(self):
        """StreetZA and StreetZC are both agents. Sequence 2 must not be dropped."""
        df = pd.DataFrame([
            {"Policy No": "8502767093", "Product": "BCBS U65 AZ", "From Month": 1, "To Month": 12, "Contract": "StreetZC", "Sequence": 1, "Agent ID": "'12047257", "Agent Name": "EUGENE  ZANGARI", "Amount": "$17.00"},
            {"Policy No": "8502767093", "Product": "BCBS U65 AZ", "From Month": 13, "To Month": 0, "Contract": "StreetZC", "Sequence": 1, "Agent ID": "'12047257", "Agent Name": "EUGENE  ZANGARI", "Amount": "$17.00"},
            {"Policy No": "8502767093", "Product": "BCBS U65 AZ", "From Month": 1, "To Month": 12, "Contract": "StreetZA", "Sequence": 2, "Agent ID": "'7750822", "Agent Name": "ALESSANDRO  ZANGARI", "Amount": "$19.00"},
            {"Policy No": "8502767093", "Product": "BCBS U65 AZ", "From Month": 13, "To Month": 0, "Contract": "StreetZA", "Sequence": 2, "Agent ID": "'7750822", "Agent Name": "ALESSANDRO  ZANGARI", "Amount": "$19.00"},
            {"Policy No": "8502767093", "Product": "BCBS U65 AZ", "From Month": 1, "To Month": 12, "Contract": "Agency", "Sequence": 3, "Agent ID": "'20411454", "Agent Name": "MAIN LINE  BENEFITS", "Amount": "$20.00"},
            {"Policy No": "8502767093", "Product": "BCBS U65 AZ", "From Month": 13, "To Month": 0, "Contract": "Agency", "Sequence": 3, "Agent ID": "'20411454", "Agent Name": "MAIN LINE  BENEFITS", "Amount": "$20.00"},
        ])
        out, summary = normalize_year_bands(df)
        self.assertEqual(summary["rows_out"], 6)
        names = set(out["Agent Name"])
        self.assertIn("ALESSANDRO  ZANGARI", names)
        self.assertIn("EUGENE  ZANGARI", names)


if __name__ == "__main__":
    unittest.main(verbosity=2)
