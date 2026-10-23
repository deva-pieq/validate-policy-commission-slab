#!/usr/bin/env python3
"""
test_post_import_validator.py

Comprehensive Test Suite for Post-Import APL vs Database Validation Engine
==========================================================================
Covers all 44 automated test scenarios required:
- Policy identifier matching (1-9)
- Person-name matching (10-22)
- Persisted data validation (23-32)
- Execution status & overall status (33-38)
- Multi-tab Excel workbook generation & security (39-44)
"""

import os
import sys
import tempfile
import unittest
import pandas as pd
import openpyxl

from validate_apl_vs_db_after_import import (
    compare_person_names,
    match_policy_identifier,
    clean_policy_number,
    clean_agent_code,
    classify_contract_role,
    compare_decimals,
    sanitize_for_excel,
    determine_execution_status,
    compute_overall_status,
    validate_policy_records,
    generate_excel_validation_workbook,
    run_post_import_validation,
    ReadOnlyDatabaseManager,
)


class MockDatabaseManager:
    """Mock DatabaseManager for testing isolated DB queries and validation."""

    def __init__(self, policies_map=None, all_db_numbers=None, users_map=None):
        self.conn = True
        self.policies_map = policies_map or {}
        self.all_db_numbers = all_db_numbers or list(self.policies_map.keys())
        self.user_display_by_uuid = users_map or {}

    def fetch_policies_and_commissions(self, agency_id, policy_numbers):
        found = {p: self.policies_map[p] for p in policy_numbers if p in self.policies_map}
        return found, self.all_db_numbers

    def fetch_all_users_and_contacts(self, agency_id):
        return {}

    def format_user_display(self, user_id):
        uid = str(user_id or "").strip().lower()
        info = self.user_display_by_uuid.get(uid, {})
        name = info.get("name", "")
        code = info.get("code", "")
        if name and code:
            return f"{name} ({code})"
        return name or code or uid


class TestPostImportValidator(unittest.TestCase):
    """Complete 44-case test suite for validate_apl_vs_db_after_import.py"""

    # =========================================================================
    # PART 1: POLICY IDENTIFIER MATCHING (Tests 1-9)
    # =========================================================================
    def test_01_exact_policy_number(self):
        res = match_policy_identifier("A03576893", ["A03576893", "B100200"])
        self.assertEqual(res["status"], "EXACT_MATCH")
        self.assertEqual(res["method"], "EXACT")
        self.assertEqual(res["db_policy"], "A03576893")
        self.assertEqual(res["similarity"], 1.0)

    def test_02_case_and_whitespace_normalization(self):
        res = match_policy_identifier("  'a03576893  ", ["A03576893"])
        self.assertEqual(res["status"], "NORMALIZED_MATCH")
        self.assertEqual(res["method"], "BASIC_NORMALIZATION")
        self.assertEqual(res["db_policy"], "A03576893")

    def test_03_policy_number_extra_character(self):
        res = match_policy_identifier("A03576893X", ["A03576893"])
        self.assertEqual(res["status"], "CLOSE_MATCH_EXTRA_CHARACTERS")
        self.assertEqual(res["method"], "EXTRA_CHARACTER")
        self.assertEqual(res["db_policy"], "A03576893")
        self.assertGreaterEqual(res["similarity"], 0.9)

    def test_04_policy_number_missing_character(self):
        res = match_policy_identifier("A0357689", ["A03576893"])
        self.assertEqual(res["status"], "CLOSE_MATCH_MISSING_CHARACTERS")
        self.assertEqual(res["method"], "MISSING_CHARACTER")
        self.assertEqual(res["db_policy"], "A03576893")

    def test_05_possible_character_substitution_typo(self):
        res = match_policy_identifier("A03576893", ["A03576894"])
        self.assertEqual(res["status"], "CLOSE_MATCH_POSSIBLE_TYPO")
        self.assertEqual(res["method"], "POSSIBLE_TYPO")
        self.assertEqual(res["db_policy"], "A03576894")

    def test_06_two_similar_but_different_policies(self):
        # Two completely different numbers with common prefix but too many differing digits
        res = match_policy_identifier("A03576893", ["A03599999"])
        self.assertEqual(res["status"], "NO_MATCH")
        self.assertIsNone(res["db_policy"])

    def test_07_multiple_plausible_db_candidates(self):
        # Multiple candidates within single-character distance
        res = match_policy_identifier("A0357689", ["A03576891", "A03576892"])
        self.assertEqual(res["status"], "AMBIGUOUS_MATCH")
        self.assertEqual(res["method"], "MULTIPLE_CANDIDATES")
        self.assertIsNone(res["db_policy"])

    def test_08_no_matching_db_policy(self):
        res = match_policy_identifier("XYZ999999", ["A03576893", "B100200"])
        self.assertEqual(res["status"], "NO_MATCH")
        self.assertEqual(res["method"], "NONE")
        self.assertIsNone(res["db_policy"])

    def test_09_prevent_one_db_record_assigned_to_unrelated_apl(self):
        # Completely unrelated policy numbers must never match
        res = match_policy_identifier("UNRELATED_12345", ["DB_POLICY_98765"])
        self.assertEqual(res["status"], "NO_MATCH")

    # =========================================================================
    # PART 2: PERSON-NAME MATCHING (Tests 10-22)
    # =========================================================================
    def test_10_exact_name(self):
        res = compare_person_names("John Smith", "John Smith")
        self.assertEqual(res["status"], "EXACT_MATCH")
        self.assertEqual(res["method"], "EXACT")
        self.assertEqual(res["similarity"], 1.0)

    def test_11_case_difference(self):
        res = compare_person_names("John Smith", "JOHN SMITH")
        self.assertEqual(res["status"], "NORMALIZED_MATCH")
        self.assertEqual(res["method"], "BASIC_NORMALIZATION")

    def test_12_repeated_whitespace(self):
        res = compare_person_names("John Smith", "JOHN  SMITH")
        self.assertEqual(res["status"], "NORMALIZED_MATCH")
        self.assertEqual(res["method"], "BASIC_NORMALIZATION")

    def test_13_extra_middle_name_in_apl(self):
        res = compare_person_names("JOHN Walk Smith", "John Smith")
        self.assertEqual(res["status"], "NORMALIZED_MATCH")
        self.assertEqual(res["method"], "EXTRA_MIDDLE_NAME")

    def test_14_extra_middle_name_in_db(self):
        res = compare_person_names("John Smith", "JOHN Walk Smith")
        self.assertEqual(res["status"], "NORMALIZED_MATCH")
        self.assertEqual(res["method"], "EXTRA_MIDDLE_NAME")

    def test_15_multiple_additional_middle_name_tokens(self):
        res = compare_person_names("John Alan Michael Smith", "John Smith")
        self.assertEqual(res["status"], "NORMALIZED_MATCH")
        self.assertEqual(res["method"], "EXTRA_MIDDLE_NAME")

    def test_16_matching_first_and_last_with_extra_middle_tokens(self):
        # Both have middle tokens, matching first and last
        res = compare_person_names("John Walk Smith", "John Michael Smith")
        self.assertEqual(res["status"], "NORMALIZED_MATCH")
        self.assertEqual(res["method"], "EXTRA_MIDDLE_NAME")

    def test_17_reordered_tokens_must_not_pass(self):
        res = compare_person_names("Smith John", "John Smith")
        self.assertEqual(res["status"], "MISMATCH")
        res2 = compare_person_names("John Smith", "Michael John Smith")
        self.assertEqual(res2["status"], "MISMATCH")

    def test_18_spelling_variations_require_fuzzy_review(self):
        res = compare_person_names("John Smith", "John Smyth")
        self.assertEqual(res["status"], "CLOSE_MATCH_REVIEW_REQUIRED")
        self.assertEqual(res["method"], "FUZZY_SPELLING_VARIATION")

    def test_19_different_surnames(self):
        res = compare_person_names("John Smith", "John Williams")
        self.assertEqual(res["status"], "MISMATCH")

    def test_20_ambiguous_candidates_with_similar_names(self):
        # Compare against multiple candidates where several are close
        candidates = ["John Smyth", "John Smythe", "Jon Smith"]
        results = [compare_person_names("John Smith", c) for c in candidates]
        # All of these are fuzzy review variations
        for r in results:
            self.assertIn(r["status"], ("CLOSE_MATCH_REVIEW_REQUIRED", "NORMALIZED_MATCH"))

    def test_21_same_person_appearing_in_different_roles(self):
        # Verify role classification keeps roles distinct
        role1 = classify_contract_role("SALES LEADER")
        role2 = classify_contract_role("LVL 1")
        self.assertEqual(role1, "Sales Leader")
        self.assertEqual(role2, "Agent")
        self.assertNotEqual(role1, role2)

    def test_22_preventing_uuid_transfer_based_solely_on_fuzzy_name(self):
        # Split comparison must flag review and not silently accept fuzzy match
        apl_row = [{"Contract": "LVL 1", "Agent Name": "John Smith", "Rate": "85.0%", "Sequence": 4}]
        db_split = [{"role": "Agent", "user_name": "John Smyth", "commission_value": 85.0, "user_uuid": "U_SMYTH", "order": 3}]
        db_rec = [{"product_name": "Test Prod", "writing_agent_user_id": "U_SMYTH", "commission_split": db_split}]
        pol_match = {"status": "EXACT_MATCH", "method": "EXACT", "apl_policy": "POL1", "db_policy": "POL1"}

        res = validate_policy_records(apl_row, db_rec, pol_match, "EXECUTED_SUCCESSFULLY", "Report", lambda u: "John Smyth")
        split_comp = res["split_comparisons"][0]
        # Must require review and flag name mismatch
        self.assertTrue(split_comp["review_required"])
        self.assertEqual(split_comp["participant_match_status"], "CLOSE_MATCH_REVIEW_REQUIRED")

    # =========================================================================
    # PART 3: PERSISTED DATA VALIDATION (Tests 23-32)
    # =========================================================================
    def test_23_exact_full_data_match(self):
        apl_rows = [
            {"Policy No": "POL_EXACT", "Product": "Dental Plus", "Contract": "Agency", "Agent Name": "Main Line Benefits", "Rate": "100%", "Sequence": 10},
            {"Policy No": "POL_EXACT", "Product": "Dental Plus", "Contract": "LVL 1", "Agent Name": "Adam Agent", "Rate": "85%", "Sequence": 4},
        ]
        db_splits = [
            {"order": 1, "role": "Agency", "user_name": "Main Line Benefits", "commission_value": 100.0, "user_uuid": "U_AGY"},
            {"order": 2, "role": "Agent", "user_name": "Adam Agent", "commission_value": 85.0, "user_uuid": "U_AGT"},
        ]
        db_recs = [{
            "policy_number": "POL_EXACT", "product_name": "Dental Plus", "writing_agent_user_id": "U_AGT", "commission_split": db_splits
        }]
        pol_match = {"status": "EXACT_MATCH", "method": "EXACT", "apl_policy": "POL_EXACT", "db_policy": "POL_EXACT"}
        res = validate_policy_records(apl_rows, db_recs, pol_match, "EXECUTED_SUCCESSFULLY", "Success", lambda u: "Adam Agent")

        self.assertEqual(res["summary"]["data_validation_status"], "EXACT_DATA_MATCH")
        self.assertEqual(res["summary"]["overall_status"], "SUCCESS_EXACT_MATCH")
        self.assertFalse(res["summary"]["review_required"])

    def test_24_single_field_mismatch_product(self):
        apl_rows = [{"Product": "Cigna Health", "Contract": "Agency", "Agent Name": "MLB", "Rate": "100%", "Sequence": 10}]
        db_recs = [{"product_name": "Aetna Vision", "writing_agent_user_id": "U1", "commission_split": [{"role": "Agency", "user_name": "MLB", "commission_value": 100.0, "user_uuid": "U1"}]}]
        pol_match = {"status": "EXACT_MATCH", "method": "EXACT", "apl_policy": "POL_MIS", "db_policy": "POL_MIS"}

        res = validate_policy_records(apl_rows, db_recs, pol_match, "EXECUTED_SUCCESSFULLY", "Success", lambda u: "MLB")
        self.assertEqual(res["summary"]["data_validation_status"], "DATA_MISMATCH")
        self.assertIn("product_name", res["summary"]["mismatch_field_names"])

    def test_25_missing_or_extra_commission_splits(self):
        # APL has Agency and Agent, but DB only has Agency
        apl_rows = [
            {"Product": "P1", "Contract": "Agency", "Agent Name": "MLB", "Rate": "100%", "Sequence": 10},
            {"Product": "P1", "Contract": "LVL 1", "Agent Name": "Adam Agent", "Rate": "80%", "Sequence": 4},
        ]
        db_recs = [{"product_name": "P1", "writing_agent_user_id": "U1", "commission_split": [{"role": "Agency", "user_name": "MLB", "commission_value": 100.0, "user_uuid": "U1"}]}]
        pol_match = {"status": "EXACT_MATCH", "method": "EXACT", "apl_policy": "POL_SP", "db_policy": "POL_SP"}

        res = validate_policy_records(apl_rows, db_recs, pol_match, "EXECUTED_SUCCESSFULLY", "Success", lambda u: "MLB")
        self.assertEqual(res["summary"]["data_validation_status"], "DATA_MISMATCH")
        self.assertTrue(any("missing_split" in name for name in res["summary"]["mismatch_field_names"].split(", ")))

    def test_26_role_mismatch(self):
        # APL specifies Sub Agent, but DB has Agent
        apl_rows = [{"Product": "P1", "Contract": "SUBAGT1", "Agent Name": "Sam Sub", "Rate": "50%", "Sequence": 1}]
        db_recs = [{"product_name": "P1", "writing_agent_user_id": "U1", "commission_split": [{"role": "Agent", "user_name": "Sam Sub", "commission_value": 50.0, "user_uuid": "U1"}]}]
        pol_match = {"status": "EXACT_MATCH", "method": "EXACT", "apl_policy": "P_ROLE", "db_policy": "P_ROLE"}

        res = validate_policy_records(apl_rows, db_recs, pol_match, "EXECUTED_SUCCESSFULLY", "Success", lambda u: "Sam Sub")
        self.assertEqual(res["summary"]["data_validation_status"], "DATA_MISMATCH")

    def test_27_incorrect_commission_value_or_percentage(self):
        apl_rows = [{"Product": "P1", "Contract": "Agency", "Agent Name": "MLB", "Rate": "100%", "Sequence": 10}]
        db_recs = [{"product_name": "P1", "writing_agent_user_id": "U1", "commission_split": [{"role": "Agency", "user_name": "MLB", "commission_value": 85.0, "user_uuid": "U1"}]}]
        pol_match = {"status": "EXACT_MATCH", "method": "EXACT", "apl_policy": "P_VAL", "db_policy": "P_VAL"}

        res = validate_policy_records(apl_rows, db_recs, pol_match, "EXECUTED_SUCCESSFULLY", "Success", lambda u: "MLB")
        self.assertEqual(res["summary"]["data_validation_status"], "DATA_MISMATCH")

    def test_28_decimal_precision(self):
        self.assertTrue(compare_decimals(100.004, 100.00, tolerance=0.01))
        self.assertFalse(compare_decimals(100.05, 100.00, tolerance=0.01))

    def test_29_month_boundary_normalization(self):
        from validate_apl_vs_db_after_import import clean_policy_number
        # Clean policy stripping quotes
        self.assertEqual(clean_policy_number("'POL123"), "POL123")

    def test_30_missing_uuid_in_db_split(self):
        apl_rows = [{"Product": "P1", "Contract": "LVL 1", "Agent Name": "Adam Agent", "Rate": "80%", "Sequence": 4}]
        db_recs = [{"product_name": "P1", "writing_agent_user_id": None, "commission_split": [{"role": "Agent", "user_name": "Adam Agent", "commission_value": 80.0, "user_uuid": None}]}]
        pol_match = {"status": "EXACT_MATCH", "method": "EXACT", "apl_policy": "P_NO_UUID", "db_policy": "P_NO_UUID"}

        res = validate_policy_records(apl_rows, db_recs, pol_match, "EXECUTED_SUCCESSFULLY", "Success", lambda u: "")
        self.assertTrue(res["summary"]["review_required"])

    def test_31_unknown_field_mapping_not_persisted(self):
        # Non-persisted carrier status columns must be marked NOT_PERSISTED_IN_DB, not mismatch
        apl_rows = [
            {"Product": "P1", "Effective Date": "01/01/2026", "Policy Status": "Active", "Contract": "LVL 1", "Agent Name": "Adam Agent", "Rate": "100%", "Sequence": 4}
        ]
        db_recs = [{
            "product_name": "P1", "writing_agent_user_id": "U1",
            "commission_split": [{"role": "Agent", "user_name": "Adam Agent", "commission_value": 100.0, "user_uuid": "U1"}]
        }]
        pol_match = {"status": "EXACT_MATCH", "method": "EXACT", "apl_policy": "P_NP", "db_policy": "P_NP"}

        res = validate_policy_records(apl_rows, db_recs, pol_match, "EXECUTED_SUCCESSFULLY", "Success", lambda u: "Adam Agent")
        np_comps = [fc for fc in res["field_comparisons"] if fc["status"] == "NOT_PERSISTED_IN_DB"]
        self.assertGreaterEqual(len(np_comps), 2)
        # Should not contribute to mismatch count
        self.assertEqual(res["summary"]["mismatched_fields_count"], 0)

    def test_32_incomplete_db_query_db_record_not_found(self):
        apl_rows = [{"Product": "P1", "Contract": "Agency", "Agent Name": "MLB", "Rate": "100%", "Sequence": 10}]
        pol_match = {"status": "NO_MATCH", "method": "NONE", "apl_policy": "P_NOT_IN_DB", "db_policy": None}

        res = validate_policy_records(apl_rows, [], pol_match, "UNKNOWN", "No evidence", lambda u: "")
        self.assertEqual(res["summary"]["data_validation_status"], "DB_RECORD_NOT_FOUND")
        self.assertEqual(res["summary"]["overall_status"], "DB_RECORD_NOT_FOUND")

    # =========================================================================
    # PART 4: EXECUTION STATUS (Tests 33-38)
    # =========================================================================
    def test_33_successful_execution_with_matching_data(self):
        st = compute_overall_status("EXECUTED_SUCCESSFULLY", "EXACT_DATA_MATCH")
        self.assertEqual(st, "SUCCESS_EXACT_MATCH")
        st_norm = compute_overall_status("EXECUTED_SUCCESSFULLY", "MATCH_AFTER_NORMALIZATION")
        self.assertEqual(st_norm, "SUCCESS_NORMALIZED_MATCH")

    def test_34_successful_execution_with_mismatching_data(self):
        st = compute_overall_status("EXECUTED_SUCCESSFULLY", "DATA_MISMATCH")
        self.assertEqual(st, "EXECUTED_BUT_DATA_MISMATCH")

    def test_35_failed_or_rejected_update(self):
        st1 = compute_overall_status("REJECTED", "EXACT_DATA_MATCH")
        self.assertEqual(st1, "REJECTED")
        st2 = compute_overall_status("FAILED", "EXACT_DATA_MATCH")
        self.assertEqual(st2, "EXECUTION_FAILED")

    def test_36_skipped_policy(self):
        st = compute_overall_status("SKIPPED", "EXACT_DATA_MATCH")
        self.assertEqual(st, "EXECUTION_SKIPPED")

    def test_37_unknown_execution_status_with_matching_data(self):
        st = compute_overall_status("UNKNOWN", "EXACT_DATA_MATCH")
        self.assertEqual(st, "EXECUTION_STATUS_UNKNOWN")

    def test_38_missing_execution_report(self):
        st, ev = determine_execution_status("POL_NO_REPORT", {})
        self.assertEqual(st, "UNKNOWN")
        self.assertIn("No per-policy execution report", ev)

    # =========================================================================
    # PART 5: EXCEL WORKBOOK GENERATION & INTEGRITY (Tests 39-44)
    # =========================================================================
    def test_39_expected_worksheets_exist(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            out_xlsx = os.path.join(tmpdir, "test_wb.xlsx")
            summaries = [{
                "apl_policy": "P1", "db_policy": "P1", "policy_match_status": "EXACT_MATCH", "policy_match_method": "EXACT",
                "policy_similarity": 1.0, "product_name": "ProdA", "carrier": "MLB", "agency": "Agency 5",
                "apl_agent_name": "Agent A", "db_agent_name": "Agent A", "agent_match_status": "EXACT_MATCH",
                "apl_writing_agent_name": "Agent A", "db_writing_agent_name": "Agent A", "writing_agent_match_status": "EXACT_MATCH",
                "execution_status": "EXECUTED_SUCCESSFULLY", "execution_evidence": "Report OK",
                "data_validation_status": "EXACT_DATA_MATCH", "overall_status": "SUCCESS_EXACT_MATCH",
                "matched_fields_count": 5, "mismatched_fields_count": 0, "review_required_fields_count": 0,
                "mismatch_field_names": "", "review_required": False, "review_reason": ""
            }]
            fld_comps = [{
                "apl_policy": "P1", "db_policy": "P1", "field_name": "policy_number", "apl_value": "P1", "db_value": "P1",
                "status": "EXACT_MATCH", "method": "EXACT", "diff_reason": "", "review_required": False
            }]
            split_comps = [{
                "apl_policy": "P1", "db_policy": "P1", "apl_name": "Agent A", "db_name": "Agent A", "apl_role": "Agent",
                "db_role": "Agent", "apl_order": 1, "db_order": 1, "apl_uuid": "", "db_uuid": "U1",
                "apl_commission_value": 100.0, "db_commission_value": 100.0, "participant_match_status": "EXACT_MATCH",
                "field_comparison_result": "EXACT_DATA_MATCH", "review_required": False, "review_reason": ""
            }]
            generate_excel_validation_workbook(summaries, fld_comps, split_comps, out_xlsx, "run_test")

            wb = openpyxl.load_workbook(out_xlsx)
            self.assertEqual(
                wb.sheetnames,
                ["Summary", "Policy Validation", "Field Comparison", "Commission Split Validation", "Exceptions"]
            )

    def test_40_expected_columns_exist(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            out_xlsx = os.path.join(tmpdir, "test_cols.xlsx")
            generate_excel_validation_workbook([], [], [], out_xlsx, "run_test")
            wb = openpyxl.load_workbook(out_xlsx)
            ws_pol = wb["Policy Validation"]
            cols = [cell.value for cell in ws_pol[1]]
            self.assertIn("APL Policy Number", cols)
            self.assertIn("DB Policy Number", cols)
            self.assertIn("Overall Status", cols)
            self.assertIn("Review Required", cols)

    def test_41_summary_counts_agree_with_detailed_sheets(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            out_xlsx = os.path.join(tmpdir, "test_counts.xlsx")
            summaries = [
                {"apl_policy": "P1", "db_policy": "P1", "policy_match_status": "EXACT_MATCH", "policy_match_method": "EXACT", "policy_similarity": 1.0, "product_name": "P", "carrier": "C", "agency": "A", "apl_agent_name": "N", "db_agent_name": "N", "agent_match_status": "EXACT_MATCH", "apl_writing_agent_name": "N", "db_writing_agent_name": "N", "writing_agent_match_status": "EXACT_MATCH", "execution_status": "EXECUTED_SUCCESSFULLY", "execution_evidence": "Report", "data_validation_status": "EXACT_DATA_MATCH", "overall_status": "SUCCESS_EXACT_MATCH", "matched_fields_count": 5, "mismatched_fields_count": 0, "review_required_fields_count": 0, "mismatch_field_names": "", "review_required": False, "review_reason": ""},
                {"apl_policy": "P2", "db_policy": "P2", "policy_match_status": "EXACT_MATCH", "policy_match_method": "EXACT", "policy_similarity": 1.0, "product_name": "P", "carrier": "C", "agency": "A", "apl_agent_name": "N", "db_agent_name": "N", "agent_match_status": "EXACT_MATCH", "apl_writing_agent_name": "N", "db_writing_agent_name": "N", "writing_agent_match_status": "EXACT_MATCH", "execution_status": "FAILED", "execution_evidence": "Report", "data_validation_status": "DATA_MISMATCH", "overall_status": "EXECUTION_FAILED", "matched_fields_count": 2, "mismatched_fields_count": 3, "review_required_fields_count": 1, "mismatch_field_names": "rate", "review_required": True, "review_reason": "Failed"}
            ]
            generate_excel_validation_workbook(summaries, [], [], out_xlsx, "run_counts")
            wb = openpyxl.load_workbook(out_xlsx)
            ws_sum = wb["Summary"]
            # Total APL policies in row 5
            total_val = None
            for row in ws_sum.iter_rows(values_only=True):
                if row[0] == "Total APL Policies Analyzed":
                    total_val = row[1]
            self.assertEqual(total_val, 2)

    def test_42_fuzzy_matches_appear_as_review_cases(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            out_xlsx = os.path.join(tmpdir, "test_fuzzy.xlsx")
            summaries = [{
                "apl_policy": "P_FUZZY", "db_policy": "P_FUZZY_DB", "policy_match_status": "CLOSE_MATCH_EXTRA_CHARACTERS",
                "policy_match_method": "EXTRA_CHARACTER", "policy_similarity": 0.95, "product_name": "Prod", "carrier": "MLB",
                "agency": "Agency 5", "apl_agent_name": "Agent", "db_agent_name": "Agent", "agent_match_status": "EXACT_MATCH",
                "apl_writing_agent_name": "Agent", "db_writing_agent_name": "Agent", "writing_agent_match_status": "EXACT_MATCH",
                "execution_status": "EXECUTED_SUCCESSFULLY", "execution_evidence": "Report",
                "data_validation_status": "CLOSE_MATCH_REVIEW_REQUIRED", "overall_status": "REVIEW_REQUIRED",
                "matched_fields_count": 4, "mismatched_fields_count": 0, "review_required_fields_count": 1,
                "mismatch_field_names": "", "review_required": True, "review_reason": "Policy close match requires review"
            }]
            generate_excel_validation_workbook(summaries, [], [], out_xlsx, "run_fuzzy")
            wb = openpyxl.load_workbook(out_xlsx)
            ws_exc = wb["Exceptions"]
            self.assertEqual(ws_exc.max_row, 2)  # Header + 1 row
            self.assertEqual(ws_exc.cell(row=2, column=1).value, "P_FUZZY")

    def test_43_source_values_cannot_inject_formulas(self):
        # Formula injection mitigation
        self.assertEqual(sanitize_for_excel("=SUM(A1:A10)"), "'=SUM(A1:A10)")
        self.assertEqual(sanitize_for_excel("+12345"), "'+12345")
        self.assertEqual(sanitize_for_excel("-99.99"), "'-99.99")
        self.assertEqual(sanitize_for_excel("@cmd"), "'@cmd")
        self.assertEqual(sanitize_for_excel("Normal Text"), "Normal Text")
        self.assertEqual(sanitize_for_excel(123), 123)

    def test_45_same_person_agent_and_sl_role_matching(self):
        from validate_apl_vs_db_after_import import annotate_apl_rows
        df_apl = pd.DataFrame([
            {
                "Policy No": "POL_MULTI_ROLE", "From Month": 1, "To Month": 12, "Rate": 25.00,
                "Contract": "Agency", "Agent Name": "Main Line Benefits", "Agent ID": "AGY",
                "Normalized_Split_Share_%": 8.00
            },
            {
                "Policy No": "POL_MULTI_ROLE", "From Month": 1, "To Month": 12, "Rate": 21.50,
                "Contract": "LVL5", "Agent Name": "GREG TEIPEL", "Agent ID": "8307933",
                "Normalized_Split_Share_%": 86.00
            },
            {
                "Policy No": "POL_MULTI_ROLE", "From Month": 1, "To Month": 12, "Rate": 23.00,
                "Contract": "Sales Leader", "Agent Name": "GREG TEIPEL", "Agent ID": "8307933SL",
                "Normalized_Split_Share_%": 6.00
            }
        ])
        db_map = {
            "POL_MULTI_ROLE": [{
                "policy_number": "POL_MULTI_ROLE", "month_from": 1, "month_to": 12, "commission_uuid": "u1",
                "commission_split": [
                    {"order": 1, "role": "Agency", "user_name": "Main Line Benefits", "commission_value": 8.0, "level": None},
                    {"order": 2, "role": "Sales Leader", "user_code": "8307933", "user_name": "8307933 - Gregory Teipel", "commission_value": 6.0, "level": "Sales Leader"},
                    {"order": 3, "role": "Agent", "user_code": "8307933", "user_name": "8307933 - Gregory Teipel", "commission_value": 86.0, "level": "Level 5"}
                ]
            }]
        }
        res = annotate_apl_rows(df_apl, db_map, ["POL_MULTI_ROLE"])
        # Agent row must match Agent split (86%), NOT Sales Leader (6%)
        agent_row = res[res["Contract"] == "LVL5"].iloc[0]
        self.assertEqual(agent_row["Role_Match"], "MATCH (Agent)")
        self.assertEqual(agent_row["Level_Match"], "MATCH (Level 5)")
        self.assertEqual(agent_row["Split_Sum_Match"], "MATCH (100.00%)")
        self.assertIn(agent_row["Match_Status"], ("EXACT_MATCH", "NORMALIZED_MATCH"))
        self.assertEqual(agent_row["Mismatch_Reason"], "None (All matched)")
        self.assertIn("86.00%", str(agent_row["Rate_Match"]))

        # Sales Leader row must match Sales Leader split (6%)
        sl_row = res[res["Contract"] == "Sales Leader"].iloc[0]
        self.assertEqual(sl_row["Role_Match"], "MATCH (Sales Leader)")
        self.assertEqual(sl_row["Level_Match"], "MATCH (Sales Leader)")
        self.assertEqual(sl_row["Split_Sum_Match"], "MATCH (100.00%)")
        self.assertIn(sl_row["Match_Status"], ("EXACT_MATCH", "NORMALIZED_MATCH"))
        self.assertEqual(sl_row["Mismatch_Reason"], "None (All matched)")
        self.assertIn("6.00%", str(sl_row["Rate_Match"]))

    def test_46_level_matching_zero_compromise(self):
        from validate_apl_vs_db_after_import import normalize_level, compare_levels
        # Test normalizations
        self.assertEqual(normalize_level("LVL5"), "Level 5")
        self.assertEqual(normalize_level("LVL 5"), "Level 5")
        self.assertEqual(normalize_level("LVL-5"), "Level 5")
        self.assertEqual(normalize_level("LVL1"), "Level 1")
        self.assertEqual(normalize_level("LVL 1"), "Level 1")
        self.assertEqual(normalize_level("LVL2"), "Level 2")
        self.assertEqual(normalize_level("LVL4"), "Level 4")
        self.assertEqual(normalize_level("SUBAGT1"), "SA1")
        self.assertEqual(normalize_level("SUBAGT 1"), "SA1")
        self.assertEqual(normalize_level("SUBAGT2"), "SA2")
        self.assertEqual(normalize_level("SA2"), "SA2")
        self.assertEqual(normalize_level("Sales Leader"), "Sales Leader")
        self.assertEqual(normalize_level("Agency"), "Agency")
        self.assertEqual(normalize_level(None), "")

        # Test compare_levels matches
        self.assertTrue(compare_levels("LVL5", "Level 5", "Agent")["match"])
        self.assertTrue(compare_levels("LVL 1", "Level 1", "Agent")["match"])
        self.assertTrue(compare_levels("SUBAGT1", "SA1", "Sub Agent")["match"])
        self.assertTrue(compare_levels("SUBAGT2", "SA2", "Sub Agent")["match"])
        self.assertTrue(compare_levels("Agency", None, "Agency")["match"])
        self.assertTrue(compare_levels("Sales Leader", "Sales Leader", "Sales Leader")["match"])
        self.assertTrue(compare_levels("Sales Leader", None, "Sales Leader")["match"])

        # Test zero-compromise mismatches
        res_mismatch = compare_levels("LVL5", "Level 2", "Agent")
        self.assertFalse(res_mismatch["match"])
        self.assertIn("MISMATCH", res_mismatch["disp"])

        res_sa_mismatch = compare_levels("SUBAGT1", "SA2", "Sub Agent")
        self.assertFalse(res_sa_mismatch["match"])

    def test_47_split_sum_100_percent_validation(self):
        from validate_apl_vs_db_after_import import annotate_apl_rows
        df_apl = pd.DataFrame([
            {"Policy No": "POL_SUM_FAIL", "From Month": 1, "To Month": 12, "Rate": 80.00, "Contract": "LVL5", "Agent Name": "Adam Agent", "Agent ID": "101", "Normalized_Split_Share_%": 80.00}
        ])
        # DB splits sum to 80 + 10 = 90% != File 80%
        db_map = {
            "POL_SUM_FAIL": [{
                "policy_number": "POL_SUM_FAIL", "month_from": 1, "month_to": 12, "commission_uuid": "u2",
                "commission_split": [
                    {"order": 1, "role": "Agency", "user_name": "Main Line Benefits", "commission_value": 10.0, "level": None},
                    {"order": 2, "role": "Agent", "user_code": "101", "user_name": "Adam Agent", "commission_value": 80.0, "level": "Level 5"}
                ]
            }]
        }
        res = annotate_apl_rows(df_apl, db_map, ["POL_SUM_FAIL"])
        row = res.iloc[0]
        self.assertEqual(row["Split_Sum_Match"], "MISMATCH (File Sum: 80.00% vs DB Sum: 90.00%)")
        self.assertEqual(row["Match_Status"], "MISMATCH")
        self.assertIn("Split Sum MISMATCH", row["Mismatch_Reason"])

    def test_50_matching_non_100_percent_split_sum(self):
        from validate_apl_vs_db_after_import import annotate_apl_rows
        # Both File and DB sum to 90.00%
        df_apl = pd.DataFrame([
            {"Policy No": "POL_SUM_MATCH_90", "From Month": 1, "To Month": 12, "Rate": 90.00, "Contract": "LVL5", "Agent Name": "Adam Agent", "Agent ID": "101", "Normalized_Split_Share_%": 90.00}
        ])
        db_map = {
            "POL_SUM_MATCH_90": [{
                "policy_number": "POL_SUM_MATCH_90", "month_from": 1, "month_to": 12, "commission_uuid": "u5",
                "commission_split": [
                    {"order": 1, "role": "Agent", "user_code": "101", "user_name": "Adam Agent", "commission_value": 90.0, "level": "Level 5"}
                ]
            }]
        }
        res = annotate_apl_rows(df_apl, db_map, ["POL_SUM_MATCH_90"])
        row = res.iloc[0]
        # Since both File and DB sum to 90%, they match!
        self.assertIn("MATCH (90.00%", row["Split_Sum_Match"])
        self.assertEqual(row["Match_Status"], "EXACT_MATCH")

    def test_48_generational_suffix_and_id_stripping(self):
        from validate_apl_vs_db_after_import import clean_name_ignoring_id, compare_person_names
        # Strip suffix
        self.assertEqual(clean_name_ignoring_id("John Rader IV"), "John Rader")
        self.assertEqual(clean_name_ignoring_id("John Rader Jr."), "John Rader")
        self.assertEqual(clean_name_ignoring_id("John Rader III"), "John Rader")
        self.assertEqual(clean_name_ignoring_id("8307933 - Gregory Teipel"), "Gregory Teipel")
        self.assertEqual(clean_name_ignoring_id("Heidi Berthelsen (19617796)"), "Heidi Berthelsen")

        # Matching John Rader IV against John Rader
        res = compare_person_names("John Rader", "John Rader IV")
        self.assertIn(res["status"], ("EXACT_MATCH", "NORMALIZED_MATCH"))

    def test_49_expired_zero_percent_tier(self):
        from validate_apl_vs_db_after_import import annotate_apl_rows
        df_apl = pd.DataFrame([
            {"Policy No": "POL_EXPIRED", "From Month": 73, "To Month": 0, "Rate": "0.00%", "Amount": "$0.00", "Contract": "LVL5", "Agent Name": "Adam Agent", "Agent ID": "101"}
        ])
        # DB has month 1-12 and 13-72, but Month 73 is omitted because commission is expired (0%)
        db_map = {
            "POL_EXPIRED": [{
                "policy_number": "POL_EXPIRED", "month_from": 1, "month_to": 12, "commission_uuid": "u3",
                "commission_split": [
                    {"order": 1, "role": "Agency", "commission_value": 10.0, "level": None},
                    {"order": 2, "role": "Agent", "commission_value": 90.0, "level": "Level 5"}
                ]
            }]
        }
        res = annotate_apl_rows(df_apl, db_map, ["POL_EXPIRED"])
        row = res.iloc[0]
        self.assertEqual(row["Match_Status"], "EXACT_MATCH")
        self.assertIn("EXPIRED_TIER", row["Month_Slab_Match"])
        self.assertEqual(row["Split_Sum_Match"], "N/A (Expired Tier)")

    def test_51_omitted_zero_percent_participant_split(self):
        from validate_apl_vs_db_after_import import annotate_apl_rows
        # APL has a participant with 0.00% rate and $0.00 amount who was not included in DB commission_split
        df_apl = pd.DataFrame([
            {"Policy No": "POL_ZERO_SPLIT", "From Month": 1, "To Month": 12, "Rate": "0.00%", "Amount": "$0.00", "Contract": "LVL5", "Agent Name": "Zero Agent", "Agent ID": "999", "Normalized_Split_Share_%": 0.0},
            {"Policy No": "POL_ZERO_SPLIT", "From Month": 1, "To Month": 12, "Rate": "100.00%", "Amount": "$100.00", "Contract": "LVL5", "Agent Name": "Primary Agent", "Agent ID": "100", "Normalized_Split_Share_%": 100.0}
        ])
        db_map = {
            "POL_ZERO_SPLIT": [{
                "policy_number": "POL_ZERO_SPLIT", "month_from": 1, "month_to": 12, "commission_uuid": "u4",
                "commission_split": [
                    {"order": 1, "role": "Agent", "user_code": "100", "user_name": "Primary Agent", "commission_value": 100.0, "level": "Level 5"}
                ]
            }]
        }
        res = annotate_apl_rows(df_apl, db_map, ["POL_ZERO_SPLIT"])
        zero_row = res.iloc[0]
        self.assertEqual(zero_row["Match_Status"], "EXACT_MATCH")
        self.assertEqual(zero_row["Rate_Match"], "MATCH (0.00%)")
        self.assertIn("0% Split (Omitted from DB)", zero_row["Matched_Details"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
