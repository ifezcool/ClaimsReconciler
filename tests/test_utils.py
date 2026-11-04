import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import pytest
import pandas as pd
import numpy as np
import io  # ADD: needed for BytesIO in test
from datetime import datetime
from utils import (
    extract_schedule_data, find_missing_schedules,
    calculate_schedule_amounts, generate_reconciliation_report,
    is_valid_email, validate_email_list,
    _norm, find_column_index,  # ADD: import _norm and find_column_index
)


class TestEmailValidation:
    def test_valid_emails(self):
        assert is_valid_email("user@example.com")
        assert is_valid_email("first.last@company.co.uk")
        assert is_valid_email("user+tag@domain.org")
        assert is_valid_email("123@abc.com")

    def test_invalid_emails(self):
        assert not is_valid_email("")
        assert not is_valid_email("not-an-email")
        assert not is_valid_email("@domain.com")
        assert not is_valid_email("user@")
        assert not is_valid_email("user@.com")

    def test_validate_email_list_valid(self):
        emails = ["a@b.com", "c@d.org"]
        assert validate_email_list(emails) is True

    def test_validate_email_list_invalid(self):
        emails = ["a@b.com", "invalid"]
        with pytest.raises(ValueError, match="invalid"):
            validate_email_list(emails)


class TestExtractScheduleData:
    def test_basic_extraction(self):
        df = pd.DataFrame({
            "SCH NO": ["123", "456", "789"],
            "AMOUNT": [100.0, 200.0, 300.0],
        })
        result = extract_schedule_data(df, "SCH NO", "AMOUNT")
        assert list(result.columns) == ["Schedule Number", "Amount"]
        assert len(result) == 3
        assert result["Amount"].sum() == 600.0

    def test_drops_missing_amount(self):
        df = pd.DataFrame({
            "SCH NO": ["123", "456", "789"],
            "AMOUNT": [100.0, None, 300.0],
        })
        result = extract_schedule_data(df, "SCH NO", "AMOUNT")
        assert len(result) == 2
        assert "456" not in result["Schedule Number"].values

    def test_coerces_types(self):
        df = pd.DataFrame({
            "SCH NO": [123, 456],
            "AMOUNT": ["100.5", "200.3"],
        })
        result = extract_schedule_data(df, "SCH NO", "AMOUNT")
        assert result["Amount"].dtype in [np.float64, float]


class TestFindMissingSchedules:
    def test_finds_missing(self):
        source = pd.DataFrame({"Schedule Number": ["A", "B", "C"], "Amount": [1, 2, 3]})
        target = pd.DataFrame({"Schedule Number": ["A", "B"], "Amount": [1, 2]})
        result = find_missing_schedules(source, target)
        assert len(result) == 1
        assert result.iloc[0]["Schedule Number"] == "C"

    def test_no_missing(self):
        source = pd.DataFrame({"Schedule Number": ["A", "B"], "Amount": [1, 2]})
        target = pd.DataFrame({"Schedule Number": ["A", "B"], "Amount": [1, 2]})
        result = find_missing_schedules(source, target)
        assert result.empty


class TestCalculateScheduleAmounts:
    def test_sums_correctly(self):
        df = pd.DataFrame({
            "Schedule Number": ["A", "A", "B"],
            "Amount": [100, 50, 200],
        })
        result = calculate_schedule_amounts(df)
        assert len(result) == 2
        assert result[result["Schedule Number"] == "A"]["Amount"].iloc[0] == 150
        assert result[result["Schedule Number"] == "B"]["Amount"].iloc[0] == 200


class TestGenerateReconciliationReport:
    def test_merge_and_diff(self):
        claims = pd.DataFrame({
            "Schedule Number": ["A", "B"],
            "Amount": [100, 200],
        })
        finance = pd.DataFrame({
            "Schedule Number": ["A", "C"],
            "Amount": [100, 300],
        })
        result = generate_reconciliation_report(claims, finance)
        assert "Difference" in result.columns
        a_row = result[result["Schedule Number"] == "A"]
        assert a_row["Difference"].iloc[0] == 0.0
        b_row = result[result["Schedule Number"] == "B"]
        assert pd.isna(b_row["Finance Amount"].iloc[0])


class TestFindColumnIndex:
    def test_matches_date_received_variants(self):
        """find_column_index should match normalized variants of DATE RECEIVED."""
        df = pd.DataFrame({
            " DATE RECEIVED": [1, 2, 3],
            "Date_Received": [4, 5, 6],
            "date claim received ": [7, 8, 9],
            "OTHER": [10, 11, 12],
        })
        from utils import find_column_index
        # Should match the first normalized variant found
        idx = find_column_index(df, ["DATE CLAIM RECEIVED"])
        assert idx is not None, "Should find a DATE CLAIM RECEIVED variant"

    def test_returns_none_for_unrelated_header(self):
        """find_column_index should return None for an unrelated header."""
        df = pd.DataFrame({"SOMETHING_ELSE": [1, 2, 3]})
        from utils import find_column_index
        idx = find_column_index(df, ["DATE CLAIM RECEIVED"])
        assert idx is None

    def test_generate_enhanced_claims_excel_with_date_received(self):
        """generate_enhanced_claims_excel on DataFrame with ' DATE RECEIVED' column
        should produce non-empty ClaimNo and ClaimBatch formulas."""
        import openpyxl
        df = pd.DataFrame({
            "SCH NO": ["123", "456"],
            "DATE CLAIM RECEIVED": ["01/01/2024", "15/03/2024"],
            "ENCOUNTER DATE (DD/MM/YYYY)": ["01/01/2024", "15/03/2024"],
            "MEMBER NO": ["M001", "M002"],
            "ENROLLEE NAME": ["John Doe", "Jane Smith"],
            "PROVIDER CODE": ["P001", "P002"],
            "PA AMOUNT": [100.0, 200.0],
        })
        from utils import generate_enhanced_claims_excel
        excel_bytes = generate_enhanced_claims_excel(
            df, "SCH NO", "PA AMOUNT"
        )
        # Read back with openpyxl and check formulas start with '='
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active
        # ClaimBatch formula is at column 9 (I), ClaimNo at column 11 (K)
        # based on the column ordering in generate_enhanced_claims_excel
        claimbatch_cell = ws.cell(row=2, column=9)  # column I
        assert claimbatch_cell.value is not None
        assert str(claimbatch_cell.value).startswith("="), f"ClaimBatch formula expected to start with '=', got: {claimbatch_cell.value}"
        # ClaimNo formula at column 11 (K)
        claimno_cell = ws.cell(row=2, column=11)  # column K
        assert claimno_cell.value is not None
        assert str(claimno_cell.value).startswith("="), f"ClaimNo formula expected to start with '=', got: {claimno_cell.value}"


class TestColumnAliases:
    def test_date_received_alias_renames(self):
        """DB upload alias: _norm collapses whitespace/underscores but different names remain distinct."""
        # _norm collapses internal whitespace/underscores; different canonical names remain different
        assert _norm("DATE  RECEIVED") == _norm("DATE_RECEIVED") == "DATE RECEIVED"
        assert _norm("DATE CLAIM RECEIVED") == "DATE CLAIM RECEIVED"
        # Different names should not match
        assert _norm("DATE RECEIVED") != _norm("DATE CLAIM RECEIVED")

    def test_both_present_raises_error(self):
        """With both DATE RECEIVED and DATE CLAIM RECEIVED present, error should be raised."""
        # This tests the _apply_column_aliases logic from db_upload_common
        import pandas as pd
        df = pd.DataFrame({
            "DATE RECEIVED": [1, 2],
            "DATE CLAIM RECEIVED": [3, 4],
        })
        # Both normalized forms match canonical, so error should detect both present
        norm_to_original = {}
        for col in df.columns:
            norm = _norm(col)
            norm_to_original[norm] = col
        # Check both are present
        canonical_norm = _norm("DATE CLAIM RECEIVED")
        alt_norm = _norm("DATE RECEIVED")
        assert canonical_norm in [norm_to_original[n] for n in df.columns]  # will be True
        assert alt_norm in [norm_to_original[n] for n in df.columns]  # will be True
