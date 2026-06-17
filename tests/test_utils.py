import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import pytest
import pandas as pd
import numpy as np
from datetime import datetime
from utils import (
    extract_schedule_data, find_missing_schedules,
    calculate_schedule_amounts, generate_reconciliation_report,
    is_valid_email, validate_email_list,
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
