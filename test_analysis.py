import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ol_analysis import (REQUIRED_COLUMNS, parse_subject, run_analysis,  # noqa: E402
                         school_markdown, student_workbook_bytes,
                         summary_workbook_bytes, validate_columns)


def make_register() -> pd.DataFrame:
    def row(cid, school, cells, gc, attempt=1):
        r = {"Attempt No": attempt, "Census ID": cid, "School Name": school,
             "Grade Count": gc}
        for col, val in zip(
            ["Religion", "Language & Literature", "English language", "Science",
             "Mathematics", "History", "1st Subject Group", "2nd Subject Group",
             "3rd Subject Group"], cells):
            r[col] = val
        return r

    nine_a = ["11TA9", "21TA9", "31TA9", "34TA9", "32TA9", "33TA9", "60TA9", "40TA9", "41TA9"]
    weak = ["11TS5", "21TS5", "31TW5", "34TW5", "32TW5", "33TW5", "60TW5", "40TW5", "41TW5"]
    absent = ["+", "+", "11TA9", "+", "+", "+", "+", "+", "+"]
    return pd.DataFrame([
        row(1, "School A", nine_a, "9A 0B 0C 0S 0W"),
        row(1, "School A", weak, "0A 0B 0C 2S 7W"),
        row(2, "School B", absent, "1A 0B 0C 0S 0W"),
        row(2, "School B", nine_a, "9A 0B 0C 0S 0W", attempt=2),   # excluded: attempt 2
        row(16179, "Excluded", nine_a, "9A 0B 0C 0S 0W"),            # excluded: census id
    ])


def test_parse_subject():
    assert parse_subject("21TA9") == ("21TA", "21", "A", "9")
    assert parse_subject("") == (None, None, None, None)
    assert parse_subject("+")[2] is None or parse_subject("+")[2] == "+"


def test_validate_columns():
    assert validate_columns(make_register()) == []
    assert "School Name" in validate_columns(pd.DataFrame({"x": [1]}))
    assert set(REQUIRED_COLUMNS) >= {"Census ID", "Grade Count"}


def test_run_analysis_counts():
    res = run_analysis(make_register())
    t = res["totals"]
    assert t["Schools"] == 2
    assert t["App"] == 3          # attempt 2 and census 16179 removed
    assert t["Abs"] == 1
    assert t["Sat"] == 2
    assert t["Q1"] == 1 and t["NQ1"] == 1
    assert t["NineA"] == 1
    assert t["Pct1"] == 50.0


def test_exclusion_toggle():
    res = run_analysis(make_register(), exclude_cid=None)
    assert res["totals"]["App"] == 4


def test_missing_columns_raises():
    with pytest.raises(ValueError):
        run_analysis(pd.DataFrame({"a": [1]}))


def test_outputs_and_exports():
    res = run_analysis(make_register())
    assert not res["subject"].empty
    assert not res["school_pivot"].empty
    assert summary_workbook_bytes(res)[:2] == b"PK"
    assert student_workbook_bytes(res)[:2] == b"PK"
    assert "Muthur Zone" in school_markdown(res)
