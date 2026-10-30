"""
OL 2025 Results Analyser - Muthur Zone (T/Mu)
=============================================
Reusable analysis engine. Used by the Streamlit app (app.py) and can also be
run from the command line to write Excel/Markdown outputs:

    python ol_analysis.py Results_Register-OL2025.xlsx
"""

from __future__ import annotations

import io
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
EXCLUDE_CID = 16179
GRADES = ["A", "B", "C", "S", "W"]
SUBJECTS = [
    "Religion", "Language & Literature", "English language",
    "Science", "Mathematics", "History",
    "1st Subject Group", "2nd Subject Group", "3rd Subject Group",
]
GRADE_COL = "Grade Count"
REQUIRED_COLUMNS = ["Attempt No", "Census ID", "School Name", GRADE_COL, *SUBJECTS]

PASS_I_MIN_CREDITS = 3
PASS_I_MIN_PASSES = 6
PASS_II_MIN_CREDITS = 3
PASS_II_MIN_PASSES = 5

BAND_BINS = [0, 40, 60, 80, 101]
BAND_LABELS = ["<40", "40–59", "60–79", "≥80"]

SUBJECT_NAMES = {
    11: "Buddhism", 12: "Saivanery", 14: "Catholicism",
    15: "Christianity", 16: "Islam",
    21: "Sinhala Language & Literature",
    22: "Tamil Language & Literature",
    31: "English Language",
    32: "Mathematics", 33: "History", 34: "Science",
    40: "Music (Oriental)", 41: "Music (Western)",
    42: "Music (Carnatic)", 43: "Art",
    44: "Dancing (Oriental)", 45: "Dancing (Bharata)",
    46: "Appreciation of English Literary Texts",
    47: "Appreciation of Sinhala Literary Texts",
    48: "Appreciation of Tamil Literary Texts",
    49: "Appreciation of Arabic Literary Texts",
    50: "Drama & Theatre (Sinhala)",
    51: "Drama & Theatre (Tamil)",
    52: "Drama & Theatre (English)",
    60: "Business & Accounting Studies",
    61: "Geography", 62: "Civic Education",
    63: "Entrepreneurship Studies",
    64: "Second Language (Sinhala)",
    65: "Second Language (Tamil)",
    66: "Pali", 67: "Sanskrit", 68: "French",
    69: "German", 70: "Hindi", 71: "Japanese",
    72: "Arabic", 73: "Korean", 74: "Chinese", 75: "Russian",
}

CATEGORY_OF_CODE = {
    **{c: "Religion" for c in [11, 12, 14, 15, 16]},
    **{c: "First Language" for c in [21, 22]},
    31: "English",
    **{c: "Core Compulsory" for c in [32, 33, 34]},
    **{c: "Basket-02" for c in range(40, 53)},
    **{c: "Basket-01" for c in range(60, 76)},
}

CAT_ORDER = ["Religion", "First Language", "English",
             "Core Compulsory", "Basket-01", "Basket-02", "Other"]

SLOT_DEFAULT_CAT = {
    "Religion": "Religion",
    "Language & Literature": "First Language",
    "English language": "English",
    "Science": "Core Compulsory",
    "Mathematics": "Core Compulsory",
    "History": "Core Compulsory",
    "1st Subject Group": "Basket-01",
    "2nd Subject Group": "Basket-02",
    "3rd Subject Group": "Basket-02",
}

PATTERN = re.compile(r"^(\d+)([TS]?)([ABCSW+])$")


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def _norm(v):
    if pd.isna(v):
        return ""
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ").strip().upper())


def parse_subject(val):
    """Return (code+medium+grade, code, grade, sba)."""
    s = _norm(val)
    if not s:
        return None, None, None, None
    sba = s[-1] if s[-1].isdigit() else None
    core = (s[:-1] if sba else s).replace(" ", "")
    m = PATTERN.match(core)
    if not m:
        return core, None, None, sba
    code, medium, grade = m.groups()
    return code + medium + grade, code, grade, sba


def _safe(col: str) -> str:
    return (col.replace(" ", "_").replace("&", "and")
               .replace("1st_", "first_").replace("2nd_", "second_")
               .replace("3rd_", "third_").lower())


def validate_columns(df: pd.DataFrame) -> list[str]:
    return [c for c in REQUIRED_COLUMNS if c not in df.columns]


def _choice_table(sub: pd.DataFrame) -> pd.DataFrame:
    t = (sub.groupby(["code_int", "Subject"], observed=True)
            .agg(Taken=("grade", "size"),
                 A=("grade", lambda s: (s == "A").sum()),
                 B=("grade", lambda s: (s == "B").sum()),
                 C=("grade", lambda s: (s == "C").sum()),
                 S=("grade", lambda s: (s == "S").sum()),
                 W=("grade", lambda s: (s == "W").sum()))
            .reset_index()
            .sort_values("Taken", ascending=False))
    t["Credits"] = t[["A", "B", "C"]].sum(axis=1)
    t["Credit%"] = (t["Credits"] / t["Taken"] * 100).round(1)
    return t


# ---------------------------------------------------------------------------
# MAIN ENGINE
# ---------------------------------------------------------------------------
def run_analysis(df: pd.DataFrame, exclude_cid: int | None = EXCLUDE_CID) -> dict:
    """Run the full analysis. Returns a dict of DataFrames (and a totals dict)."""
    missing = validate_columns(df)
    if missing:
        raise ValueError("Register is missing required column(s): " + ", ".join(missing))

    OL = df[df["Attempt No"] == 1]
    if exclude_cid is not None:
        OL = OL[OL["Census ID"] != exclude_cid]
    OL = OL.copy().reset_index(drop=True)

    # ---- parse subject cells
    for col in SUBJECTS:
        s = _safe(col)
        parsed = OL[col].apply(parse_subject)
        OL[f"{s}_cg"] = parsed.str[0]
        OL[f"{s}_code"] = parsed.str[1]
        OL[f"{s}_grade"] = parsed.str[2]
        OL[f"{s}_sba"] = parsed.str[3]
    grade_cols = [f"{_safe(c)}_grade" for c in SUBJECTS]

    # ---- tallies + grade count validation
    for g in GRADES:
        OL[f"count_{g}"] = (OL[grade_cols] == g).sum(axis=1)
    OL["count_plus"] = (OL[grade_cols] == "+").sum(axis=1)

    gc = OL[GRADE_COL].astype(str).str.findall(r"\d+").apply(
        lambda xs: (xs + ["0"] * 5)[:5])
    gc_df = pd.DataFrame(gc.tolist(), index=OL.index,
                         columns=[f"gc_{g}" for g in GRADES]).astype(int)
    OL = pd.concat([OL, gc_df], axis=1)
    for g in GRADES:
        OL[f"check_{g}"] = OL[f"count_{g}"] == OL[f"gc_{g}"]
    OL["GradeCount_OK"] = OL[[f"check_{g}" for g in GRADES]].all(axis=1)

    # ---- pass logic
    lang = OL["language_and_literature_grade"]
    maths = OL["mathematics_grade"]
    lang_pass = lang.isin(["A", "B", "C", "S"])
    maths_pass = maths.isin(["A", "B", "C", "S"])
    lang_absent = lang.isna() | (lang == "+")
    maths_absent = maths.isna() | (maths == "+")
    both_absent = lang_absent & maths_absent

    credits = OL["count_A"] + OL["count_B"] + OL["count_C"]
    total_passes = credits + OL["count_S"]
    OL["Credits"] = credits
    OL["Passes"] = total_passes

    OL["Pass_I"] = np.where(
        ~both_absent & (credits >= PASS_I_MIN_CREDITS)
        & (total_passes >= PASS_I_MIN_PASSES) & lang_pass & maths_pass,
        "P", "F")
    OL["Pass_II"] = np.where(
        ~both_absent & (credits >= PASS_II_MIN_CREDITS)
        & (total_passes >= PASS_II_MIN_PASSES) & (lang_pass | maths_pass),
        "P", "F")
    OL.loc[both_absent, ["Pass_I", "Pass_II"]] = np.nan

    # ---- diagnose Pass II delta vs register Grade Count
    credits_gc = OL["gc_A"] + OL["gc_B"] + OL["gc_C"]
    passes_gc = credits_gc + OL["gc_S"]
    pass2_gc = (~both_absent & (credits_gc >= PASS_II_MIN_CREDITS)
                & (passes_gc >= PASS_II_MIN_PASSES) & (lang_pass | maths_pass))
    discrep = OL[pass2_gc & (OL["Pass_II"] == "F")]
    mism = OL[~OL["GradeCount_OK"]]
    diagnostics = {
        "pass2_parsed": int((OL["Pass_II"] == "P").sum()),
        "pass2_register": int(pass2_gc.sum()),
        "plus_per_subject": {c: int((OL[f"{_safe(c)}_grade"] == "+").sum()) for c in SUBJECTS},
        "mismatch_per_grade": {g: int((~OL[f"check_{g}"]).sum()) for g in GRADES},
    }

    # ---- school summary
    grp = OL.groupby(["Census ID", "School Name"], as_index=False)
    school = grp.agg(
        App=("Census ID", "size"),
        Sat=("Pass_I", lambda s: s.notna().sum()),
        Abs=("Pass_I", lambda s: s.isna().sum()),
        Q1=("Pass_I", lambda s: (s == "P").sum()),
        NQ1=("Pass_I", lambda s: (s == "F").sum()),
        Q2=("Pass_II", lambda s: (s == "P").sum()),
        NQ2=("Pass_II", lambda s: (s == "F").sum()),
        NineA=("count_A", lambda s: (s == 9).sum()),
        NineF=("count_W", lambda s: (s == 9).sum()),
    )
    denom = (school["Q1"] + school["NQ1"]).replace(0, np.nan)
    school["ratio1"] = school["Q1"] / denom
    school["ratio2"] = school["Q2"] / denom
    school["Pct1"] = (school["ratio1"] * 100).round(2)
    school["Pct2"] = (school["ratio2"] * 100).round(2)
    school["Rk1"] = school["ratio1"].rank(ascending=False, method="min").astype("Int64")
    school["Rk2"] = school["ratio2"].rank(ascending=False, method="min").astype("Int64")
    school["Delta"] = school["Q2"] - school["Q1"]
    school["Band1"] = pd.cut(school["Pct1"], BAND_BINS, labels=BAND_LABELS, right=False)
    school["Band2"] = pd.cut(school["Pct2"], BAND_BINS, labels=BAND_LABELS, right=False)
    school = school[[
        "Census ID", "School Name", "App", "Sat", "Abs", "NineA", "NineF",
        "Q1", "NQ1", "Pct1", "Rk1", "Q2", "NQ2", "Pct2", "Rk2", "Delta",
        "Band1", "Band2",
    ]].sort_values("Rk1").reset_index(drop=True)

    # ---- zone totals
    totals = {k: int(school[k].sum()) for k in
              ["App", "Sat", "Abs", "NineA", "NineF", "Q1", "NQ1", "Q2", "NQ2"]}
    totals = {"Schools": len(school), **totals}
    d1 = totals["Q1"] + totals["NQ1"]
    d2 = totals["Q2"] + totals["NQ2"]
    totals["Pct1"] = round(totals["Q1"] / d1 * 100, 2) if d1 else 0.0
    totals["Pct2"] = round(totals["Q2"] / d2 * 100, 2) if d2 else 0.0
    totals["Delta"] = totals["Q2"] - totals["Q1"]

    bands = pd.DataFrame({
        "Band": BAND_LABELS,
        "Pass I": [int((school["Band1"] == b).sum()) for b in BAND_LABELS],
        "Pass II": [int((school["Band2"] == b).sum()) for b in BAND_LABELS],
    })

    # ---- long format (candidate x subject)
    parts = []
    for slot in SUBJECTS:
        s = _safe(slot)
        parts.append(pd.DataFrame({
            "Census ID": OL["Census ID"].values,
            "School Name": OL["School Name"].values,
            "code": OL[f"{s}_code"].values,
            "grade": OL[f"{s}_grade"].values,
            "sba": OL[f"{s}_sba"].values,
            "slot": slot,
        }))
    long = pd.concat(parts, ignore_index=True)
    long["code_int"] = pd.to_numeric(long["code"], errors="coerce")

    def _classify(r):
        g = r["grade"]
        if pd.notna(g) and g in GRADES:
            return "real"
        if g == "+":
            return "plus_set"
        if pd.isna(r["code"]) and pd.isna(g):
            return "blank"
        return "parse_fail"

    long["status"] = long.apply(_classify, axis=1)

    real = long[long["status"] == "real"].copy()
    real["Subject"] = real["code_int"].map(SUBJECT_NAMES)
    real["Subject"] = real.apply(
        lambda r: r["Subject"] if pd.notna(r["Subject"])
        else (f"Code {int(r['code_int'])} (unknown)" if pd.notna(r["code_int"]) else "Unknown"),
        axis=1)
    real["Category"] = real["code_int"].map(CATEGORY_OF_CODE)
    real["Category"] = real["Category"].fillna(real["slot"].map(SLOT_DEFAULT_CAT)).fillna("Other")
    real["Category"] = pd.Categorical(real["Category"], categories=CAT_ORDER, ordered=True)
    real["sba_num"] = pd.to_numeric(real["sba"], errors="coerce")

    def _grade_agg(prefix_size="Sat"):
        return dict(
            **{prefix_size: ("grade", "size")},
            A=("grade", lambda s: (s == "A").sum()),
            B=("grade", lambda s: (s == "B").sum()),
            C=("grade", lambda s: (s == "C").sum()),
            S=("grade", lambda s: (s == "S").sum()),
            W=("grade", lambda s: (s == "W").sum()),
        )

    subject_tbl = (real.groupby(["Category", "code_int", "Subject"], observed=True)
                   .agg(**_grade_agg(), Avg_SBA=("sba_num", "mean")).reset_index())
    subject_tbl["Credits"] = subject_tbl[["A", "B", "C"]].sum(axis=1)
    subject_tbl["Passes"] = subject_tbl["Credits"] + subject_tbl["S"]
    subject_tbl["Credit%"] = (subject_tbl["Credits"] / subject_tbl["Sat"] * 100).round(1)
    subject_tbl["Pass%"] = (subject_tbl["Passes"] / subject_tbl["Sat"] * 100).round(1)
    subject_tbl["Avg_SBA"] = subject_tbl["Avg_SBA"].round(1)
    subject_tbl = subject_tbl.sort_values(["Category", "Sat"], ascending=[True, False])

    cat_tbl = (real.groupby("Category", observed=True)
               .agg(Subjects=("code_int", "nunique"), **_grade_agg()).reset_index())
    cat_tbl["Credits"] = cat_tbl[["A", "B", "C"]].sum(axis=1)
    cat_tbl["Passes"] = cat_tbl["Credits"] + cat_tbl["S"]
    cat_tbl["Credit%"] = (cat_tbl["Credits"] / cat_tbl["Sat"] * 100).round(1)
    cat_tbl["Pass%"] = (cat_tbl["Passes"] / cat_tbl["Sat"] * 100).round(1)

    notset = (long[long["status"] != "real"]
              .assign(Category=lambda d: d["code_int"].map(CATEGORY_OF_CODE)
                      .fillna(d["slot"].map(SLOT_DEFAULT_CAT)).fillna("Other"))
              .groupby(["slot", "Category", "status"], observed=True)
              .size().reset_index(name="N"))

    religion = _choice_table(real[real["Category"] == "Religion"])
    basket1 = _choice_table(real[real["Category"] == "Basket-01"])
    basket2 = _choice_table(real[real["Category"] == "Basket-02"])

    school_subject = (real.groupby(["School Name", "Subject", "Category"], observed=True)
                      .agg(Sat=("grade", "size"),
                           A=("grade", lambda s: (s == "A").sum()),
                           B=("grade", lambda s: (s == "B").sum()),
                           C=("grade", lambda s: (s == "C").sum()),
                           S=("grade", lambda s: (s == "S").sum()),
                           W=("grade", lambda s: (s == "W").sum()))
                      .reset_index())
    school_subject["Passes"] = school_subject[["A", "B", "C", "S"]].sum(axis=1)
    school_subject["Credits"] = school_subject[["A", "B", "C"]].sum(axis=1)
    school_subject["Pass%"] = (school_subject["Passes"] / school_subject["Sat"] * 100).round(1)
    school_subject["Credit%"] = (school_subject["Credits"] / school_subject["Sat"] * 100).round(1)

    school_pivot = school_subject.pivot_table(
        index="School Name", columns="Subject", values="Sat",
        aggfunc="sum", fill_value=0)

    return {
        "students": OL, "mismatches": mism, "pass2_delta": discrep,
        "school": school, "totals": totals, "bands": bands,
        "subject": subject_tbl, "category": cat_tbl, "notset": notset,
        "religion": religion, "basket1": basket1, "basket2": basket2,
        "school_subject": school_subject, "school_pivot": school_pivot,
        "diagnostics": diagnostics,
    }


# ---------------------------------------------------------------------------
# EXPORT HELPERS
# ---------------------------------------------------------------------------
def _clean_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if isinstance(out[c].dtype, pd.CategoricalDtype):
            out[c] = out[c].astype(str)
    return out


def summary_workbook_bytes(res: dict) -> bytes:
    """School + subject summaries (no student-level data) as one .xlsx."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        _clean_for_excel(res["school"]).to_excel(w, "By School", index=False)
        pd.DataFrame([res["totals"]]).to_excel(w, "Totals", index=False)
        res["bands"].to_excel(w, "Bands", index=False)
        _clean_for_excel(res["subject"]).to_excel(w, "By Subject", index=False)
        _clean_for_excel(res["category"]).to_excel(w, "By Category", index=False)
        res["religion"].to_excel(w, "Religion Choice", index=False)
        res["basket1"].to_excel(w, "Basket-01 Choice", index=False)
        res["basket2"].to_excel(w, "Basket-02 Choice", index=False)
        res["school_pivot"].to_excel(w, "School x Subject")
        _clean_for_excel(res["school_subject"]).to_excel(w, "School Subject Detail", index=False)
        res["notset"].to_excel(w, "Not-Set", index=False)
    return buf.getvalue()


def student_workbook_bytes(res: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        _clean_for_excel(res["students"]).to_excel(w, "All Candidates", index=False)
        if len(res["mismatches"]):
            _clean_for_excel(res["mismatches"]).to_excel(w, "GradeCount Mismatches", index=False)
        if len(res["pass2_delta"]):
            _clean_for_excel(res["pass2_delta"]).to_excel(w, "PassII Delta", index=False)
    return buf.getvalue()


def school_markdown(res: dict) -> str:
    totals = ", ".join(f"{k}={v}" for k, v in res["totals"].items())
    return "\n".join([
        "## Muthur Zone — O/L 2025 Summary", "",
        _clean_for_excel(res["school"]).to_markdown(index=False, floatfmt=".2f"), "",
        "**Zone totals:** " + totals, "",
        res["bands"].to_markdown(index=False),
    ])


if __name__ == "__main__":
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("Results_Register-OL2025.xlsx")
    result = run_analysis(pd.read_excel(path))
    Path("OL_school_summary.xlsx").write_bytes(summary_workbook_bytes(result))
    Path("OL_student_level.xlsx").write_bytes(student_workbook_bytes(result))
    Path("OL_school_summary.md").write_text(school_markdown(result), encoding="utf-8")
    print("Wrote OL_school_summary.xlsx, OL_student_level.xlsx, OL_school_summary.md")
    print(result["totals"])
