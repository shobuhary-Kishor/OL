"""Muthur Zone GEC(O/L) Dashboard - Streamlit app."""

from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from ol_analysis import (BAND_LABELS, GRADES, run_analysis, school_markdown,
                         student_workbook_bytes, summary_workbook_bytes,
                         validate_columns)

# ==========================================
# PAGE CONFIG
# ==========================================
st.set_page_config(
    page_title="Muthur Zone GEC(O/L) Dashboard",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.markdown(
    """
    <style>
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; padding-left: 20px; padding-right: 20px; }
    </style>
    """,
    unsafe_allow_html=True,
)

_matches = sorted(Path(".").glob("Results_Register-*.xlsx"))
DEFAULT_REGISTER = _matches[0] if _matches else Path("Results_Register.xlsx")
GRADE_COLORS = {"A": "#2ca02c", "B": "#98df8a", "C": "#aec7e8",
                "S": "#ffbb78", "W": "#d62728"}


# ==========================================
# OPTIONAL PASSWORD GATE (set APP_PASSWORD in Streamlit secrets)
# ==========================================
def check_password() -> bool:
    try:
        expected = st.secrets.get("APP_PASSWORD")
    except Exception:
        expected = None
    if not expected:
        return True
    if st.session_state.get("authed"):
        return True
    st.title("🔒 Muthur Zone GEC(O/L)")
    pw = st.text_input("Password", type="password")
    if pw:
        if pw == expected:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


if not check_password():
    st.stop()


# ==========================================
# DATA LOADING (upload or local file)
# ==========================================
@st.cache_data(show_spinner="Reading register…")
def read_register(source) -> pd.DataFrame:
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    return pd.read_excel(source)


@st.cache_data(show_spinner="Analysing results…")
def analyse(df: pd.DataFrame, exclude_cid: int | None):
    return run_analysis(df, exclude_cid)


st.sidebar.title("🎓 Muthur Zone GEC(O/L)")
st.sidebar.markdown("---")

uploaded = st.sidebar.file_uploader(
    "Upload Results Register (.xlsx)", type=["xlsx"],
    help="The file is processed in memory and is not stored.")
use_exclude = st.sidebar.checkbox("Exclude Census ID 16179", value=True)
exclude_cid = 16179 if use_exclude else None

if uploaded is not None:
    raw = read_register(uploaded.getvalue())
elif DEFAULT_REGISTER.exists():
    raw = read_register(DEFAULT_REGISTER)
    st.sidebar.caption(f"Using local file: {DEFAULT_REGISTER.name}")
else:
    st.title("🎓 Muthur Zone GEC(O/L) Dashboard")
    st.info("👈 Upload the **Results Register (.xlsx)** in the sidebar to begin.")
    st.stop()

missing = validate_columns(raw)
if missing:
    st.error("The register is missing required column(s): " + ", ".join(missing))
    st.stop()

res = analyse(raw, exclude_cid)
school = res["school"]
totals = res["totals"]

st.sidebar.markdown("---")
page = st.sidebar.radio("📊 Dashboard Sections", [
    "Zone Overview",
    "School Performance",
    "School Drill-down",
    "Compare Schools",
    "Subject Analysis",
    "Basket & Religion Choices",
    "Student Explorer",
    "Data Quality & Diagnostics",
    "Downloads",
])


def fmt_table(df: pd.DataFrame, formats: dict):
    existing = {k: v for k, v in formats.items() if k in df.columns}
    return df.style.format(existing, na_rep="–")


# ==========================================
# PAGE: Zone Overview
# ==========================================
if page == "Zone Overview":
    st.header("🌍 Zone Overview")
    st.caption(f"{totals['Schools']} schools · first-attempt candidates only")

    c = st.columns(5)
    c[0].metric("Total Applicants", f"{totals['App']:,}")
    c[1].metric("Satisfied", f"{totals['Sat']:,}")
    c[2].metric("Absent", f"{totals['Abs']:,}")
    c[3].metric("9 A's", f"{totals['NineA']:,}")
    c[4].metric("9 F's / W's", f"{totals['NineF']:,}")

    c = st.columns(4)
    c[0].metric("Pass I Rate", f"{totals['Pct1']:.1f}%")
    c[1].metric("Pass II Rate", f"{totals['Pct2']:.1f}%")
    c[2].metric("Qualified (Pass I)", f"{totals['Q1']:,}")
    c[3].metric("Qualified (Pass II)", f"{totals['Q2']:,}")

    st.markdown("---")
    left, right = st.columns(2)

    with left:
        st.subheader("📈 Performance Bands")
        bands = res["bands"]
        fig = go.Figure()
        fig.add_bar(x=bands["Band"], y=bands["Pass I"],
                    name="Pass I", marker_color="#1f77b4")
        fig.add_bar(x=bands["Band"], y=bands["Pass II"],
                    name="Pass II", marker_color="#ff7f0e")
        fig.update_layout(barmode="group", xaxis_title="Band (%)",
                          yaxis_title="Number of Schools")
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("🎯 Zone Grade Distribution")
        cat = res["category"]
        gsum = {g: int(cat[g].sum()) for g in GRADES}
        fig = px.pie(names=list(gsum), values=list(gsum.values()), hole=0.45,
                     color=list(gsum), color_discrete_map=GRADE_COLORS)
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("🏆 Top & Bottom Schools (Pass I %)")
    n = st.slider("Schools to show", 3, 15, 5)
    ranked = school.dropna(subset=["Pct1"])
    t1, t2 = st.columns(2)
    with t1:
        st.markdown("**Top**")
        st.dataframe(ranked.nsmallest(n, "Rk1")[["Rk1", "School Name", "App", "Pct1", "Pct2"]],
                     hide_index=True, use_container_width=True)
    with t2:
        st.markdown("**Bottom**")
        st.dataframe(ranked.nlargest(n, "Rk1")[["Rk1", "School Name", "App", "Pct1", "Pct2"]],
                     hide_index=True, use_container_width=True)

# ==========================================
# PAGE: School Performance
# ==========================================
elif page == "School Performance":
    st.header("🏫 School Performance")

    st.sidebar.markdown("---")
    st.sidebar.subheader("Filters")
    band_opts = [b for b in BAND_LABELS if b in set(
        school["Band1"].dropna().astype(str))]
    band_filter = st.sidebar.multiselect(
        "Pass I Band", band_opts, default=band_opts)
    min_app = st.sidebar.number_input(
        "Minimum applicants", 0, int(school["App"].max()), 0)
    name_q = st.sidebar.text_input("School name contains")

    f = school[school["Band1"].astype(str).isin(
        band_filter) | school["Band1"].isna()]
    f = f[f["App"] >= min_app]
    if name_q:
        f = f[f["School Name"].str.contains(name_q, case=False, na=False)]
    st.caption(f"Showing {len(f)} of {len(school)} schools")

    st.subheader("Pass I vs Pass II Qualification Rates")
    plot_df = f.dropna(subset=["Pct1", "Pct2"]).assign(
        Band1=lambda d: d["Band1"].astype(str))
    fig = px.scatter(plot_df, x="Pct1", y="Pct2", size="App", color="Band1",
                     hover_name="School Name", hover_data=["Q1", "Q2", "Rk1", "Rk2"],
                     labels={"Pct1": "Pass I %",
                             "Pct2": "Pass II %", "App": "Applicants"},
                     category_orders={"Band1": BAND_LABELS})
    fig.add_shape(type="line", x0=0, x1=100, y0=0, y1=100,
                  line=dict(dash="dash", color="gray"))
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Ranking")
    metric = st.radio("Rank by", ["Pct1", "Pct2"], horizontal=True,
                      format_func=lambda x: "Pass I %" if x == "Pct1" else "Pass II %")
    rk = f.dropna(subset=[metric]).sort_values(metric)
    fig = px.bar(rk, x=metric, y="School Name", orientation="h", color=metric,
                 color_continuous_scale="RdYlGn", range_color=[0, 100])
    fig.update_layout(height=max(400, 22 * len(rk)), yaxis_title=None)
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("School × Subject Candidate Distribution")
    pivot = res["school_pivot"]
    pivot = pivot[pivot.index.isin(f["School Name"])]
    fig = px.imshow(pivot, aspect="auto", color_continuous_scale="YlOrRd",
                    labels=dict(x="Subject", y="School", color="Candidates"))
    fig.update_layout(height=max(500, 22 * len(pivot)))
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("School Summary Table")
    st.dataframe(
        fmt_table(f, {"Pct1": "{:.1f}%", "Pct2": "{:.1f}%", "App": "{:,}", "Sat": "{:,}",
                      "Abs": "{:,}", "NineA": "{:,}", "NineF": "{:,}", "Q1": "{:,}",
                      "NQ1": "{:,}", "Q2": "{:,}", "NQ2": "{:,}"}),
        use_container_width=True, hide_index=True)
    st.download_button("⬇️ Download filtered table (CSV)",
                       f.to_csv(index=False).encode("utf-8"), "school_summary_filtered.csv")

# ==========================================
# PAGE: School Drill-down
# ==========================================
elif page == "School Drill-down":
    st.header("🔎 School Drill-down")
    name = st.selectbox("Choose a school", school["School Name"].tolist())
    row = school[school["School Name"] == name].iloc[0]

    c = st.columns(5)
    c[0].metric("Applicants", f"{int(row['App']):,}")
    c[1].metric("Pass I", f"{row['Pct1']:.1f}%" if pd.notna(row["Pct1"]) else "–",
                f"Rank {row['Rk1']}" if pd.notna(row["Rk1"]) else None, delta_color="off")
    c[2].metric("Pass II", f"{row['Pct2']:.1f}%" if pd.notna(row["Pct2"]) else "–",
                f"Rank {row['Rk2']}" if pd.notna(row["Rk2"]) else None, delta_color="off")
    c[3].metric("9 A's", int(row["NineA"]))
    c[4].metric("Absent", int(row["Abs"]))

    zone_gap = (row["Pct1"] - totals["Pct1"]
                ) if pd.notna(row["Pct1"]) else None
    if zone_gap is not None:
        st.caption(
            f"Pass I is {zone_gap:+.1f} pts vs zone average ({totals['Pct1']:.1f}%).")

    ss = res["school_subject"]
    ss = ss[ss["School Name"] == name].sort_values("Sat", ascending=False)
    st.subheader("Subject-wise results")
    zone_subj = res["subject"].set_index("Subject")["Credit%"]
    ss = ss.assign(Zone_Credit=ss["Subject"].map(zone_subj))
    fig = go.Figure()
    fig.add_bar(x=ss["Subject"], y=ss["Credit%"],
                name="This school", marker_color="#1f77b4")
    fig.add_bar(x=ss["Subject"], y=ss["Zone_Credit"],
                name="Zone", marker_color="#bbbbbb")
    fig.update_layout(barmode="group", yaxis_title="Credit %",
                      xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)

    fig = px.bar(ss.melt(id_vars="Subject", value_vars=GRADES, var_name="Grade", value_name="N"),
                 x="Subject", y="N", color="Grade", color_discrete_map=GRADE_COLORS,
                 title="Grade distribution by subject")
    fig.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(ss.drop(columns=["School Name"]),
                 hide_index=True, use_container_width=True)

# ==========================================
# PAGE: Compare Schools
# ==========================================
elif page == "Compare Schools":
    st.header("⚖️ Compare Schools")
    names = st.multiselect("Select 2–6 schools", school["School Name"].tolist(),
                           default=school["School Name"].head(3).tolist(), max_selections=6)
    if len(names) < 2:
        st.info("Pick at least two schools.")
    else:
        sub = school[school["School Name"].isin(names)]
        long = sub.melt(id_vars="School Name", value_vars=["Pct1", "Pct2"],
                        var_name="Metric", value_name="Percent")
        long["Metric"] = long["Metric"].map(
            {"Pct1": "Pass I %", "Pct2": "Pass II %"})
        st.plotly_chart(px.bar(long, x="School Name", y="Percent", color="Metric", barmode="group"),
                        use_container_width=True)

        ss = res["school_subject"]
        ss = ss[ss["School Name"].isin(names)]
        subj_opts = sorted(ss["Subject"].unique())
        pick = st.multiselect("Subjects", subj_opts, default=subj_opts[:8])
        ss = ss[ss["Subject"].isin(pick)]
        st.plotly_chart(
            px.bar(ss, x="Subject", y="Credit%",
                   color="School Name", barmode="group"),
            use_container_width=True)
        radar = ss.pivot_table(
            index="Subject", columns="School Name", values="Credit%")
        if len(radar) >= 3:
            fig = go.Figure()
            for s in radar.columns:
                fig.add_trace(go.Scatterpolar(r=radar[s].fillna(0).tolist() + [radar[s].fillna(0).iloc[0]],
                                              theta=radar.index.tolist() +
                                              [radar.index[0]],
                                              name=s, fill="toself", opacity=0.5))
            fig.update_layout(polar=dict(radialaxis=dict(range=[0, 100])))
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(sub.drop(columns=["Census ID"]),
                     hide_index=True, use_container_width=True)

# ==========================================
# PAGE: Subject Analysis
# ==========================================
elif page == "Subject Analysis":
    st.header("📚 Subject Analysis")
    tab1, tab2, tab3 = st.tabs(
        ["All Subjects", "Category Totals", "Grade Distribution"])
    subjects = res["subject"].assign(
        Category=lambda d: d["Category"].astype(str))

    with tab1:
        cats = st.multiselect("Categories", subjects["Category"].unique().tolist(),
                              default=subjects["Category"].unique().tolist())
        s = subjects[subjects["Category"].isin(cats)]
        fig = px.bar(s, x="Subject", y=["Credit%", "Pass%"], barmode="group",
                     labels={"value": "Percentage (%)", "variable": "Metric"})
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(fmt_table(s, {"Credit%": "{:.1f}%", "Pass%": "{:.1f}%", "Avg_SBA": "{:.1f}"}),
                     use_container_width=True, hide_index=True)

    with tab2:
        categories = res["category"].assign(
            Category=lambda d: d["Category"].astype(str))
        fig = px.bar(categories, x="Category", y=["Credit%", "Pass%"], barmode="group",
                     labels={"value": "Percentage (%)", "variable": "Metric"})
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(fmt_table(categories, {"Credit%": "{:.1f}%", "Pass%": "{:.1f}%"}),
                     use_container_width=True, hide_index=True)

    with tab3:
        top = st.slider("Number of subjects (by candidates)", 5, 40, 15)
        s = subjects.nlargest(top, "Sat")
        m = s.melt(id_vars="Subject", value_vars=GRADES,
                   var_name="Grade", value_name="N")
        fig = px.bar(m, x="Subject", y="N", color="Grade", color_discrete_map=GRADE_COLORS,
                     barmode="stack")
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
        if subjects["Avg_SBA"].notna().any():
            st.subheader("Average SBA mark vs Credit %")
            st.plotly_chart(px.scatter(subjects, x="Avg_SBA", y="Credit%", size="Sat",
                                       hover_name="Subject", color="Category"),
                            use_container_width=True)

# ==========================================
# PAGE: Basket & Religion Choices
# ==========================================
elif page == "Basket & Religion Choices":
    st.header("🧺 Subject Choices (Baskets & Religion)")
    tabs = st.tabs(["Religion", "Basket-01", "Basket-02"])
    for tab, key, title in zip(tabs, ["religion", "basket1", "basket2"],
                               ["Religion", "Basket-01", "Basket-02"]):
        with tab:
            choices = res[key]
            if choices.empty:
                st.info("No data for this group.")
                continue
            c1, c2 = st.columns(2)
            top = choices.head(10)
            c1.plotly_chart(px.bar(top, x="Subject", y="Taken", title=f"Top 10 in {title}",
                                   labels={"Taken": "Candidates"}), use_container_width=True)
            c2.plotly_chart(px.pie(choices, names="Subject", values="Taken",
                                   title="Share of candidates", hole=0.4),
                            use_container_width=True)
            st.dataframe(fmt_table(choices, {"Credit%": "{:.1f}%"}),
                         use_container_width=True, hide_index=True)

# ==========================================
# PAGE: Student Explorer
# ==========================================
elif page == "Student Explorer":
    st.header("🧑‍🎓 Student Explorer")
    st.warning("This page shows individual candidate records. Protect access with "
               "`APP_PASSWORD` when deploying publicly.")
    stu = res["students"]

    c = st.columns(4)
    sch = c[0].multiselect("School", sorted(stu["School Name"].unique()))
    p1 = c[1].multiselect("Pass I", ["P", "F"])
    p2 = c[2].multiselect("Pass II", ["P", "F"])
    min_a = c[3].slider("Min. A grades", 0, 9, 0)
    q = st.text_input("Search any text column (name, index no., etc.)")

    f = stu
    if sch:
        f = f[f["School Name"].isin(sch)]
    if p1:
        f = f[f["Pass_I"].isin(p1)]
    if p2:
        f = f[f["Pass_II"].isin(p2)]
    f = f[f["count_A"] >= min_a]
    if q:
        text_cols = f.select_dtypes(include="object").columns
        mask = f[text_cols].apply(lambda s: s.astype(
            str).str.contains(q, case=False, na=False)).any(axis=1)
        f = f[mask]

    raw_cols = [c for c in raw.columns if c in f.columns]
    show_cols = raw_cols + ["count_A", "count_B", "count_C", "count_S", "count_W",
                            "Credits", "Passes", "Pass_I", "Pass_II"]
    show_cols = list(dict.fromkeys(show_cols))
    st.caption(f"{len(f):,} candidates match")
    st.dataframe(f[show_cols], hide_index=True, use_container_width=True)
    st.download_button("⬇️ Download results (CSV)", f[show_cols].to_csv(index=False).encode("utf-8"),
                       "students_filtered.csv")

# ==========================================
# PAGE: Data Quality & Diagnostics
# ==========================================
elif page == "Data Quality & Diagnostics":
    st.header("🔍 Data Quality & Diagnostics")
    diag = res["diagnostics"]
    c = st.columns(2)
    c[0].metric("Pass II (parsed cells)", diag["pass2_parsed"])
    c[1].metric("Pass II (register Grade Count)", diag["pass2_register"],
                diag["pass2_register"] - diag["pass2_parsed"], delta_color="inverse")

    tab1, tab2, tab3 = st.tabs(
        ["Grade Count Mismatches", "Pass II Delta", "Not-Set / Unparseable"])
    with tab1:
        st.markdown(
            "Candidates whose parsed grades differ from the register's 'Grade Count'.")
        st.bar_chart(pd.Series(diag["mismatch_per_grade"], name="Mismatches"))
        mism = res["mismatches"]
        if len(mism):
            st.warning(f"{len(mism)} mismatched records.")
            st.dataframe(mism[[c for c in raw.columns if c in mism.columns] +
                              [f"count_{g}" for g in GRADES] + [f"gc_{g}" for g in GRADES]],
                         use_container_width=True, hide_index=True)
        else:
            st.success("✅ No grade-count mismatches.")
    with tab2:
        delta = res["pass2_delta"]
        if len(delta):
            st.warning(
                f"{len(delta)} candidates qualify by register Grade Count but not by parsed cells.")
            st.dataframe(delta[["School Name", "Language & Literature", "Mathematics", "Grade Count",
                                "count_A", "count_B", "count_C", "count_S",
                                "gc_A", "gc_B", "gc_C", "gc_S"]],
                         use_container_width=True, hide_index=True)
        else:
            st.success("✅ No Pass II discrepancies.")
    with tab3:
        st.markdown(
            "Cells that were not set ('+'), blank, or could not be parsed.")
        st.bar_chart(
            pd.Series(diag["plus_per_subject"], name="'+' occurrences"))
        notset = res["notset"]
        if len(notset):
            st.dataframe(notset, use_container_width=True, hide_index=True)
        else:
            st.success("✅ Nothing unparseable.")

# ==========================================
# PAGE: Downloads
# ==========================================
elif page == "Downloads":
    st.header("⬇️ Downloads")
    st.markdown("Generated from the currently loaded register.")
    c = st.columns(3)
    c[0].download_button("📊 School & subject summary (.xlsx)", summary_workbook_bytes(res),
                         "OL_summary.xlsx",
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c[1].download_button("🧑‍🎓 Student-level results (.xlsx)", student_workbook_bytes(res),
                         "OL_student_level.xlsx",
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    try:
        c[2].download_button("📝 School summary (.md)",
                             school_markdown(res), "OL_school_summary.md")
    except ImportError:
        c[2].caption("Install `tabulate` for Markdown export.")
