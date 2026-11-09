"""
Delivery report builder: web front end for report.py
Upload a tournament's production logs -> client PDF report + internal payroll and data-issues file.
"""
import shutil, tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import report

HERE = Path(__file__).parent
SAMPLES = {
    "Sample week (messy, as logged)": HERE / "sample",
    "Sample week with extra mistakes (test file)": HERE / "sample_mistakes" / "logs_with_mistakes.xlsx",
}
FILES = {
    "production_log.csv": "Production log",
    "faults.csv": "Fault / incident log",
    "distribution.xlsx": "Distribution & audience",
    "crew_timesheet.csv": "Crew timesheet",
    "entry_list.csv": "Tournament entry list",
}

st.set_page_config(page_title="Delivery report builder", page_icon="📄", layout="wide")
st.markdown("""<style>
div[data-testid="stMetricValue"]{font-size:1.7rem}
.small{color:#6b7280;font-size:.9rem}
</style>""", unsafe_allow_html=True)

st.title("Tournament delivery report builder")
st.markdown(
    "After a tournament, the production team's logs are spread over several files, typed by different people. "
    "This tool reads them, **checks every value**, and produces two things: a **client report (PDF)** for the rights holder, "
    "and an **internal file** with crew overtime for payroll and a to-do list of every data problem.\n\n"
    "Obvious typos are read automatically and listed. Anything uncertain is **left out of the numbers and flagged**, never guessed."
)
st.caption("Demo: the sample tournament, players, crew and audience figures are invented.")

# ------------------------------------------------------------------ input
st.subheader("1 · Choose the logs")
mode = st.radio("Source", ["Try a sample", "Upload my own logs"], horizontal=True, label_visibility="collapsed")

source = None
work = Path(st.session_state.setdefault("work", tempfile.mkdtemp(prefix="delivery_")))

if mode == "Try a sample":
    pick = st.selectbox("Sample", list(SAMPLES))
    source = SAMPLES[pick]
    if pick.endswith("(test file)"):
        st.markdown('<p class="small">All logs imported into Google Sheets as one workbook, then filled with typical mistakes: '
                    "round typos (R33, QQF, “Final”), player and crew name typos, a duplicated row edited by hand, "
                    "broadcaster typos, two viewer numbers with an extra digit, and shifts dated outside the tournament week.</p>",
                    unsafe_allow_html=True)
    with st.expander("Download the sample files to see what goes in"):
        files = [source] if source.is_file() else sorted(p for p in source.iterdir() if p.is_file())
        for f in files + ([source.parent / "entry_list.csv"] if source.is_file() else []):
            st.download_button(f"⬇ {f.name}", f.read_bytes(), file_name=f.name, key=f"dl_{f.parent.name}_{f.name}")
else:
    st.markdown("Upload **one Excel workbook** with the four logs as tabs (e.g. downloaded from Google Sheets), "
                "**or** the four separate files. The entry list (one player per line) is optional; "
                "without it the sample tournament's list is used.")
    up = st.file_uploader("Logs", type=["csv", "xlsx"], accept_multiple_files=True)
    if up:
        folder = work / "upload"
        shutil.rmtree(folder, ignore_errors=True); folder.mkdir(parents=True)
        names = {u.name.lower(): u for u in up}
        for u in up: (folder / u.name).write_bytes(u.getvalue())
        if not (folder / "entry_list.csv").exists():
            shutil.copy(HERE / "sample" / "entry_list.csv", folder / "entry_list.csv")
        workbooks = [folder / u.name for u in up if u.name.lower().endswith(".xlsx") and u.name.lower() != "distribution.xlsx"]
        missing = [f for f in FILES if f != "entry_list.csv" and not (folder / f).exists()]
        if workbooks:
            source = workbooks[0]
        elif missing:
            st.warning("Still missing: " + ", ".join(FILES[m] + f" ({m})" for m in missing))
        else:
            source = folder

# ------------------------------------------------------------------ run
if source and st.button("Build the report", type="primary"):
    out = work / "out"
    shutil.rmtree(out, ignore_errors=True)
    try:
        with st.spinner("Reading and checking the logs…"):
            s, matches, faults, dist, shifts, issues = report.build(Path(source), out, pdf=False)
        pdf_bytes, pdf_error = None, None
        with st.spinner("Making the PDF…"):
            try:
                report.write_pdf(out / "delivery_report.html", out / "delivery_report.pdf")
                pdf_bytes = (out / "delivery_report.pdf").read_bytes()
            except Exception as e:
                pdf_error = str(e)
        st.session_state["result"] = dict(
            s=s, issues=[vars(i) for i in issues.items], html=(out / "delivery_report.html").read_text(encoding="utf-8"),
            xlsx=(out / "internal_ops.xlsx").read_bytes(), pdf=pdf_bytes, pdf_error=pdf_error,
            overtime=pd.read_excel(out / "internal_ops.xlsx", sheet_name="Overtime (payroll)"))
    except SystemExit as e:
        st.error(f"Couldn't read the logs: {e}")
        st.session_state.pop("result", None)
    except Exception as e:
        st.error(f"Something in the files couldn't be read ({type(e).__name__}: {e}). Check the column headers match the sample files.")
        st.session_state.pop("result", None)

# ------------------------------------------------------------------ results
res = st.session_state.get("result")
if res:
    s = res["s"]
    st.subheader("2 · Result")
    c = st.columns(4)
    c[0].metric("Matches produced", s["matches"])
    c[1].metric("Hours on air", f"{s['on_air_min'] / 60:.1f}")
    c[2].metric("On-air reliability", f"{s['uptime_pct']:.2f}%")
    c[3].metric("Highest peak audience", f"{s['viewers_peak']:,}".replace(",", " "))

    lv = pd.Series([i["level"] for i in res["issues"]]).value_counts()
    st.markdown(f"**Data check:** {lv.get('ERROR', 0)} left out of the numbers · {lv.get('WARN', 0)} held back to confirm · "
                f"{lv.get('INFO', 0)} typos read automatically")

    d = st.columns(2)
    if res["pdf"]:
        d[0].download_button("⬇ Client report (PDF)", res["pdf"], file_name="delivery_report.pdf", mime="application/pdf", type="primary")
    else:
        d[0].download_button("⬇ Client report (HTML, print to PDF)", res["html"], file_name="delivery_report.html", mime="text/html")
        st.caption(f"PDF engine not available here, so the report is offered as a printable web page. ({res['pdf_error'][:120]})")
    d[1].download_button("⬇ Internal file: overtime + data issues (Excel)", res["xlsx"], file_name="internal_ops.xlsx")

    t1, t2, t3 = st.tabs(["Client report", "Data issues (to-do)", "Overtime (payroll)"])
    with t1:
        st.caption("What the rights holder receives. No payroll or internal file details.")
        components.html(res["html"], height=1100, scrolling=True)
    with t2:
        st.caption("For whoever owns the logs: what was read automatically, what was held back, and what must be fixed.")
        df = pd.DataFrame(res["issues"]).rename(columns={"level": "Level", "source": "File", "row": "Row", "field": "Field",
                                                         "value": "Value", "problem": "Problem", "effect": "Effect on report"})
        show = st.multiselect("Show", ["ERROR", "WARN", "INFO"], default=["ERROR", "WARN"],
                              help="ERROR: left out of the numbers. WARN: held back or needs a check. INFO: typo read automatically.")
        st.dataframe(df[df["Level"].isin(show)].sort_values(["Level", "File", "Row"]), hide_index=True, width="stretch")
    with t3:
        st.caption("Internal only. Minutes past the planned end of shift, per person per day, paid at double rate.")
        st.dataframe(res["overtime"], hide_index=True, width="stretch")

st.divider()
st.markdown('<p class="small">Built with Claude Code by Jānis Saulītis · engine: <code>report.py</code> with 12 automated tests '
            "against the hidden clean data, including a file deliberately filled with mistakes.</p>", unsafe_allow_html=True)
