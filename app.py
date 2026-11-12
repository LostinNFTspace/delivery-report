"""
Delivery report builder: web front end for report.py
Upload a tournament's production logs -> client PDF report + internal payroll and data-issues file.
"""
import os, shutil, tempfile
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

st.set_page_config(page_title="Tournament report from messy logs", page_icon="📄", layout="wide")
st.markdown("""<style>
div[data-testid="stMetricValue"]{font-size:1.7rem}
.small{opacity:.7;font-size:.9rem}
.messy td,.messy th{padding:4px 10px;font-size:.92rem}
.messy code{font-size:.9rem}
/* keep the page a readable width on big monitors */
.block-container{max-width:1240px}
div[data-testid="stImage"] img{border:1px solid rgba(128,128,128,.35);border-radius:6px}
</style>""", unsafe_allow_html=True)


# --- demo strip: a 1-minute video and a short how-to, so a first-time visitor knows what this is
PORTFOLIO = "https://janis-saulitis-portfolio.netlify.app"
DEMO_VIDEO = os.environ.get("DEMO_VIDEO", PORTFOLIO + "/media/report.mp4")


@st.dialog("Client report in 1 minute", width="large")
def show_demo():
    st.video(DEMO_VIDEO, autoplay=True)


def demo_strip():
    bar = st.container(horizontal=True)
    build = bar.button("▶  Build the sample report", type="primary")
    bar.space("stretch")
    right = bar.container(horizontal=True, width="content")
    if right.button("▶ 1-min demo", type="primary", key="demo_btn"):
        show_demo()
    with right.popover("How it works"):
        st.markdown('1. Click **Build the sample report**: five messy log files are checked in about 10 seconds.\n2. Read the data check: what was fixed, what is held back to confirm, what was left out.\n3. Download the client PDF, plus the internal Excel with crew overtime and a to-do list of data problems.\n4. Try your own files under **Use my own logs**.')
    return build


st.title("Turn messy production logs into a client report")
st.markdown("**Five messy log files become one checked client report.** Typos are fixed and listed, anything uncertain is held back, never guessed.")
build_clicked = demo_strip()

work = Path(st.session_state.setdefault("work", tempfile.mkdtemp(prefix="delivery_")))
run_source = None
if build_clicked:
    run_source = SAMPLES["Sample week with extra mistakes (test file)"]
st.caption("Demo: the tournament, players, crew and audience figures are invented. Takes about 10 seconds.")

left, right = st.columns([1, 1], gap="large", vertical_alignment="top")
with left:
    st.markdown("**What comes in**: real entries from the sample logs")
    st.markdown("""<table class="messy">
<tr><th>Typed in the logs</th><th>What the tool does</th></tr>
<tr><td><code>11.18</code> · <code>1330</code> · <code>13</code></td><td>three ways to write a time → read as 11:18, 13:30, 13:00</td></tr>
<tr><td><code>R33</code></td><td>no such round → corrected to R32</td></tr>
<tr><td><code>Tenniss TV</code></td><td>broadcaster typo → Tennis TV</td></tr>
<tr><td><code>4.629</code> viewers</td><td>dot as thousands separator → 4 629</td></tr>
<tr><td><code>130 966</code> viewers</td><td>20× the usual audience for that round, likely an extra digit → <b>held back</b>, not shown as a record</td></tr>
<tr><td>off air <code>19:12</code>, on air <code>19:29</code></td><td>impossible → <b>left out</b> and flagged to fix</td></tr>
</table>""", unsafe_allow_html=True)
with right:
    st.markdown("**What comes out**: a client PDF report (page 1 shown)")
    st.image(str(HERE / "assets" / "report_preview.png"), width="stretch")
    st.markdown('<p class="small">plus an internal Excel file: crew overtime for payroll and a to-do list of every data problem.</p>',
                unsafe_allow_html=True)


# ------------------------------------------------------------------ own logs / other sample
with st.expander("Use my own logs, or the other sample"):
    mode = st.radio("Source", ["Try a sample", "Upload my own logs"], horizontal=True, label_visibility="collapsed")
    source = None
    if mode == "Try a sample":
        pick = st.selectbox("Sample", list(SAMPLES))
        source = SAMPLES[pick]
        files = [source] if source.is_file() else sorted(p for p in source.iterdir() if p.is_file())
        st.markdown('<p class="small">Download the sample files to see what goes in:</p>', unsafe_allow_html=True)
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
    if source and st.button("Build the report"):
        run_source = source

# ------------------------------------------------------------------ run
if run_source:
    out = work / "out"
    shutil.rmtree(out, ignore_errors=True)
    try:
        with st.spinner("Reading and checking the logs…"):
            s, matches, faults, dist, shifts, issues = report.build(Path(run_source), out, pdf=False)
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
        st.toast("Report built: scroll down to see it", icon="✅")
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
    st.subheader("Result")
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
