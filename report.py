"""
Tournament delivery report
==========================
Messy production logs in  ->  two outputs:
  output/delivery_report.pdf   client report for the rights holder
  output/internal_ops.xlsx     internal: crew overtime for payroll + every data issue found

Rule: obvious formatting is read automatically (and listed); anything that can't be read
safely is left out of the numbers and flagged. Nothing is silently guessed.

Usage:  python report.py [input_folder] [output_folder]
"""
import csv, re, sys, unicodedata
from dataclasses import dataclass, field
from datetime import datetime, date, timedelta
from pathlib import Path
from html import escape

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

EVENT = "Baltic Coast Challenger 100 — Jūrmala"
PRODUCED_COURTS = ["Centre Court", "Court 1"]
ALL_COURTS = ["Centre Court", "Court 1", "Court 2", "Court 3"]
ROSTER = ["Marta Ozola", "Henrik Vass", "Ilze Kalniņa", "Pavel Novik", "Noor Haddad", "Owen Pryce",
          "Jonas Reuter", "Aino Lehto", "Viktor Lind", "Mila Horvat"]
ROUND_ORDER = ["R32", "R16", "QF", "SF", "F"]
ROUNDS = ["R64", "R32", "R16", "QF", "SF", "F"]
ROUND_ALIASES = {"final": "F", "fin": "F", "semifinal": "SF", "semi final": "SF", "semi": "SF",
                 "quarterfinal": "QF", "quarter final": "QF", "quarter": "QF"}
BROADCASTERS = ["Tennis TV", "Sportacentrs.com", "TV4", "Eurosport 2"]
# Official entry list: read from entry_list.csv next to the logs (the tournament publishes it)
PLAYERS = []
VIEWER_OUTLIER = 4  # a match with more than 4x its round's typical audience is held back until confirmed


# ---------------------------------------------------------------- issues
@dataclass
class Issue:
    level: str      # ERROR = left out of the numbers, WARN = read but check, INFO = auto-corrected
    source: str     # which file
    row: int        # row number in that file (header = 1)
    field: str
    value: str
    problem: str
    effect: str = ""


@dataclass
class Issues:
    items: list = field(default_factory=list)
    def add(self, *a, **k): self.items.append(Issue(*a, **k))
    def of(self, level): return [i for i in self.items if i.level == level]


# ---------------------------------------------------------------- readers for messy values
def fold(s):
    return unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().lower().strip()

def edit_distance(a, b):
    d = [[i + j if i * j == 0 else 0 for j in range(len(b) + 1)] for i in range(len(a) + 1)]
    for i in range(1, len(a) + 1):
        for j in range(1, len(b) + 1):
            d[i][j] = min(d[i-1][j] + 1, d[i][j-1] + 1, d[i-1][j-1] + (a[i-1] != b[j-1]))
    return d[-1][-1]

def closest(value, known, max_diff):
    """The one known value a typo most likely means, or None if none/ambiguous."""
    f = re.sub(r"\s+", " ", fold(value))
    exact = [k for k in known if re.sub(r"\s+", " ", fold(k)) == f]
    if len(exact) == 1: return exact[0]
    scored = sorted((edit_distance(f.replace(" ", ""), fold(k).replace(" ", "")), k) for k in known)
    if scored and scored[0][0] <= max_diff and (len(scored) == 1 or scored[1][0] > scored[0][0]):
        return scored[0][1]
    return None

def read_court(v):
    t = fold(v)
    if t in ("centre court", "center court", "centre", "center", "cc"): return "Centre Court"
    m = re.fullmatch(r"(?:court|c)\s*(\d+)", t)
    if m: return f"Court {int(m.group(1))}" if f"Court {int(m.group(1))}" in ALL_COURTS else None
    return closest(v, ALL_COURTS, 2)

def read_time_note(v):
    """-> ("HH:MM", note) or (None, ""). note explains anything that wasn't a plain HH:MM."""
    t = re.sub(r"\s+", "", str(v or "")).lower()
    note = ""
    m = re.fullmatch(r"(\d{1,2})[:.;,](\d{2})(?::\d{2})?", t) or re.fullmatch(r"(\d{1,2})(\d{2})", t)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
    elif m := re.fullmatch(r"(\d{1,2})[.,](\d)", t):   # 15.5: a spreadsheet dropped the trailing zero of 15.50
        h, mi = int(m.group(1)), int(m.group(2)) * 10
        note = "a spreadsheet probably dropped the trailing zero"
    elif m := re.fullmatch(r"(\d{1,2})", t):             # 13: "13.00" turned into a number
        h, mi = int(m.group(1)), 0
        note = "hour only (a spreadsheet probably turned 13.00 into 13)"
    else:
        return None, ""
    if h > 23 or mi > 59: return None, ""
    return f"{h:02d}:{mi:02d}", note

def read_time(v):
    return read_time_note(v)[0]

def read_round(v):
    """-> (round, note) ; unknown -> (None, reason)."""
    t = re.sub(r"\s+", " ", fold(v))
    if t.upper() in ROUNDS: return t.upper(), ""
    if t in ROUND_ALIASES: return ROUND_ALIASES[t], ""
    r = closest(t, ROUNDS, 1)
    return (r, f'read as "{r}"') if r else (None, "round not recognised")

def read_player(v):
    """Player name checked against the entry list -> (name, note)."""
    p = closest(v, PLAYERS, 2)
    return (p, "" if p == str(v).strip() else f'read as "{p}"') if p else (str(v).strip(), "player not on the entry list")

def read_match(v):
    """'A v. B' with each player checked against the entry list -> (match, notes)."""
    parts = re.split(r"\s+v(?:s)?\.?\s+", str(v or "").strip())
    if len(parts) != 2: return str(v).strip(), ["match name not in 'Player v. Player' form"]
    (a, na), (b, nb) = read_player(parts[0]), read_player(parts[1])
    notes = [f"{p}: {n}" for p, n in ((parts[0], na), (parts[1], nb)) if n]
    return f"{a} v. {b}", notes

def read_date(v):
    t = str(v or "").strip()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try: return datetime.strptime(t, fmt).date()
        except ValueError: pass
    return None

def read_seconds(v):
    """'40 sek', '45 s', '2 min', '180', '1:30' -> (seconds, note) ; unreadable -> (None, reason)."""
    t = str(v or "").strip().lower()
    if t in ("", "n/a", "na", "-", "?"): return None, "duration not recorded"
    if m := re.fullmatch(r"(\d+(?:[.,]\d+)?)\s*(s|sec|sek|secs|seconds|sekundes)", t): return round(float(m.group(1).replace(",", "."))), ""
    if m := re.fullmatch(r"(\d+(?:[.,]\d+)?)\s*(m|min|mins|minutes|minūtes)", t): return round(float(m.group(1).replace(",", ".")) * 60), ""
    if m := re.fullmatch(r"(\d+):(\d{2})", t): return int(m.group(1)) * 60 + int(m.group(2)), ""
    if m := re.fullmatch(r"\d+", t): return int(t), "no unit given, read as seconds"
    return None, "duration not readable"

def read_number(v):
    """12450, '12 450', '12.450' (thousands), 'n/a' -> (int|None, note)."""
    if isinstance(v, (int, float)): return int(v), ""
    t = str(v or "").strip().replace(" ", " ")
    if t.lower() in ("", "n/a", "na", "-", "?"): return None, "not recorded"
    if re.fullmatch(r"-?\d{1,3}( \d{3})+", t): return int(t.replace(" ", "")), ""
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t): return int(t.replace(".", "")), f'"{t}" read as {t.replace(".", "")} (dot as thousands separator)'
    if re.fullmatch(r"-?\d+", t): return int(t), ""
    return None, "not a number"

def minutes(a, b):
    return int((datetime.strptime(b, "%H:%M") - datetime.strptime(a, "%H:%M")).total_seconds() // 60)


# ---------------------------------------------------------------- loading
def load_production(path, issues):
    rows, seen = [], {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for n, r in enumerate(csv.DictReader(f, delimiter=";"), start=2):
            key = tuple(r.values())
            if key in seen:
                issues.add("INFO", path.name, n, "row", r["Match"], f"exact duplicate of row {seen[key]}", "removed")
                continue
            seen[key] = n
            match, mnotes = read_match(r["Match"])
            for note in mnotes:
                issues.add("WARN" if "not on the entry list" in note or "form" in note else "INFO", path.name, n, "Match", r["Match"], note)
            rnd, rnote = read_round(r["Rnd"])
            if rnote: issues.add("INFO" if rnd else "WARN", path.name, n, "Rnd", r["Rnd"], rnote, "" if rnd else "shown as typed")
            m = {"row": n, "match": match, "round": rnd or r["Rnd"].strip(), "notes": r["Notes"].strip()}
            m["date"] = read_date(r["Date"])
            if not m["date"]:
                issues.add("ERROR", path.name, n, "Date", r["Date"], "date not readable", "match left out"); continue
            court = read_court(r["Court"])
            if not court:
                issues.add("ERROR", path.name, n, "Court", r["Court"], "unknown court", "match left out"); continue
            if court != r["Court"].strip():
                issues.add("INFO", path.name, n, "Court", r["Court"], f'read as "{court}"')
            m["court"] = court
            director = closest(r["Director"], ROSTER, 3)
            if director and director != r["Director"].strip():
                issues.add("INFO", path.name, n, "Director", r["Director"], f'read as "{director}"')
            m["director"] = director or r["Director"].strip()
            for col, key in (("Sched.", "scheduled"), ("On air", "on_air"), ("First serve", "first_serve"),
                             ("Match end", "match_end"), ("Off air", "off_air")):
                m[key], tnote = read_time_note(r[col])
                if tnote: issues.add("INFO", path.name, n, col, r[col], f'read as {m[key]} ({tnote})')
                if m[key] is None:
                    issues.add("ERROR", path.name, n, col, r[col], "time missing or not readable",
                               "on-air time for this match left out" if key in ("on_air", "off_air") else "left blank in report")
            # Order must make sense: on air <= first serve < match end <= off air
            if m["on_air"] and m["off_air"] and minutes(m["on_air"], m["off_air"]) <= 0:
                issues.add("ERROR", path.name, n, "Off air", r["Off air"], f'off air ({m["off_air"]}) is before on air ({m["on_air"]})',
                           "on-air time for this match left out until corrected")
                m["off_air"] = None
            if m["first_serve"] and m["match_end"] and minutes(m["first_serve"], m["match_end"]) <= 0:
                issues.add("ERROR", path.name, n, "Match end", r["Match end"], "match end is before first serve", "match length left out")
                m["match_end"] = None
            twin = next((x for x in rows if x["date"] == m["date"] and fold(x["match"]) == fold(m["match"])), None)
            if twin:
                same = all(twin[k] == m[k] for k in ("court", "scheduled", "on_air", "first_serve", "match_end", "off_air"))
                issues.add("INFO" if same else "ERROR", path.name, n, "row", r["Match"],
                           f"same match as row {twin['row']}" + ("" if same else " but with different details"),
                           "second copy removed" if same else "second copy removed; check which row is correct")
                continue
            rows.append(m)
    rows.sort(key=lambda m: (m["date"], m["court"], m["scheduled"] or ""))
    for i, m in enumerate(rows, 1): m["id"] = f"M{i:02d}"
    return rows

def load_faults(path, matches, issues):
    out = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for n, r in enumerate(csv.DictReader(f), start=2):
            d, court = read_date(r["date"]), read_court(r["court"])
            desc = r["what happened"].strip()
            if court not in PRODUCED_COURTS:
                issues.add("WARN", path.name, n, "court", r["court"], "fault logged on a court we don't produce", "left out"); continue
            t = r["time"].strip()
            secs, note = read_seconds(r["duration"])
            if rng := re.fullmatch(r"(\S+)\s*[-–]\s*(\S+)", t):  # "14:35-14:40"
                a, b = read_time(rng.group(1)), read_time(rng.group(2))
                if a and b and secs is None:
                    secs, note = minutes(a, b) * 60, f"duration taken from time range {t}"
                t = a or t
            start = read_time(t)
            match = next((m for m in matches if m["date"] == d and m["court"] == court and start and m["on_air"] and m["off_air"]
                          and m["on_air"] <= start <= m["off_air"]), None)
            if not match:
                issues.add("WARN", path.name, n, "time", r["time"], "no produced match on air at that time", "kept in fault list, not linked to a match")
            if secs is None:
                issues.add("WARN", path.name, n, "duration", r["duration"], note, "fault listed, duration not counted")
            elif note:
                issues.add("INFO", path.name, n, "duration", r["duration"], note)
            out.append({"row": n, "date": d, "court": court, "start": start, "seconds": secs, "desc": desc,
                        "match": match["id"] if match else None})
    return out

def load_distribution(path, matches, issues):
    by_name = {fold(m["match"]): m for m in matches}  # same match even if accents differ (Kalējs / Kalejs)
    ws = openpyxl.load_workbook(path, data_only=True).active
    out = {}
    for n, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        name, takes, viewers, clips = row[:4]
        canon, mnotes = read_match(name)
        m = by_name.get(fold(canon))
        if m:
            for note in mnotes: issues.add("INFO", path.name, n, "Match", str(name), note)
        if not m:
            issues.add("WARN", path.name, n, "Match", str(name), "match not in production log", "left out"); continue
        v, note = read_number(viewers)
        if v is not None and v < 0:
            issues.add("ERROR", path.name, n, "Peak viewers", str(viewers), "negative viewer number", "viewers for this match left out"); v = None
        elif v is None:
            issues.add("WARN", path.name, n, "Peak viewers", str(viewers), note, "viewers for this match left out")
        elif note:
            issues.add("INFO", path.name, n, "Peak viewers", str(viewers), note)
        c, _ = read_number(clips)
        names = []
        for t in [t.strip() for t in re.split(r"[;,]", str(takes or "")) if t.strip()]:
            b = closest(t, BROADCASTERS, 2)
            if b and b != t: issues.add("INFO", path.name, n, "Broadcasters", t, f'read as "{b}"')
            if not b: issues.add("WARN", path.name, n, "Broadcasters", t, "broadcaster not on the known list", "shown as typed")
            names.append(b or t)
        out[m["id"]] = {"takes": names, "viewers": v, "clips": c or 0, "row": n, "raw": str(viewers)}
    # Plausibility: a typo like 13096 -> 130966 must not become a record audience in a client report
    by_round = {}
    for mid, d in out.items():
        rnd = next(m["round"] for m in matches if m["id"] == mid)
        if d["viewers"] is not None: by_round.setdefault(rnd, []).append(d["viewers"])
    for mid, d in out.items():
        rnd = next(m["round"] for m in matches if m["id"] == mid)
        vals = sorted(by_round.get(rnd, []))
        if d["viewers"] is None or len(vals) < 3: continue
        typical = vals[len(vals) // 2]
        if d["viewers"] > VIEWER_OUTLIER * typical:
            issues.add("WARN", path.name, d["row"], "Peak viewers", d["raw"],
                       f"{d['viewers']:,} is {d['viewers'] / typical:.0f}x the typical {rnd} audience ({typical:,}), possibly an extra digit",
                       "viewers for this match left out until confirmed")
            d["viewers"] = None
    return out

def load_shifts(path, issues, event_days):
    out = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for n, r in enumerate(csv.DictReader(f), start=2):
            name = closest(r["Name"], ROSTER, 3)
            if not name:
                issues.add("ERROR", path.name, n, "Name", r["Name"], "crew name not recognised", "shift left out"); continue
            if name != r["Name"].strip():
                issues.add("INFO", path.name, n, "Name", r["Name"], f'read as "{name}"')
            d, planned, actual = read_date(r["Date"]), read_time(r["Planned end"]), read_time(r["Actual end"])
            bad = next(((c, r[c]) for c, v in (("Date", d), ("Planned end", planned), ("Actual end", actual)) if not v), None)
            if bad:
                issues.add("ERROR", path.name, n, bad[0], bad[1], "not readable", "shift left out, not paid until corrected"); continue
            if d not in event_days:
                near = [e for e in event_days if abs((e - d).days) in (10, 1) or e.strftime("%d")[-1] == d.strftime("%d")[-1]]
                hint = f" (maybe {near[0].strftime('%d.%m')}?)" if len(near) == 1 else ""
                issues.add("ERROR", path.name, n, "Date", r["Date"], f"date is outside the tournament week{hint}", "shift left out, not paid until corrected"); continue
            out.append({"name": name, "date": d, "overtime": max(0, minutes(planned, actual))})
    return out


# ---------------------------------------------------------------- numbers
def summarise(matches, faults, dist):
    timed = [m for m in matches if m["on_air"] and m["off_air"]]
    on_air_min = sum(minutes(m["on_air"], m["off_air"]) for m in timed)
    covered = [m for m in matches if m["on_air"] and m["first_serve"] and m["on_air"] <= m["first_serve"]]
    late = [m for m in matches if m["on_air"] and m["first_serve"] and m["on_air"] > m["first_serve"]]
    counted = [f for f in faults if f["seconds"] is not None]
    fault_sec = sum(f["seconds"] for f in counted)
    viewers = [d["viewers"] for d in dist.values() if d["viewers"] is not None]
    takes = {}
    for d in dist.values():
        for t in d["takes"]: takes[t] = takes.get(t, 0) + 1
    days = sorted({m["date"] for m in matches})
    per_day = [{"date": d,
                "matches": sum(1 for m in matches if m["date"] == d),
                "on_air_h": round(sum(minutes(m["on_air"], m["off_air"]) for m in timed if m["date"] == d) / 60, 1),
                "peak": max((dist[m["id"]]["viewers"] or 0) for m in matches if m["date"] == d and m["id"] in dist)} for d in days]
    return {
        "matches": len(matches), "timed": len(timed), "on_air_min": on_air_min,
        "full_coverage": len(covered), "late": late,
        "faults": faults, "faults_counted": len(counted), "fault_sec": fault_sec,
        "uptime_pct": 100 * (1 - fault_sec / (on_air_min * 60)) if on_air_min else 0,
        "viewers_total": sum(viewers), "viewers_peak": max(viewers) if viewers else 0,
        "viewers_known": len(viewers), "clips": sum(d["clips"] for d in dist.values()),
        "takes": dict(sorted(takes.items(), key=lambda kv: -kv[1])), "per_day": per_day,
        "first_day": days[0], "last_day": days[-1],
    }


# ---------------------------------------------------------------- client report (HTML -> PDF)
def fmt_int(n): return f"{n:,}".replace(",", " ")
def fmt_day(d): return d.strftime("%a %d %b")

def bar_chart(rows, key, label, unit=""):
    """Small inline SVG bar chart, no external libraries."""
    w, h, pad, top = 520, 210, 10, 30
    mx = max(r[key] for r in rows) or 1
    bw = (w - pad) / len(rows)
    bars = []
    for i, r in enumerate(rows):
        bh = (h - top - 30) * r[key] / mx
        x, y = pad + i * bw + bw * 0.18, h - 30 - bh
        val = fmt_int(r[key]) if isinstance(r[key], int) else r[key]
        bars.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw*0.64:.1f}" height="{bh:.1f}" rx="2" fill="var(--accent)"/>'
                    f'<text x="{x + bw*0.32:.1f}" y="{y - 5:.1f}" class="v">{val}{unit}</text>'
                    f'<text x="{x + bw*0.32:.1f}" y="{h - 8}" class="l">{r["date"].strftime("%a")}</text>')
    return (f'<figure><figcaption>{escape(label)}</figcaption><svg viewBox="0 0 {w} {h}" role="img" aria-label="{escape(label)}">'
            f'<line x1="{pad}" y1="{h-30}" x2="{w}" y2="{h-30}" stroke="var(--rule)"/>{"".join(bars)}</svg></figure>')

def client_html(s, matches, dist, issues):
    late_txt = "".join(f"<li>{m['id']} {escape(m['match'])} ({fmt_day(m['date'])}, {m['court']}): on air {m['on_air']}, "
                       f"first serve {m['first_serve']} — {minutes(m['first_serve'], m['on_air'])} min of play missed. "
                       f"{escape(m['notes'] or '')}</li>" for m in s["late"])
    fault_rows = "".join(
        f"<tr><td>{fmt_day(f['date'])}</td><td>{f['start'] or ''}</td><td>{f['court']}</td><td>{f['match'] or '—'}</td>"
        f"<td>{escape(f['desc'])}</td><td class='n'>{'not recorded' if f['seconds'] is None else (str(f['seconds']) + ' s')}</td></tr>"
        for f in sorted(s["faults"], key=lambda f: (f["date"], f["start"] or "")))
    match_rows = "".join(
        f"<tr><td>{m['id']}</td><td class='nw'>{m['date'].strftime('%a %d.%m')}</td><td class='nw'>{m['court'].replace('Centre Court', 'Centre')}</td><td>{m['round']}</td><td>{escape(m['match'])}</td>"
        f"<td class='nw'>{m['on_air'] or '—'}–{m['off_air'] or '?'}</td>"
        f"<td class='n'>{fmt_int(dist[m['id']]['viewers']) if m['id'] in dist and dist[m['id']]['viewers'] is not None else '—'}</td>"
        f"<td>{escape(', '.join(dist[m['id']]['takes'])) if m['id'] in dist else ''}</td></tr>" for m in matches)
    takes = "".join(f"<tr><td>{escape(k)}</td><td class='n'>{v}</td></tr>" for k, v in s["takes"].items())
    # Client-facing notes: only what changes a number in this report, described by match, not by file row.
    # The full technical list (file, row, value) goes to the internal workbook.
    notes = []
    for m in matches:
        if not (m["on_air"] and m["off_air"]):
            notes.append(f"{m['id']} {m['match']} ({fmt_day(m['date'])}): on-air end time under review; this match is not included in hours on air.")
    for f in s["faults"]:
        if f["seconds"] is None:
            notes.append(f"Incident {fmt_day(f['date'])} {f['start']} ({f['desc'].lower()}): duration not recorded; listed but not counted in incident time.")
    for m in matches:
        if m["id"] in dist and dist[m["id"]]["viewers"] is None:
            notes.append(f"{m['id']} {m['match']}: stream audience figure not available; left out of audience totals.")
    note_items = "".join(f"<li>{escape(n)}</li>" for n in notes)
    hours = s["on_air_min"] / 60

    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="color-scheme" content="light"><title>Delivery report — {escape(EVENT)}</title><style>
:root{{--ink:#14181f;--muted:#5b6472;--rule:#d9dde3;--accent:#0b6e4f;--bad:#b42318;--bg:#ffffff;--soft:#f3f5f7}}
@page{{size:A4;margin:16mm 15mm 18mm}}
*{{box-sizing:border-box}} body{{font:10pt/1.45 "Helvetica Neue",Arial,sans-serif;color:var(--ink);background:var(--bg);margin:0}}
header{{border-bottom:3px solid var(--accent);padding-bottom:10px;margin-bottom:16px}}
.kicker{{font-size:8.5pt;letter-spacing:.12em;text-transform:uppercase;color:var(--accent);font-weight:700}}
h1{{font-size:20pt;margin:4px 0 2px}} h2{{font-size:12pt;margin:20px 0 8px;border-bottom:1px solid var(--rule);padding-bottom:4px}}
.sub{{color:var(--muted)}} .kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:12px 0}}
.kpi{{background:var(--soft);border-radius:6px;padding:10px}} .kpi b{{display:block;font-size:17pt;line-height:1.1}}
.kpi span{{font-size:8.5pt;color:var(--muted)}} table{{width:100%;border-collapse:collapse;font-size:8.5pt}}
th{{text-align:left;color:var(--muted);font-weight:600;border-bottom:1px solid var(--rule);padding:4px 5px}}
td{{border-bottom:1px solid var(--soft);padding:4px 5px;vertical-align:top}} .n{{text-align:right;white-space:nowrap}}
figure{{margin:0}} figcaption{{font-size:8.5pt;color:var(--muted);margin-bottom:2px}}
svg text.v{{font-size:17px;font-weight:600;text-anchor:middle;fill:var(--ink)}} svg text.l{{font-size:16px;text-anchor:middle;fill:var(--muted)}}
.charts{{display:grid;grid-template-columns:1fr 1fr;gap:14px}} .lead{{font-size:10.5pt}}
.notes li{{margin-bottom:3px}} .bad{{color:var(--bad)}} section{{break-inside:avoid}} .nw{{white-space:nowrap}} .log table{{font-size:8pt}} .log section{{break-inside:auto}} section.log{{break-inside:auto}}
footer{{margin-top:18px;color:var(--muted);font-size:8pt}}
</style></head><body>
<header><div class="kicker">Production delivery report · Demo with fictional event and data</div><h1>{escape(EVENT)}</h1>
<div class="sub">{fmt_day(s['first_day'])} – {fmt_day(s['last_day'])} 2026 · Courts produced: {', '.join(PRODUCED_COURTS)}</div></header>

<p class="lead">All {s['matches']} main-draw singles matches on the produced courts were delivered live.
{s['full_coverage']} of {s['matches']} were on air before first serve; on-air reliability was
{s['uptime_pct']:.2f}% with {s['faults_counted']} logged incidents totalling {s['fault_sec'] // 60} min {s['fault_sec'] % 60} s.</p>

<div class="kpis">
<div class="kpi"><b>{s['matches']}</b><span>matches produced live</span></div>
<div class="kpi"><b>{hours:.1f} h</b><span>on air ({s['timed']} of {s['matches']} matches timed)</span></div>
<div class="kpi"><b>{s['uptime_pct']:.2f}%</b><span>on-air reliability</span></div>
<div class="kpi"><b>{fmt_int(s['viewers_peak'])}</b><span>highest peak stream audience</span></div>
</div>

<div class="charts">{bar_chart(s['per_day'], 'on_air_h', 'Hours on air per day', ' h')}{bar_chart(s['per_day'], 'peak', 'Highest peak stream audience per day')}</div>

<section><h2>Coverage</h2>
<p>{s['full_coverage']} of {s['matches']} matches were on air before first serve.{' The exception:' if s['late'] else ''}</p>
<ul>{late_txt}</ul></section>

<section><h2>Incidents</h2>
<p>{len(s['faults'])} incidents logged; {s['faults_counted']} with a recorded duration, totalling {s['fault_sec'] // 60} min {s['fault_sec'] % 60} s
against {hours:.1f} h on air.</p>
<table><tr><th>Day</th><th>Time</th><th>Court</th><th>Match</th><th>What happened</th><th class="n">Duration</th></tr>{fault_rows}</table></section>

<section><h2>Distribution</h2>
<p>{s['clips']} highlight clips delivered. Peak stream audience recorded for {s['viewers_known']} of {s['matches']} matches.</p>
<table style="width:55%"><tr><th>Broadcaster / platform</th><th class="n">Matches taken</th></tr>{takes}</table></section>

<section class="log"><h2>Match log</h2>
<table><tr><th>#</th><th>Day</th><th>Court</th><th>Rnd</th><th>Match</th><th>On air</th><th class="n">Peak viewers</th><th>Taken by</th></tr>{match_rows}</table></section>

<section><h2>Notes on this report</h2>
<p>Figures come from the production, fault and distribution logs. Items that could not be verified were left out of the totals rather than estimated:</p>
<ul class="notes">{note_items or '<li>None.</li>'}</ul></section>
<footer>Demo report: event, matches, crew, incidents and audience figures are invented for illustration. Prepared by the live production team · {fmt_int(s['matches'])} matches · generated from logs dated {s['first_day'].isoformat()} to {s['last_day'].isoformat()}</footer>
</body></html>"""


def write_pdf(html_path, pdf_path):
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = None
        # Playwright's own browser locally; the system Chromium on hosting (Streamlit Cloud installs it from packages.txt)
        for exe in (None, "/usr/bin/chromium", "/usr/bin/chromium-browser"):
            try:
                b = p.chromium.launch(executable_path=exe) if exe else p.chromium.launch()
                break
            except Exception as e:
                last = e
        if b is None: raise RuntimeError(f"No browser available to make the PDF: {last}")
        page = b.new_page()
        page.goto(html_path.resolve().as_uri())
        page.pdf(path=str(pdf_path), format="A4", print_background=True, prefer_css_page_size=True,
                 display_header_footer=True, header_template="<span></span>",
                 footer_template='<div style="font-size:7pt;color:#5b6472;width:100%;text-align:center">Page <span class="pageNumber"></span> of <span class="totalPages"></span></div>')
        b.close()


# ---------------------------------------------------------------- internal workbook
def write_internal(path, shifts, issues):
    wb = openpyxl.Workbook()
    head = Font(bold=True, color="FFFFFF"); fill = PatternFill("solid", fgColor="0B6E4F")
    ws = wb.active; ws.title = "Overtime (payroll)"
    days = sorted({s["date"] for s in shifts})
    ws.append(["Person"] + [d.strftime("%a %d.%m") for d in days] + ["Total (min)", "Total", "Pay rate"])
    for name in ROSTER:
        per = [sum(s["overtime"] for s in shifts if s["name"] == name and s["date"] == d) for d in days]
        if not any(s["name"] == name for s in shifts): continue
        tot = sum(per)
        ws.append([name] + per + [tot, f"{tot // 60} h {tot % 60:02d} min", "Double rate"])
    ws2 = wb.create_sheet("Data issues")
    ws2.append(["Level", "File", "Row", "Field", "Value", "Problem", "Effect on report"])
    order = {"ERROR": 0, "WARN": 1, "INFO": 2}
    for i in sorted(issues.items, key=lambda i: (order[i.level], i.source, i.row)):
        ws2.append([i.level, i.source, i.row, i.field, i.value, i.problem, i.effect])
    for sheet in (ws, ws2):
        for c in sheet[1]: c.font, c.fill = head, fill
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = min(48, max(10, max(len(str(c.value or "")) for c in col) + 2))
        sheet.freeze_panes = "B2"
    wb.save(path)


# ---------------------------------------------------------------- run
def split_workbook(xlsx, folder):
    """One .xlsx with tabs named roughly production_log / faults / distribution / crew_timesheet -> the four input files."""
    import datetime as dt
    folder.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    def cell(v):
        if isinstance(v, dt.datetime):  # a real date, or a time over 24h that Sheets stored as a 1899/1900 date
            return v.strftime("%d.%m.%Y") if v.time() == dt.time(0) and v.year > 1900 else v.strftime("%Y-%m-%d %H:%M")
        if isinstance(v, dt.time): return v.strftime("%H:%M")
        if isinstance(v, dt.date): return v.strftime("%d.%m.%Y")
        if isinstance(v, float) and v.is_integer(): return str(int(v))
        return "" if v is None else str(v)
    tabs = {"production_log": None, "faults": None, "distribution": None, "crew_timesheet": None}
    for ws in wb.worksheets:
        key = closest(ws.title.replace("_", " "), [k.replace("_", " ") for k in tabs], 4)
        if key: tabs[key.replace(" ", "_")] = ws
    missing = [k for k, ws in tabs.items() if ws is None]
    if missing: raise SystemExit(f"Workbook is missing tab(s): {', '.join(missing)}")
    rows = lambda ws: [[cell(c) for c in r] for r in ws.iter_rows(values_only=True) if any(c is not None for c in r)]
    for name, delim in (("production_log", ";"), ("faults", ","), ("crew_timesheet", ",")):
        with open(folder / f"{name}.csv", "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f, delimiter=delim).writerows(rows(tabs[name]))
    out_wb = openpyxl.Workbook(); ws = out_wb.active
    for r in tabs["distribution"].iter_rows(values_only=True):
        if any(c is not None for c in r): ws.append(list(r))
    out_wb.save(folder / "distribution.xlsx")
    return folder


def load_entry_list(path):
    names = []
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        if line.strip() and line.strip().lower() != "player" and line.strip() not in names: names.append(line.strip())
    PLAYERS[:] = names


def build(inp, out, pdf=True):
    out.mkdir(exist_ok=True)
    entry = (inp.parent if inp.suffix.lower() == ".xlsx" else inp) / "entry_list.csv"
    if inp.suffix.lower() == ".xlsx":
        inp = split_workbook(inp, out / "input_from_workbook")
    if not entry.exists(): raise SystemExit(f"entry_list.csv not found next to the logs ({entry})")
    load_entry_list(entry)
    issues = Issues()
    matches = load_production(inp / "production_log.csv", issues)
    faults = load_faults(inp / "faults.csv", matches, issues)
    dist = load_distribution(inp / "distribution.xlsx", matches, issues)
    shifts = load_shifts(inp / "crew_timesheet.csv", issues, {m["date"] for m in matches})
    s = summarise(matches, faults, dist)
    html = client_html(s, matches, dist, issues)
    (out / "delivery_report.html").write_text(html, encoding="utf-8")
    if pdf: write_pdf(out / "delivery_report.html", out / "delivery_report.pdf")
    write_internal(out / "internal_ops.xlsx", shifts, issues)
    return s, matches, faults, dist, shifts, issues


if __name__ == "__main__":
    inp = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "input"
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).parent / "output"
    s, *_, issues = build(inp, out)
    print(f"{s['matches']} matches, {s['on_air_min'] / 60:.1f} h on air, reliability {s['uptime_pct']:.2f}%")
    print(f"Issues: {len(issues.of('ERROR'))} errors, {len(issues.of('WARN'))} warnings, {len(issues.of('INFO'))} auto-read")
    for i in issues.items: print(f"  {i.level:5} {i.source} r{i.row} {i.field}={i.value!r}: {i.problem} {('→ ' + i.effect) if i.effect else ''}")
