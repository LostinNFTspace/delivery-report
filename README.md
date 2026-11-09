# Tournament delivery report builder

After a tournament, the production team's logs (production log, fault log, distribution sheet, crew timesheet) are typed by different people in different formats. This tool reads them, checks every value, and produces:

- **Client report (PDF)** for the rights holder: matches produced, hours on air, on-air reliability, incidents, distribution and audience, with honest notes on anything left out.
- **Internal file (Excel)**: crew overtime for payroll and a to-do list of every data problem by file, row and value.

Obvious typos are read automatically and listed. Anything uncertain is left out of the numbers and flagged, never guessed.

Demo data only: the tournament, players, crew and audience figures are invented.

- `app.py` web page (Streamlit)
- `report.py` the engine (reading, checks, numbers, report)
- `sample/` a messy tournament week; `sample_mistakes/` the same logs imported into Google Sheets with extra deliberate mistakes

Built with Claude Code by Jānis Saulītis.
