# Zoom Attendance Analyzer

A simple Python script to process Zoom meeting CSV reports and generate formatted Excel attendance sheets.

---

## Folder Structure

```
your_folder/
├── analyze_attendance.py       ← the script
├── raw_reports/                ← put your Zoom CSVs here
│   ├── zoomus_meeting_report_xxxx.csv
│   ├── zoomus_meeting_report_yyyy.csv
│   └── ...
└── attendance_reports/         ← generated Excel reports go here (auto-created)
    ├── attendance_2026-02-16_1301.xlsx
    └── ...
```

---

## Setup (one time)

```bash
pip install pandas openpyxl
```

---

## Usage

**Interactive** (pick a file from a menu):
```bash
python analyze_attendance.py
```

**Specific file:**
```bash
python analyze_attendance.py raw_reports/zoomus_meeting_report_93385877760.csv
```

**Process all CSVs at once:**
```bash
python analyze_attendance.py --all
```

---

## Output Excel Sheets

Each generated `.xlsx` has:

| Sheet | Contents |
|-------|----------|
| **Class Stats** | Date, duration, attendance rate summary |
| **All Students** | Every student, color-coded by status |
| **Present** | Students who attended ≥50% of class |
| **Absent or Brief** | Students below the threshold |
| **Ghost Joins** | ≤2 min joins (Zoom glitches / accidental joins) |
| **Raw Data** | Original CSV unchanged |

Color coding: 🟢 Present, 🔴 Absent/Brief, ⬜ Ghost Join

---

## Configuration

Edit the `CONFIG` block at the top of `analyze_attendance.py`:

```python
CONFIG = {
    "professor": "beth bonsignore",          # exact name (case-insensitive)
    "exclude": ["beth bonsignore", "jay walter", ],  # professor + TAs
    "min_attendance_pct": 50,                # % threshold for "present"
    "late_threshold_min": 10,                # minutes after start = late
    "early_leave_threshold_min": 10,         # minutes before end = left early
    "ghost_join_max_min": 2,                 # sessions ≤ this = Zoom glitch
    "raw_reports_dir": "raw_reports",        # input folder
    "output_dir": "attendance_reports",      # output folder
}
```

---

## Flags in the report

- **late +Nmin** — joined more than 10 min after class started
- **left early -Nmin** — left more than 10 min before class ended
- **reconnected Nx** — had multiple real sessions (dropped and rejoined)
- **ghost join (≤2min)** — joined for under 2 minutes, likely a Zoom glitch
