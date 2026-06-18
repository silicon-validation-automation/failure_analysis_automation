[README.md](https://github.com/user-attachments/files/29101643/README.md)
# 🔍 Failure Analyzer

Automated test-failure analysis tool that parses log files and produces a
polished **Excel workbook** and an **interactive HTML report** — zero
configuration needed for common log formats.

---

## ✨ Features

| Feature | Details |
|---|---|
| **Smart parser** | Handles timestamped logs, pytest/unittest output, and custom patterns |
| **Auto-categorization** | Groups failures into Timeout, Assertion, Connection, NullRef, Permission, Memory, Syntax, FileNotFound |
| **Excel workbook** | 4 sheets: All Failures · By Category (with bar chart) · Per-Log Summary · Executive Summary |
| **Interactive HTML** | Live search, doughnut + bar charts (Chart.js), colour-coded badges |
| **Severity scoring** | HIGH / MEDIUM / LOW per log file based on failure count |
| **Custom regex** | Pass `--pattern` to match proprietary log formats |
| **Recursive scan** | Processes an entire folder of `.log` / `.txt` / `.out` files |
| **Deduplication** | Suppresses repeated identical failures within the same file |

---

## 📦 Installation

```bash
git clone https://github.com/<you>/failure-analyzer.git
cd failure-analyzer
pip install openpyxl
```

> Python 3.10+ recommended.  Only external dependency is `openpyxl`.

---

## 🚀 Usage

```bash
# Analyse a single log file
python failure_analyzer.py --path run.log

# Analyse an entire folder
python failure_analyzer.py --path logs/

# Custom output name
python failure_analyzer.py --path logs/ --output sprint_42_failures

# Custom failure pattern (regex)
python failure_analyzer.py --path logs/ --pattern "FAIL|CRITICAL|BROKEN"

# Excel only (skip HTML)
python failure_analyzer.py --path logs/ --no-html

# HTML only (skip Excel)
python failure_analyzer.py --path logs/ --no-excel
```

---

## 📂 Output

```
failure_report.xlsx   ← Excel workbook (4 sheets)
failure_report.html   ← Interactive HTML dashboard
```

### Excel sheets

| Sheet | Contents |
|---|---|
| `All_Failures` | Every failure with category, reason, file, line number, timestamp |
| `By_Category` | Aggregated counts with bar chart |
| `Per_Log_Summary` | Per-file failure count, categories hit, severity rating |
| `Summary` | Executive KPIs (total failures, top category, highest-risk log) |

---

## 🧩 Supported Log Formats (out of the box)

```
# Timestamped structured
[2024-01-15 10:32:01] Test Failed: Test: login_flow - reason: timeout

# pytest / unittest style
FAILED tests/test_api.py::test_response_code
ERROR  TestDatabase.test_connect  AssertionError

# CI pipeline style
FAIL  build_step  (exit code 1)
```

Pass `--pattern <regex>` to support any other format.

---

## 🗂️ Project Structure

```
failure-analyzer/
├── failure_analyzer.py   ← main script
├── README.md
├── requirements.txt
└── sample_logs/
    ├── app.log
    └── integration.log
```

---

## 📸 Sample Output

**HTML Dashboard**
- KPI cards (logs scanned, total failures, unique tests, categories)
- Doughnut chart — failures by category
- Bar chart — failures per log file
- Searchable, colour-coded failures table

**Excel Workbook**
- Navy header rows, red-highlighted failure cells
- Alternating row colours
- Auto-filter on all data sheets
- Embedded bar chart on the By_Category sheet

---

## 🔧 Extending

### Add a new failure category
```python
# In failure_analyzer.py → FAILURE_CATEGORIES dict
"MyCategory": re.compile(r"my_error_pattern", re.I),
```

### Plug into CI/CD (GitHub Actions example)
```yaml
- name: Analyse test failures
  run: python failure_analyzer.py --path test-results/ --output ci_report

- name: Upload reports
  uses: actions/upload-artifact@v4
  with:
    name: failure-reports
    path: |
      ci_report.xlsx
      ci_report.html
```

---

## 📄 License

MIT — free to use, modify, and distribute.
