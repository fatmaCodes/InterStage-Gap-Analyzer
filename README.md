# Gap Report Automation Script

## Overview

This project provides a robust automation solution for analyzing stage-wise time gaps in operational workflows. It processes raw Excel data, applies business-specific time rules, and generates a structured, formatted Excel report with actionable insights.

The script is designed to handle real-world constraints such as working hours, holidays, lunch breaks, and special half-day schedules, ensuring accurate and meaningful gap calculations.

---

## Key Features

* **Accurate Time Gap Calculation**

  * Excludes non-working hours (before 09:30 AM and after 06:30 PM)
  * Excludes lunch break (01:30 PM – 02:30 PM)
  * Skips Sundays automatically
  * Ignores configured public holidays
  * Supports custom half-day working schedules

* **Data Preprocessing & Cleaning**

  * Handles inconsistent formats in date and time fields
  * Ensures proper sorting and grouping by Opportunity Number
  * Filters only relevant and latest stage records

* **Advanced Reporting**

  * Stage-wise sheet generation
  * Summary dashboard with navigation hyperlinks
  * Conditional highlighting for critical delays (Gap Days > 1)
  * Vendor count integration from external dataset

* **Excel Output Enhancements**

  * Styled headers and formatted tables
  * Auto-adjusted column widths
  * Internal navigation (hyperlinks between sheets)
  * Clean and structured layout for business use

---

## Tech Stack

* **Python 3.x**
* **Pandas** – Data processing and transformation
* **OpenPyXL** – Excel report generation and styling
* **Tkinter** – File selection dialog

---

## Input Requirements

### 1. Primary Excel File

The input file must contain the following columns:

* `Opp No`
* `Stage`
* `Stage Date`
* `Created Time`
* `Item Name 1`

Additional columns are optional and will be ignored if not required.

---

### 2. Vendor Mapping File

An external Excel file is required for vendor count mapping:

**Expected structure:**

* Sheet Name: `Sheet2`
* Column 1: `Row Labels` (Product Name)
* Column 2: Vendor Count

**Path configuration (update in script if needed):**

```python
C:\\Users\\it4\\Desktop\\Report\\GAP REPORT\\Product Wise Vendor Details.xlsx
```

---

## Business Logic

The gap calculation engine follows strict operational rules:

* Only working hours are considered
* Non-working periods are excluded:

  * Sundays
  * Defined holidays
  * Lunch break
* Supports dynamic handling of:

  * Multi-day gaps
  * Partial working days
  * Custom half-day closing times

Gap output is split into:

* **Gap Days** (based on 8-hour workday = 480 minutes)
* **Gap Hours** (HH:MM format)

---

## Output

The script generates a new Excel file:

```
<original_filename>_OP.xlsx
```

### Output Structure:

* **Gap Report (Summary Sheet)**

  * Lists all stages
  * Clickable navigation to detailed sheets

* **Stage-wise Sheets**

  * Filtered records per stage
  * Vendor count column
  * Highlighted delays

---

## How to Run

1. Install dependencies:

```bash
pip install pandas openpyxl
```

2. Run the script:

```bash
python gap_report.py
```

3. Select the input Excel file when prompted.

4. The processed report will be generated and opened automatically.

---

## Configuration

You can modify the following directly in the script:

* Working hours:

```python
WORK_START, WORK_END = time(9, 30), time(18, 30)
```

* Lunch break:

```python
LUNCH_START, LUNCH_END = time(13, 30), time(14, 30)
```

* Holidays:

```python
HOLIDAYS = [ ... ]
```

* Half-days:

```python
HALF_DAYS = { date: closing_time }
```

---

## Use Cases

* Sales pipeline delay analysis
* Operational workflow tracking
* SLA compliance monitoring
* Vendor response time evaluation

---

## Limitations

* Requires consistent column naming in input files
* Vendor mapping file path is currently hardcoded
* GUI is limited to file selection only

---

## Future Enhancements

* Configurable UI for business rules
* Dynamic vendor file selection
* Dashboard visualization (charts)
* Packaging as standalone executable (.exe)

---

## Author

Developed as a business-oriented automation tool for improving operational visibility and efficiency.

---
