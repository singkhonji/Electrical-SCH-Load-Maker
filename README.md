# Electrical SCH Load Maker

A Python toolset for parsing electrical single-line diagram DXF files and generating formatted Load Schedule spreadsheets (Excel `.xlsx`).

---

## Overview

This project automates the extraction of electrical panel data from AutoCAD DXF schematic drawings and produces a professional Load Schedule in Excel format — ready for M&E Engineers, QS/estimators, and project documentation.

**Workflow:**

```
DXF Schematic  ──►  parse_dxf_load_table.py  ──►  generate_load_schedule.py  ──►  Excel Load Schedule
```

---

## Features

- Reads **AutoCAD DXF** files (AC1032 / AutoCAD 2018+)
- Extracts all `TEXT` and `MTEXT` entities with spatial coordinates
- Reconstructs panel structure:
  - Panel designation, supply source, location
  - Incomer cable, main circuit breaker, busbar rating
  - All outgoing circuit data (circuit no., CB rating, cable size, route, status)
  - Panel accessories (SPD, ELR, ZCT, PM, BMS)
- Outputs a fully formatted **Excel Load Schedule** with:
  - Section A — Incomer
  - Section B — Outgoing Circuits (Feeders to Sub-Boards)
  - Section C — Panel Accessories & Metering
  - Auto-calculated Max Demand formula (`=Connected Load × Demand Factor`)
  - Professional colour coding, merged cells, freeze panes, print settings

---

## Project Structure

```
Electrical-SCH-Load-Maker/
├── parse_dxf_load_table.py      # DXF parser — extracts entities from schematic
├── generate_load_schedule.py    # Excel builder — generates formatted load schedule
├── requirements.txt             # Python dependencies
└── README.md
```

---

## Requirements

- Python 3.8+
- [ezdxf](https://ezdxf.readthedocs.io/) — DXF file parsing
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel file generation

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Usage

### Step 1 — Parse the DXF

Update the file paths in `parse_dxf_load_table.py`:

```python
DXF_FILE  = r"path/to/your/schematic.dxf"
XLSX_FILE = r"path/to/output/LoadTable.xlsx"
```

Run:

```bash
python parse_dxf_load_table.py
```

This will print all extracted text entities (with XY coordinates) to the console and produce an initial Excel output.

---

### Step 2 — Generate the Formatted Load Schedule

Update the output path in `generate_load_schedule.py`:

```python
OUT = r"path/to/output/LoadSchedule_Generated.xlsx"
```

Populate the `ACTIVE_CIRCUITS`, `SPARE_CIRCUITS`, `ACCESSORIES`, and `NOTES` data structures with values extracted from Step 1 (or from the DXF directly).

Run:

```bash
python generate_load_schedule.py
```

---

## Output Example

The generated Excel file follows this layout:

| Row(s) | Content |
|--------|---------|
| 1 | Title bar — Panel name |
| 2 | Project subtitle |
| 3–5 | Panel info (Board Designation, Supply Source, Location, CB, Cable, Busbar) |
| 7 | **Section A** — Incomer |
| 8 | Column headers |
| 9 | Incomer circuit row |
| 10 | **Section B** — Outgoing Circuits |
| 11 | Column headers |
| 12+ | One row per outgoing circuit (active / spare) |
| 20 | **Section C** — Panel Accessories & Metering |
| 28 | Notes |

**Column headers (Sections A & B):**

> Circuit No. | Description / Load Name | CB Rating (A) | CB Type | No. of Poles | Breaking Cap. | Connected Load (kW) | Demand Factor | Max Demand (kW) | Current (A) | Cable Size | Cable Route | Remarks | Status

---

## Tested With

- `D1-MDB-COM-FOH1A-A.dxf` — Main Distribution Board, Common Area, DayOne CTP Building D
- AutoCAD DXF version: AC1032 (AutoCAD 2018)
- Python 3.14 / ezdxf 1.4.4 / openpyxl 3.1.5

---

## Roadmap

- [ ] Auto-map circuit data directly from DXF coordinates (no manual CIRCUITS dict needed)
- [ ] Support multiple panels / sheets in one run
- [ ] Auto-calculate demand factor and max demand from load data
- [ ] CLI interface with `argparse`
- [ ] Export to PDF via `xlwings` or `reportlab`
- [ ] Support for Thai-language panel schedules

---

## License

MIT License — free to use and modify.

---

## Author

Developed for M&E electrical documentation automation.  
Contributions and issues welcome.
