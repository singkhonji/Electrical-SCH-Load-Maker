"""
extract_single_db.py
Extracts a single Distribution Board load schedule from a DXF file
and writes a formatted Excel load schedule.

Usage:  python extract_single_db.py input/D1-DB-COM-FOHRA-A.dxf
Output: output/<board_name>_LoadSchedule.xlsx
"""

import os, re, sys
import ezdxf
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

# ── helpers ──────────────────────────────────────────────────────────────────

def load_ents(path):
    doc = ezdxf.readfile(path)
    msp = doc.modelspace()
    ents = []
    for e in msp:
        if e.dxftype() == 'TEXT':
            ents.append({'x': round(e.dxf.insert.x), 'y': round(e.dxf.insert.y),
                         'text': (e.dxf.text or '').strip()})
        elif e.dxftype() == 'MTEXT':
            ents.append({'x': round(e.dxf.insert.x), 'y': round(e.dxf.insert.y),
                         'text': (e.plain_text() or '').strip()})
    return [e for e in ents if e['text']]


def near_x(candidates, x, tol=3000):
    """Return text of the entity in candidates closest to x within tol."""
    best = min(candidates, key=lambda e: abs(e['x'] - x), default=None)
    if best and abs(best['x'] - x) <= tol:
        return best['text']
    return ''


def band(ents, y_center, half=2000):
    return [e for e in ents if abs(e['y'] - y_center) <= half]


def first_match(ents, pattern, flags=re.I):
    for e in sorted(ents, key=lambda e: -e['y']):
        if re.search(pattern, e['text'], flags):
            return e['text']
    return ''


def y_of(ents, keyword):
    for e in sorted(ents, key=lambda e: -e['y']):
        if keyword.lower() in e['text'].lower():
            return e['y']
    return None


# ── circuit number parsing ────────────────────────────────────────────────────

# Pattern for the FIRST label in a group:  "LT-1R", "L1-2B", "P-1R", "AC-2Y"
_FIRST_CKT = re.compile(r'^([A-Z][A-Z0-9]*)[-/](\d+)([RYB])$', re.I)
# Pattern for continuation labels:  "1Y", "2B", "3R"
_CONT_CKT  = re.compile(r'^(\d+)([RYB])$', re.I)


def parse_circuit_ents(circuit_ents):
    """
    Sort by X, walk left→right assigning group prefixes.
    Returns list of dicts: {x, y, circuit_no, group, num, phase}
    """
    results = []
    current_group = ''
    for e in sorted(circuit_ents, key=lambda e: e['x']):
        t = e['text'].strip()
        m = _FIRST_CKT.match(t)
        if m:
            current_group = m.group(1).upper()
            num   = int(m.group(2))
            phase = m.group(3).upper()
        else:
            m2 = _CONT_CKT.match(t)
            if m2:
                num   = int(m2.group(1))
                phase = m2.group(2).upper()
            else:
                continue  # skip non-circuit entities
        full_id = f"{current_group}-{num}{phase}" if current_group else t
        results.append({'x': e['x'], 'y': e['y'],
                        'circuit_no': full_id,
                        'group': current_group,
                        'num': num, 'phase': phase})
    return results


# ── data extraction ───────────────────────────────────────────────────────────

def extract_board(ents):
    """Parse all board-level and per-circuit data from entity list."""

    # Board name: top-most entity matching board name pattern
    board_name = ''
    for e in sorted(ents, key=lambda e: -e['y']):
        if re.match(r'^D\d+-[A-Z]', e['text'], re.I):
            board_name = e['text']
            break

    # FOR / FROM / incomer / busbar / feed cable
    for_desc     = first_match(ents, r'^FOR\s+', re.I)
    from_supply  = first_match(ents, r'^FROM\s+', re.I)
    incomer_cb   = first_match(ents, r'\d+A\s+TPN\s+MCCB', re.I)
    busbar       = first_match(ents, r'BUSBAR', re.I)
    feed_cable   = first_match(ents, r'^\d+\s*x\s*1C', re.I)

    # ── Find key Y levels by their header text ────────────────────────────────
    y_ckt_hdr  = y_of(ents, "CIRCUIT NO")       # label row
    y_kw_hdr   = y_of(ents, "CAPACITY")         # kW header
    y_room_hdr = y_of(ents, "ROOM / EQUIPMENT") # description header

    if y_ckt_hdr is None:
        raise ValueError("Cannot find 'CIRCUIT NO' label in DXF")

    # Circuit entities: small band just ABOVE the header label
    ckt_band  = band(ents, y_ckt_hdr, 500)
    ckt_ents  = [e for e in ckt_band
                 if _FIRST_CKT.match(e['text']) or _CONT_CKT.match(e['text'])]
    circuits  = parse_circuit_ents(ckt_ents)

    if not circuits:
        raise ValueError("No circuit entities found near CIRCUIT NO row")

    # ── Data bands (scan ±half around header) ─────────────────────────────────
    # kW values sit ~30-100 units above kW header
    kw_ents   = band(ents, y_kw_hdr + 50, 300) if y_kw_hdr else []
    # Room/equipment descriptions: entities above the header
    desc_ents = band(ents, y_room_hdr + 300, 500) if y_room_hdr else []

    # Cable: scan for cable-pattern texts (largest cluster by Y)
    cable_ents = [e for e in ents
                  if re.search(r'\d+mm[²2^]', e['text'], re.I)]
    # MCB/RCCB: scan for breaker texts
    mcb_ents   = [e for e in ents
                  if re.search(r'\d+A\s+(?:SPN|TPN|4P)\s+(?:MCB|RCCB|MCCB)', e['text'], re.I)]

    # Isolate circuit-level MCBs (SPN/single-pole → per circuit)
    spn_mcb    = [e for e in mcb_ents if re.search(r'SPN', e['text'], re.I)]
    # Group-level RCCBs
    rccb_ents  = [e for e in mcb_ents if re.search(r'RCCB', e['text'], re.I)]

    # Cable texts nearest to circuit X
    # Use a Y-cluster: group cable_ents by rounded Y, pick largest cluster
    if cable_ents:
        from collections import Counter
        y_rounded = [round(e['y'] / 500) * 500 for e in cable_ents]
        dominant_y = Counter(y_rounded).most_common(1)[0][0]
        cable_band = [e for e in cable_ents if abs(round(e['y'] / 500) * 500 - dominant_y) < 1]
    else:
        cable_band = []

    # ── Assign per-circuit data ────────────────────────────────────────────────
    for c in circuits:
        x = c['x']
        c['kw']   = near_x(kw_ents,   x, tol=2000)
        c['desc'] = near_x(desc_ents,  x, tol=5000)
        c['cb']   = near_x(spn_mcb,    x, tol=2000)
        c['cable']= near_x(cable_band, x, tol=2000)
        # Find RCCB for this circuit (nearest RCCB by X)
        c['rccb'] = near_x(rccb_ents,  x, tol=15000)

    return {
        'name':         board_name,
        'for_desc':     for_desc,
        'from_supply':  from_supply,
        'incomer_cb':   incomer_cb,
        'busbar':       busbar,
        'feed_cable':   feed_cable,
        'circuits':     circuits,
    }


# ── Excel writer ──────────────────────────────────────────────────────────────

HDR_FILL  = PatternFill("solid", fgColor="1F3864")   # dark navy
GRP_FILL  = PatternFill("solid", fgColor="2E75B6")   # blue
ALT_FILL  = PatternFill("solid", fgColor="D6E4F7")   # light blue
WHT_FILL  = PatternFill("solid", fgColor="FFFFFF")
YEL_FILL  = PatternFill("solid", fgColor="FFF2CC")   # pale yellow for info rows

HDR_FONT  = Font(name='Calibri', bold=True, color="FFFFFF", size=11)
GRP_FONT  = Font(name='Calibri', bold=True, color="FFFFFF", size=10)
COL_FONT  = Font(name='Calibri', bold=True, color="FFFFFF", size=10)
DAT_FONT  = Font(name='Calibri', size=10)
INFO_FONT = Font(name='Calibri', bold=True, size=10)

THIN  = Side(style='thin',   color='B0B0B0')
MED   = Side(style='medium', color='888888')
THICK = Side(style='thick',  color='1F3864')

def thin_border():
    return Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def med_border():
    return Border(left=MED, right=MED, top=MED, bottom=MED)

def center(wrap=False):
    return Alignment(horizontal='center', vertical='center', wrap_text=wrap)

def left(wrap=True):
    return Alignment(horizontal='left', vertical='center', wrap_text=wrap)


COLS = ['A','B','C','D','E']
#       No  Desc kW  CB  Cable

COL_WIDTHS = [18, 42, 12, 28, 42]
COL_HEADERS = ['CIRCUIT NO.', 'DESCRIPTION', 'INSTALLED\nCAPACITY (kW)',
               'CB RATING', 'CABLE SIZE']


def write_cell(ws, row, col_letter, value, font=None, fill=None,
               alignment=None, border=None):
    cell = ws[f"{col_letter}{row}"]
    cell.value = value
    if font:      cell.font      = font
    if fill:      cell.fill      = fill
    if alignment: cell.alignment = alignment
    if border:    cell.border    = border
    return cell


def write_excel(data, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = data['name'][:31]

    # column widths
    for i, w in enumerate(COL_WIDTHS):
        ws.column_dimensions[COLS[i]].width = w

    row = 1

    # ── Title row ─────────────────────────────────────────────────────────────
    ws.merge_cells(f"A{row}:E{row}")
    c = ws[f"A{row}"]
    c.value     = f"LOAD SCHEDULE  –  {data['name']}"
    c.font      = Font(name='Calibri', bold=True, color="FFFFFF", size=13)
    c.fill      = HDR_FILL
    c.alignment = center()
    c.border    = med_border()
    ws.row_dimensions[row].height = 22
    row += 1

    # ── Info rows ──────────────────────────────────────────────────────────────
    info_rows = [
        ("FOR",         data['for_desc']),
        ("FROM",        data['from_supply']),
        ("INCOMER CB",  data['incomer_cb']),
        ("BUSBAR",      data['busbar']),
        ("FEED CABLE",  data['feed_cable']),
    ]
    for label, val in info_rows:
        ws.merge_cells(f"B{row}:E{row}")
        lc = ws[f"A{row}"]
        lc.value     = label
        lc.font      = INFO_FONT
        lc.fill      = YEL_FILL
        lc.alignment = left(wrap=False)
        lc.border    = thin_border()
        vc = ws[f"B{row}"]
        vc.value     = val
        vc.font      = DAT_FONT
        vc.fill      = WHT_FILL
        vc.alignment = left()
        vc.border    = thin_border()
        ws.row_dimensions[row].height = 16 if len(val) < 80 else 30
        row += 1

    row += 1  # spacer

    # ── Column headers ────────────────────────────────────────────────────────
    for col_l, hdr in zip(COLS, COL_HEADERS):
        c = ws[f"{col_l}{row}"]
        c.value     = hdr
        c.font      = COL_FONT
        c.fill      = HDR_FILL
        c.alignment = center(wrap=True)
        c.border    = med_border()
    ws.row_dimensions[row].height = 30
    row += 1

    # ── Circuit rows, grouped by sub-group ────────────────────────────────────
    circuits = data['circuits']
    # Collect ordered groups (preserve order of first appearance)
    seen_groups = []
    for c in circuits:
        if c['group'] not in seen_groups:
            seen_groups.append(c['group'])

    for grp in seen_groups:
        grp_circuits = [c for c in circuits if c['group'] == grp]

        # Sub-group header row (show RCCB info)
        rccb_text = grp_circuits[0]['rccb'] if grp_circuits else ''
        ws.merge_cells(f"A{row}:E{row}")
        c = ws[f"A{row}"]
        c.value     = f"GROUP  {grp}   —   {rccb_text}" if rccb_text else f"GROUP  {grp}"
        c.font      = GRP_FONT
        c.fill      = GRP_FILL
        c.alignment = left(wrap=False)
        c.border    = med_border()
        ws.row_dimensions[row].height = 16
        row += 1

        # Circuit rows
        for i, ckt in enumerate(grp_circuits):
            fill = ALT_FILL if i % 2 == 0 else WHT_FILL
            vals = [ckt['circuit_no'], ckt['desc'], ckt['kw'],
                    ckt['cb'], ckt['cable']]
            for col_l, val in zip(COLS, vals):
                c = ws[f"{col_l}{row}"]
                c.value     = val
                c.font      = DAT_FONT
                c.fill      = fill
                c.border    = thin_border()
                c.alignment = center() if col_l in ('A','C') else left()
            ws.row_dimensions[row].height = 14
            row += 1

        row += 1  # blank row between groups

    # ── Freeze panes ──────────────────────────────────────────────────────────
    ws.freeze_panes = f"A{8}"

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    print(f"Saved → {out_path}")


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_single_db.py <input.dxf>")
        sys.exit(1)

    dxf_path = sys.argv[1]
    _base    = os.path.dirname(os.path.abspath(__file__))
    stem     = os.path.splitext(os.path.basename(dxf_path))[0]
    out_path = os.path.join(_base, "output", f"{stem}_LoadSchedule.xlsx")

    print(f"Reading: {os.path.abspath(dxf_path)}")
    ents = load_ents(dxf_path)
    print(f"  {len(ents)} text entities loaded")

    data = extract_board(ents)
    print(f"  Board: {data['name']}")
    print(f"  Circuits found: {len(data['circuits'])}")

    # Print summary
    from collections import Counter
    group_counts = Counter(c['group'] for c in data['circuits'])
    for grp, cnt in sorted(group_counts.items()):
        print(f"    Group {grp}: {cnt} circuit columns")

    write_excel(data, out_path)


if __name__ == '__main__':
    main()
