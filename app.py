"""
11.41B Status Report Generator
Flask web app that reads MDB and history_space files and generates
a Word document report from the standard template.
"""
import os
import io
import re
import copy
import glob
from datetime import datetime

import pyodbc
import openpyxl
from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = Flask(__name__)

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR     = r'M:\BRSSD\CSS\11.41B\2026'
TEMPLATE_PATH = r'M:\BRSSD\CSS\11.41B\11.41B_WM1668_D72317_Draft - sample.docx'

# ── Helpers ───────────────────────────────────────────────────────────────────

def ordinal_suffix(n: int) -> str:
    """Return the ordinal suffix for an integer (e.g. 1→'st', 2→'nd', ...)."""
    if 11 <= (n % 100) <= 13:
        return 'th'
    return {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')


def ordinal(n: int) -> str:
    return f'{n}{ordinal_suffix(n)}'


def _numeric_key(v: str):
    """Sort provision values numerically: '9.7' < '9.13' < '9.14'."""
    try:
        parts = v.split('.')
        return tuple(int(p) for p in parts)
    except ValueError:
        return (0, v)


def _make_run_xml(
    text: str,
    bold: bool = False,
    highlight: str = None,
    superscript: bool = False,
    font_name: str = 'Calibri',
    sz_cs: int = 22,
) -> object:
    """
    Create a ``w:r`` OxmlElement.
    Matches the body-text run style of the template:
      rFonts(Calibri) + optional b/bCs + optional vertAlign(superscript)
      + optional highlight + szCs(22).
    """
    r   = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    # Font family (ascii + hAnsi + cs)
    if font_name:
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:cs'),    font_name)
        rPr.append(rFonts)

    # Bold
    if bold:
        rPr.append(OxmlElement('w:b'))
        rPr.append(OxmlElement('w:bCs'))

    # Superscript
    if superscript:
        va = OxmlElement('w:vertAlign')
        va.set(qn('w:val'), 'superscript')
        rPr.append(va)

    # Highlight
    if highlight:
        hl = OxmlElement('w:highlight')
        hl.set(qn('w:val'), highlight)
        rPr.append(hl)

    # Complex-script size (matches template's szCs val="22")
    if sz_cs:
        szCs = OxmlElement('w:szCs')
        szCs.set(qn('w:val'), str(sz_cs))
        rPr.append(szCs)

    r.append(rPr)

    # w:t
    t = OxmlElement('w:t')
    t.text = text
    if text != text.strip() or not text:
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(t)
    return r


# ── Data readers ──────────────────────────────────────────────────────────────

def get_mdb_data(notice_id: str) -> dict:
    """
    Open the MDB at BASE_DIR/<notice_id>/<notice_id>.mdb (read-only) and
    return:
      com_el: list of dicts with ntc_id, adm, sat_name, type, d_rcv
      still_exist: 'YES' if provn has any non-O agree_st, else 'NO'
    """
    mdb_path = os.path.join(BASE_DIR, notice_id, f'{notice_id}.mdb')
    if not os.path.exists(mdb_path):
        raise FileNotFoundError(f'MDB file not found: {mdb_path}')

    conn_str = (
        f'Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};'
        f'DBQ={mdb_path};ReadOnly=True;'
    )
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # com_el ─────────────────────────────────────────────────────────────────
    cursor.execute('SELECT ntc_id, adm, sat_name, long_nom, d_rcv FROM com_el')
    com_el_rows = []
    for row in cursor.fetchall():
        d_rcv = row[4].strftime('%d.%m.%Y') if row[4] else ''
        com_el_rows.append({
            'ntc_id':   str(row[0]) if row[0] is not None else '',
            'adm':      (row[1] or '').strip(),
            'sat_name': (row[2] or '').strip(),
            'type':     'G' if row[3] is not None else 'N',
            'd_rcv':    d_rcv,
        })

    # provn.coord_prov: YES if any entry contains '11.41' ────────────────────
    cursor.execute("SELECT COUNT(*) FROM provn WHERE coord_prov LIKE '%11.41%'")
    has_1141 = cursor.fetchone()[0]
    still_exist = 'YES' if has_1141 > 0 else 'NO'

    conn.close()
    return {'com_el': com_el_rows, 'still_exist': still_exist}


def get_history_data(notice_id: str) -> dict:
    """
    Find the **(history_space).xlsx** file under
    BASE_DIR/BR_TEXT_RESULTS/<notice_id>/Status Review/
    and return:
      provn_targets: sorted list of distinct provn_target strings
      adms:          sorted list of distinct adm strings
    """
    status_dir = os.path.join(
        BASE_DIR, 'BR_TEXT_RESULTS', notice_id, 'Status Review'
    )
    if not os.path.isdir(status_dir):
        raise FileNotFoundError(
            f'Status Review folder not found: {status_dir}'
        )

    # Find the file (may be .xlsx or .xls)
    matches = glob.glob(os.path.join(status_dir, '*(history_space).*'))
    if not matches:
        raise FileNotFoundError(
            f'No (history_space) file found in: {status_dir}'
        )
    xlsx_path = matches[0]

    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    ws = wb.active

    # Row 1 = sheet title  (e.g. 'COMPASS-80E')
    # Row 2 = column headers
    # Row 3+ = data
    header_row = [
        (str(h).strip() if h is not None else '')
        for h in next(ws.iter_rows(min_row=2, max_row=2, values_only=True))
    ]

    def col_idx(name: str, default: int) -> int:
        try:
            return header_row.index(name)
        except ValueError:
            return default

    provn_target_idx = col_idx('provn_target', 8)
    adm_idx          = col_idx('adm', 9)

    provn_targets: set = set()
    adms:          set = set()

    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[0] is None:
            break
        pt  = str(row[provn_target_idx]).strip() if row[provn_target_idx] is not None else ''
        adm = str(row[adm_idx]).strip()          if row[adm_idx]          is not None else ''
        if pt:
            provn_targets.add(pt)
        if adm:
            adms.add(adm)

    wb.close()

    return {
        'provn_targets': sorted(provn_targets, key=_numeric_key),
        'adms':          sorted(adms),
    }


# ── Document builders ─────────────────────────────────────────────────────────

def _update_meeting_para(para, wm_num: str, meeting_date: str) -> None:
    """
    Replace the 'Presented at the 1668th meeting on 4 June 2026' paragraph.
    Rebuilds as three runs so the ordinal suffix (th/st/nd/rd) is superscript:
      Run 1: 'Presented at the {wm_num}'  – bold, Calibri
      Run 2: '{suffix}'                   – bold, Calibri, superscript
      Run 3: ' meeting on {date}'         – bold, Calibri
    """
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    p_elem = para._p

    # Remove all existing w:r elements
    for r in list(p_elem.findall(f'{{{W}}}r')):
        p_elem.remove(r)

    # Determine the ordinal suffix
    try:
        wm_int = int(wm_num)
        suffix = ordinal_suffix(wm_int)
        num_str = str(wm_int)
    except (ValueError, TypeError):
        suffix  = 'th'
        num_str = str(wm_num) if wm_num else 'xxxx'

    # Run 1 – 'Presented at the <number>' (bold)
    p_elem.append(_make_run_xml(
        f'Presented at the {num_str}', bold=True
    ))
    # Run 2 – superscript ordinal suffix (bold + superscript)
    p_elem.append(_make_run_xml(
        suffix, bold=True, superscript=True
    ))
    # Run 3 – rest of the heading (bold)
    p_elem.append(_make_run_xml(
        f' meeting on {meeting_date}', bold=True
    ))


def _replace_provn_section(para, provn_targets: list) -> None:
    """
    In the 'obligation to coordinate' paragraph, find the section between
    '|X| to ' and 'or removed from' and replace it with highlighted runs
    for each distinct provn_target (format: 'No. <value> |O|').
    """
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    runs = para.runs

    # Find the run containing '|X| to' and the run containing 'or removed from'
    start_after = None   # index of the '|X| to ' run
    end_before  = None   # index of the 'or removed from' run

    for i, run in enumerate(runs):
        txt = run.text
        if '|X| to' in txt:
            start_after = i
        elif start_after is not None and 'or removed from' in txt:
            end_before = i
            break

    if start_after is None or end_before is None:
        return   # pattern not found; leave paragraph untouched

    # Collect the w:r XML elements that need to be deleted
    runs_to_delete = [runs[i]._r for i in range(start_after + 1, end_before)]

    p_elem   = para._p
    ref_elem = runs[start_after]._r

    # Delete old provn runs
    for r_elem in runs_to_delete:
        p_elem.remove(r_elem)

    # Insert new highlighted runs right after ref_elem
    ref_idx = list(p_elem).index(ref_elem)
    insert_pos = ref_idx + 1

    new_elems = []
    for j, target in enumerate(provn_targets):
        if j > 0:
            # separator between items
            new_elems.append(_make_run_xml(', ', highlight='yellow'))
        new_elems.append(_make_run_xml('No. ',    bold=False, highlight='yellow'))
        new_elems.append(_make_run_xml(target,    bold=True,  highlight='yellow'))
        new_elems.append(_make_run_xml(' |O|',    bold=False, highlight='yellow'))

    for offset, elem in enumerate(new_elems):
        p_elem.insert(insert_pos + offset, elem)


def _update_footer(doc, wm_num: str, d_num: str) -> None:
    """
    Update document number occurrences in all section footers.
    Replaces 'WM<digits>' with 'WM<wm_num>' and standalone 'D<digits>'
    (not preceded by a letter) with 'D<d_num>' in each footer run.
    """
    for section in doc.sections:
        for para in section.footer.paragraphs:
            for run in para.runs:
                t = run.text
                if wm_num:
                    t = re.sub(r'WM\d+', f'WM{wm_num}', t)
                if d_num:
                    t = re.sub(r'(?<![A-Za-z])D\d+', f'D{d_num}', t)
                run.text = t


def _set_cell_text(cell, text: str) -> None:
    """
    Clear a table cell and set plain text.
    Preserves the run properties (rPr) from the cloned template row so that
    font (Calibri via minorHAnsi theme) and size (10 pt) are maintained.
    """
    para   = cell.paragraphs[0]
    W      = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    p_elem = para._p

    existing_runs = p_elem.findall(f'{{{W}}}r')

    # Capture run properties from the first existing run
    existing_rPr = None
    if existing_runs:
        rPr_list = existing_runs[0].findall(f'{{{W}}}rPr')
        if rPr_list:
            existing_rPr = copy.deepcopy(rPr_list[0])

    # Remove all existing runs
    for r in existing_runs:
        p_elem.remove(r)

    # Build a new run
    r = OxmlElement('w:r')
    if existing_rPr is not None:
        # Re-use the cloned rPr (preserves theme font + size from template)
        r.append(existing_rPr)
    else:
        # Fallback: set Calibri explicitly
        rPr    = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Calibri')
        rFonts.set(qn('w:hAnsi'), 'Calibri')
        rFonts.set(qn('w:cs'),    'Calibri')
        rPr.append(rFonts)
        r.append(rPr)

    t = OxmlElement('w:t')
    t.text = text
    if text != text.strip():
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(t)
    p_elem.append(r)


def _replace_table_data(table, notice_data_list: list) -> None:
    """
    Remove all sample data rows (row index >= 1) and add one row per com_el
    record across all notices, cloning the first data row's XML for formatting.
    notice_data_list: list of (mdb_data, history_data) tuples.
    """
    # Save the first data row as formatting template (before deletion)
    template_tr = None
    if len(table.rows) > 1:
        template_tr = copy.deepcopy(table.rows[1]._tr)

    # Delete all data rows (keep header row 0)
    while len(table.rows) > 1:
        table._tbl.remove(table.rows[1]._tr)

    for mdb_data, history_data in notice_data_list:
        adm_coord = ', '.join(history_data['adms'])
        provn_col = ', '.join(
            f'No. {pt} |O|' for pt in history_data['provn_targets']
        )

        for record in mdb_data['com_el']:
            # Clone template row or add a new one
            if template_tr is not None:
                new_tr = copy.deepcopy(template_tr)
                table._tbl.append(new_tr)
                new_row = table.rows[-1]
            else:
                new_row = table.add_row()

            values = [
                record['ntc_id'],        # Notice ID
                record['adm'],           # ADM
                record['sat_name'],      # STATION
                record['type'],          # TYPE (G/N)
                record['d_rcv'],         # Date of Receipt
                adm_coord,               # ADM Coordination completed
                '',                      # ADM Coordination no longer required
                '',                      # Findings review required (13A)
                mdb_data['still_exist'], # 11.41 still exist
                provn_col,               # Provision updated from 11.41|X|
            ]
            for i, val in enumerate(values):
                if i < len(new_row.cells):
                    _set_cell_text(new_row.cells[i], val)


def generate_report(
    notice_ids: list,
    wm_num: str,
    d_num: str,
    meeting_date: str,
) -> tuple:
    """
    Generate the Word report for one or more notice IDs.
    Returns (BytesIO buffer, filename string).
    """
    # ── Gather data for all notices ───────────────────────────────────────────
    notice_data_list = []
    all_provn_targets: set = set()

    for notice_id in notice_ids:
        mdb_data     = get_mdb_data(notice_id)
        history_data = get_history_data(notice_id)
        notice_data_list.append((mdb_data, history_data))
        all_provn_targets.update(history_data['provn_targets'])

    combined_provn_targets = sorted(all_provn_targets, key=_numeric_key)

    # ── Load a fresh copy of the template ────────────────────────────────────
    doc = Document(TEMPLATE_PATH)

    # ── 1. Update meeting header paragraph ───────────────────────────────────
    for para in doc.paragraphs:
        if 'meeting on' in para.text:
            _update_meeting_para(para, wm_num, meeting_date)
            break

    # ── 2. Replace provn_target section in coordination paragraph ────────────
    for para in doc.paragraphs:
        if 'coordination agreement with the concerned administrations' in para.text:
            _replace_provn_section(para, combined_provn_targets)
            break

    # ── 3. Replace table rows ─────────────────────────────────────────────────
    if doc.tables:
        _replace_table_data(doc.tables[0], notice_data_list)

    # ── 4. Update footer with document numbers ────────────────────────────────
    _update_footer(doc, wm_num, d_num)

    # ── 5. Determine output filename ─────────────────────────────────────────
    wm_part = wm_num  if wm_num  else 'xxxx'
    d_part  = d_num   if d_num   else 'xxxxx'
    filename = f'11.41B_WM{wm_part}_D{d_part}_Draft.docx'

    # ── 6. Save to buffer and return ─────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, filename


# ── Flask routes ──────────────────────────────────────────────────────────────

@app.route('/', methods=['GET'])
def index():
    # Enumerate available notice IDs from the base directory
    notice_ids = []
    if os.path.isdir(BASE_DIR):
        for name in sorted(os.listdir(BASE_DIR)):
            folder = os.path.join(BASE_DIR, name)
            mdb    = os.path.join(folder, f'{name}.mdb')
            if os.path.isdir(folder) and os.path.isfile(mdb):
                notice_ids.append(name)
    return render_template('index.html', notice_ids=notice_ids)


@app.route('/generate', methods=['POST'])
def generate():
    # Support both multi-select (getlist) and comma-separated text input
    raw_ids    = request.form.getlist('notice_id')
    notice_ids = []
    for raw in raw_ids:
        for nid in raw.split(','):
            nid = nid.strip()
            if nid:
                notice_ids.append(nid)

    wm_num       = (request.form.get('wm_num', '')    or '').strip()
    d_num        = (request.form.get('d_num', '')     or '').strip()
    meeting_date = (request.form.get('meeting_date', '') or '').strip()

    if not notice_ids:
        return jsonify({'error': 'At least one Notice ID is required.'}), 400

    # Default meeting date to today
    if not meeting_date:
        meeting_date = datetime.today().strftime('%-d %B %Y') if os.name != 'nt' \
            else datetime.today().strftime('%#d %B %Y')

    try:
        buf, filename = generate_report(notice_ids, wm_num, d_num, meeting_date)
    except FileNotFoundError as exc:
        return jsonify({'error': str(exc)}), 404
    except Exception as exc:
        return jsonify({'error': f'Unexpected error: {exc}'}), 500

    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            'application/vnd.openxmlformats-officedocument'
            '.wordprocessingml.document'
        ),
    )


if __name__ == '__main__':
    app.run(debug=True, port=5050)
