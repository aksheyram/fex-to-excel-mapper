import re
import zipfile
from io import BytesIO
from pathlib import Path
from collections import Counter, defaultdict

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


WF_KEYWORDS = {
    'IF', 'THEN', 'ELSE', 'AND', 'OR', 'NOT', 'EQ', 'NE', 'GT', 'LT', 'GE', 'LE',
    'CONTAINS', 'MISSING', 'LIKE', 'IN', 'IS', 'END', 'DEFINE', 'FILE', 'TABLE',
    'SUM', 'BY', 'WHERE', 'ON', 'COMPUTE', 'PRINT', 'LIST', 'JOIN', 'TO', 'UNIQUE',
    'AS', 'SET', 'DEFAULT', 'INCLUDE', 'FORMAT', 'NOPRINT', 'SUMMARIZE',
    'COLUMN', 'TOTAL', 'PAGE', 'NUM', 'OFF', 'STYLE', 'ENDSTYLE', 'PCHOLD',
    'HTMLCSS', 'UNITS', 'PAGESIZE', 'LEFTMARGIN', 'RIGHTMARGIN', 'TOPMARGIN',
    'BOTTOMMARGIN', 'SQUEEZE', 'ORIENTATION', 'PORTRAIT', 'LANDSCAPE', 'FONT',
    'SIZE', 'BOLD', 'ITALIC', 'NORMAL', 'COLOR', 'BACKCOLOR', 'BORDER', 'JUSTIFY',
    'CENTER', 'LEFT', 'RIGHT', 'WIDTH', 'LINE', 'OBJECT', 'TEXT', 'FIELD', 'ITEM',
    'TYPE', 'REPORT', 'TITLE', 'HEADING', 'FOOTING', 'SUBHEAD', 'SUBFOOT',
    'SUBTOTAL', 'GRANDTOTAL', 'ACROSSVALUE', 'ACROSSTITLE', 'TABHEADING',
    'TABFOOTING', 'SILVER', 'RGB', 'LIGHT',
}

THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP_TOP = Alignment(wrap_text=True, vertical='top')

COLORS = {
    'header': ('1F4E79', 'FFFFFF'),
    'source': ('DDEBF7', '000000'),
    'define': ('E2EFDA', '000000'),
    'compute': ('FFF2CC', '000000'),
    'by_real': ('EBF3FB', '000000'),
    'by_calc': ('F4ECFA', '000000'),
    # Group ID column colors
    'group_unique': ('F2F2F2', '888888'),   # grey for unique/solo FEX files
    'group_dup': ('FFE699', '7B5B00'),       # amber for grouped duplicates
}

FIELD_TYPE_COLORS = {
    'Source Field (DB Column)': COLORS['source'],
    'Calculated - DEFINE': COLORS['define'],
    'Calculated - COMPUTE': COLORS['compute'],
    'BY Field (Real)': COLORS['by_real'],
    'BY Field (Calculated)': COLORS['by_calc'],
}

# Sheet1 now has 'Group ID' as the last column
DETAIL_HEADERS = [
    'Folder',
    'Fex name',
    'Field Type',
    'Field Role',
    'Formula Step',
    'Multiple Formula (Y/N)',
    'Field Name',
    'Source/Table',
    'Used In',
    'Formula',
    'Raw Source Field',
    'Group ID',          # NEW
]
DETAIL_WIDTHS = [38, 28, 22, 24, 14, 22, 30, 25, 14, 60, 40, 14]

SHEET2_HEADERS = ['Source/Table', 'Field Name']
SHEET2_WIDTHS = [35, 35]

# Sheet3 - duplicate/group summary
SHEET3_HEADERS = [
    'Group ID',
    'FEX Count in Group',
    'Folder',
    'Fex Name',
    'Source Tables',
    'Total Fields',
]
SHEET3_WIDTHS = [14, 20, 38, 35, 50, 14]


def _fill(bg):
    return PatternFill('solid', start_color=bg)


def _font(fg, bold=False):
    return Font(name='Arial', size=9, color=fg, bold=bold)


def _write_cell(ws, row, col, val, bg='FFFFFF', fg='000000', bold=False):
    c = ws.cell(row=row, column=col, value=val)
    c.fill = _fill(bg)
    c.font = _font(fg, bold)
    c.alignment = WRAP_TOP
    c.border = BORDER


def setup_sheet(ws, headers, widths):
    hbg, hfg = COLORS['header']
    for col, (h, w) in enumerate(zip(headers, widths), 1):
        _write_cell(ws, 1, col, h, hbg, hfg, bold=True)
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 22


def clear_sheet(ws):
    if ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)


def raw_db_fields(formula, defined_names):
    tokens = re.findall(r'\b([A-Z][A-Z0-9_]{2,})\b', formula)
    return sorted({t for t in tokens if t not in WF_KEYWORDS and t not in defined_names})


def strip_comments(text):
    lines = [
        line for line in text.splitlines()
        if not line.strip().startswith('-*') and not line.strip().startswith('-!')
    ]
    return '\n'.join(lines)


def classify_field_role(field_name, source_name_set, calculated_name_set):
    in_source = field_name in source_name_set
    in_calc = field_name in calculated_name_set

    if in_source and in_calc:
        return 'Both DB Source and Calculated'
    if in_source:
        return 'DB Source Only'
    if in_calc:
        return 'Calculated Only'
    return ''


def extract_hold_names(text):
    hold_names = set()
    for name in re.findall(r'ON\s+TABLE\s+HOLD\s+AS\s+([A-Za-z_]\w*)', text, re.IGNORECASE):
        hold_names.add(name.upper())
    hold_names.add('HOLD')
    return hold_names


def is_hold_like_table(table_name, explicit_hold_names=None):
    if not table_name:
        return False
    t = str(table_name).strip().upper()
    if not t:
        return False
    explicit_hold_names = explicit_hold_names or set()
    if t.startswith('&'):
        return True
    if t in explicit_hold_names:
        return True
    if t.startswith('HOLD'):
        return True
    if t.startswith('HLD'):
        return True
    if 'HOLD' in t:
        return True
    return False


def parse_fex(fex_text):
    text = strip_comments(fex_text)
    explicit_hold_names = extract_hold_names(text)

    result = {
        'sources': [],
        'define_fields': [],
        'compute_fields': [],
        'sum_real': [],
        'sum_calc': [],
        'by_real': [],
        'by_calc': [],
        'source_fields': [],
        'calculated_counts': {},
        'source_name_set': set(),
        'calculated_name_set': set(),
        'real_sources': [],
        'hold_names': explicit_hold_names,
    }

    table_src = re.findall(r'TABLE\s+FILE\s+(\S+)', text, re.IGNORECASE)
    define_src = re.findall(r'DEFINE\s+FILE\s+(\S+)', text, re.IGNORECASE)
    result['sources'] = list(dict.fromkeys(table_src + define_src))

    result['real_sources'] = [
        s for s in result['sources']
        if not is_hold_like_table(s, explicit_hold_names)
    ]

    primary = result['sources'][0] if result['sources'] else ''
    def_src = define_src[0] if define_src else primary

    for block in re.findall(r'DEFINE\s+FILE\s+\S+\s*(.*?)END', text, re.IGNORECASE | re.DOTALL):
        for fname, fmt, formula in re.findall(
            r'([A-Za-z_]\w*)\s*/\s*([A-Za-z0-9%.]+)\s*=\s*(.*?);',
            block,
            re.DOTALL
        ):
            result['define_fields'].append({
                'field': fname,
                'format': fmt,
                'formula': ' '.join(formula.split()),
                'source': def_src
            })

    defined_names = {f['field'] for f in result['define_fields']}
    raw_set = set()

    for f in result['define_fields']:
        f['raw_fields'] = raw_db_fields(f['formula'], defined_names)
        for r in f['raw_fields']:
            raw_set.add((r, f['source']))

    for tbl_src, block in re.findall(
        r'TABLE\s+FILE\s+(\S+)\s*(.*?)(?=\nEND\b|\Z)',
        text,
        re.IGNORECASE | re.DOTALL
    ):
        for fname, fmt, formula, alias in re.findall(
            r'COMPUTE\s+([A-Za-z_]\w*)\s*/\s*([A-Za-z0-9%.]+)\s*=\s*(.*?);(?:\s*AS\s*[\'"]([^\'"]*)[\'"])?',
            block,
            re.IGNORECASE | re.DOTALL
        ):
            fc = ' '.join(formula.split())
            raws = raw_db_fields(fc, defined_names)
            result['compute_fields'].append({
                'field': fname,
                'format': fmt,
                'formula': fc,
                'alias': alias or '',
                'source': tbl_src,
                'raw_fields': raws
            })
            for r in raws:
                raw_set.add((r, tbl_src))

        ss = re.search(
            r'(?:SUM|PRINT)\b(.*?)(?=\nBY\b|\nWHERE\b|\nON\s+TABLE\b|\Z)',
            block,
            re.IGNORECASE | re.DOTALL
        )
        if ss:
            for line in ss.group(1).splitlines():
                line = line.strip()
                if not line or re.match(r'COMPUTE\b', line, re.IGNORECASE):
                    continue
                m = re.match(r'([A-Za-z_]\w*)', line)
                if not m:
                    continue
                fn = m.group(1)
                if fn.upper() in ('BY', 'WHERE', 'ON', 'SUM', 'PRINT', 'END', 'COMPUTE'):
                    continue
                if fn in defined_names:
                    result['sum_calc'].append({'field': fn, 'source': tbl_src})
                else:
                    result['sum_real'].append({'field': fn, 'source': tbl_src})
                    raw_set.add((fn, tbl_src))

        for m in re.finditer(r'\bBY\s+([A-Za-z_]\w*)', block, re.IGNORECASE):
            fn = m.group(1)
            if fn.upper() in ('TABLE', 'ON', 'END'):
                continue
            if fn in defined_names:
                result['by_calc'].append({'field': fn, 'source': tbl_src})
            else:
                result['by_real'].append({'field': fn, 'source': tbl_src})
                raw_set.add((fn, tbl_src))

    seen = set()
    for fn, src in sorted(raw_set):
        if (fn, src) not in seen:
            seen.add((fn, src))
            result['source_fields'].append({'field': fn, 'source': src})

    calculated_names = [f['field'] for f in result['define_fields']] + [f['field'] for f in result['compute_fields']]
    result['calculated_counts'] = dict(Counter(calculated_names))
    result['calculated_name_set'] = set(calculated_names)
    result['source_name_set'] = {f['field'] for f in result['source_fields']}

    return result


# ---------------------------------------------------------------------------
# GROUPING LOGIC
# ---------------------------------------------------------------------------

def compute_fex_fingerprint(parsed):
    """
    A fingerprint uniquely represents what a FEX 'does':
      - the set of real (non-HOLD) source tables
      - the set of all field names (source + calculated)

    Two FEX files with the same fingerprint are considered duplicates /
    versions of the same logical report.
    """
    real_tables = frozenset(t.upper() for t in parsed['real_sources'])

    all_fields = frozenset(
        f['field'].upper() for f in parsed['source_fields']
    ) | frozenset(
        f['field'].upper() for f in parsed['define_fields']
    ) | frozenset(
        f['field'].upper() for f in parsed['compute_fields']
    )

    return (real_tables, all_fields)


def build_group_map(fex_fingerprints):
    """
    fex_fingerprints: list of (folder, fex_name, fingerprint)

    Returns:
        group_map: dict  (folder, fex_name) -> group_id string
        groups:    dict  fingerprint -> list of (folder, fex_name)
    """
    groups = defaultdict(list)
    for folder, fex_name, fp in fex_fingerprints:
        groups[fp].append((folder, fex_name))

    # Assign group IDs only to fingerprints that appear more than once
    group_map = {}
    group_counter = 1

    for fp, members in groups.items():
        if len(members) > 1:
            gid = f'GRP-{group_counter:03d}'
            group_counter += 1
            for key in members:
                group_map[key] = gid
        else:
            group_map[members[0]] = 'Unique'

    return group_map, groups


# ---------------------------------------------------------------------------

def detail_row_values(folder, fex_name, field_type, field_role, formula_step, multiple_formula,
                      field_name, source_table, used_in, formula, raw, group_id):
    return [
        folder,
        fex_name,
        field_type,
        field_role,
        formula_step,
        multiple_formula,
        field_name,
        source_table,
        used_in,
        formula,
        raw,
        group_id,       # NEW
    ]


def append_rows(ws_detail, parsed, folder, fex_name, sheet2_pairs, group_id):
    row = ws_detail.max_row + 1

    source_name_set = parsed['source_name_set']
    calculated_name_set = parsed['calculated_name_set']
    calculated_counts = parsed['calculated_counts']
    step_counter = defaultdict(int)
    real_sources = parsed['real_sources']
    hold_names = parsed['hold_names']

    # Pick Group ID cell styling
    if group_id == 'Unique':
        grp_bg, grp_fg = COLORS['group_unique']
    else:
        grp_bg, grp_fg = COLORS['group_dup']

    def get_multiple_formula_flag(field_name):
        if field_name in calculated_counts:
            return 'Y' if calculated_counts[field_name] > 1 else 'N'
        return ''

    def add_detail_and_sheet2(field_type, field_name, source_table, used_in, formula='', raw='', formula_step=''):
        nonlocal row

        field_role = classify_field_role(field_name, source_name_set, calculated_name_set)
        multiple_formula = get_multiple_formula_flag(field_name)

        bg, fg = FIELD_TYPE_COLORS.get(field_type, ('FFFFFF', '000000'))

        detail_values = detail_row_values(
            folder, fex_name, field_type, field_role, formula_step, multiple_formula,
            field_name, source_table, used_in, formula, raw, group_id
        )

        # Write all columns except Group ID with the field-type color
        for col, val in enumerate(detail_values[:-1], 1):
            _write_cell(ws_detail, row, col, val, bg, fg)

        # Group ID column gets its own color to stand out
        _write_cell(ws_detail, row, len(detail_values), group_id, grp_bg, grp_fg, bold=(group_id != 'Unique'))
        row += 1

        if not field_name:
            return

        if source_table and not is_hold_like_table(source_table, hold_names):
            sheet2_pairs.add((source_table, field_name))
        else:
            for real_table in real_sources:
                sheet2_pairs.add((real_table, field_name))

    source_names_only = {f['field'] for f in parsed['source_fields']}

    for f in parsed['source_fields']:
        add_detail_and_sheet2(
            'Source Field (DB Column)',
            f['field'],
            f['source'],
            'DB Source'
        )

    for f in parsed['define_fields']:
        step_counter[f['field']] += 1
        add_detail_and_sheet2(
            'Calculated - DEFINE',
            f['field'],
            f['source'],
            'DEFINE FILE',
            f['formula'],
            ', '.join(f['raw_fields']),
            step_counter[f['field']]
        )

    for f in parsed['compute_fields']:
        step_counter[f['field']] += 1
        add_detail_and_sheet2(
            'Calculated - COMPUTE',
            f['field'],
            f['source'],
            'TABLE/COMPUTE',
            f['formula'],
            ', '.join(f['raw_fields']),
            step_counter[f['field']]
        )

    for f in parsed['by_real']:
        if f['field'] in source_names_only:
            continue
        add_detail_and_sheet2(
            'BY Field (Real)',
            f['field'],
            f['source'],
            'BY'
        )

    for f in parsed['by_calc']:
        add_detail_and_sheet2(
            'BY Field (Calculated)',
            f['field'],
            f['source'],
            'BY'
        )


def write_sheet2(ws_sheet2, sheet2_pairs):
    setup_sheet(ws_sheet2, SHEET2_HEADERS, SHEET2_WIDTHS)
    for row_num, (table_name, field_name) in enumerate(
        sorted(sheet2_pairs, key=lambda x: (str(x[0]).lower(), str(x[1]).lower())),
        start=2
    ):
        _write_cell(ws_sheet2, row_num, 1, table_name)
        _write_cell(ws_sheet2, row_num, 2, field_name)


def write_sheet3(ws_sheet3, groups, group_map, fex_parsed_map):
    """
    Sheet3 - Duplicate / Group Summary

    One row per FEX file that belongs to a named group (GRP-xxx).
    Unique FEX files are listed at the bottom under 'Unique'.
    Columns: Group ID | FEX Count in Group | Folder | Fex Name | Source Tables | Total Fields
    """
    setup_sheet(ws_sheet3, SHEET3_HEADERS, SHEET3_WIDTHS)

    row = 2

    # Sort: grouped first (GRP-xxx), then Unique
    sorted_groups = sorted(
        groups.items(),
        key=lambda kv: (
            '0' if len(kv[1]) > 1 else '1',    # grouped before unique
            group_map.get(kv[1][0], 'Unique')   # then by group ID
        )
    )

    for fp, members in sorted_groups:
        real_tables_fp, all_fields_fp = fp
        group_size = len(members)
        gid = group_map.get(members[0], 'Unique')

        if gid == 'Unique':
            bg, fg, bold = COLORS['group_unique'][0], COLORS['group_unique'][1], False
        else:
            bg, fg, bold = COLORS['group_dup'][0], COLORS['group_dup'][1], True

        tables_str = ', '.join(sorted(real_tables_fp))
        field_count = len(all_fields_fp)

        for folder, fex_name in sorted(members):
            _write_cell(ws_sheet3, row, 1, gid, bg, fg, bold=bold)
            _write_cell(ws_sheet3, row, 2, group_size, bg, fg)
            _write_cell(ws_sheet3, row, 3, folder, bg, fg)
            _write_cell(ws_sheet3, row, 4, fex_name, bg, fg)
            _write_cell(ws_sheet3, row, 5, tables_str, bg, fg)
            _write_cell(ws_sheet3, row, 6, field_count, bg, fg)
            row += 1


def read_uploaded_fex(uploaded_file):
    return uploaded_file.read().decode('utf-8', errors='replace')


def collect_fex_from_zip(uploaded_zip):
    files = []
    with zipfile.ZipFile(uploaded_zip, 'r') as zf:
        for name in zf.namelist():
            if name.lower().endswith('.fex'):
                try:
                    content = zf.read(name).decode('utf-8', errors='replace')
                    files.append((str(Path(name).parent), Path(name).name, content))
                except Exception as e:
                    raise ValueError(f"Failed reading {name}: {e}")
    return files


def prepare_workbook(template_bytes):
    wb = load_workbook(BytesIO(template_bytes))

    if 'Legend' in wb.sheetnames:
        wb.remove(wb['Legend'])

    ws_detail = wb[wb.sheetnames[0]]
    ws_detail.title = 'Sheet1'
    clear_sheet(ws_detail)
    setup_sheet(ws_detail, DETAIL_HEADERS, DETAIL_WIDTHS)

    if 'Sheet2' in wb.sheetnames:
        ws_sheet2 = wb['Sheet2']
        clear_sheet(ws_sheet2)
    else:
        ws_sheet2 = wb.create_sheet('Sheet2')

    setup_sheet(ws_sheet2, SHEET2_HEADERS, SHEET2_WIDTHS)

    if 'Sheet3' in wb.sheetnames:
        ws_sheet3 = wb['Sheet3']
        clear_sheet(ws_sheet3)
    else:
        ws_sheet3 = wb.create_sheet('Sheet3')

    ws_sheet3.title = 'Duplicate Groups'
    setup_sheet(ws_sheet3, SHEET3_HEADERS, SHEET3_WIDTHS)

    return wb, ws_detail, ws_sheet2, ws_sheet3


def build_output_workbook(template_bytes, fex_items):
    wb, ws_detail, ws_sheet2, ws_sheet3 = prepare_workbook(template_bytes)

    sheet2_pairs = set()
    errors = []

    # -------------------------------------------------------------------
    # PASS 1: Parse all FEX files and compute fingerprints
    # -------------------------------------------------------------------
    total = len(fex_items)
    progress = st.progress(0)
    status = st.empty()

    parsed_results = []         # list of (folder, fex_name, parsed)
    fex_fingerprints = []       # list of (folder, fex_name, fingerprint)

    status.text("Pass 1 of 2: Parsing FEX files...")
    for idx, (folder, fex_name, content) in enumerate(fex_items, start=1):
        try:
            parsed = parse_fex(content)
            fp = compute_fex_fingerprint(parsed)
            parsed_results.append((folder, fex_name, parsed))
            fex_fingerprints.append((folder, fex_name, fp))
        except Exception as e:
            errors.append(f"{fex_name}: {e}")
            parsed_results.append((folder, fex_name, None))
            fex_fingerprints.append((folder, fex_name, None))

        progress.progress(idx / total * 0.5)

    # -------------------------------------------------------------------
    # Build group map from fingerprints (skip failed parses)
    # -------------------------------------------------------------------
    valid_fingerprints = [
        (folder, fex_name, fp)
        for folder, fex_name, fp in fex_fingerprints
        if fp is not None
    ]
    group_map, groups = build_group_map(valid_fingerprints)

    # -------------------------------------------------------------------
    # PASS 2: Write rows with group IDs
    # -------------------------------------------------------------------
    status.text("Pass 2 of 2: Writing output...")
    for idx, (folder, fex_name, parsed) in enumerate(parsed_results, start=1):
        if parsed is None:
            continue
        try:
            gid = group_map.get((folder, fex_name), 'Unique')
            append_rows(ws_detail, parsed, folder, fex_name, sheet2_pairs, gid)
        except Exception as e:
            errors.append(f"{fex_name} (write): {e}")

        progress.progress(0.5 + idx / total * 0.5)

    write_sheet2(ws_sheet2, sheet2_pairs)

    # fex_parsed_map not needed in write_sheet3 currently but kept for extensibility
    write_sheet3(ws_sheet3, groups, group_map, fex_parsed_map={})

    # Count duplicate groups (groups with more than one member)
    dup_group_count = sum(1 for members in groups.values() if len(members) > 1)
    dup_fex_count = sum(len(members) for members in groups.values() if len(members) > 1)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output, errors, len(sheet2_pairs), dup_group_count, dup_fex_count


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="WebFOCUS FEX to Excel Mapper", layout="wide")
st.title("WebFOCUS FEX to Excel Mapper")

st.markdown("Upload your template and either FEX files or a ZIP containing FEX files.")

col1, col2 = st.columns(2)

with col1:
    template_file = st.file_uploader("Upload Template XLSX", type=["xlsx"])
    mode = st.radio("Input Type", ["Multiple FEX Files", "ZIP File"], horizontal=True)

with col2:
    if mode == "Multiple FEX Files":
        uploaded_fex_files = st.file_uploader("Upload one or more .fex files", type=["fex"], accept_multiple_files=True)
        uploaded_zip = None
    else:
        uploaded_zip = st.file_uploader("Upload ZIP file", type=["zip"])
        uploaded_fex_files = []

output_name = st.text_input("Output file name", value="Ulbrich_output.xlsx")

with st.expander("Grouping Info", expanded=False):
    st.markdown(
        """
        **How duplicate grouping works:**

        Two FEX files are placed in the same group when they share **exactly** the same:
        - Set of real DB source tables (HOLD files are excluded)
        - Set of all field names (source fields + DEFINE + COMPUTE)

        This identifies reports that are effectively copies of each other (e.g., created
        because WebFOCUS forces a copy when a second user edits a shared report).

        - **Sheet1** gets a new `Group ID` column. Grouped files show `GRP-001`, `GRP-002` etc.
          in amber. Unique files show `Unique` in grey.
        - **Sheet3 (Duplicate Groups)** lists every group with member FEX files, their shared
          source tables, and total field count.
        """
    )

if st.button("Run Mapping", type="primary"):
    if not template_file:
        st.error("Please upload the template Excel file.")
    else:
        fex_items = []

        try:
            if mode == "Multiple FEX Files":
                if not uploaded_fex_files:
                    st.error("Please upload at least one .fex file.")
                    st.stop()

                for f in uploaded_fex_files:
                    content = read_uploaded_fex(f)
                    fex_items.append(("uploaded_files", f.name, content))

            else:
                if not uploaded_zip:
                    st.error("Please upload a ZIP file.")
                    st.stop()

                fex_items = collect_fex_from_zip(uploaded_zip)

                if not fex_items:
                    st.error("No .fex files found inside the ZIP.")
                    st.stop()

            output_stream, errors, sheet2_count, dup_group_count, dup_fex_count = build_output_workbook(
                template_file.getvalue(),
                fex_items
            )

            st.success(
                f"Completed. Processed **{len(fex_items)}** file(s). "
                f"Sheet2 unique source/field pairs: **{sheet2_count}**. "
                f"Duplicate groups found: **{dup_group_count}** "
                f"(covering **{dup_fex_count}** FEX files)."
            )

            st.download_button(
                label="Download Output Excel",
                data=output_stream,
                file_name=output_name if output_name.lower().endswith(".xlsx") else f"{output_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            if errors:
                st.warning(f"{len(errors)} file(s) had errors.")
                with st.expander("View Error Log"):
                    for err in errors:
                        st.text(err)

        except Exception as e:
            st.error(str(e))
