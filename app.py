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
}

FIELD_TYPE_COLORS = {
    'Source Field (DB Column)': COLORS['source'],
    'Calculated - DEFINE': COLORS['define'],
    'Calculated - COMPUTE': COLORS['compute'],
    'BY Field (Real)': COLORS['by_real'],
    'BY Field (Calculated)': COLORS['by_calc'],
}

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
    'Raw Source Field'
]

DETAIL_WIDTHS = [38, 28, 22, 24, 14, 22, 30, 25, 14, 60, 40]

UNIQUE_FIELD_HEADERS = [
    'Field Type',
    'Field Role',
    'Field Name',
    'Source/Table',
    'Used In',
    'Formula',
    'Raw Source Field'
]

UNIQUE_FIELD_WIDTHS = [22, 24, 30, 25, 14, 60, 40]

TABLE_HEADERS = ['Source/Table']
TABLE_WIDTHS = [35]


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


def parse_fex(fex_text):
    text = strip_comments(fex_text)

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
    }

    table_src = re.findall(r'TABLE\s+FILE\s+(\S+)', text, re.IGNORECASE)
    define_src = re.findall(r'DEFINE\s+FILE\s+(\S+)', text, re.IGNORECASE)
    result['sources'] = list(dict.fromkeys(table_src + define_src))
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


def ensure_legend(wb):
    if 'Legend' in wb.sheetnames:
        return

    lg = wb.create_sheet('Legend')
    lg['A1'] = 'Color Legend'
    lg['A1'].font = Font(name='Arial', bold=True, size=11)

    items = [
        ('Source Field (DB Column)', COLORS['source'], 'Actual database columns from the source system'),
        ('Calculated - DEFINE', COLORS['define'], 'Fields derived in DEFINE FILE block'),
        ('Calculated - COMPUTE', COLORS['compute'], 'Fields computed inline inside TABLE block'),
        ('BY Field (Real)', COLORS['by_real'], 'Real DB column used as BY'),
        ('BY Field (Calculated)', COLORS['by_calc'], 'Calculated field used as BY'),
    ]

    for i, (label, (bg, fg), desc) in enumerate(items, 3):
        c = lg.cell(row=i, column=1, value=label)
        c.fill = _fill(bg)
        c.font = Font(name='Arial', size=9, color=fg, bold=True)
        lg.cell(row=i, column=2, value=desc).font = Font(name='Arial', size=9)

    lg.column_dimensions['A'].width = 30
    lg.column_dimensions['B'].width = 60


def detail_row_values(folder, fex_name, field_type, field_role, formula_step, multiple_formula,
                      field_name, source_table, used_in, formula, raw):
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
        raw
    ]


def unique_field_row_values(field_type, field_role, field_name, source_table, used_in, formula, raw):
    return [
        field_type,
        field_role,
        field_name,
        source_table,
        used_in,
        formula,
        raw
    ]


def append_rows(ws_detail, parsed, folder, fex_name, unique_fields_set, table_names_set):
    row = ws_detail.max_row + 1

    source_name_set = parsed['source_name_set']
    calculated_name_set = parsed['calculated_name_set']
    calculated_counts = parsed['calculated_counts']
    step_counter = defaultdict(int)

    def get_multiple_formula_flag(field_name):
        if field_name in calculated_counts:
            return 'Y' if calculated_counts[field_name] > 1 else 'N'
        return ''

    def add_detail_and_unique(field_type, field_name, source_table, used_in, formula='', raw='', formula_step=''):
        nonlocal row

        field_role = classify_field_role(field_name, source_name_set, calculated_name_set)
        multiple_formula = get_multiple_formula_flag(field_name)

        bg, fg = FIELD_TYPE_COLORS.get(field_type, ('FFFFFF', '000000'))

        detail_values = detail_row_values(
            folder, fex_name, field_type, field_role, formula_step, multiple_formula,
            field_name, source_table, used_in, formula, raw
        )

        for col, val in enumerate(detail_values, 1):
            _write_cell(ws_detail, row, col, val, bg, fg)
        row += 1

        unique_key = tuple(unique_field_row_values(
            field_type, field_role, field_name, source_table, used_in, formula, raw
        ))
        unique_fields_set.add(unique_key)

        if source_table:
            table_names_set.add(source_table)

    source_names_only = {f['field'] for f in parsed['source_fields']}

    for f in parsed['source_fields']:
        add_detail_and_unique(
            'Source Field (DB Column)',
            f['field'],
            f['source'],
            'DB Source'
        )

    for f in parsed['define_fields']:
        step_counter[f['field']] += 1
        add_detail_and_unique(
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
        add_detail_and_unique(
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
        add_detail_and_unique(
            'BY Field (Real)',
            f['field'],
            f['source'],
            'BY'
        )

    for f in parsed['by_calc']:
        add_detail_and_unique(
            'BY Field (Calculated)',
            f['field'],
            f['source'],
            'BY'
        )


def write_unique_fields_sheet(ws, unique_fields_set):
    setup_sheet(ws, UNIQUE_FIELD_HEADERS, UNIQUE_FIELD_WIDTHS)
    sorted_rows = sorted(unique_fields_set, key=lambda x: (
        str(x[3]).lower(),  # source/table
        str(x[2]).lower(),  # field name
        str(x[0]).lower(),  # field type
        str(x[4]).lower(),  # used in
        str(x[5]).lower(),  # formula
    ))

    for row_num, values in enumerate(sorted_rows, start=2):
        field_type = values[0]
        bg, fg = FIELD_TYPE_COLORS.get(field_type, ('FFFFFF', '000000'))
        for col_num, val in enumerate(values, start=1):
            _write_cell(ws, row_num, col_num, val, bg, fg)


def write_table_names_sheet(ws, table_names_set):
    setup_sheet(ws, TABLE_HEADERS, TABLE_WIDTHS)
    for row_num, table_name in enumerate(sorted(table_names_set, key=lambda x: str(x).lower()), start=2):
        _write_cell(ws, row_num, 1, table_name)


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

    ws_detail = wb[wb.sheetnames[0]]
    ws_detail.title = 'Detailed Fields'
    setup_sheet(ws_detail, DETAIL_HEADERS, DETAIL_WIDTHS)

    if len(wb.sheetnames) >= 2:
        ws_unique = wb[wb.sheetnames[1]]
        ws_unique.title = 'Unique Fields'
    else:
        ws_unique = wb.create_sheet('Unique Fields')

    if len(wb.sheetnames) >= 3:
        ws_tables = wb[wb.sheetnames[2]]
        ws_tables.title = 'Tables'
    else:
        ws_tables = wb.create_sheet('Tables')

    for extra_name in wb.sheetnames[3:]:
        pass

    setup_sheet(ws_unique, UNIQUE_FIELD_HEADERS, UNIQUE_FIELD_WIDTHS)
    setup_sheet(ws_tables, TABLE_HEADERS, TABLE_WIDTHS)
    ensure_legend(wb)

    return wb, ws_detail, ws_unique, ws_tables


def build_output_workbook(template_bytes, fex_items):
    wb, ws_detail, ws_unique, ws_tables = prepare_workbook(template_bytes)

    unique_fields_set = set()
    table_names_set = set()
    errors = []

    total = len(fex_items)
    progress = st.progress(0)
    status = st.empty()

    for idx, (folder, fex_name, content) in enumerate(fex_items, start=1):
        try:
            parsed = parse_fex(content)
            append_rows(ws_detail, parsed, folder, fex_name, unique_fields_set, table_names_set)
        except Exception as e:
            errors.append(f"{fex_name}: {e}")

        progress.progress(idx / total)
        status.text(f"Processing {idx} of {total}: {fex_name}")

    write_unique_fields_sheet(ws_unique, unique_fields_set)
    write_table_names_sheet(ws_tables, table_names_set)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output, errors, len(unique_fields_set), len(table_names_set)


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

            output_stream, errors, unique_field_count, table_count = build_output_workbook(
                template_file.getvalue(),
                fex_items
            )

            st.success(
                f"Completed. Processed {len(fex_items)} file(s). "
                f"Sheet2 unique fields: {unique_field_count}. "
                f"Sheet3 tables: {table_count}."
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
