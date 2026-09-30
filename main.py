from copy import copy
import os
import re
import openpyxl
import unicodedata
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
SKIP_SHEETS = ['Abiinfo', '2027.Kevadsemetri põhi']

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, 'output_files')

CAL_SHEET = '2027.Kevadsemetri põhi'
GREY = 'D9D9D9'  # Color for "teacher has no classes for this group"
SHORT_DAY = 'E2EFDA'  # Color for "Lühendatud tööpäev"
HEADER_BLUE = '0070C0'  # Default table header color

# Explicit mapping for known spelling variations (preserving TT and specific notes)
TEACHER_ALIASES = {
    'Andu Rämmer': 'Andu Rämmer',
    'Andu Rämmer (1 paar)': 'Andu Rämmer',
    'Elvira Küün': 'Elvira Küün',
    'Elvira  Küün': 'Elvira Küün',
    'Kayleigh Kleiva': 'Kayleigh Kleiva',
    'Kayleigh Mari Kleiva': 'Kayleigh Kleiva',
    'Kayleigh Mary Kleiva': 'Kayleigh Kleiva',
    'Külli Kallas TT': 'Külli Kallas TT',
    'Külli Kallas': 'Külli Kallas TT',
    'Merju Lomp': 'Merju Lomp-Alev',
    'Merju Lomp-Alev': 'Merju Lomp-Alev',
    'Viivika Roostar TT': 'Viivika Roostar TT',
    'Viivika Roostar': 'Viivika Roostar TT',
    'Viivika Roostar (TT)': 'Viivika Roostar TT',
    'Vilja Vendelin Reigo': 'Vilja Vendelin-Reigo TT',
    'Vilja Vendelin Reigo TT': 'Vilja Vendelin-Reigo TT',
    'Vilja Vendelin-Reigo': 'Vilja Vendelin-Reigo TT',
    'Vilja Vendelin-Reigo TT': 'Vilja Vendelin-Reigo TT',
}

# ---------------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------------
def normalize(text):
    """Converts text to lowercase, removes extra spaces and non-breaking spaces."""
    text = unicodedata.normalize('NFC', str(text)).replace('\xa0', ' ')
    return ' '.join(text.split()).casefold()

def standardize_teacher_name(name):
    """Maps variant teacher names to a single canonical name using TEACHER_ALIASES."""
    if not name or str(name).lower() == 'nan':
        return ''

    cleaned = ' '.join(str(name).split())

    # 1. Direct or normalized lookup in the alias dictionary
    if cleaned in TEACHER_ALIASES:
        return TEACHER_ALIASES[cleaned]

    norm_cleaned = normalize(cleaned)
    for alias, canonical in TEACHER_ALIASES.items():
        if normalize(alias) == norm_cleaned:
            return canonical

    # 2. Fallback: strip parentheses like "(1 paar)" but KEEP "TT" if present
    base_name = re.sub(r'\s*\(.*?\)', '', cleaned).strip()

    if base_name in TEACHER_ALIASES:
        return TEACHER_ALIASES[base_name]

    norm_base = normalize(base_name)
    for alias, canonical in TEACHER_ALIASES.items():
        if normalize(alias) == norm_base:
            return canonical

    return base_name


def tidy(value):
    """Converts float values with a zero decimal part into integers (3.0 -> 3)."""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def find_col(df, prefix):
    """Finds the actual column name by its starting substring."""
    for col in df.columns:
        if normalize(col).startswith(prefix):
            return col
    return None


def fill(rgb):
    """Returns a cell fill with the specified RGB color."""
    return PatternFill(start_color=rgb, end_color=rgb, fill_type='solid')


def parse_teacher_names(cell_value):
    """Splits a cell containing teacher names and standardizes them."""
    if not cell_value or str(cell_value).strip() == '':
        return []
    text = str(cell_value).replace(' ja ', ',').replace(';', ',')
    names = [n.strip() for n in text.split(',')]
    standardized = [
        standardize_teacher_name(n)
        for n in names
        if n and n.lower() != 'nan'
    ]
    return [n for n in standardized if n]


# ---------------------------------------------------------------------------
# Block 1: Automatic Scanning of All Teachers and Data Collection
# ---------------------------------------------------------------------------
def get_all_teachers_map(file_path):
    """Scans all sheets and collects teaching loads."""
    xls = pd.ExcelFile(file_path)
    all_rows = []

    for sheet_name in xls.sheet_names:
        if sheet_name in SKIP_SHEETS:
            continue

        df = pd.read_excel(file_path, sheet_name=sheet_name).fillna('')

        col_resp = find_col(df, 'vastutav')
        col_other = find_col(df, 'õppejõud kes on veel')
        col_group = find_col(df, 'õpperühm')
        col_group2 = find_col(df, 'kelle tunniplaaniga')
        col_code = find_col(df, 'aine kood')
        col_name = find_col(df, 'aine nimetus')
        col_eap = find_col(df, 'aine maht')
        col_pairs = find_col(df, 'kontaktpaare')
        col_status = find_col(df, 'lepinguline')

        if col_resp is None:
            continue

        for _, row in df.iterrows():
            resp_list = parse_teacher_names(row[col_resp])
            other_list = (
                parse_teacher_names(row[col_other])
                if col_other is not None
                else []
            )

            for teacher in resp_list:
                all_rows.append({
                    'Teacher': teacher,
                    'Sheet': sheet_name,
                    'Grupp': tidy(row[col_group]) if col_group is not None else '',
                    'Teised_ruhmad': (
                        tidy(row[col_group2]) if col_group2 is not None else ''
                    ),
                    'Kood': tidy(row[col_code]) if col_code is not None else '',
                    'Nimetus': tidy(row[col_name]) if col_name is not None else '',
                    'EAP': tidy(row[col_eap]) if col_eap is not None else '',
                    'Kontaktpaarid': (
                        tidy(row[col_pairs]) if col_pairs is not None else ''
                    ),
                    'Staatus': tidy(row[col_status]) if col_status is not None else '',
                    'Roll': 'vastutav',
                })

            for teacher in other_list:
                all_rows.append({
                    'Teacher': teacher,
                    'Sheet': sheet_name,
                    'Grupp': tidy(row[col_group]) if col_group is not None else '',
                    'Teised_ruhmad': (
                        tidy(row[col_group2]) if col_group2 is not None else ''
                    ),
                    'Kood': tidy(row[col_code]) if col_code is not None else '',
                    'Nimetus': tidy(row[col_name]) if col_name is not None else '',
                    'EAP': tidy(row[col_eap]) if col_eap is not None else '',
                    'Kontaktpaarid': (
                        tidy(row[col_pairs]) if col_pairs is not None else ''
                    ),
                    'Staatus': tidy(row[col_status]) if col_status is not None else '',
                    'Roll': 'kaasõppejõud',
                })

    return pd.DataFrame(all_rows)


# ---------------------------------------------------------------------------
# Block 2: Helper Functions for Excel Generation
# ---------------------------------------------------------------------------
def teacher_status(df):
    """Determines teacher status."""
    for s in df['Staatus']:
        s = normalize(s)
        if s.startswith('tt') or s.startswith('tunnitas'):
            return 'TT'
        if s.startswith('l'):
            return 'lepinguline'
    return 'status not specified'


def parse_groups(text):
    """Extracts study groups from text."""
    return {
        (p, int(y))
        for p, y in re.findall(r'([A-ZÕÄÖÜ]{2,4})\s*(\d)', str(text))
    }


def teacher_groups(df):
    """Collects all unique pairs of (program, year) for a teacher."""
    groups = set()
    for col in ('Grupp', 'Teised_ruhmad'):
        for text in df[col]:
            groups |= parse_groups(text)
    return groups


def read_study_modes(abi):
    """Reads study modes from the Abiinfo sheet."""
    modes = {}
    for row in abi.iter_rows():
        code, kind = (
            str(row[1].value or '').strip(),
            str(row[4].value or '').strip().casefold(),
        )
        if re.fullmatch(r'[A-ZÕÄÖÜ]{2,5}', code) and kind.endswith('õpe'):
            modes[code] = kind
    return modes


def read_legend(cal, modes):
    """Reads calendar coloring rules based on the legend row and extracts color mapping."""
    codes = '|'.join(sorted(modes, key=len, reverse=True))
    pattern = re.compile(
        r'(' + codes + r')\s*(\d(?:\s*,\s*\d(?![A-Za-zÕÄÖÜõäöü]))*)'
    )
    rules = {}
    color_to_groups = {}

    for cell in cal[2]:
        text = ' '.join(str(cell.value or '').split())
        if not text or not cell.fill.fill_type or cell.fill.fgColor.type != 'rgb':
            continue
        rgb = cell.fill.fgColor.rgb[-6:]
        toks = {
            (m.group(1), int(y))
            for m in pattern.finditer(text)
            for y in re.findall(r'\d', m.group(2))
        }
        if toks:
            rules[rgb] = ('groups', toks)
            color_to_groups[rgb] = toks
        elif 'eksam' in text.casefold():
            rules[rgb] = ('session', None)
        else:
            rules[rgb] = ('always', None)
    return rules, color_to_groups


def rule_applies(rule, groups, modes):
    """Checks whether a calendar day should remain colored for a given teacher."""
    kind, toks = rule
    if kind == 'always':
        return True
    if kind == 'session':
        return any(modes.get(p) == 'sessioonõpe' for p, _ in groups)
    return bool(groups & toks)


def get_subject_color(group_text, color_to_groups):
    """Determines if the subject's group matches any legend color rule."""
    subj_groups = parse_groups(group_text)
    if not subj_groups:
        return None
    for rgb, allowed_groups in color_to_groups.items():
        if subj_groups & allowed_groups:
            return rgb
    return None


# ---------------------------------------------------------------------------
# Block 3: Generation of Individual Excel File for a Teacher
# ---------------------------------------------------------------------------
def create_teacher_excel(teacher_title, subjects_df, master_path):
    wb = openpyxl.load_workbook(master_path)
    cal, abi = wb[CAL_SHEET], wb['Abiinfo']
    modes = read_study_modes(abi)
    rules, color_to_groups = read_legend(cal, modes)
    groups = teacher_groups(subjects_df)

    ws = wb.create_sheet('2026_27 Kevadsemestri põhi')
    n = len(subjects_df)
    offset = n + 1

    # File header (Teacher name and status)
    ws['A1'] = f'{teacher_title} — {teacher_status(subjects_df)}'
    ws['A1'].font = Font(name='Times New Roman', size=14, bold=True)

    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    base_font = Font(name='Times New Roman', size=11)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    headers = {
        1: 'Rühmad',
        2: 'Õppeaine kood',
        3: 'Õppeaine nimetus',
        9: 'EAP',
        10: 'Soovituslik paaride arv (90 min)',
        11: (
            'Kontakttundide soov 90 min (kui soovite vähem/rohkem kontakttunde'
            ' tunniplaani, siis palun siia täpset arvu)'
        ),
        12: 'Mitu paari järjest võib panna?',
        13: 'Kas aine lõpeb eksamiga (E) või arvestusega (A)?',
        14: 'Eksami ajad',
        15: 'Arvestuse aeg',
        16: 'Soovitud auditoorium',
        17: (
            'Lisainfo\ne-õpe, väljasõit, ettevõtete külastus, külalisõppejõud. Kui'
            ' õppeainel on mitu õppejõudu, siis palun märkida siia kuidas on'
            ' õppetöö korraldatud jne'
        ),
    }
    for col in range(1, 23):
        c = ws.cell(row=2, column=col, value=headers.get(col))
        c.font = Font(name='Times New Roman', size=11, bold=True)
        c.fill, c.alignment, c.border = fill(HEADER_BLUE), center, border
    ws.row_dimensions[2].height = 135

    # Filling subject rows for the teacher and coloring based on the legend
    for i, (_, subj) in enumerate(subjects_df.iterrows()):
        r = 3 + i
        group_val = subj['Grupp']
        subj_color = get_subject_color(group_val, color_to_groups)

        values = {
            1: group_val,
            2: subj['Kood'],
            3: subj['Nimetus'],
            9: subj['EAP'],
            10: subj['Kontaktpaarid'],
            17: 'Co-lecturer' if subj['Roll'] == 'kaasõppejõud' else '',
        }
        for col in range(1, 23):
            c = ws.cell(row=r, column=col, value=values.get(col))
            c.font, c.border = base_font, border
            c.alignment = left if col in (3, 17) else center
            if subj_color:
                c.fill = fill(subj_color)

        ws.row_dimensions[r].height = 30

    # Merging cells for names and additional information
    for r in range(2, 3 + n):
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        ws.merge_cells(start_row=r, start_column=17, end_row=r, end_column=22)

    # Copying the calendar
    for row in cal.iter_rows(min_row=2, max_row=cal.max_row):
        for c in row:
            d = ws.cell(row=c.row + offset, column=c.column, value=c.value)
            if c.has_style:
                d._style = copy(c._style)
    for mr in cal.merged_cells.ranges:
        if mr.min_row >= 2:
            ws.merge_cells(
                start_row=mr.min_row + offset,
                end_row=mr.max_row + offset,
                start_column=mr.min_col,
                end_column=mr.max_col,
            )
    for r, dim in cal.row_dimensions.items():
        if r >= 2 and dim.height:
            ws.row_dimensions[r + offset].height = dim.height
    for key, dim in cal.column_dimensions.items():
        ws.column_dimensions[key].width = dim.width

    # Adding the 'Lühendatud tööpäev' legend element
    legend_row = 2 + offset
    exam = next(c for c in ws[legend_row] if 'eksam' in str(c.value or '').casefold())
    col = exam.column + 3
    label = ws.cell(row=legend_row, column=col, value='Lühendatud tööpäev')
    label._style = copy(exam._style)
    label.fill = fill(SHORT_DAY)
    ws.cell(row=legend_row, column=col + 1)._style = copy(exam._style)
    ws.cell(row=legend_row, column=col + 1).fill = fill(SHORT_DAY)
    ws.merge_cells(
        start_row=legend_row,
        start_column=col,
        end_row=legend_row,
        end_column=col + 1,
    )
    ws.row_dimensions[legend_row].height = 72

    # Coloring calendar (other groups' days are greyed out)
    last_row = cal.max_row + offset
    for row in ws.iter_rows(min_row=legend_row + 1, max_row=last_row):
        for c in row:
            if not c.fill.fill_type or c.fill.fgColor.type != 'rgb':
                continue
            rule = rules.get(c.fill.fgColor.rgb[-6:])
            if rule and not rule_applies(rule, groups, modes):
                c.fill = fill(GREY)

    # Copying reference table from Abiinfo
    start = last_row + 1
    for r in range(17, 23):
        for col in range(2, 7):
            s = abi.cell(row=r, column=col)
            d = ws.cell(row=start + r - 17, column=col - 1, value=s.value)
            if s.has_style:
                d._style = copy(s._style)
    ws.merge_cells(
        start_row=start + 5, start_column=2, end_row=start + 5, end_column=5
    )

    # Cleaning up extra sheets and saving
    for name in wb.sheetnames:
        if name != ws.title:
            del wb[name]
    ws.sheet_view.zoomScale = 80

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    safe_title_filename = re.sub(r'[\\/*?:"<>|]', '_', teacher_title)
    save_path = os.path.join(
        OUTPUT_DIR, f'{safe_title_filename}_Kevadsemester2026_27.xlsx'
    )
    try:
        wb.save(save_path)
        print(f'File successfully created: {save_path}')
    except PermissionError:
        print(
            f"Unable to save '{save_path}'. Please close the file in Excel and try"
            ' again.'
        )


# ---------------------------------------------------------------------------
# Block 4: Execution Process
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    input_file = os.path.join(
        BASE_DIR,
        'input_files',
        (
            'Kevadsemester 2026_27, Programmijuhtide poolt sisu semestri'
            ' planeerimiseks_Yuliya Nevar, IT.xlsx'
        ),
    )

    print('Scanning the master table and automatically collecting teachers...')

    df_all_teachers = get_all_teachers_map(input_file)

    if df_all_teachers.empty:
        print('No teachers were found in the table.')
    else:
        unique_teachers = df_all_teachers['Teacher'].unique()
        print(
            f'Found unique teachers after merging aliases: {len(unique_teachers)}'
            f' ({list(unique_teachers)}).'
        )
        print('Starting generation of individual files for each...\n')

        for teacher_name, group_df in df_all_teachers.groupby('Teacher'):
            print(
                f'-> Creating file for teacher: "{teacher_name}" (subject rows:'
                f' {len(group_df)})'
            )
            create_teacher_excel(teacher_name, group_df, input_file)

        print(
            '\nAll tasks completed! Individual files have been saved in the'
            ' output_files folder.'
        )