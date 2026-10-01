import os
import re
import sys
import unicodedata
from datetime import date, datetime, timedelta
from collections import defaultdict, Counter

import openpyxl
from openpyxl.utils import get_column_letter as L


def norm(v):
    s = unicodedata.normalize('NFC', str(v or '')).replace('\xa0', ' ')
    return ' '.join(s.split()).casefold()


def master_rows(path):
    """Plain scan: every row below a header row that contains 'Aine kood' + 'Vastutav'."""
    wb = openpyxl.load_workbook(path)
    rows = []
    for ws in wb:
        cols = None
        for row in ws.iter_rows():
            vals = [norm(c.value) for c in row]
            if any(v.startswith('aine kood') for v in vals) and any(v.startswith('vastutav') for v in vals):
                cols = {}
                for i, v in enumerate(vals):
                    for key, pre in (('code', 'aine kood'), ('name', 'aine nimetus'),
                                     ('lead', 'vastutav'), ('co', 'õppejõud kes on veel')):
                        if v.startswith(pre) and key not in cols:
                            cols[key] = i
                continue
            if not cols:
                continue
            g = lambda k: vals[cols[k]] if k in cols and cols[k] < len(vals) else ''
            if (g('code') or g('name')) and not g('name').startswith('kokku'):
                rows.append(dict(sheet=ws.title, row=row[0].row, code=g('code'), name=g('name'),
                                 people=g('lead') + ' | ' + g('co'), lead=g('lead'), co=g('co')))
    return rows


def teacher_in_text(name, text):
    """first name + surname (last word, first part of a hyphen) must both occur in the text."""
    words = [w for w in norm(name).split() if w != 'tt']
    first, last_word = words[0], words[-1]
    surnames = {last_word} | set(last_word.split('-'))  # 'lill-saar' -> also 'lill', 'saar'
    return first in text and any(s in text for s in surnames)


def read_file(path):
    ws = openpyxl.load_workbook(path).active
    name = str(ws['A1'].value).split(' — ')[0].strip()
    legend = next(r for r in range(3, 200) if norm(ws.cell(r, 1).value) == 'õppetööd ei toimu')
    subj = [(norm(ws.cell(r, 2).value), norm(ws.cell(r, 3).value), norm(ws.cell(r, 1).value))
            for r in range(3, legend)]
    return ws, name, legend, subj


def main():
    master, out_dir = sys.argv[1], sys.argv[2]
    sample = sys.argv[3] if len(sys.argv) > 3 else None
    rows = master_rows(master)
    files = sorted(f for f in os.listdir(out_dir) if f.endswith('.xlsx') and not f.startswith('_'))
    ok = True

    # 1) completeness
    print(f'[1] subject rows in the shared workbook: {len(rows)}, teacher files: {len(files)}')
    covered = defaultdict(int)
    problems = 0
    for f in files:
        ws, name, legend, subj = read_file(os.path.join(out_dir, f))
        expected = Counter()
        for r in rows:
            if teacher_in_text(name, r['people']):
                expected[(r['code'], r['name'])] += 1
                covered[(r['sheet'], r['row'])] += 1
        actual = Counter((c, n) for c, n, _ in subj)
        if expected != actual:
            problems += 1
            print('   MISMATCH', name, '\n     missing:', dict(expected - actual), '\n     extra  :',
                  dict(actual - expected))
    no_teacher = [r for r in rows if not r['lead'].strip(' ?-') and not r['co'].strip(' ?-')]
    uncovered = [r for r in rows if (r['sheet'], r['row']) not in covered and r not in no_teacher]
    print(f'    teachers with a mismatch: {problems}')
    print(f'    rows without any teacher in the table (cannot go to a file): {len(no_teacher)}')
    print(f'    rows WITH a teacher text that are in no file: {len(uncovered)}')
    for r in uncovered:
        print('      ', r['sheet'], r['row'], r['code'], r['name'][:40], '|', r['people'])
    ok &= problems == 0 and not uncovered

    # 2) duplicates
    names = Counter()
    dup_rows = 0
    for f in files:
        _, name, _, subj = read_file(os.path.join(out_dir, f))
        names[norm(name)] += 1
        c = Counter(subj)
        dup_rows += sum(v - 1 for v in c.values() if v > 1)
    dups = [n for n, v in names.items() if v > 1]
    print(f'[2] files with the same teacher name: {len(dups)}; repeated subject rows inside files: {dup_rows}')
    ok &= not dups and dup_rows == 0

    # 3) dates
    bad, checked, years = [], 0, Counter()
    for f in files:
        ws, name, legend, _ = read_file(os.path.join(out_dir, f))
        off = legend - 2
        day_rows = {}
        for r in range(legend, ws.max_row + 1):
            lab = str(ws.cell(r, 1).value or '').strip()
            if lab in ('E', 'T', 'K', 'N', 'R', 'L', 'P') and lab not in day_rows:
                day_rows[lab] = r - 1
        days = ['E', 'T', 'K', 'N', 'R', 'L', 'P']
        base = ws.cell(day_rows['E'], 3).value.date()
        if base.weekday() != 0:
            bad.append((f, 'first date is not a Monday', base))
        for lab, r in day_rows.items():
            for col in range(3, ws.max_column + 1):
                v = ws.cell(r, col).value
                if isinstance(v, datetime):
                    checked += 1
                    years[v.year] += 1
                    exp = base + timedelta(days=7 * (col - 3) + days.index(lab))
                    if v.date() != exp:
                        bad.append((f, ws.cell(r, col).coordinate, v.date(), exp))
    print(f'[3] date cells checked: {checked}, years: {dict(years)}, wrong: {len(bad)}  (first Monday = {base})')
    for b in bad[:5]:
        print('    ', b)
    ok &= not bad

    # 4) sample
    if sample:
        ex = openpyxl.load_workbook(sample).active
        # compare with the first file whose layout is comparable (same number of subject rows as the sample)
        nm = next((f for f in files if read_file(os.path.join(out_dir, f))[2] == 8), None)
        if nm:
            ws, name, legend, subj = read_file(os.path.join(out_dir, nm))
            same_offset = (legend == 8)
            print(f'[4] sample vs {nm}: legend row {legend} (sample 8) -> rows are comparable: {same_offset}')
            diffs = Counter()
            for r in range(1, 65):
                for c in range(1, 26):
                    a, b = ex.cell(r, c), ws.cell(r, c)
                    if r <= 7 or r >= 8:
                        if (a.font.name or '').strip('"') != (b.font.name or '').strip('"') and (a.value or b.value):
                            diffs['font name'] += 1
                        if a.font.sz != b.font.sz and (a.value or b.value):
                            diffs['font size'] += 1
                        if a.number_format != b.number_format and isinstance(a.value, datetime):
                            diffs['date format'] += 1
            print('    style differences on filled cells:', dict(diffs) or 'none')
            em = {str(m) for m in ex.merged_cells.ranges if m.min_col < 24}
            om = {str(m) for m in ws.merged_cells.ranges}
            print('    merged ranges only in sample :', sorted(em - om))
            print('    merged ranges only in output :', sorted(om - em))
            ew = {L(i): ex.column_dimensions[L(i)].width for i in range(1, 25)}
            ow = {L(i): ws.column_dimensions[L(i)].width for i in range(1, 25)}
            print('    column widths sample:', {k: round(v, 1) for k, v in ew.items() if k in 'ABCDEFGHIJ'})
            print('    column widths output:', {k: round(v, 1) for k, v in ow.items() if k in 'ABCDEFGHIJ'})
    print('\nRESULT:', 'ALL CHECKS PASSED' if ok else 'THERE ARE PROBLEMS (see above)')


if __name__ == '__main__':
    main()