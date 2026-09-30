# -*- coding: utf-8 -*-
"""Parse 'software format.xls' (accounting-software GSTR-2B reconciliation export)
into a unified pickle for report generation.

Input : <repo root>/software format.xls   (legacy .xls, OLE2 - read with xlrd)
Output: <this folder>/software_data.pkl

Layout of the file (both sheets):
  row 0 : company header   row 1 : 'GSTR-2B Reconciliation Details Month : All'
  row 3 : 'As per Records' / 'As per GSTR-2B' group headers
  row 4 : column headers
  row 5+ : data
  last row : totals (no Status) - skipped
Columns (invoice sheet):
  c0 Sr | c1 Status | c2 Party | c3 GSTIN
  Books  : c4 Period, c5 Invoice No, c6 POS, c7 Date, c8 Value, c9 Taxable,
           c10 TAX, c11 IGST, c12 CGST, c13 SGST, c14 CESS, c15 CFS, c16 RC, c17 Remark
  GSTR-2B: c18..c32 same fields + c29 '3B Status', c30 'R1 Date'
  Diff   : c33..c44 (block = GSTR-2B minus Books, verified for every row)
Columns (note sheet): same + 'Note Type' after Date, 'Note Value' instead of
  'Invoice Value'; blocks shifted by one extra column.
"""
import xlrd, pickle, re, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC  = os.path.join(HERE, '..', 'software format.xls')

MONTHS = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
          'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}

def period_m(p):
    """'Apr,2025' -> '2025-04';  ',' / '' / None -> None."""
    if p is None: return None
    p = str(p).strip()
    m = re.match(r'^([A-Za-z]{3}),?\s*(\d{4})$', p)
    if not m: return None
    mon, yr = MONTHS.get(m.group(1).title()), int(m.group(2))
    return f'{yr:04d}-{mon:02d}' if mon else None

def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0

def txt(v):
    s = str(v).strip()
    return None if s in ('', ' ', ',') else s

rows = []
wb = xlrd.open_workbook(SRC)

# ---------------------------------------------------------------- invoice sheet
s = wb.sheet_by_name('invoice')
for r in range(5, s.nrows):
    st = str(s.cell_value(r, 1)).strip()
    if not st:                      # totals row / empty
        continue
    rows.append(dict(
        kind='Invoice', sr=int(num(s.cell_value(r, 0))), status=st,
        party=txt(s.cell_value(r, 2)) or '', gstin=txt(s.cell_value(r, 3)) or '',
        b_period=txt(s.cell_value(r, 4)),  b_doc=txt(s.cell_value(r, 5)),
        b_pos=txt(s.cell_value(r, 6)),     b_date=txt(s.cell_value(r, 7)),
        b_notype=None,
        b_value=num(s.cell_value(r, 8)),   b_taxable=num(s.cell_value(r, 9)),
        b_tax=num(s.cell_value(r, 10)), b_igst=num(s.cell_value(r, 11)),
        b_cgst=num(s.cell_value(r, 12)), b_sgst=num(s.cell_value(r, 13)),
        b_cess=num(s.cell_value(r, 14)), b_cfs=txt(s.cell_value(r, 15)),
        b_rc=txt(s.cell_value(r, 16)),
        t_period=txt(s.cell_value(r, 18)), t_doc=txt(s.cell_value(r, 19)),
        t_pos=txt(s.cell_value(r, 20)),    t_date=txt(s.cell_value(r, 21)),
        t_notype=None,
        t_value=num(s.cell_value(r, 22)),  t_taxable=num(s.cell_value(r, 23)),
        t_tax=num(s.cell_value(r, 24)), t_igst=num(s.cell_value(r, 25)),
        t_cgst=num(s.cell_value(r, 26)), t_sgst=num(s.cell_value(r, 27)),
        t_cess=num(s.cell_value(r, 28)),
        t_r1=txt(s.cell_value(r, 30)), t_rc=txt(s.cell_value(r, 31)),
        d_value=num(s.cell_value(r, 37)), d_taxable=num(s.cell_value(r, 38)),
        d_tax=num(s.cell_value(r, 39)), d_igst=num(s.cell_value(r, 40)),
        d_cgst=num(s.cell_value(r, 41)), d_sgst=num(s.cell_value(r, 42)),
        d_cess=num(s.cell_value(r, 43)),
        d_flag=txt(s.cell_value(r, 36)),
    ))

# ---------------------------------------------------------------- note sheet
s = wb.sheet_by_name('note')
for r in range(5, s.nrows):
    st = str(s.cell_value(r, 1)).strip()
    if not st:
        continue
    rows.append(dict(
        kind='Note', sr=int(num(s.cell_value(r, 0))), status=st,
        party=txt(s.cell_value(r, 2)) or '', gstin=txt(s.cell_value(r, 3)) or '',
        b_period=txt(s.cell_value(r, 4)),  b_doc=txt(s.cell_value(r, 5)),
        b_pos=txt(s.cell_value(r, 6)),     b_date=txt(s.cell_value(r, 7)),
        b_notype=txt(s.cell_value(r, 8)),
        b_value=num(s.cell_value(r, 9)),   b_taxable=num(s.cell_value(r, 10)),
        b_tax=num(s.cell_value(r, 11)), b_igst=num(s.cell_value(r, 12)),
        b_cgst=num(s.cell_value(r, 13)), b_sgst=num(s.cell_value(r, 14)),
        b_cess=num(s.cell_value(r, 15)), b_cfs=txt(s.cell_value(r, 16)),
        b_rc=txt(s.cell_value(r, 17)),
        t_period=txt(s.cell_value(r, 19)), t_doc=txt(s.cell_value(r, 20)),
        t_pos=txt(s.cell_value(r, 21)),    t_date=txt(s.cell_value(r, 22)),
        t_notype=txt(s.cell_value(r, 23)),
        t_value=num(s.cell_value(r, 24)),  t_taxable=num(s.cell_value(r, 25)),
        t_tax=num(s.cell_value(r, 26)), t_igst=num(s.cell_value(r, 27)),
        t_cgst=num(s.cell_value(r, 28)), t_sgst=num(s.cell_value(r, 29)),
        t_cess=num(s.cell_value(r, 30)),
        t_r1=txt(s.cell_value(r, 32)), t_rc=txt(s.cell_value(r, 33)),
        d_value=num(s.cell_value(r, 39)), d_taxable=num(s.cell_value(r, 40)),
        d_tax=num(s.cell_value(r, 41)), d_igst=num(s.cell_value(r, 42)),
        d_cgst=num(s.cell_value(r, 43)), d_sgst=num(s.cell_value(r, 44)),
        d_cess=num(s.cell_value(r, 45)),
        d_flag=txt(s.cell_value(r, 38)),
    ))

# ---------------------------------------------------------------- derive fields
for x in rows:
    x['b_m'] = period_m(x['b_period'])
    x['t_m'] = period_m(x['t_period'])
    x['b_has'] = x['b_doc'] is not None or x['b_value'] != 0
    x['t_has'] = x['t_doc'] is not None or x['t_value'] != 0
    x['b_gst']  = x['b_igst'] + x['b_cgst'] + x['b_sgst'] + x['b_cess']
    x['t_gst']  = x['t_igst'] + x['t_cgst'] + x['t_sgst'] + x['t_cess']
    x['d_gst']  = x['d_igst'] + x['d_cgst'] + x['d_sgst'] + x['d_cess']
    x['d_any']  = any(abs(x[k]) > 0.005 for k in
                      ('d_value','d_taxable','d_igst','d_cgst','d_sgst'))
    # fields that differ between the two sides (for matched rows)
    fl = []
    if x['status'] in ('Matched','Partly Mat'):
        if (x['b_date'] or '') != (x['t_date'] or ''): fl.append('Date')
        if (x['b_pos'] or '') != (x['t_pos'] or ''): fl.append('POS')
        if (x['b_period'] or '') != (x['t_period'] or ''): fl.append('Period')
    x['field_diffs'] = fl

# ---------------------------------------------------------------- sanity checks
n = len(rows)
nir  = sum(x['status'] == 'Not in Rec' for x in rows)
nitb = sum(x['status'] == 'Not in 2B' for x in rows)
pair = sum(x['status'] in ('Matched','Partly Mat') for x in rows)
assert nir + nitb + pair == n, (nir, nitb, pair, n)
assert all((x['b_has'] and x['t_has']) or x['status'] in ('Not in Rec','Not in 2B')
           for x in rows), 'side-presence inconsistent with status'
# difference block must equal 2B - Books
for x in rows:
    for k in ('value','taxable','igst','cgst','sgst','cess'):
        assert abs(x['d_'+k] - (x['t_'+k] - x['b_'+k])) <= 0.01, (x['sr'], k)

out = dict(
    company='DRIVE TRUCKING PRIVATE LIMITED',
    gstin='24AAJCD4457C1ZV',
    fy='2025-26',
    rows=rows,
    counts=dict(total=n, matched=sum(x['status']=='Matched' for x in rows),
                partly=sum(x['status']=='Partly Mat' for x in rows),
                not_in_rec=nir, not_in_2b=nitb,
                invoices=sum(x['kind']=='Invoice' for x in rows),
                notes=sum(x['kind']=='Note' for x in rows)),
)
pickle.dump(out, open(os.path.join(HERE, 'software_data.pkl'), 'wb'))
print(f'parsed {n} rows  (invoices {out["counts"]["invoices"]}, notes {out["counts"]["notes"]})')
print('status mix:', {k: out["counts"][k] for k in
      ('matched','partly','not_in_rec','not_in_2b')})
print('saved', os.path.join(HERE, 'software_data.pkl'))
