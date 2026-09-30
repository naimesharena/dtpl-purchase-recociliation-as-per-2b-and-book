# -*- coding: utf-8 -*-
"""Build a reconciliation in the SOFTWARE FORMAT (see 'software format.xls')
but computed from OUR uploaded files:

    gstr-2B.xls  +  PURCHASE-HR/PL/VL.xls  +  DEBIT NOTE-HR/PL/VL.xls

using the main project's matching engine (reconcile.py pickles in /tmp).

Output: software-recon/GSTR-2B_Reconciliation_Software_Format_FY2025-26.xlsx
Sheets: 'invoice', 'note', 'invoice_ims', 'note_ims' (exact software layout,
headers copied verbatim from the sample file) + '0. Executive Summary'.

Conventions (verified against the sample file):
  * rows = book documents (PURCHASE / DEBIT NOTE registers) plus
    GSTR-2B documents not found in books (status 'Not in Rec');
    book documents not found in GSTR-2B get 'Not in 2B';
    matched pairs are 'Matched' (amounts equal) or 'Partly Mat' (differ).
  * difference block (rightmost) = GSTR-2B minus Books, per row.
  * last row = totals (books side and 2B side separately).
"""
import os, sys, pickle
import pandas as pd
import xlrd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SAMPLE = os.path.join(ROOT, 'software format.xls')
OUT = os.path.join(HERE, 'GSTR-2B_Reconciliation_Software_Format_FY2025-26.xlsx')

# ------------------------------------------------------------------ inputs
def load_pkl(name):
    p = f'/tmp/{name}.pkl'
    if not os.path.exists(p):
        sys.exit(f'{p} missing - run: python reconcile.py (from repo root)')
    return pickle.load(open(p, 'rb'))

bk   = load_pkl('bk_final')
g    = load_pkl('g_final')
dn   = load_pkl('dn_final')
cdnr = load_pkl('cdnr_final')

# sample header rows (0-4) copied verbatim
swb = xlrd.open_workbook(SAMPLE)
HDR = {}
for name in ['invoice', 'note', 'invoice_ims', 'note_ims']:
    s = swb.sheet_by_name(name)
    HDR[name] = [[s.cell_value(r, c) for c in range(s.ncols)] for r in range(5)]
IMS_TOTAL = {n: [swb.sheet_by_name(n).cell_value(5, c) for c in range(swb.sheet_by_name(n).ncols)]
             for n in ('invoice_ims', 'note_ims')}
COMPANY = swb.sheet_by_name('invoice').cell_value(0, 0)

# ------------------------------------------------------------------ helpers
STATE = {'01':'Jammu & Kashmir','02':'Himachal Pradesh','03':'Punjab','04':'Chandigarh',
         '05':'Uttarakhand','06':'Haryana','07':'Delhi','08':'Rajasthan','09':'Uttar Pradesh',
         '10':'Bihar','12':'Jharkhand','13':'West Bengal','14':'Sikkim','15':'Arunachal Pradesh',
         '16':'Nagaland','17':'Manipur','18':'Mizoram','19':'Tripura','20':'Meghalaya',
         '21':'Assam','22':'Odisha','23':'Chandigarh','24':'Gujarat','27':'Maharashtra',
         '29':'Madhya Pradesh','30':'Chhattisgarh','32':'Karnataka','33':'Kerala',
         '34':'Tamil Nadu','36':'Andhra Pradesh','37':'Telangana','38':'Puducherry',
         '39':'Lakshadweep','40':'Andhra Pradesh'}

def state_of(gstin):
    gstin = s_(gstin)
    return STATE.get(gstin[:2], '')

def pos2name(p):
    p = s_(p)
    return p.split('-')[-1].title() if p else ''

MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

def period_label(m):
    if not m: return ','
    y, mo = str(m).split('-')
    return f'{MON[int(mo)-1]},{y}'

def ddt(v):
    """Timestamp or 'dd-mm-yyyy' -> 'dd-Mon-yy'"""
    if v is None or (isinstance(v, float) and pd.isna(v)): return ''
    if isinstance(v, str):
        v = v.strip()
        if not v: return ''
        try: v = pd.to_datetime(v, dayfirst=True)
        except Exception: return v
    if isinstance(v, pd.Timestamp):
        return v.strftime('%d-%b-%y')
    return str(v)

def s_(v):
    """string-or-empty, NaN safe"""
    try:
        if v is None or pd.isna(v): return ''
    except Exception: pass
    s = str(v).strip()
    return s if s and s.lower() != 'nan' else ''

def r2(v):
    try:
        if pd.isna(v): return 0.0
    except Exception: pass
    return round(float(v), 2)

def d2(v):
    """2-decimal or '' """
    try:
        if v is None or pd.isna(v): return ''
    except Exception: pass
    return round(float(v), 2)

def eq(a, b, tol):
    return abs(float(a or 0) - float(b or 0)) <= tol

# ------------------------------------------------------------------ invoice rows
used2b = set(bk.loc[bk['g_idx'] >= 0, 'g_idx'].tolist())

inv_rows = []   # (party_key, row_list)
for i, b in bk.iterrows():
    gi = int(b['g_idx'])
    t = g.loc[gi] if gi >= 0 else None
    if gi >= 0:
        if eq(b['Gross Total'], t['gross'], 0.5) and eq(b['taxable'], t['taxable'], 0.5) \
           and eq(b['gst_total'], t['gst_total'], 0.05):
            st = 'Matched'
        else:
            st = 'Partly Mat'
        b_remark = ''
        t_remark = []
        if t['sheet'] == 'B2B(Rejected)': t_remark.append('Rejected on IMS')
        if str(t['itc_avail']).strip() != 'Yes': t_remark.append('ITC Not Available')
        t_remark = ' | '.join(t_remark)
        row = [1, st, s_(b['Particulars']), s_(b['gst']),
               period_label(b['book_month']), s_(b['inv']), state_of(b['gst']), ddt(b['Date']),
               d2(b['Gross Total']), d2(b['taxable']), d2(b['gst_total']),
               d2(b['IGST TAX']), d2(b['CGST TAX']), d2(b['SGST TAX']), d2(0.0), 'No', 'No', b_remark,
               period_label(t['g2b_month']), s_(t['inv']), pos2name(t['pos']), ddt(t['inv_date']),
               d2(t['gross']), d2(t['taxable']), d2(t['gst_total']),
               d2(t['igst']), d2(t['cgst']), d2(t['sgst']), d2(t['cess']), '',
               s_(t['filing_date']), 'Yes' if t['rcm'] else 'No', t_remark,
               '', '', '',
               'Mismatch' if (ddt(b['Date']) != ddt(t['inv_date'])) else '',
               r2((t['gross'] or 0) - (b['Gross Total'] or 0)),
               r2((t['taxable'] or 0) - (b['taxable'] or 0)),
               r2((t['gst_total'] or 0) - (b['gst_total'] or 0)),
               r2((t['igst'] or 0) - (b['IGST TAX'] or 0)),
               r2((t['cgst'] or 0) - (b['CGST TAX'] or 0)),
               r2((t['sgst'] or 0) - (b['SGST TAX'] or 0)),
               r2((t['cess'] or 0) - 0.0), '']
    else:
        row = [1, 'Not in 2B', s_(b['Particulars']), s_(b['gst']),
               period_label(b['book_month']), s_(b['inv']), state_of(b['gst']), ddt(b['Date']),
               d2(b['Gross Total']), d2(b['taxable']), d2(b['gst_total']),
               d2(b['IGST TAX']), d2(b['CGST TAX']), d2(b['SGST TAX']), d2(0.0), 'No', 'No', '',
               ',', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
               '', '', '', '',
               r2(-(b['Gross Total'] or 0)), r2(-(b['taxable'] or 0)),
               r2(-(b['gst_total'] or 0)), r2(-(b['IGST TAX'] or 0)),
               r2(-(b['CGST TAX'] or 0)), r2(-(b['SGST TAX'] or 0)), r2(0.0), '']
    inv_rows.append(((str(b['Particulars']).strip().lower(), s_(b['inv'])), row))

for i, t in g.iterrows():
    if i in used2b: continue
    t_remark = []
    if t['sheet'] == 'B2B(Rejected)': t_remark.append('Rejected on IMS')
    if str(t['itc_avail']).strip() != 'Yes': t_remark.append('ITC Not Available')
    t_remark = ' | '.join(t_remark)
    row = [1, 'Not in Rec', s_(t['trade_name']), s_(t['gstin']),
           ',', ' ', '', '', 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 'No', 'No', '',
           period_label(t['g2b_month']), s_(t['inv']), pos2name(t['pos']), ddt(t['inv_date']),
           d2(t['gross']), d2(t['taxable']), d2(t['gst_total']),
           d2(t['igst']), d2(t['cgst']), d2(t['sgst']), d2(t['cess']), '',
           s_(t['filing_date']), 'Yes' if t['rcm'] else 'No', t_remark,
           '', '', '', '',
           r2(t['gross'] or 0), r2(t['taxable'] or 0), r2(t['gst_total'] or 0),
           r2(t['igst'] or 0), r2(t['cgst'] or 0), r2(t['sgst'] or 0), r2(t['cess'] or 0), '']
    inv_rows.append(((str(t['trade_name']).strip().lower(), s_(t['inv'])), row))

inv_rows.sort(key=lambda x: x[0])
for n, (k, row) in enumerate(inv_rows, 1):
    row[0] = n

# ------------------------------------------------------------------ note rows
usedcdn = set(dn.loc[dn['c_idx'] >= 0, 'c_idx'].tolist())
nt_rows = []
for i, b in dn.iterrows():
    ci = int(b['c_idx'])
    t = cdnr.loc[ci] if ci >= 0 else None
    b_tax = (b['Gross Total'] or 0) - (b['gst_total'] or 0)
    if ci >= 0:
        if eq(b['Gross Total'], t['gross'], 0.5) and eq(b_tax, t['taxable'], 0.5) \
           and eq(b['gst_total'], t['gst_total'], 0.05):
            st = 'Matched'
        else:
            st = 'Partly Mat'
        t_remark = []
        if t['sheet'] == 'B2B(Rejected)': t_remark.append('Rejected on IMS')
        if str(t['itc_avail']).strip() != 'Yes': t_remark.append('ITC Not Available')
        t_remark = ' | '.join(t_remark)
        row = [1, st, s_(b['Particulars']), s_(b['gst']),
               period_label(b['book_month']), s_(b['note']), state_of(b['gst']), ddt(b['Date']),
               'Debit', d2(b['Gross Total']), d2(b_tax), d2(b['gst_total']),
               d2(b['IGST TAX']), d2(b['CGST TAX']), d2(b['SGST TAX']), d2(0.0), 'No', 'No', '',
               period_label(t['g2b_month']), s_(t['note']), pos2name(t['pos']), ddt(t['note_date']),
               'Credit', d2(t['gross']), d2(t['taxable']), d2(t['gst_total']),
               d2(t['igst']), d2(t['cgst']), d2(t['sgst']), d2(t['cess']), '',
               s_(t['filing_date']), 'Yes' if t['rcm'] else 'No', t_remark,
               '', '', '',
               'Mismatch' if (ddt(b['Date']) != ddt(t['note_date'])) else '',
               r2((t['gross'] or 0) - (b['Gross Total'] or 0)),
               r2((t['taxable'] or 0) - b_tax),
               r2((t['gst_total'] or 0) - (b['gst_total'] or 0)),
               r2((t['igst'] or 0) - (b['IGST TAX'] or 0)),
               r2((t['cgst'] or 0) - (b['CGST TAX'] or 0)),
               r2((t['sgst'] or 0) - (b['SGST TAX'] or 0)),
               r2((t['cess'] or 0) - 0.0), '']
    else:
        row = [1, 'Not in 2B', s_(b['Particulars']), s_(b['gst']),
               period_label(b['book_month']), s_(b['note']), state_of(b['gst']), ddt(b['Date']),
               'Debit', d2(b['Gross Total']), d2(b_tax), d2(b['gst_total']),
               d2(b['IGST TAX']), d2(b['CGST TAX']), d2(b['SGST TAX']), d2(0.0), 'No', 'No', '',
               ',', ' ', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
               '', '', '', '',
               r2(-(b['Gross Total'] or 0)), r2(-b_tax), r2(-(b['gst_total'] or 0)),
               r2(-(b['IGST TAX'] or 0)), r2(-(b['CGST TAX'] or 0)), r2(-(b['SGST TAX'] or 0)), r2(0.0), '']
    nt_rows.append(((str(b['Particulars']).strip().lower(), s_(b['note'])), row))

for i, t in cdnr.iterrows():
    if i in usedcdn: continue
    t_remark = []
    if t['sheet'] == 'B2B(Rejected)': t_remark.append('Rejected on IMS')
    if str(t['itc_avail']).strip() != 'Yes': t_remark.append('ITC Not Available')
    t_remark = ' | '.join(t_remark)
    row = [1, 'Not in Rec', s_(t['trade_name']), s_(t['gstin']),
           ',', ' ', '', '', '', '', '', 0.0, '', '', '', '', 'No', 'No', '',
           period_label(t['g2b_month']), s_(t['note']), pos2name(t['pos']), ddt(t['note_date']),
           'Credit', d2(t['gross']), d2(t['taxable']), d2(t['gst_total']),
           d2(t['igst']), d2(t['cgst']), d2(t['sgst']), d2(t['cess']), '',
           s_(t['filing_date']), 'Yes' if t['rcm'] else 'No', t_remark,
           '', '', '', '',
           r2(t['gross'] or 0), r2(t['taxable'] or 0), r2(t['gst_total'] or 0),
           r2(t['igst'] or 0), r2(t['cgst'] or 0), r2(t['sgst'] or 0), r2(t['cess'] or 0), '']
    nt_rows.append(((str(t['trade_name']).strip().lower(), s_(t['note'])), row))

nt_rows.sort(key=lambda x: x[0])
for n, (k, row) in enumerate(nt_rows, 1):
    row[0] = n

# ------------------------------------------------------------------ write workbook
wb = Workbook()
wb.remove(wb.active)

def write_sheet(name, data_rows, totals_cols_b, totals_cols_t, hdr_idx, tot_idx):
    ws = wb.create_sheet(name)
    for r in HDR[name]:
        ws.append(r)
    for row in data_rows:
        ws.append(row)
    # totals row
    tot = [''] * len(HDR[name][0])
    for c in totals_cols_b: tot[c] = r2(sum(x[c] for x in data_rows if isinstance(x[c], (int, float))))
    for c in totals_cols_t: tot[c] = r2(sum(x[c] for x in data_rows if isinstance(x[c], (int, float))))
    ws.append(tot)
    # column widths
    for c in range(len(HDR[name][0])):
        w = max([len(str(HDR[name][4][c] or ''))] +
                [len(str(x[c])) for x in data_rows[:400] if x[c] is not None])
        ws.column_dimensions[get_column_letter(c + 1)].width = min(max(w + 2, 8), 50)
    ws.freeze_panes = 'A6'

# invoice: books amt cols 8-14, 2B amt cols 22-28
write_sheet('invoice', [r for _, r in sorted(inv_rows, key=lambda x: x[1][0])],
            range(8, 15), range(22, 29), None, None)
# note: books amt cols 9-15, 2B amt cols 24-30
write_sheet('note', [r for _, r in sorted(nt_rows, key=lambda x: x[1][0])],
            range(9, 16), range(24, 31), None, None)
for n in ('invoice_ims', 'note_ims'):
    ws = wb.create_sheet(n)
    for r in HDR[n]:
        ws.append(r)
    ws.append(IMS_TOTAL[n])
    for c in range(len(HDR[n][0])):
        ws.column_dimensions[get_column_letter(c + 1)].width = min(max(len(str(HDR[n][4][c] or '')) + 2, 8), 50)

# ------------------------------------------------------------------ executive summary
from collections import Counter
st_inv = Counter(r[1] for _, r in inv_rows)
st_nt  = Counter(r[1] for _, r in nt_rows)
bk_gst  = round(float(bk['gst_total'].sum()), 2)
dn_gst  = round(float(dn['gst_total'].sum()), 2)
g_gst   = round(float(g['gst_total'].sum()), 2)
cdn_gst = round(float(cdnr['gst_total'].sum()), 2)
bk_tax  = round(float(bk['taxable'].sum()), 2)
g_tax   = round(float(g['taxable'].sum()), 2)
bk_net  = round(bk_gst - dn_gst, 2)
g_net   = round(g_gst - cdn_gst, 2)
diff    = round(g_net - bk_net, 2)

ws = wb.create_sheet('0. Executive Summary')
def put(r, c, v): ws.cell(row=r, column=c, value=v)
put(1, 1, 'GSTR-2B Reconciliation (Software Format) - FY 2025-26')
put(2, 1, str(COMPANY))
put(3, 1, 'Built from: gstr-2B.xls + PURCHASE-HR/PL/VL.xls + DEBIT NOTE-HR/PL/VL.xls (main project matching engine, reconcile.py)')
put(4, 1, 'Layout follows the sample software export: software format.xls. Sheets: invoice, note, invoice_ims, note_ims.')
put(5, 1, 'Status: Matched = amounts equal | Partly Mat = matched with differences | Not in Rec = in GSTR-2B, not in Books | Not in 2B = in Books, not in GSTR-2B')
put(6, 1, 'Difference block (rightmost) = GSTR-2B minus Books, per row. Last row of each sheet = totals.')
put(8, 1, 'Documents')
put(9, 1, 'Invoices (Books 9,629 + GSTR-2B not in Books 138)'); put(9, 3, len(inv_rows))
put(10, 1, 'Debit/Credit notes (Books 61 + GSTR-2B not in Books 6)'); put(10, 3, len(nt_rows))
put(12, 1, 'Status mix')
for j, k in enumerate(['Matched', 'Partly Mat', 'Not in Rec', 'Not in 2B']):
    put(13 + j, 1, f'Invoices - {k}'); put(13 + j, 3, st_inv.get(k, 0))
for j, k in enumerate(['Matched', 'Partly Mat', 'Not in Rec', 'Not in 2B']):
    put(18 + j, 1, f'Notes - {k}'); put(18 + j, 3, st_nt.get(k, 0))
put(23, 1, 'ITC totals (Rs)')
rows_s = [
    ('ITC as per Books - invoices (taxable | total GST)', f'{bk_tax:,.2f}  |  {bk_gst:,.2f}'),
    ('Less: Debit notes in Books (supplier credit notes)', f'  |  -{dn_gst:,.2f}'),
    ('ITC as per Books (NET)', f'  |  {bk_net:,.2f}'),
    ('ITC as per GSTR-2B - invoices (taxable | total GST)', f'{g_tax:,.2f}  |  {g_gst:,.2f}'),
    ('Less: Credit notes in GSTR-2B', f'  |  -{cdn_gst:,.2f}'),
    ('ITC as per GSTR-2B (NET)', f'  |  {g_net:,.2f}'),
    ('NET DIFFERENCE (GSTR-2B minus Books)', f'  |  {diff:,.2f}'),
]
for j, (a, b) in enumerate(rows_s):
    put(24 + j, 1, a); put(24 + j, 3, b)
for c, w in {1: 62, 2: 4, 3: 44}.items():
    ws.column_dimensions[get_column_letter(c)].width = w

wb.save(OUT)
print('written:', OUT)
print(f'invoice rows: {len(inv_rows)}  note rows: {len(nt_rows)}')
print('invoice status:', dict(st_inv))
print('note status:   ', dict(st_nt))
print(f'Books invoices GST {bk_gst:,.2f} | DN GST {dn_gst:,.2f} | Books net {bk_net:,.2f}')
print(f'2B invoices GST {g_gst:,.2f} | CDN GST {cdn_gst:,.2f} | 2B net {g_net:,.2f}')
print(f'NET DIFF (2B - Books): {diff:,.2f}')

# ------------------------------------------------------------------ verify vs main project
import openpyxl
wb2 = openpyxl.load_workbook(OUT)
ws2 = wb2['invoice']
last = list(ws2.iter_rows(values_only=True))[-1]
tot_bk_gst = round((last[11] or 0) + (last[12] or 0) + (last[13] or 0), 2)
tot_2b_gst = round((last[25] or 0) + (last[26] or 0) + (last[27] or 0), 2)
assert abs(tot_bk_gst - bk_gst) < 0.05, (tot_bk_gst, bk_gst)
assert abs(tot_2b_gst - g_gst) < 0.05, (tot_2b_gst, g_gst)
ws3 = wb2['note']
lastn = list(ws3.iter_rows(values_only=True))[-1]
tot_bk_n = round((lastn[12] or 0) + (lastn[13] or 0) + (lastn[14] or 0), 2)
tot_2b_n = round((lastn[27] or 0) + (lastn[28] or 0) + (lastn[29] or 0), 2)
assert abs(tot_bk_n - dn_gst) < 0.05, (tot_bk_n, dn_gst)
assert abs(tot_2b_n - cdn_gst) < 0.05, (tot_2b_n, cdn_gst)
print('VERIFY OK: totals rows equal register/2B sums (invoice GST books 2B; note GST books 2B)')
print('  invoice totals:', tot_bk_gst, tot_2b_gst)
print('  note totals:   ', tot_bk_n, tot_2b_n)
