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

# ------------------------------------------------------------------ monthly bridge
# per entry, per month:
#   books  L1 ITC as per Books (book month)            [DN contributes negative]
#   tm_in  L2 +2B amount  (2B month, booked in a different book month)
#   tm_out L3 -books amt (book month, in 2B a different month)
#   vdiff  L4 +(2B-books) (same month, value difference)
#   nitb   L5 -books amt (in Books, not in GSTR-2B)
#   nir    L6 +2B amount (in GSTR-2B, not in Books)
MONTHS = ['2025-04','2025-05','2025-06','2025-07','2025-08','2025-09',
          '2025-10','2025-11','2025-12','2026-01','2026-02','2026-03']
MLBL = ['Apr-25','May-25','Jun-25','Jul-25','Aug-25','Sep-25',
        'Oct-25','Nov-25','Dec-25','Jan-26','Feb-26','Mar-26']
Z = ['taxable','igst','cgst','sgst']
KEYS = ['books','tm_in','tm_out','vdiff','nitb','nir','g2b_raw']
agg = {m: {k: dict.fromkeys(Z, 0.0) for k in KEYS} for m in MONTHS}
det = []   # A2 detail rows

def put(m, k, t, ig, cg, sg):
    if m in agg:
        agg[m][k]['taxable'] += t; agg[m][k]['igst'] += ig
        agg[m][k]['cgst'] += cg; agg[m][k]['sgst'] += sg

def add(line, m, x, t, ig, cg, sg, remarks, doc, ddate, supplier, gstin):
    det.append(dict(Month=m, Line=line, Type=x['type'], Status=x['status'],
                    Supplier=supplier, GSTIN=gstin,
                    **{'Doc No': s_(doc), 'Doc Date': ddate or ''},
                    **{'Book Period': period_label(x['b_m']) if x['b_m'] else '',
                       'GSTR-2B Period': period_label(x['t_m']) if x['t_m'] else ''},
                    Taxable=round(t, 2), IGST=round(ig, 2), CGST=round(cg, 2),
                    SGST=round(sg, 2), **{'Total GST': round(ig + cg + sg, 2)},
                    Remarks=remarks))

def bridge_entry(kind, b_m, t_m, b_amt, t_amt, status, supplier, gstin,
                 b_doc, b_date, t_doc, t_date):
    """b_amt / t_amt: (taxable, igst, cgst, sgst), signed (DN = negative books
    side handled by caller passing negative t_amt for the 2B CDN)."""
    x = dict(type=kind, status=status, b_m=b_m, t_m=t_m)
    t, ig, cg, sg = b_amt
    tt, tig, tcg, tsg = t_amt
    if b_m is not None and t_m is not None:  # matched pair
        put(b_m, 'books', t, ig, cg, sg)
        add('ITC as per Books (As per Records)', b_m, x, t, ig, cg, sg,
            'matched - same month' if t_m == b_m else f'matched - in GSTR-2B month {t_m}',
            b_doc, b_date, supplier, gstin)
        if t_m == b_m:
            if abs(tt - t) + abs(tig - ig) + abs(tcg - cg) + abs(tsg - sg) > 1e-9:
                put(b_m, 'vdiff', tt - t, tig - ig, tcg - cg, tsg - sg)
                add('Add/Less: value difference (Books vs GSTR-2B, same month)', b_m, x,
                    tt - t, tig - ig, tcg - cg, tsg - sg,
                    'value difference (GSTR-2B minus Books)', b_doc, b_date, supplier, gstin)
        else:
            put(b_m, 'tm_out', -t, -ig, -cg, -sg)
            put(t_m, 'tm_in', tt, tig, tcg, tsg)
            add('Less: booked in Books this month, in GSTR-2B a different month', b_m, x,
                -t, -ig, -cg, -sg, f'in GSTR-2B month {t_m} (month difference)',
                b_doc, b_date, supplier, gstin)
            add('Add: in GSTR-2B this month, booked in Books a different month', t_m, x,
                tt, tig, tcg, tsg, f'booked in Books month {b_m} (month difference)',
                t_doc, t_date, supplier, gstin)
        put(t_m, 'g2b_raw', tt, tig, tcg, tsg)
    elif b_m is not None:                    # in Books, not in GSTR-2B
        put(b_m, 'books', t, ig, cg, sg)
        put(b_m, 'nitb', -t, -ig, -cg, -sg)
        add('ITC as per Books (As per Records)', b_m, x, t, ig, cg, sg,
            'not found in GSTR-2B', b_doc, b_date, supplier, gstin)
        add('Less: in Books this month, not found in GSTR-2B', b_m, x,
            -t, -ig, -cg, -sg, 'in Books, not found in GSTR-2B', b_doc, b_date, supplier, gstin)
    else:                                    # in GSTR-2B, not in Books
        put(t_m, 'g2b_raw', tt, tig, tcg, tsg)
        put(t_m, 'nir', tt, tig, tcg, tsg)
        add('Add: in GSTR-2B this month, not found in Books', t_m, x,
            tt, tig, tcg, tsg, 'in GSTR-2B, not found in Books', t_doc, t_date, supplier, gstin)

for i, b in bk.iterrows():
    gi = int(b['g_idx'])
    if gi >= 0:
        t = g.loc[gi]
        st = ('Matched' if eq(b['Gross Total'], t['gross'], 0.5) and eq(b['taxable'], t['taxable'], 0.5)
              and eq(b['gst_total'], t['gst_total'], 0.05) else 'Partly Mat')
        bridge_entry('Invoice', b['book_month'], t['g2b_month'],
                     (b['taxable'], b['IGST TAX'], b['CGST TAX'], b['SGST TAX']),
                     (t['taxable'], t['igst'], t['cgst'], t['sgst']),
                     st, s_(b['Particulars']), s_(b['gst']),
                     b['inv'], ddt(b['Date']), t['inv'], ddt(t['inv_date']))
    else:
        bridge_entry('Invoice', b['book_month'], None,
                     (b['taxable'], b['IGST TAX'], b['CGST TAX'], b['SGST TAX']), (0, 0, 0, 0),
                     'Not in 2B', s_(b['Particulars']), s_(b['gst']),
                     b['inv'], ddt(b['Date']), '', '')
for i, t in g.iterrows():
    if i in used2b: continue
    bridge_entry('Invoice', None, t['g2b_month'], (0, 0, 0, 0),
                 (t['taxable'], t['igst'], t['cgst'], t['sgst']),
                 'Not in Rec', s_(t['trade_name']), s_(t['gstin']),
                 '', '', t['inv'], ddt(t['inv_date']))
for i, b in dn.iterrows():
    ci = int(b['c_idx'])
    b_amt = ((b['Gross Total'] or 0) - (b['gst_total'] or 0),
             -(b['IGST TAX'] or 0), -(b['CGST TAX'] or 0), -(b['SGST TAX'] or 0))
    if ci >= 0:
        t = cdnr.loc[ci]
        st = ('Matched' if eq(b['Gross Total'], t['gross'], 0.5)
              and eq(b_amt[0], t['taxable'], 0.5) and eq(b['gst_total'], t['gst_total'], 0.05)
              else 'Partly Mat')
        bridge_entry('Debit Note (Books)', b['book_month'], t['g2b_month'], b_amt,
                     (-t['taxable'], -t['igst'], -t['cgst'], -t['sgst']),
                     st, s_(b['Particulars']), s_(b['gst']),
                     b['note'], ddt(b['Date']), t['note'], ddt(t['note_date']))
    else:
        bridge_entry('Debit Note (Books)', b['book_month'], None, b_amt, (0, 0, 0, 0),
                     'Not in 2B', s_(b['Particulars']), s_(b['gst']),
                     b['note'], ddt(b['Date']), '', '')
for i, t in cdnr.iterrows():
    if i in usedcdn: continue
    bridge_entry('Credit Note (2B)', None, t['g2b_month'], (0, 0, 0, 0),
                 (-t['taxable'], -t['igst'], -t['cgst'], -t['sgst']),
                 'Not in Rec', s_(t['trade_name']), s_(t['gstin']),
                 '', '', t['note'], ddt(t['note_date']))

LINES = [
    ('books',  'ITC as per Books (As per Records)'),
    ('tm_in',  'Add: in GSTR-2B this month, booked in Books a different month'),
    ('tm_out', 'Less: booked in Books this month, in GSTR-2B a different month'),
    ('vdiff',  'Add/Less: value difference (Books vs GSTR-2B, same month)'),
    ('nitb',   'Less: in Books this month, not found in GSTR-2B'),
    ('nir',    'Add: in GSTR-2B this month, not found in Books'),
]
# tie-out: books + adjustments must equal raw GSTR-2B every month
raw2b_chk = {m: dict.fromkeys(Z, 0.0) for m in MONTHS}
for x_ in g.itertuples():
    if x_.g2b_month in raw2b_chk:
        for k in Z: raw2b_chk[x_.g2b_month][k] += getattr(x_, k)
for x_ in cdnr.itertuples():
    if x_.g2b_month in raw2b_chk:
        for k in Z: raw2b_chk[x_.g2b_month][k] -= getattr(x_, k)
for m in MONTHS:
    for k in Z: agg[m]['g2b_raw'][k] = raw2b_chk[m][k]

worst = 0.0
for m in MONTHS:
    for k in Z:
        s = sum(agg[m][L][k] for L, _ in LINES)
        worst = max(worst, abs(s - raw2b_chk[m][k]))
assert worst < 0.01, f'bridge tie-out failed, worst dev {worst}'

def line_vals(m, key):
    v = agg[m][key]
    return [round(v['taxable'], 2), round(v['igst'], 2), round(v['cgst'], 2),
            round(v['sgst'], 2), round(v['igst'] + v['cgst'] + v['sgst'], 2)]

# raw Books totals (invoices less debit notes) per month
COLS_BK = {'taxable': 'taxable', 'igst': 'IGST TAX', 'cgst': 'CGST TAX', 'sgst': 'SGST TAX'}
COLS_DN = {'igst': 'IGST TAX', 'cgst': 'CGST TAX', 'sgst': 'SGST TAX'}
bk_m = bk.groupby('book_month')[list(COLS_BK.values())].sum()
dn_m = dn.groupby('book_month')[list(COLS_DN.values())].sum()
dn_tax_m = (dn.groupby('book_month')['Gross Total'].sum()
            - dn.groupby('book_month')['gst_total'].sum())

def rawbk_row(m):
    out = {}
    for k in Z:
        b = float(bk_m.loc[m, COLS_BK[k]]) if m in bk_m.index else 0.0
        if k == 'taxable':
            d = float(dn_tax_m.loc[m]) if m in dn_tax_m.index else 0.0
        else:
            d = float(dn_m.loc[m, COLS_DN[k]]) if m in dn_m.index else 0.0
        out[k] = b - d
    return out

def final_diff_row(m):
    if m is None:
        rb = {k: float(bk[COLS_BK[k]].sum())
              - (float(dn['Gross Total'].sum() - dn['gst_total'].sum()) if k == 'taxable'
                 else float(dn[COLS_DN[k]].sum())) for k in Z}
    else:
        rb = rawbk_row(m)
    t2b = {k: sum(agg[mm]['g2b_raw'][k] for mm in MONTHS) if m is None else raw2b_chk[m][k] for k in Z}
    bd = [round(t2b[k] - rb[k], 2) for k in Z]
    bd.append(round(sum(bd[1:]), 2))
    return bd

sheetA_rows = []
for m in MONTHS:
    sheetA_rows.append([f'Month: {MLBL[MONTHS.index(m)]} ({m})'] + [''] * 5)
    for key, lab in LINES:
        sheetA_rows.append([lab] + line_vals(m, key))
    sheetA_rows.append(['= ITC as per GSTR-2B (month)'] + line_vals(m, 'g2b_raw'))
    sheetA_rows.append(['Final Difference (GSTR-2B minus Books)'] + final_diff_row(m))
    sheetA_rows.append([''] * 6)
sheetA_rows.append(['Month: TOTAL (FY 2025-26)'] + [''] * 5)
for key in [k for k, _ in LINES] + ['g2b_raw']:
    t = [round(sum(agg[m][key][k] for m in MONTHS), 2) for k in Z]
    t.append(round(sum(t[1:]), 2))
    sheetA_rows.append([dict(books='ITC as per Books (As per Records)',
                             tm_in='Add: in GSTR-2B this month, booked in Books a different month',
                             tm_out='Less: booked in Books this month, in GSTR-2B a different month',
                             vdiff='Add/Less: value difference (Books vs GSTR-2B, same month)',
                             nitb='Less: in Books this month, not found in GSTR-2B',
                             nir='Add: in GSTR-2B this month, not found in Books',
                             g2b_raw='= ITC as per GSTR-2B (month)')[key]] + t)
sheetA_rows.append(['Final Difference (GSTR-2B minus Books)'] + final_diff_row(None))

dfA2 = pd.DataFrame(det)
dfA2 = dfA2.sort_values(['Month', 'Line', 'Supplier', 'Doc No']).reset_index(drop=True)

# bridge summary numbers (Total GST)
def gst_of(key):
    return round(sum(agg[m][key]['igst'] + agg[m][key]['cgst'] + agg[m][key]['sgst'] for m in MONTHS), 2)
BRIDGE_GST = {k: gst_of(k) for k in KEYS}

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
put(4, 1, 'Layout follows the sample software export: software format.xls. Sheets: invoice, note, invoice_ims, note_ims, A. Monthly Bridge, A2. Bridge Line Details.')
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
tm_gst = round(BRIDGE_GST['tm_in'] + BRIDGE_GST['tm_out'], 2)
chk = round(BRIDGE_GST['nir'] + BRIDGE_GST['nitb'] + BRIDGE_GST['vdiff'] + tm_gst, 2)
put(32, 1, 'Where the difference comes from (Rs, Total GST) - ties to NET DIFFERENCE below')
put(33, 1, 'In GSTR-2B, not found in Books (Not in Rec)'); put(33, 3, BRIDGE_GST['nir'])
put(34, 1, 'In Books, not found in GSTR-2B (Not in 2B)'); put(34, 3, BRIDGE_GST['nitb'])
put(35, 1, 'Net value difference on matched pairs (GSTR-2B minus Books)'); put(35, 3, BRIDGE_GST['vdiff'])
put(36, 1, 'Month differences (booked in a different month than GSTR-2B)'); put(36, 3, tm_gst)
put(37, 1, 'Check: sum of the four rows above = NET DIFFERENCE'); put(37, 3, chk)
put(39, 1, 'Monthly reconciliation: sheet A. Monthly Bridge (Books -> GSTR-2B month by month, ties out exactly);')
put(40, 1, 'every line is drillable in sheet A2. Bridge Line Details (filter by Month + Line).')
for c, w in {1: 62, 2: 4, 3: 44}.items():
    ws.column_dimensions[get_column_letter(c)].width = w

# ------------------------------------------------------------------ A. Monthly Bridge + A2
wsA = wb.create_sheet('A. Monthly Bridge')
wsA.append(['Particulars', 'Taxable', 'IGST', 'CGST', 'SGST/UTGST', 'Total GST'])
for row in sheetA_rows:
    wsA.append(row)
for c in range(1, 7):
    w = max(len(str(wsA.cell(row=r, column=c).value or '')) for r in range(1, wsA.max_row + 1))
    wsA.column_dimensions[get_column_letter(c)].width = min(max(w + 2, 10), 62)
wsA.freeze_panes = 'A3'

cols2 = ['Month', 'Line', 'Type', 'Status', 'Supplier', 'GSTIN', 'Doc No', 'Doc Date',
         'Book Period', 'GSTR-2B Period', 'Taxable', 'IGST', 'CGST', 'SGST', 'Total GST', 'Remarks']
ws2 = wb.create_sheet('A2. Bridge Line Details')
ws2.append(cols2)
for _, r in dfA2.iterrows():
    ws2.append([r[c] for c in cols2])
for c, name in enumerate(cols2, 1):
    w = max([len(name)] + [len(str(v)) for v in dfA2[name].astype(str).head(500)])
    ws2.column_dimensions[get_column_letter(c)].width = min(max(w + 2, 8), 55)
ws2.freeze_panes = 'A2'

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

# A2 detail rows must sum to every component line of sheet A, every month
wsA2 = wb2['A. Monthly Bridge']
rowsA = list(wsA2.iter_rows(values_only=True))
blocks = {}
cur = None
for r in rowsA:
    if r[0] and str(r[0]).startswith('Month:'):
        cur = str(r[0]).split(':')[1].strip().split(' (')[0]
    elif r[0] and cur:
        blocks.setdefault(cur, {})[str(r[0])] = r[5]
COMP = [lab for _, lab in LINES]
ok = True
for m, ml in zip(MONTHS, MLBL):
    sub = dfA2[dfA2['Month'] == m].groupby('Line')['Total GST'].sum().round(2)
    for l, v in blocks[ml].items():
        if l not in COMP:
            continue
        if abs(sub.get(l, 0.0) - v) > 0.02:
            ok = False
            print('TIE-OUT FAIL', ml, l, sub.get(l, 0.0), v)
    comp_sum = sum(blocks[ml][l] for l in COMP)
    if abs(comp_sum - blocks[ml]['= ITC as per GSTR-2B (month)']) > 0.02:
        ok = False
        print('DERIVED ROW FAIL', ml, comp_sum, blocks[ml]['= ITC as per GSTR-2B (month)'])
print('A2 tie-out vs A (all 12 months, all component lines + derived rows):', 'OK' if ok else 'FAIL')
