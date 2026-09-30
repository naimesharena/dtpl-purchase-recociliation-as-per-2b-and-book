# -*- coding: utf-8 -*-
"""Generate the reconciliation deliverables from 'software format.xls' data.

Same report design as the main project (monthly bridge with drill-down,
unmatched details, value differences) but computed entirely from the
accounting-software GSTR-2B reconciliation export.

Input : software_data.pkl (run read_software.py first)
Output: Software_Format_Reconciliation_FY2025-26.xlsx / .md
"""
import pickle, os, numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
D = pickle.load(open(os.path.join(HERE, 'software_data.pkl'), 'rb'))
R = D['rows']

XL = os.path.join(HERE, 'Software_Format_Reconciliation_FY2025-26.xlsx')
MD = os.path.join(HERE, 'Software_Format_Reconciliation_FY2025-26.md')

MONTHS = ['2025-04','2025-05','2025-06','2025-07','2025-08','2025-09',
          '2025-10','2025-11','2025-12','2026-01','2026-02','2026-03']
MLBL = ['Apr-25','May-25','Jun-25','Jul-25','Aug-25','Sep-25',
        'Oct-25','Nov-25','Dec-25','Jan-26','Feb-26','Mar-26']
Z = ['taxable','igst','cgst','sgst']

def doc_type(x):
    if x['kind'] == 'Note':
        return (x['t_notype'] or x['b_notype'] or 'Note').title() + ' Note'
    return 'Invoice'

# ----------------------------------------------------------------- monthly bridge
# per entry, per month, per bridge-line:
#   books  L1 ITC as per Books (book month)
#   tm_in  L2 +2B amount  (2B month, booked in books a different month)
#   tm_out L3 -books amt  (book month, in 2B a different month)
#   vdiff  L4 +(2B-books) (same month, value difference)
#   nitb   L5 -books amt  (in books, not in 2B)
#   nir    L6 +2B amount  (in 2B, not in books)
KEYS = ['books','tm_in','tm_out','vdiff','nitb','nir','g2b_raw']
agg = {m: {k: dict.fromkeys(Z, 0.0) for k in KEYS} for m in MONTHS}

def put(m, k, t, ig, cg, sg):
    if m in agg:
        agg[m][k]['taxable'] += t; agg[m][k]['igst'] += ig
        agg[m][k]['cgst'] += cg; agg[m][k]['sgst'] += sg

det = []   # A2 detail rows

def add(line, m, x, t, ig, cg, sg, remarks, doc, ddate, amt_side='books'):
    det.append(dict(Month=m, Line=line, Type=doc_type(x), Status=x['status'],
                    Supplier=x['party'], GSTIN=x['gstin'],
                    **{'Doc No': doc or '', 'Doc Date': ddate or ''},
                    **{'Book Period': x['b_period'] or '', 'GSTR-2B Period': x['t_period'] or ''},
                    Taxable=round(t,2), IGST=round(ig,2), CGST=round(cg,2),
                    SGST=round(sg,2), **{'Total GST': round(ig+cg+sg,2)},
                    Remarks=remarks))

for x in R:
    st, b_m, t_m = x['status'], x['b_m'], x['t_m']
    if st in ('Matched','Partly Mat'):
        if not b_m or not t_m:
            continue
        put(b_m, 'books', x['b_taxable'], x['b_igst'], x['b_cgst'], x['b_sgst'])
        if t_m == b_m:
            if x['d_any']:
                put(b_m, 'vdiff', x['d_taxable'], x['d_igst'], x['d_cgst'], x['d_sgst'])
                add('Add/Less: value difference (Books vs GSTR-2B, same month)', b_m, x,
                    x['d_taxable'], x['d_igst'], x['d_cgst'], x['d_sgst'],
                    'value difference (GSTR-2B minus Books)', x['b_doc'], x['b_date'])
            add('ITC as per Books (As per Records)', b_m, x,
                x['b_taxable'], x['b_igst'], x['b_cgst'], x['b_sgst'],
                'matched - same month in Books and GSTR-2B', x['b_doc'], x['b_date'])
        else:
            put(b_m, 'tm_out', -x['b_taxable'], -x['b_igst'], -x['b_cgst'], -x['b_sgst'])
            put(t_m, 'tm_in',  x['t_taxable'], x['t_igst'], x['t_cgst'], x['t_sgst'])
            add('ITC as per Books (As per Records)', b_m, x,
                x['b_taxable'], x['b_igst'], x['b_cgst'], x['b_sgst'],
                f'matched - in GSTR-2B month {t_m}', x['b_doc'], x['b_date'])
            add('Less: booked in Books this month, in GSTR-2B a different month', b_m, x,
                -x['b_taxable'], -x['b_igst'], -x['b_cgst'], -x['b_sgst'],
                f'in GSTR-2B month {t_m} (month difference)', x['b_doc'], x['b_date'])
            add('Add: in GSTR-2B this month, booked in Books a different month', t_m, x,
                x['t_taxable'], x['t_igst'], x['t_cgst'], x['t_sgst'],
                f'booked in Books month {b_m} (month difference)', x['t_doc'], x['t_date'])
    elif st == 'Not in 2B':
        if not b_m: continue
        put(b_m, 'books', x['b_taxable'], x['b_igst'], x['b_cgst'], x['b_sgst'])
        put(b_m, 'nitb', -x['b_taxable'], -x['b_igst'], -x['b_cgst'], -x['b_sgst'])
        add('ITC as per Books (As per Records)', b_m, x,
            x['b_taxable'], x['b_igst'], x['b_cgst'], x['b_sgst'],
            'not found in GSTR-2B', x['b_doc'], x['b_date'])
        add('Less: in Books this month, not found in GSTR-2B', b_m, x,
            -x['b_taxable'], -x['b_igst'], -x['b_cgst'], -x['b_sgst'],
            'in Books, not found in GSTR-2B', x['b_doc'], x['b_date'])
    elif st == 'Not in Rec':
        if not t_m: continue
        put(t_m, 'g2b_raw', x['t_taxable'], x['t_igst'], x['t_cgst'], x['t_sgst'])
        put(t_m, 'nir', x['t_taxable'], x['t_igst'], x['t_cgst'], x['t_sgst'])
        add('Add: in GSTR-2B this month, not found in Books', t_m, x,
            x['t_taxable'], x['t_igst'], x['t_cgst'], x['t_sgst'],
            'in GSTR-2B, not found in Books', x['t_doc'], x['t_date'])

# raw 2B / raw books per month (independent totals for tie-out)
raw2b  = {m: dict.fromkeys(Z, 0.0) for m in MONTHS}
rawbk  = {m: dict.fromkeys(Z, 0.0) for m in MONTHS}
for x in R:
    if x['t_m'] in raw2b:
        for k in Z: raw2b[x['t_m']][k] += x['t_'+k]
    if x['b_m'] in rawbk:
        for k in Z: rawbk[x['b_m']][k] += x['b_'+k]
for m in MONTHS:
    for k in Z: agg[m]['g2b_raw'][k] = raw2b[m][k]

LINES = [
    ('books',  'ITC as per Books (As per Records)'),
    ('tm_in',  'Add: in GSTR-2B this month, booked in Books a different month'),
    ('tm_out', 'Less: booked in Books this month, in GSTR-2B a different month'),
    ('vdiff',  'Add/Less: value difference (Books vs GSTR-2B, same month)'),
    ('nitb',   'Less: in Books this month, not found in GSTR-2B'),
    ('nir',    'Add: in GSTR-2B this month, not found in Books'),
]

# tie-out: books + adjustments must equal raw GSTR-2B every month
worst = 0.0
for m in MONTHS:
    for k in Z:
        s = sum(agg[m][L][k] for L, _ in LINES)
        worst = max(worst, abs(s - raw2b[m][k]))
assert worst < 0.01, f'bridge tie-out failed, worst dev {worst}'

def line_vals(m, key):
    v = agg[m][key]
    return [round(v['taxable'],2), round(v['igst'],2), round(v['cgst'],2),
            round(v['sgst'],2), round(v['igst']+v['cgst']+v['sgst'],2)]

sheetA_rows = []
for m in MONTHS:
    sheetA_rows.append([f'Month: {MLBL[MONTHS.index(m)]} ({m})'] + ['']*5)
    for key, lab in LINES:
        sheetA_rows.append([lab] + line_vals(m, key))
    g2b = line_vals(m, 'g2b_raw')
    sheetA_rows.append(['= ITC as per GSTR-2B (month)'] + g2b)
    bd = [round(raw2b[m][k]-rawbk[m][k], 2) for k in Z]
    bd.append(round(sum(bd[1:]), 2))
    sheetA_rows.append(['Final Difference (GSTR-2B minus Books)'] + bd)
    sheetA_rows.append(['']*6)
sheetA_rows.append(['Month: TOTAL (FY 2025-26)'] + ['']*5)
tot = {}
for key in [k for k, _ in LINES] + ['g2b_raw']:
    t = [round(sum(agg[m][key][k] for m in MONTHS), 2) for k in Z]
    t.append(round(sum(t[1:]), 2))
    tot[key] = t
    sheetA_rows.append([dict(books='ITC as per Books (As per Records)',
                            tm_in='Add: in GSTR-2B this month, booked in Books a different month',
                            tm_out='Less: booked in Books this month, in GSTR-2B a different month',
                            vdiff='Add/Less: value difference (Books vs GSTR-2B, same month)',
                            nitb='Less: in Books this month, not found in GSTR-2B',
                            nir='Add: in GSTR-2B this month, not found in Books',
                            g2b_raw='= ITC as per GSTR-2B (month)')[key]] + t)
bd = [round(sum(raw2b[m][k]-rawbk[m][k] for m in MONTHS), 2) for k in Z]
bd.append(round(sum(bd[1:]), 2))
sheetA_rows.append(['Final Difference (GSTR-2B minus Books)'] + bd)

dfA2 = pd.DataFrame(det)
dfA2 = dfA2.sort_values(['Month','Line','Supplier','Doc No']).reset_index(drop=True)
dfA2['#'] = range(1, len(dfA2)+1)
dfA2 = dfA2.drop(columns=['#'])

# ----------------------------------------------------------------- B. unmatched
def blank(v, has):
    return v if has else ''
brows = []
for x in sorted(R, key=lambda r: (r['status'], r['party'], r['t_doc'] or r['b_doc'] or '')):
    if x['status'] not in ('Not in Rec','Not in 2B'): continue
    brows.append([x['status'], x['party'], x['gstin'], doc_type(x),
        blank(x['b_period'], x['b_has']), blank(x['b_doc'], x['b_has']),
        blank(x['b_date'], x['b_has']), blank(round(x['b_value'],2), x['b_has']),
        blank(round(x['b_taxable'],2), x['b_has']), blank(round(x['b_igst'],2), x['b_has']),
        blank(round(x['b_cgst'],2), x['b_has']), blank(round(x['b_sgst'],2), x['b_has']),
        blank(round(x['b_tax'],2), x['b_has']),
        blank(x['t_period'], x['t_has']), blank(x['t_doc'], x['t_has']),
        blank(x['t_date'], x['t_has']), blank(x['t_notype'], x['t_has']),
        blank(round(x['t_value'],2), x['t_has']), blank(round(x['t_taxable'],2), x['t_has']),
        blank(round(x['t_igst'],2), x['t_has']), blank(round(x['t_cgst'],2), x['t_has']),
        blank(round(x['t_sgst'],2), x['t_has']), blank(round(x['t_tax'],2), x['t_has']),
        blank(x['t_r1'], x['t_has']),
        round(x['d_value'],2), round(x['d_taxable'],2),
        round(x['d_igst'],2), round(x['d_cgst'],2), round(x['d_sgst'],2)])
dfB = pd.DataFrame(brows, columns=[
    'Status','Supplier','GSTIN','Type',
    'Books Period','Books Doc No','Books Date','Books Value','Books Taxable',
    'Books IGST','Books CGST','Books SGST','Books TAX',
    'GSTR-2B Period','GSTR-2B Doc No','GSTR-2B Date','GSTR-2B Type','GSTR-2B Value',
    'GSTR-2B Taxable','GSTR-2B IGST','GSTR-2B CGST','GSTR-2B SGST','GSTR-2B TAX','R1 Date',
    'Diff Value','Diff Taxable','Diff IGST','Diff CGST','Diff SGST'])

# ----------------------------------------------------------------- C. value diffs
crows = []
for x in R:
    if x['status'] in ('Matched','Partly Mat') and x['d_any']:
        crows.append([x['status'], x['party'], x['gstin'], doc_type(x), x['b_doc'],
            x['b_period'], x['b_date'], round(x['b_value'],2), round(x['b_taxable'],2),
            round(x['b_igst'],2), round(x['b_cgst'],2), round(x['b_sgst'],2),
            x['t_period'], x['t_date'], round(x['t_value'],2), round(x['t_taxable'],2),
            round(x['t_igst'],2), round(x['t_cgst'],2), round(x['t_sgst'],2),
            round(x['d_value'],2), round(x['d_taxable'],2), round(x['d_igst'],2),
            round(x['d_cgst'],2), round(x['d_sgst'],2),
            ', '.join(x['field_diffs']) or 'Amounts'])
dfC = pd.DataFrame(crows, columns=[
    'Status','Supplier','GSTIN','Type','Doc No',
    'Book Period','Book Date','Book Value','Book Taxable','Book IGST','Book CGST','Book SGST',
    '2B Period','2B Date','2B Value','2B Taxable','2B IGST','2B CGST','2B SGST',
    'Diff Value','Diff Taxable','Diff IGST','Diff CGST','Diff SGST','Mismatch'])
dfC = dfC.sort_values(['Supplier','Doc No']).reset_index(drop=True)

# ----------------------------------------------------------------- D. field mismatches
drows = []
for x in R:
    if x['status'] == 'Partly Mat' and not x['d_any']:
        drows.append([x['status'], x['party'], x['gstin'], doc_type(x), x['b_doc'],
            x['b_period'], x['t_period'], x['b_date'], x['t_date'],
            x['b_pos'], x['t_pos'], ', '.join(x['field_diffs']) or '(amounts & dates equal)',
            round(x['b_value'],2), round(x['b_taxable'],2), round(x['b_gst'],2),
            round(x['t_value'],2), round(x['t_taxable'],2), round(x['t_gst'],2)])
dfD = pd.DataFrame(drows, columns=[
    'Status','Supplier','GSTIN','Type','Doc No','Book Period','2B Period',
    'Book Date','2B Date','Book POS','2B POS','Mismatched Fields',
    'Book Value','Book Taxable','Book GST','2B Value','2B Taxable','2B GST'])
dfD = dfD.sort_values(['Supplier','Doc No']).reset_index(drop=True)

# ----------------------------------------------------------------- E. matched clean
erows = []
for x in R:
    if x['status'] == 'Matched' and not x['d_any']:
        erows.append([x['party'], x['gstin'], x['b_doc'], x['b_period'],
            x['b_date'], round(x['b_value'],2), round(x['b_taxable'],2),
            round(x['b_igst'],2), round(x['b_cgst'],2), round(x['b_sgst'],2),
            round(x['b_gst'],2)])
dfE = pd.DataFrame(erows, columns=[
    'Supplier','GSTIN','Doc No','Period','Invoice Date','Value','Taxable',
    'IGST','CGST','SGST','Total GST'])
dfE = dfE.sort_values(['Supplier','Doc No']).reset_index(drop=True)

# ----------------------------------------------------------------- summary numbers
def sumf(rows, side, k):
    return round(sum(x[side+k] for x in rows), 2)
inv = [x for x in R if x['kind']=='Invoice']
note = [x for x in R if x['kind']=='Note']
S = {}
S['bk_inv_tax']  = sumf(inv, 'b_', 'taxable');  S['bk_inv_gst']  = sumf(inv, 'b_', 'gst')
S['t_inv_tax']   = sumf(inv, 't_', 'taxable');  S['t_inv_gst']   = sumf(inv, 't_', 'gst')
S['bk_note_tax'] = sumf(note, 'b_', 'taxable'); S['bk_note_gst'] = sumf(note, 'b_', 'gst')
S['t_note_tax']  = sumf(note, 't_', 'taxable'); S['t_note_gst']  = sumf(note, 't_', 'gst')
S['bk_tax'] = S['bk_inv_tax']+S['bk_note_tax']; S['bk_gst'] = S['bk_inv_gst']+S['bk_note_gst']
S['t_tax']  = S['t_inv_tax']+S['t_note_tax'];   S['t_gst']  = S['t_inv_gst']+S['t_note_gst']
S['diff_tax'] = round(S['t_tax']-S['bk_tax'],2); S['diff_gst'] = round(S['t_gst']-S['bk_gst'],2)

nir_rows = [x for x in R if x['status']=='Not in Rec']
nitb_rows = [x for x in R if x['status']=='Not in 2B']
vd_rows = [x for x in R if x['status'] in ('Matched','Partly Mat') and x['d_any']]
tm_rows = [x for x in R if x['status'] in ('Matched','Partly Mat') and x['b_m'] and x['t_m'] and x['b_m'] != x['t_m']]
S['nir_gst'] = round(sum(x['t_gst'] for x in nir_rows),2)
S['nitb_gst'] = round(sum(x['b_gst'] for x in nitb_rows),2)
S['vd_gst'] = round(sum(x['d_gst'] for x in vd_rows),2)
S['tm_gst'] = round(sum(abs(x['t_gst']) for x in tm_rows),2)

def top(rows, side, n=10):
    g = {}
    for x in rows:
        p = x['party']
        g[p] = g.get(p, [0, 0.0])
        g[p][0] += 1; g[p][1] += x[side+'gst']
    return sorted(g.items(), key=lambda kv: -kv[1][1])[:n]
TOP_NIR = top(nir_rows, 't_')
TOP_NITB = top(nitb_rows, 'b_')

# ----------------------------------------------------------------- executive summary sheet
ex = [
    ['GST ITC Reconciliation - Software Format (GSTR-2B export)  |  FY 2025-26', ''],
    [f'Company: {D["company"]}   ({D["gstin"]})', ''],
    ['Source: software format.xls  (GSTR-2B Reconciliation Details, Month: All - accounting software export)', ''],
    ['Note: matching/status as reported by the software; this workbook adds the monthly bridge, drill-down and tie-out.', ''],
    ['', ''],
    ['Rows', ''],
    ['Total rows (invoices + notes)', D['counts']['total']],
    ['  Invoices', D['counts']['invoices']],
    ['  Credit/debit notes', D['counts']['notes']],
    ['', ''],
    ['Status mix (as per software)', ''],
    ['Matched', D['counts']['matched']],
    ['Partly Matched', D['counts']['partly']],
    ['Not in Rec (in GSTR-2B, not in Books)', D['counts']['not_in_rec']],
    ['Not in 2B (in Books, not in GSTR-2B)', D['counts']['not_in_2b']],
    ['', ''],
    ['ITC totals (Rs)', ''],
    ['ITC as per Books - invoices (taxable | total GST)', f"{S['bk_inv_tax']:,.2f}  |  {S['bk_inv_gst']:,.2f}"],
    ['ITC as per Books - notes', f"{S['bk_note_tax']:,.2f}  |  {S['bk_note_gst']:,.2f}"],
    ['ITC as per GSTR-2B - invoices (taxable | total GST)', f"{S['t_inv_tax']:,.2f}  |  {S['t_inv_gst']:,.2f}"],
    ['ITC as per GSTR-2B - credit notes (taxable | total GST)', f"{S['t_note_tax']:,.2f}  |  {S['t_note_gst']:,.2f}"],
    ['NET difference (GSTR-2B minus Books)', f"{S['diff_tax']:,.2f}  |  {S['diff_gst']:,.2f}"],
    ['', ''],
    ['Where the difference comes from (Rs, Total GST)', ''],
    ['In GSTR-2B, not found in Books (Not in Rec)', S['nir_gst']],
    ['In Books, not found in GSTR-2B (Not in 2B)', -S['nitb_gst']],
    ['Net value difference on matched pairs (GSTR-2B minus Books)', S['vd_gst']],
    ['Month differences (booked in a different month than GSTR-2B; no FY impact)', S['tm_gst']],
    ['', ''],
    ['Top suppliers - in GSTR-2B, not in Books (by Total GST)', ''],
]
ex += [[p, f'{c} docs  |  GST {g:,.2f}'] for p, (c, g) in TOP_NIR]
ex += [['', ''], ['Top suppliers - in Books, not in GSTR-2B (by Total GST)', '']]
ex += [[p, f'{c} docs  |  GST {g:,.2f}'] for p, (c, g) in TOP_NITB]
ex += [
    ['', ''],
    ['Notes', ''],
    ['1. Bridge identity (verified): Books + month-diff + value-diff + (Not in 2B) + (Not in Rec) = GSTR-2B, for every month.', ''],
    ['2. All 58 GSTR-2B credit notes are marked "Not in Rec" - the software records contain no debit note entries (books side of the note sheet is empty). The company maintains a separate DEBIT NOTE register (see the main project reconciliation), where these book-side debit note entries exist; part of this "Not in Rec" pool is that timing/registration gap.', ''],
    ['3. CESS is zero throughout. R1 Date = date the supplier invoice was reported in GSTR-1 (reference only).', ''],
    ['4. "Partly Matched" = matched on invoice no/GSTIN but with differences (amounts and/or date/POS) - see sheets C and D.', ''],
]
dfEX = pd.DataFrame(ex, columns=['Particulars','Amount / Details'])

# ----------------------------------------------------------------- write excel
with pd.ExcelWriter(XL, engine='openpyxl') as W:
    dfEX.to_excel(W, sheet_name='0. Executive Summary', index=False)
    pd.DataFrame(sheetA_rows, columns=['Particulars','Taxable','IGST','CGST','SGST/UTGST','Total GST']).to_excel(
        W, sheet_name='A. Monthly Bridge', index=False, startrow=1)
    dfA2.to_excel(W, sheet_name='A2. Bridge Line Details', index=False)
    dfB.to_excel(W, sheet_name='B. Unmatched Details', index=False)
    dfC.to_excel(W, sheet_name='C. Value Differences', index=False)
    dfD.to_excel(W, sheet_name='D. Field Mismatches', index=False)
    dfE.to_excel(W, sheet_name='E. Matched (Clean)', index=False)
    for ws in W.book.worksheets:
        for col in ws.columns:
            try:
                w = max(len(str(c.value)) if c.value is not None else 0 for c in col[:300])
                ws.column_dimensions[col[0].column_letter].width = min(max(w+2, 10), 60)
            except Exception:
                pass

print('excel written:', XL)
print(f'A2 detail rows: {len(dfA2)}   B unmatched: {len(dfB)}   C value diffs: {len(dfC)}'
      f'   D field mismatches: {len(dfD)}   E matched clean: {len(dfE)}')

# ----------------------------------------------------------------- markdown
def md_table(headers, rowsdata, money_cols=()):
    out = ['| ' + ' | '.join(headers) + ' |',
           '|' + '|'.join(['---']*len(headers)) + '|']
    for r in rowsdata:
        cells = []
        for i, v in enumerate(r):
            if i in money_cols and isinstance(v, (int, float)):
                cells.append(f'{v:,.2f}')
            elif v is None or (isinstance(v, float) and np.isnan(v)):
                cells.append('')
            else:
                cells.append(str(v).replace('|','/').replace('\n',' ')[:120])
        out.append('| ' + ' | '.join(cells) + ' |')
    return '\n'.join(out)

md = []
md.append('# GST ITC Reconciliation - Software Format (GSTR-2B export) - FY 2025-26\n')
md.append(f'**Company:** {D["company"]}  ({D["gstin"]})  ')
md.append('**Source:** `software format.xls` - accounting-software "GSTR-2B Reconciliation Details, Month: All" export. ')
md.append('Matching/status as reported by the software; this report adds the monthly bridge, drill-down and tie-out. ')
md.append('Companion Excel: `Software_Format_Reconciliation_FY2025-26.xlsx`.\n')

md.append('## 1. Status mix (as per software)\n')
md.append(md_table(['Status','Rows'], [
    ['Matched', D['counts']['matched']],
    ['Partly Matched', D['counts']['partly']],
    ['Not in Rec (in GSTR-2B, not in Books)', D['counts']['not_in_rec']],
    ['Not in 2B (in Books, not in GSTR-2B)', D['counts']['not_in_2b']],
    ['**Total** (invoices + notes)', D['counts']['total']],
]))
md.append('\n\n## 2. ITC totals (Rs)\n')
md.append(md_table(['Particulars','Taxable','Total GST'], [
    ['ITC as per Books - invoices', S['bk_inv_tax'], S['bk_inv_gst']],
    ['ITC as per Books - notes', S['bk_note_tax'], S['bk_note_gst']],
    ['ITC as per GSTR-2B - invoices', S['t_inv_tax'], S['t_inv_gst']],
    ['ITC as per GSTR-2B - credit notes', S['t_note_tax'], S['t_note_gst']],
    ['**NET difference (GSTR-2B minus Books)**', S['diff_tax'], S['diff_gst']],
], money_cols=(1,2)))
md.append('\nWhere the difference comes from (Total GST):\n')
md.append(md_table(['Item','Total GST'], [
    ['In GSTR-2B, not found in Books (Not in Rec)', S['nir_gst']],
    ['In Books, not found in GSTR-2B (Not in 2B)', -S['nitb_gst']],
    ['Net value difference on matched pairs', S['vd_gst']],
], money_cols=(1,)))

md.append('\n\n## 3. Monthly Bridge\n')
md.append('Identity (verified for every month, max deviation < 0.01): ')
md.append('`Books + month differences + value differences - (Not in 2B) + (Not in Rec) = GSTR-2B`\n')
for i, m in enumerate(MONTHS):
    md.append(f'### {MLBL[i]} ({m})\n')
    rowsdata = []
    for key, lab in LINES:
        rowsdata.append([lab] + line_vals(m, key))
    rowsdata.append(['= ITC as per GSTR-2B (month)'] + line_vals(m, 'g2b_raw'))
    bd = [round(raw2b[m][k]-rawbk[m][k], 2) for k in Z]
    bd.append(round(sum(bd[1:]), 2))
    rowsdata.append(['Final Difference (GSTR-2B minus Books)'] + bd)
    md.append(md_table(['Particulars','Taxable','IGST','CGST','SGST','Total GST'],
                       rowsdata, money_cols=(1,2,3,4,5)))
    md.append('')

md.append('## 4. Key exceptions\n')
md.append('### Top suppliers - in GSTR-2B, not in Books (by Total GST)\n')
md.append(md_table(['Supplier','Docs','Total GST'],
                   [[p, c, g] for p,(c,g) in TOP_NIR], money_cols=(2,)))
md.append('\n### Top suppliers - in Books, not in GSTR-2B (by Total GST)\n')
md.append(md_table(['Supplier','Docs','Total GST'],
                   [[p, c, g] for p,(c,g) in TOP_NITB], money_cols=(2,)))
md.append(f'\nValue differences on matched pairs: **{len(dfC)}** rows (net GST {S["vd_gst"]:,.2f}); '
          f'field-only mismatches: **{len(dfD)}** rows. See sheets C/D of the Excel.\n')

md.append('## 5. Notes\n')
md.append('1. All 58 GSTR-2B credit notes are "Not in Rec" - the software records contain no debit note entries; the book-side debit note entries live in the separate DEBIT NOTE register (main project reconciliation), so part of this pool is that timing/registration gap.')
md.append('2. CESS is zero throughout; R1 Date is the supplier GSTR-1 reporting date (reference only).')
md.append('3. "Partly Matched" = matched on invoice no/GSTIN with differences (amounts and/or date/POS).')
md.append('4. Every bridge line can be drilled down in sheet A2 (filter by Month + Line); totals tie out exactly to sheet A.\n')

open(MD, 'w').write('\n'.join(md))
print('markdown written:', MD)

# ----------------------------------------------------------------- verification
import openpyxl
wb = openpyxl.load_workbook(XL)
rowsA = list(wb['A. Monthly Bridge'].iter_rows(values_only=True))
blocks = {}; cur = None
for r in rowsA:
    if r[0] and str(r[0]).startswith('Month:'):
        cur = str(r[0]).split(':')[1].strip().split(' (')[0]
    elif r[0] and cur:
        blocks.setdefault(cur, {})[str(r[0])] = r[5]
df = pd.read_excel(XL, sheet_name='A2. Bridge Line Details')
COMP = [lab for _, lab in LINES]   # six component lines (A2 rows exist for these)
ok = True
for m, ml in zip(MONTHS, MLBL):
    sub = df[df['Month']==m].groupby('Line')['Total GST'].sum().round(2)
    for l, v in blocks[ml].items():
        if l not in COMP:            # '= 2B' and 'Final Difference' are derived rows
            continue
        if abs(sub.get(l, 0.0) - v) > 0.02:
            ok = False; print('TIE-OUT FAIL', ml, l, sub.get(l,0.0), v)
    # derived row must equal sum of the six component rows
    comp_sum = sum(blocks[ml][l] for l in COMP)
    if abs(comp_sum - blocks[ml]['= ITC as per GSTR-2B (month)']) > 0.02:
        ok = False; print('DERIVED ROW FAIL', ml, comp_sum, blocks[ml]['= ITC as per GSTR-2B (month)'])
print('A2 tie-out vs A (all months, all component lines) + derived rows:', 'OK' if ok else 'FAIL')
