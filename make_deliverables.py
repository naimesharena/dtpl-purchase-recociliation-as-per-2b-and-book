# -*- coding: utf-8 -*-
"""Write final deliverables: Excel workbook + Markdown report."""
import pandas as pd
import numpy as np
import pickle

bk   = pickle.load(open('/tmp/bk_final.pkl', 'rb'))
dn   = pickle.load(open('/tmp/dn_final.pkl', 'rb'))
g    = pickle.load(open('/tmp/g_final.pkl', 'rb'))
cdnr = pickle.load(open('/tmp/cdnr_final.pkl', 'rb'))

MONTHS = ['2025-04','2025-05','2025-06','2025-07','2025-08','2025-09',
          '2025-10','2025-11','2025-12','2026-01','2026-02','2026-03']
MLBL = ['Apr-25','May-25','Jun-25','Jul-25','Aug-25','Sep-25',
        'Oct-25','Nov-25','Dec-25','Jan-26','Feb-26','Mar-26']
Z = ['taxable','igst','cgst','sgst']

def base(s): return s.replace('DUPLICATE booking - ', '')
bk = bk.copy(); bk['b_base'] = bk['status'].map(base); bk['dup'] = bk['status'].str.startswith('DUPLICATE')
dn = dn.copy(); dn['d_base'] = dn['status'].map(base); dn['dup'] = dn['status'].str.startswith('DUPLICATE')
bk['b_tax'] = bk['taxable'].fillna(0); bk['b_ig'] = bk['IGST TAX'].fillna(0)
bk['b_cg'] = bk['CGST TAX'].fillna(0); bk['b_sg'] = bk['SGST TAX'].fillna(0)
bk['b_gst'] = bk['gst_total'].fillna(0)
bk['b_gm']  = bk['g_idx'].map(lambda i: g.at[i,'g2b_month'] if i >= 0 else '')
bk['b_gtax']= bk['g_idx'].map(lambda i: g.at[i,'taxable'] if i >= 0 else 0.0)
bk['b_gig'] = bk['g_idx'].map(lambda i: g.at[i,'igst']  if i >= 0 else 0.0)
bk['b_gcg'] = bk['g_idx'].map(lambda i: g.at[i,'cgst']  if i >= 0 else 0.0)
bk['b_gsg'] = bk['g_idx'].map(lambda i: g.at[i,'sgst']  if i >= 0 else 0.0)
bk['b_ggst']= bk['g_idx'].map(lambda i: g.at[i,'gst_total'] if i >= 0 else 0.0)
bk['b_ginv']= bk['g_idx'].map(lambda i: g.at[i,'inv'] if i >= 0 else '')
dn['d_ig'] = dn['IGST TAX'].fillna(0).abs(); dn['d_cg'] = dn['CGST TAX'].fillna(0).abs()
dn['d_sg'] = dn['SGST TAX'].fillna(0).abs(); dn['d_gst'] = dn['d_ig']+dn['d_cg']+dn['d_sg']
dn['d_tax'] = dn['Gross Total'].fillna(0) - dn['d_gst']
dn['c_gm'] = dn['c_idx'].map(lambda i: cdnr.at[i,'g2b_month'] if i >= 0 else '')
g['used']  = g.index.isin(set(bk[bk['g_idx']>=0]['g_idx'].tolist()))
def gcat(r):
    if r['sheet'] == 'B2B(Rejected)': return 'In GSTR-2B (Rejected on IMS), not in Books'
    if r['rcm']: return 'In GSTR-2B (RCM), not in Books'
    if str(r.get('itc_avail','')).strip() == 'ITC Not Available': return 'ITC Not Available'
    return 'In GSTR-2B, not in Books'
g['cat'] = g.apply(gcat, axis=1)
cdnr['used'] = cdnr.index.isin(set(dn[dn['c_idx']>=0]['c_idx'].tolist()))

# ------------------------------------------------------------------ monthly bridge
K8 = ['books','g2b_new','bk_only','dn_cn_adj','tm_out','tm_in','other','g2b']
rows = {m: {k: dict.fromkeys(Z, 0.0) for k in K8} for m in MONTHS}
def put(m, k, t, ig, cg, sg):
    if m in rows:
        rows[m][k]['taxable'] += t; rows[m][k]['igst'] += ig
        rows[m][k]['cgst'] += cg; rows[m][k]['sgst'] += sg
for r in bk.itertuples():
    put(r.book_month, 'books', r.b_tax, r.b_ig, r.b_cg, r.b_sg)
for r in dn.itertuples():
    put(r.book_month, 'books', -r.d_tax, -r.d_ig, -r.d_cg, -r.d_sg)
for r in g.itertuples():
    put(r.g2b_month, 'g2b', r.taxable, r.igst, r.cgst, r.sgst)
for r in cdnr.itertuples():
    put(r.g2b_month, 'g2b', -r.taxable, -r.igst, -r.cgst, -r.sgst)
for r in bk.itertuples():
    if r.g_idx < 0:
        put(r.book_month, 'bk_only', -r.b_tax, -r.b_ig, -r.b_cg, -r.b_sg); continue
    if r.b_gm == r.book_month:
        t = r.b_gtax - r.b_tax; ig = r.b_gig - r.b_ig; cg = r.b_gcg - r.b_cg; sg = r.b_gsg - r.b_sg
        if abs(t)+abs(ig)+abs(cg)+abs(sg) > 1e-9:
            lab = 'g2b_new' if (r.b_gst == 0 and r.b_ggst > 0) else 'other'
            put(r.book_month, lab, t, ig, cg, sg)
    else:
        put(r.book_month, 'tm_out', -r.b_tax, -r.b_ig, -r.b_cg, -r.b_sg)
        put(r.b_gm, 'tm_in', r.b_gtax, r.b_gig, r.b_gcg, r.b_gsg)
for r in g.itertuples():
    if not r.used:
        put(r.g2b_month, 'g2b_new', r.taxable, r.igst, r.cgst, r.sgst)
for r in dn.itertuples():
    if r.c_idx < 0:
        put(r.book_month, 'bk_only', r.d_tax, r.d_ig, r.d_cg, r.d_sg); continue
    ci = r.c_idx
    c = cdnr.loc[ci]
    if r.c_gm == r.book_month:
        t = r.d_tax - c['taxable']; ig = r.d_ig - c['igst']; cg = r.d_cg - c['cgst']; sg = r.d_sg - c['sgst']
        if abs(t)+abs(ig)+abs(cg)+abs(sg) > 1e-9:
            put(r.book_month, 'dn_cn_adj', t, ig, cg, sg)
    else:
        put(r.book_month, 'dn_cn_adj', r.d_tax, r.d_ig, r.d_cg, r.d_sg)
        put(r.c_gm, 'dn_cn_adj', -c['taxable'], -c['igst'], -c['cgst'], -c['sgst'])
for r in cdnr.itertuples():
    if not r.used:
        put(r.g2b_month, 'g2b_new', -r.taxable, -r.igst, -r.cgst, -r.sgst)

BLABELS = [
    ('ITC as per Books', 'books'),
    ('Add: ITC in GSTR-2B not booked in Books', 'g2b_new'),
    ('Less: ITC booked in Books not in GSTR-2B', 'bk_only'),
    ('Add/Less: Timing difference - previous month ITC received in current month', 'tm_in'),
    ('Add/Less: Timing difference - current month ITC in subsequent month', 'tm_out'),
    ('Add/Less: Debit/Credit Note adjustments', 'dn_cn_adj'),
    ('Other differences', 'other'),
    ('ITC as per GSTR-2B', 'g2b'),
]
def bridge_block(m):
    out = []
    for lab, k in BLABELS:
        v = rows[m][k]
        out.append([lab, v['taxable'], v['igst'], v['cgst'], v['sgst'], v['igst']+v['cgst']+v['sgst']])
    lhs = {c: sum(out[i][j] for i in range(7)) for j, c in enumerate(Z, 1)}
    g2b = out[7]
    fin = [lhs[c]-g2b[j+1] for j, c in enumerate(Z)]
    fin.append(lhs['igst']+lhs['cgst']+lhs['sgst'] - (g2b[2]+g2b[3]+g2b[4]))
    out.append(['Final Difference', *fin])
    return out
bridge = {m: bridge_block(m) for m in MONTHS}
TOTBLK = []
for i, (lab, k) in enumerate(BLABELS):
    t = [sum(rows[m][k][c] for m in MONTHS) for c in Z]
    TOTBLK.append([lab, *t, t[1]+t[2]+t[3]])
lhs = {c: sum(TOTBLK[i][j] for i in range(7)) for j, c in enumerate(Z, 1)}
TOTBLK.append(['Final Difference',
               lhs['taxable']-TOTBLK[7][1], lhs['igst']-TOTBLK[7][2],
               lhs['cgst']-TOTBLK[7][3], lhs['sgst']-TOTBLK[7][4],
               lhs['igst']+lhs['cgst']+lhs['sgst']-(TOTBLK[7][2]+TOTBLK[7][3]+TOTBLK[7][4])])

# ------------------------------------------------------------------ stats
matched = bk[bk['g_idx'] >= 0]
st = dict(
    books_gst=bk['b_gst'].sum(), dn_gst=dn['d_gst'].sum(),
    g_gst=g['gst_total'].sum(), cn_gst=cdnr['gst_total'].sum(),
    bk_only=bk.loc[bk['g_idx']<0, 'b_gst'].sum(),
    g_only=g.loc[~g['used'], 'gst_total'].sum(),
    dn_only=dn.loc[dn['c_idx']<0, 'd_gst'].sum(),
    cn_only=cdnr.loc[~cdnr['used'], 'gst_total'].sum(),
    notcl=matched.loc[matched['b_gst'].eq(0) & matched['b_ggst'].gt(0), 'b_ggst'].sum(),
    mm=(matched['b_gst'] - matched['b_ggst']).sum() - 0.0,
    tm_sub=matched.loc[matched['b_gm'] > matched['book_month'], 'b_gst'].sum(),
    tm_ear=matched.loc[matched['b_gm'] < matched['book_month'], 'b_ggst'].sum(),
    dn_tm=dn.loc[(dn['c_idx']>=0) & (dn['c_gm']!='') & (dn['c_gm']!=dn['book_month']), 'd_gst'].sum(),
    dn_cn_val=0.0,
)
st['dn_cn_val'] = (dn.loc[dn['c_idx']>=0,'d_gst']
                   - dn['c_idx'].map(lambda i: cdnr.at[i,'gst_total'] if i>=0 else 0.0)).sum()
st['timing'] = st['tm_sub'] + st['tm_ear'] + st['dn_tm']
st['books_net'] = st['books_gst'] - st['dn_gst']
st['g2b_net'] = st['g_gst'] - st['cn_gst']
st['diff'] = st['g2b_net'] - st['books_net']
print('stats ok:', {k: round(v,2) for k, v in st.items()})

# ------------------------------------------------------------------ sheet B: invoice-wise (Books vs GSTR-2B side by side)
dfB = bk[['branch','Date','Particulars','gst','inv','Supplier Invoice Date','book_month',
          'taxable','IGST TAX','CGST TAX','SGST TAX','gst_total','Gross Total',
          'b_gtax','b_gig','b_gcg','b_gsg','b_ggst','b_gm','b_ginv','dup','status','reason']].copy()
dfB.columns = ['Branch','Book Date','Supplier','Supplier GSTIN','Invoice No (Books)',
               'Invoice Date (per books)','Book Month',
               'Taxable (Books)','IGST (Books)','CGST (Books)','SGST (Books)','Total GST (Books)',
               'Gross (Books)',
               'Taxable (GSTR-2B)','IGST (GSTR-2B)','CGST (GSTR-2B)','SGST (GSTR-2B)',
               'Total GST (GSTR-2B)','GSTR-2B Month','Invoice No (GSTR-2B)',
               'Duplicate Booking','Status','Remarks']
for c in ['Taxable (Books)','IGST (Books)','CGST (Books)','SGST (Books)','Total GST (Books)',
          'Gross (Books)','Taxable (GSTR-2B)','IGST (GSTR-2B)','CGST (GSTR-2B)','SGST (GSTR-2B)',
          'Total GST (GSTR-2B)']:
    dfB[c] = dfB[c].round(2)
dfB['Diff Taxable (Books - 2B)'] = (dfB['Taxable (Books)'] - dfB['Taxable (GSTR-2B)']).round(2)
dfB['Diff IGST (Books - 2B)'] = (dfB['IGST (Books)'] - dfB['IGST (GSTR-2B)']).round(2)
dfB['Diff CGST (Books - 2B)'] = (dfB['CGST (Books)'] - dfB['CGST (GSTR-2B)']).round(2)
dfB['Diff SGST (Books - 2B)'] = (dfB['SGST (Books)'] - dfB['SGST (GSTR-2B)']).round(2)
dfB['Diff Total GST (Books - 2B)'] = (dfB['Total GST (Books)'] - dfB['Total GST (GSTR-2B)']).round(2)
dfB['Duplicate Booking'] = np.where(dfB['Duplicate Booking'], 'Yes', '')
dfB = dfB.sort_values(['Status','Supplier','Book Date'])
dfB.insert(0, 'S.No', np.arange(1, len(dfB)+1))
dfB = dfB[['S.No','Branch','Book Date','Supplier','Supplier GSTIN',
           'Invoice No (Books)','Invoice No (GSTR-2B)','Invoice Date (per books)',
           'Book Month','GSTR-2B Month',
           'Taxable (Books)','IGST (Books)','CGST (Books)','SGST (Books)',
           'Total GST (Books)','Gross (Books)',
           'Taxable (GSTR-2B)','IGST (GSTR-2B)','CGST (GSTR-2B)','SGST (GSTR-2B)',
           'Total GST (GSTR-2B)',
           'Diff Taxable (Books - 2B)','Diff IGST (Books - 2B)','Diff CGST (Books - 2B)',
           'Diff SGST (Books - 2B)','Diff Total GST (Books - 2B)',
           'Duplicate Booking','Status','Remarks']]

g_rem = g[~g['used']].copy()
dfB2 = g_rem[['trade_name','gst','inv','inv_date_dt','taxable','igst','cgst','sgst',
              'gst_total','gross','g2b_month','cat']].copy()
dfB2.columns = ['Supplier','Supplier GSTIN','Invoice No (GSTR-2B)','Invoice Date (2B)',
                'Taxable (GSTR-2B)','IGST (GSTR-2B)','CGST (GSTR-2B)','SGST (GSTR-2B)',
                'Total GST (GSTR-2B)','Gross (2B)','GSTR-2B Month','Status/Reason']
dfB2['Remarks'] = dfB2.pop('Status/Reason')
dfB2['Branch'] = ''; dfB2['Book Date'] = ''; dfB2['Book Month'] = ''
dfB2['Invoice No (Books)'] = ''
for c in ['Taxable (Books)','IGST (Books)','CGST (Books)','SGST (Books)','Total GST (Books)','Gross (Books)']:
    dfB2[c] = 0.0
for c in ['Taxable (GSTR-2B)','IGST (GSTR-2B)','CGST (GSTR-2B)','SGST (GSTR-2B)','Total GST (GSTR-2B)','Gross (2B)']:
    dfB2[c] = dfB2[c].round(2)
dfB2['Diff Taxable (Books - 2B)'] = -dfB2['Taxable (GSTR-2B)']
dfB2['Diff IGST (Books - 2B)'] = -dfB2['IGST (GSTR-2B)']
dfB2['Diff CGST (Books - 2B)'] = -dfB2['CGST (GSTR-2B)']
dfB2['Diff SGST (Books - 2B)'] = -dfB2['SGST (GSTR-2B)']
dfB2['Diff Total GST (Books - 2B)'] = -dfB2['Total GST (GSTR-2B)']
dfB2['Duplicate Booking'] = ''
dfB2 = dfB2.reindex(dfB.columns, axis=1)

# ------------------------------------------------------------------ sheet C: exceptions
def exc_table(df, cols, names):
    d = df[cols].copy(); d.columns = names
    d['Total GST (Books)'] = d.get('Total GST (Books)', 0)
    return d
inv_exc = []
sec = []
bk_only = bk[(bk['g_idx']<0) & (bk['b_gst']>0)].sort_values('b_gst', ascending=False)
sec.append(('1. ITC booked in Books but NOT in GSTR-2B (supplier not reported / to follow up)', bk_only,
            ['branch','Date','Particulars','gst','inv','book_month','b_tax','b_ig','b_cg','b_sg','b_gst','status','reason'],
            ['Branch','Book Date','Supplier','GSTIN','Invoice No','Book Month','Taxable','IGST','CGST','SGST','Total GST','Status','Remarks']))
bk_only0 = bk[(bk['g_idx']<0) & (bk['b_gst'].eq(0))]
sec.append(('2. Booked in Books without GST split, and not present in GSTR-2B (no ITC claimed, not reported)', bk_only0,
            ['branch','Date','Particulars','gst','inv','book_month','b_tax','status'],
            ['Branch','Book Date','Supplier','GSTIN','Invoice No','Book Month','Amount (no GST)','Status']))
g_only = g[~g['used']].sort_values('gst_total', ascending=False)
sec.append(('3. ITC in GSTR-2B but NOT booked in Books (incl. RCM / Rejected on IMS)', g_only,
            ['trade_name','gst','inv','inv_date_dt','taxable','igst','cgst','sgst','gst_total','g2b_month','cat'],
            ['Supplier','GSTIN','Invoice No','Inv Date','Taxable','IGST','CGST','SGST','Total GST','2B Month','Category']))
mm = bk[bk['b_base'].str.startswith('GST Mismatch')].sort_values('b_gst', ascending=False)
sec.append(('4. GST amount mismatch - same invoice, GST value differs (to be verified with supplier)', mm,
            ['branch','Date','Particulars','gst','inv','b_ginv','book_month','b_gm','b_tax','b_gtax','b_ig','b_gig','b_cg','b_gcg','b_sg','b_gsg','b_gst','b_ggst','status','reason'],
            ['Branch','Book Date','Supplier','GSTIN','Inv (Books)','Inv (2B)','Book Mth','2B Mth','Taxable (Bk)','Taxable (2B)','IGST (Bk)','IGST (2B)','CGST (Bk)','CGST (2B)','SGST (Bk)','SGST (2B)','GST (Bk)','GST (2B)','Status','Remarks']))
nc = bk[bk['b_base'].str.startswith('ITC in GSTR-2B but NOT claimed')].sort_values('b_ggst', ascending=False)
sec.append(('5. ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split)', nc,
            ['branch','Date','Particulars','gst','inv','book_month','b_gm','b_tax','b_gtax','b_gig','b_gcg','b_gsg','b_ggst','status','reason'],
            ['Branch','Book Date','Supplier','GSTIN','Invoice No','Book Mth','2B Mth','Booked Amt (no GST)','Taxable (2B)','IGST (2B)','CGST (2B)','SGST (2B)','ITC in 2B not claimed','Status','Remarks']))
dt = bk[bk['b_base'].str.startswith('Matched - invoice date differs')].sort_values('b_gst', ascending=False)
sec.append(('6. Matched - invoice date differs (GST & values agree)', dt,
            ['branch','Date','Particulars','gst','inv','b_ginv','Supplier Invoice Date','book_month','b_gm','b_gst','status'],
            ['Branch','Book Date','Supplier','GSTIN','Inv (Books)','Inv (2B)','Inv Date (Books)','Book Mth','2B Mth','Total GST','Status']))
du = bk[bk['dup']].sort_values('b_gst', ascending=False)
sec.append(('7. Duplicate booking in Books (same invoice booked more than once)', du,
            ['branch','Date','Particulars','gst','inv','book_month','b_gst','status','reason'],
            ['Branch','Book Date','Supplier','GSTIN','Invoice No','Book Mth','Total GST','Status','Remarks']))
un = bk[(bk['b_base'].str.startswith('In Books, not in GSTR-2B')) & (bk['b_gst'].eq(0))]
sec.append(('8. Other / to be verified (small value, blank invoice numbers etc.)', un,
            ['branch','Date','Particulars','gst','inv','book_month','b_gst','status'],
            ['Branch','Book Date','Supplier','GSTIN','Invoice No','Book Mth','Total GST','Status']))
dn_exc_rows = []
for title, d, cols, names in [
    ('Debit/Credit Note exceptions', dn[~dn['d_base'].str.startswith('Credit Note - Matched')].copy(),
     ['branch','Date','Particulars','gst','note','book_month','c_gm','d_ig','d_cg','d_sg','d_gst','status','reason'],
     ['Branch','Date','Supplier','GSTIN','Note No','Book Mth','CN 2B Mth','IGST','CGST','SGST','Total GST','Status','Remarks'])]:
    sec.append((title, d, cols, names))

# ------------------------------------------------------------------ sheet E: timing details
tm = matched[(matched['b_gm'] != matched['book_month'])].copy()
dfE = tm[['branch','Date','Particulars','gst','inv','b_ginv','Supplier Invoice Date',
          'book_month','b_gm','b_tax','b_gtax','b_ig','b_gig','b_cg','b_gcg','b_sg','b_gsg',
          'b_gst','b_ggst','status','reason']].copy()
dfE.columns = ['Branch','Book Date','Supplier','Supplier GSTIN','Invoice No (Books)',
               'Invoice No (2B)','Invoice Date','Book Month','GSTR-2B Month','Taxable (Books)',
               'Taxable (2B)','IGST (Books)','IGST (2B)','CGST (Books)','CGST (2B)','SGST (Books)',
               'SGST (2B)','Total GST (Books)','Total GST (2B)','Direction','Reason']
dfE['Direction'] = np.where(dfE['Book Month'] < dfE['GSTR-2B Month'],
                            'ITC booked in month, reflected in subsequent GSTR-2B',
                            'ITC in earlier GSTR-2B, invoice booked in later month')
dfE['Difference (GST)'] = (dfE['Total GST (Books)'] - dfE['Total GST (2B)']).round(2)
dntm = dn[(dn['c_idx']>=0) & (dn['c_gm']!='') & (dn['c_gm']!=dn['book_month'])]
if len(dntm):
    d = dntm[['branch','Date','Particulars','gst','note','book_month','c_gm','d_ig','d_cg','d_sg','d_gst','status']].copy()
    d.columns = ['Branch','Book Date','Supplier','Supplier GSTIN','Note No (Books)','Book Month',
                 'GSTR-2B Month','IGST (Books)','CGST (Books)','SGST (Books)','Total GST (Books)','Status']
    d['Invoice No (Books)'] = d['Note No (Books)']; d = d.drop(columns='Note No (Books)')
    d = d.reindex(dfE.columns, axis=1)
    d['Reason'] = 'Credit note timing difference'
    dfE = pd.concat([dfE, d], ignore_index=True)

# ------------------------------------------------------------------ sheet D: DN register
dfD = dn[['branch','Date','Particulars','gst','note','book_month','c_gm','Gross Total',
          'IGST TAX','CGST TAX','SGST TAX','d_gst','status','reason']].copy()
dfD.columns = ['Branch','Date','Supplier','Supplier GSTIN','Debit Note No','Book Month',
               'CN Month (2B)','Gross Value','IGST','CGST','SGST','Total GST (as booked)',
               'Status','Remarks']

# ------------------------------------------------------------------ bridge line-item details
BLINE = {key: lab for lab, key in BLABELS}
LINE_ORDER = {key: i for i, (lab, key) in enumerate(BLABELS)}

def bridge_detail():
    recs = []
    def add(line, m, rtype, branch, date, sup, gst, inv_b, inv_g, bm, gm, t, ig, cg, sg, note=''):
        if m not in MONTHS: return
        recs.append({'Month': m, 'Line': BLINE.get(line, 'Final Difference'), '_o': LINE_ORDER.get(line, 99),
                     'Type': rtype, 'Branch': branch, 'Date': date, 'Supplier': sup, 'GSTIN': gst,
                     'Invoice/Note No (Books)': inv_b, 'Invoice/Note No (2B)': inv_g,
                     'Book Month': bm, 'GSTR-2B Month': gm,
                     'Taxable': round(t, 2), 'IGST': round(ig, 2), 'CGST': round(cg, 2),
                     'SGST': round(sg, 2), 'Total GST': round(ig + cg + sg, 2), 'Remarks': note})
    for r in bk.itertuples():
        add('books', r.book_month, 'Invoice', r.branch, r.Date, r.Particulars, r.gst,
            r.inv, '', r.book_month, r.b_gm, r.b_tax, r.b_ig, r.b_cg, r.b_sg)
    for r in dn.itertuples():
        add('books', r.book_month, 'Debit Note', r.branch, r.Date, r.Particulars, r.gst,
            r.note, '', r.book_month, r.c_gm, -r.d_tax, -r.d_ig, -r.d_cg, -r.d_sg, 'reduces ITC')
    for r in g.itertuples():
        add('g2b', r.g2b_month, 'Invoice (2B)', '', r.inv_date_dt, r.trade_name, r.gst,
            '', r.inv, '', r.g2b_month, r.taxable, r.igst, r.cgst, r.sgst,
            r.cat if not r.used else '')
    for r in cdnr.itertuples():
        add('g2b', r.g2b_month, 'Credit Note (2B)', '', r.note_date, r.trade_name, r.gst,
            '', r.note, '', r.g2b_month, -r.taxable, -r.igst, -r.cgst, -r.sgst, 'reduces ITC')
    for r in bk.itertuples():
        if r.g_idx < 0:
            add('bk_only', r.book_month, 'Invoice', r.branch, r.Date, r.Particulars, r.gst,
                r.inv, '', r.book_month, '', -r.b_tax, -r.b_ig, -r.b_cg, -r.b_sg, r.b_base)
            continue
        if r.b_gm == r.book_month:
            t = r.b_gtax - r.b_tax; ig = r.b_gig - r.b_ig; cg = r.b_gcg - r.b_cg; sg = r.b_gsg - r.b_sg
            if abs(t) + abs(ig) + abs(cg) + abs(sg) > 1e-9:
                if r.b_gst == 0 and r.b_ggst > 0:
                    add('g2b_new', r.book_month, 'Invoice (2B)', r.branch, r.Date, r.Particulars,
                        r.gst, r.inv, r.b_ginv, r.book_month, r.b_gm, t, ig, cg, sg,
                        'ITC in 2B not claimed in books (booked w/o GST split)')
                else:
                    add('other', r.book_month, 'Invoice (2B)', r.branch, r.Date, r.Particulars,
                        r.gst, r.inv, r.b_ginv, r.book_month, r.b_gm, t, ig, cg, sg,
                        'GST value mismatch / rounding (2B - Books)')
        else:
            # both months are affected (mirrors the bridge exactly)
            if r.b_gm > r.book_month:
                add('tm_out', r.book_month, 'Invoice', r.branch, r.Date, r.Particulars, r.gst,
                    r.inv, r.b_ginv, r.book_month, r.b_gm,
                    -r.b_tax, -r.b_ig, -r.b_cg, -r.b_sg,
                    'ITC reflected in subsequent GSTR-2B month ' + r.b_gm)
            else:
                add('tm_out', r.book_month, 'Invoice', r.branch, r.Date, r.Particulars, r.gst,
                    r.inv, r.b_ginv, r.book_month, r.b_gm,
                    -r.b_tax, -r.b_ig, -r.b_cg, -r.b_sg,
                    'ITC was in earlier GSTR-2B month ' + r.b_gm)
            add('tm_in', r.b_gm, 'Invoice (2B)', r.branch, r.Date, r.Particulars, r.gst,
                r.inv, r.b_ginv, r.book_month, r.b_gm,
                r.b_gtax, r.b_gig, r.b_gcg, r.b_gsg,
                'booked in ' + ('earlier' if r.b_gm > r.book_month else 'subsequent') + ' month ' + r.book_month)
    for r in g.itertuples():
        if not r.used:
            add('g2b_new', r.g2b_month, 'Invoice (2B)', '', r.inv_date_dt, r.trade_name, r.gst,
                '', r.inv, '', r.g2b_month, r.taxable, r.igst, r.cgst, r.sgst, r.cat)
    for r in dn.itertuples():
        if r.c_idx < 0:
            add('bk_only', r.book_month, 'Debit Note', r.branch, r.Date, r.Particulars, r.gst,
                r.note, '', r.book_month, '', r.d_tax, r.d_ig, r.d_cg, r.d_sg,
                'debit note in books, no credit note in GSTR-2B')
            continue
        ci = r.c_idx; c = cdnr.loc[ci]
        if r.c_gm == r.book_month:
            t = r.d_tax - c['taxable']; ig = r.d_ig - c['igst']; cg = r.d_cg - c['cgst']; sg = r.d_sg - c['sgst']
            if abs(t) + abs(ig) + abs(cg) + abs(sg) > 1e-9:
                add('dn_cn_adj', r.book_month, 'Debit Note', r.branch, r.Date, r.Particulars, r.gst,
                    r.note, c['note'], r.book_month, r.c_gm, t, ig, cg, sg,
                    'DN/CN value difference (DN - CN)')
        else:
            add('dn_cn_adj', r.book_month, 'Debit Note', r.branch, r.Date, r.Particulars, r.gst,
                r.note, c['note'], r.book_month, r.c_gm, r.d_tax, r.d_ig, r.d_cg, r.d_sg,
                'credit note appears in GSTR-2B month ' + r.c_gm)
            dirn = 'earlier' if r.book_month < r.c_gm else 'subsequent'
            add('dn_cn_adj', r.c_gm, 'Credit Note (2B)', r.branch, r.Date, r.Particulars, r.gst,
                r.note, c['note'], r.book_month, r.c_gm,
                -c['taxable'], -c['igst'], -c['cgst'], -c['sgst'],
                'debit note booked in %s month %s (DN/CN timing difference)' % (dirn, r.book_month))
    for r in cdnr.itertuples():
        if not r.used:
            add('g2b_new', r.g2b_month, 'Credit Note (2B)', '', r.note_date, r.trade_name, r.gst,
                '', r.note, '', r.g2b_month, -r.taxable, -r.igst, -r.cgst, -r.sgst,
                'credit note in GSTR-2B, no debit note in books')
    for m in MONTHS:
        add('final', m, 'Check', '', None, '', '', '', '', '', '', 0, 0, 0, 0,
            'bridge closes (zero by construction)')
    df = pd.DataFrame(recs)
    df = df.sort_values(['Month', '_o', 'Line', 'Supplier', 'Total GST'],
                        ascending=[True, True, True, True, False])
    return df.drop(columns='_o')

# ------------------------------------------------------------------ Excel
XL = 'GST_ITC_Reconciliation_FY2025-26.xlsx'
with pd.ExcelWriter(XL, engine='openpyxl') as W:
    # A. monthly bridge
    r = 0
    sheetA_rows = []
    for m in MONTHS:
        sheetA_rows.append([f'Month: {MLBL[MONTHS.index(m)]} ({m})', '', '', '', '', ''])
        for row in bridge[m]: sheetA_rows.append(row)
        sheetA_rows.append(['', '', '', '', '', ''])
    sheetA_rows.append(['Month: TOTAL (FY 2025-26)', '', '', '', '', ''])
    sheetA_rows.extend(TOTBLK)
    pd.DataFrame(sheetA_rows, columns=['Particulars','Taxable','IGST','CGST','SGST/UTGST','Total GST']).to_excel(
        W, sheet_name='A. Monthly Bridge', index=False, startrow=1)
    # A2. bridge line-item details (every invoice behind each bridge line, month by month)
    bridge_detail().to_excel(W, sheet_name='A2. Bridge Line Details', index=False)
    # B. invoice-wise
    dfB.to_excel(W, sheet_name='B. Invoice Reconciliation', index=False)
    dfB2.to_excel(W, sheet_name='B. Invoice Reconciliation', index=False,
                  startrow=len(dfB)+3)
    # C. exceptions
    er = []
    for title, d, cols, names in sec:
        er.append(['### ' + title, '', '', '', '', '', '', '', '', '', '', ''])
        x = d[cols].copy(); x.columns = names
        for c in x.columns:
            if x[c].dtype.kind == 'f': x[c] = x[c].round(2)
        for _, row in x.iterrows(): er.append(list(row.values))
        gcol = 'Total GST' if 'Total GST' in x.columns else 'ITC in 2B not claimed'
        if gcol in x.columns:
            er.append([f'Subtotal - {title.split(".")[0].strip()}', *['']*(len(x.columns)-2),
                       x[gcol].sum().round(2) if x[gcol].dtype.kind=='f' else ''])
        er.append(['', '', '', '', '', '', '', '', '', '', '', ''])
    pd.DataFrame(er).to_excel(W, sheet_name='C. Exception Report', index=False, header=False)
    # D. DN
    dfD.to_excel(W, sheet_name='D. Debit-Credit Notes', index=False)
    # E. timing
    dfE.to_excel(W, sheet_name='E. Timing Differences', index=False)
    # F. summary
    s = pd.DataFrame([
        ['ITC as per Books (gross, invoices)', st['books_gst']],
        ['Less: Debit notes in books (supplier credit notes)', -st['dn_gst']],
        ['ITC as per Books (NET)', st['books_net']],
        ['ITC as per GSTR-2B (gross, invoices)', st['g_gst']],
        ['Less: Credit notes in GSTR-2B', -st['cn_gst']],
        ['ITC as per GSTR-2B (NET)', st['g2b_net']],
        ['NET DIFFERENCE (GSTR-2B minus Books)', st['diff']],
        ['', ''],
        ['Booked in Books, not in GSTR-2B (ITC at risk - supplier follow up)', -st['bk_only']],
        ['ITC in GSTR-2B, not booked in Books', st['g_only']],
        ['ITC in GSTR-2B not claimed in Books (booked w/o GST split)', st['notcl']],
        ['Debit notes in Books without matching credit note in 2B', st['dn_only']],
        ['Credit notes in 2B without matching debit note in Books', -st['cn_only']],
        ['Net GST difference on matched invoices (mismatch/rounding)', -st['mm'] - st['notcl']],
        ['Net DN/CN value difference on matched pairs', st['dn_cn_val']],
        ['', ''],
        ['Timing differences (no impact on FY total, month-wise only)', st['timing']],
        ['  - current month ITC appearing in subsequent GSTR-2B month', -st['tm_sub']],
        ['  - previous month ITC received in current GSTR-2B month', st['tm_ear']],
        ['  - debit/credit note timing', -st['dn_tm']],
    ], columns=['Particulars','Amount (Rs)'])
    s.to_excel(W, sheet_name='F. Summary', index=False)
    for ws in W.book.worksheets:
        for col in ws.columns:
            try:
                w = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(w+2, 10), 60)
            except Exception:
                pass

print('excel written:', XL)

# ------------------------------------------------------------------ Markdown
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
                cells.append(str(v).replace('|', '/').replace('\n', ' ')[:120])
        out.append('| ' + ' | '.join(cells) + ' |')
    return '\n'.join(out)

L = []
A = L.append
A('# GSTR-2B vs Books ITC Reconciliation - Drive Trucking Pvt Ltd - FY 2025-26')
A('')
A('**Period:** 01-Apr-2025 to 31-Mar-2026 | **Branches:** HR / PL / VL | **Sources:** GSTR-2B (B2B + B2BA + CDN), Purchase Register, Debit Note Register')
A('')
A('**Matching basis:** supplier GSTIN + invoice number (fuzzy where mistyped), invoice date, taxable value, IGST/CGST/SGST and gross amount (rupee-level rounding tolerance Rs 1). Duplicate bookings in books are flagged, not removed.')
A('')
A('**Note on debit/credit notes:** the "Debit Note" register in the books records the supplier credit notes (ITC reductions); these are matched against the credit notes (CDN) reported by the supplier in GSTR-2B. Updated DEBIT NOTE-HR/PL/VL files (with corrected Gross Total values) are used in this report.')
A('')
A('## 1. Headline position (Rs)')
A('')
b_ig = bk['b_ig'].sum()-dn['d_ig'].sum(); b_cg = bk['b_cg'].sum()-dn['d_cg'].sum()
b_sg = bk['b_sg'].sum()-dn['d_sg'].sum()
g_ig = g['igst'].sum()-cdnr['igst'].sum(); g_cg = g['cgst'].sum()-cdnr['cgst'].sum()
g_sg = g['sgst'].sum()-cdnr['sgst'].sum()
A(md_table(['Particulars','IGST','CGST','SGST/UTGST','Total GST'],
           [['ITC as per Books (net of debit notes)', b_ig, b_cg, b_sg, st['books_net']],
            ['ITC as per GSTR-2B (net of credit notes)', g_ig, g_cg, g_sg, st['g2b_net']],
            ['Difference (GSTR-2B minus Books)', g_ig-b_ig, g_cg-b_cg, g_sg-b_sg, st['diff']]],
           money_cols=(1,2,3,4)))
A('')
A('## 2. Section A - Month-wise Summary Reconciliation')
A('')
A('Format: Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST. The bridge converts ITC as per Books into ITC as per GSTR-2B month by month; "Final Difference" is zero in every month by construction (all differences are explained).')
A('')
A('**Line-item drill-down:** every rupee of each bridge line is traced to the underlying invoice/note in Excel sheet **A2. Bridge Line Details** (Month | Line | Type | Supplier | GSTIN | Invoice No (Books) | Invoice No (2B) | Book Month | 2B Month | Taxable | IGST | CGST | SGST | Total GST | Remarks). Filter by Month + Line to see exactly which invoices sit behind each bridge row - e.g. Line = "Add: ITC in GSTR-2B not booked in Books" + Month = Sep-25 lists the Parth credit notes and the 2B invoices not present in books for that month.')
A('')
for m in MONTHS + ['TOTAL']:
    A(f'### {MLBL[MONTHS.index(m)]} ' if m != 'TOTAL' else '### TOTAL (FY 2025-26)')
    A('')
    blk = TOTBLK if m == 'TOTAL' else bridge[m]
    A(md_table(['Particulars','Taxable','IGST','CGST','SGST/UTGST','Total GST'],
               [[r[0], r[1], r[2], r[3], r[4], r[5]] for r in blk],
               money_cols=(1,2,3,4,5)))
    A('')
A('## 3. Section B - Invoice-wise Reconciliation (summary)')
A('')
A('Complete invoice-wise listing (9,629 book rows + 139 unmatched GSTR-2B rows, with status, amounts, differences and remarks) is in the Excel workbook, sheet **B. Invoice Reconciliation**.')
A('')
mix = bk['status'].value_counts()
A(md_table(['Status','Rows'], [[i, int(j)] for i, j in mix.items()]))
A('')
A(f"Debit note register: 61 entries - " +
  '; '.join(f"{i}: {int(j)}" for i, j in dn['status'].value_counts().items()))
A('')
A('## 4. Section C - Exception / Action Report')
A('')
for title, d, cols, names in sec:
    A(f'### {title}  **({len(d)} rows)**')
    A('')
    x = d[cols].copy(); x.columns = names
    money = set()
    for c in x.columns:
        if x[c].dtype.kind == 'f':
            x[c] = x[c].round(2); money.add(c)
    if len(x) > 200: x = x.head(200)
    A(md_table(list(x.columns), x.values.tolist(),
               money_cols=tuple(i for i, c in enumerate(x.columns) if c in money)))
    A('')
A('## 5. Debit / Credit Note reconciliation')
A('')
A('Full listing in Excel sheet **D. Debit/Credit Notes**.')
A('')
A(md_table(['Status','Notes'], [[i, int(j)] for i, j in dn['status'].value_counts().items()]))
A('')
A(f"Debit notes as per books: IGST {dn['d_ig'].sum():,.2f}, CGST {dn['d_cg'].sum():,.2f}, SGST {dn['d_sg'].sum():,.2f}, Total {dn['d_gst'].sum():,.2f}. Credit notes as per GSTR-2B: IGST {cdnr['igst'].sum():,.2f}, CGST {cdnr['cgst'].sum():,.2f}, SGST {cdnr['sgst'].sum():,.2f}, Total {cdnr['gst_total'].sum():,.2f}.")
A('')
A('## 6. Timing Differences (month-wise movement only - no impact on FY total)')
A('')
A(f"Total timing movement: Rs {st['timing']:,.2f} = current-month ITC appearing in subsequent GSTR-2B month Rs {st['tm_sub']:,.2f} + previous-month ITC received in current GSTR-2B month Rs {st['tm_ear']:,.2f} + debit/credit-note timing Rs {st['dn_tm']:,.2f}. Full listing in Excel sheet **E. Timing Differences**.")
A('')
A('## 7. Why Books and GSTR-2B differ - final explanation')
A('')
A(f"ITC as per books (net) is Rs {st['books_net']:,.2f}; ITC as per GSTR-2B (net) is Rs {st['g2b_net']:,.2f} - GSTR-2B is higher by Rs {st['diff']:,.2f}. The difference is fully explained:")
A('')
A(md_table(['#','Component','Amount (Rs)','Nature'],
           [
            ['1','ITC in GSTR-2B not booked in Books (incl. RCM and invoices rejected on IMS)', st['g_only'], 'Supplier reported but not in books - verify and book ITC where valid'],
            ['2','ITC in GSTR-2B not claimed in Books (invoices booked without GST split)', st['notcl'], 'Books carry no ITC on these invoices - claim/regularise ITC'],
            ['3','Credit notes in GSTR-2B without matching debit note in Books', -st['cn_only'], 'ITC reduction in 2B with no book entry - investigate'],
            ['4','ITC booked in Books not in GSTR-2B (supplier not reported / RCM / to follow up)', -st['bk_only'], 'ITC at risk - obtain supplier GSTR-1 report'],
            ['5','Debit notes in Books without matching credit note in GSTR-2B', st['dn_only'], 'ITC reduction booked with no 2B credit note - verify'],
            ['6','Net GST difference on matched invoices (value mismatches + rounding)', -st['mm'] - st['notcl'], 'Gross mismatches listed in Exception Report item 4 - verify with suppliers'],
            ['7','Net debit/credit-note value difference on matched pairs', st['dn_cn_val'], 'Minor DN/CN value differences - see DN section'],
           ], money_cols=(2,)))
A('')
A(f"**Timing vs other:** Of the total difference, **Rs {st['timing']:,.2f} of invoice-level movement is timing only** (booked in one month, ITC in another - it cancels within the year and does not create a loss or excess). The balance of **Rs {st['diff']:,.2f} is net of all timing effects** and is dominated by (a) ITC in GSTR-2B not booked/claimed (+Rs {st['g_only']+st['notcl']:,.2f}) largely offset by (b) ITC booked in books not appearing in GSTR-2B (−Rs {st['bk_only']:,.2f}, mostly Daimler India invoices not yet reported by the supplier and several small suppliers) and (c) credit notes in 2B without matching debit notes (−Rs {st['cn_only']:,.2f}).")
A('')
A('**Action priorities:** (1) follow up suppliers for the book-side ITC not in GSTR-2B (largest: Daimler India ~Rs 4.8L IGST on PO-number invoices); (2) regularise the 49 invoices booked without GST split where 2B shows ITC (Rs 4.3L); (3) verify the 40 GST-value mismatches with suppliers; (4) match the 6 credit notes without book debit notes (Rs 4.1L, mostly Parth Cement CN/02-04 and Sunrise); (5) book/reverse the 13 invoices rejected on IMS as per final IMS position.')
open('GST_ITC_Reconciliation_FY2025-26.md', 'w').write('\n'.join(L))
print('markdown written')
