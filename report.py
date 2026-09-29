# -*- coding: utf-8 -*-
"""Build the final GST ITC reconciliation deliverables (Excel + Markdown) from
the pipeline pickles produced by reconcile.py."""
import pandas as pd
import numpy as np
import pickle

bk  = pickle.load(open('/tmp/bk_final.pkl', 'rb'))
dn  = pickle.load(open('/tmp/dn_final.pkl', 'rb'))
g   = pickle.load(open('/tmp/g_final.pkl', 'rb'))
cdnr = pickle.load(open('/tmp/cdnr_final.pkl', 'rb'))

MONTHS = ['2025-04','2025-05','2025-06','2025-07','2025-08','2025-09',
          '2025-10','2025-11','2025-12','2026-01','2026-02','2026-03']
MLBL   = ['Apr-25','May-25','Jun-25','Jul-25','Aug-25','Sep-25',
          'Oct-25','Nov-25','Dec-25','Jan-26','Feb-26','Mar-26']
Z = dict(taxable=0.0, igst=0.0, cgst=0.0, sgst=0.0)

def vec(tax, ig, cg, sg):
    return dict(taxable=float(tax or 0), igst=float(ig or 0),
                cgst=float(cg or 0), sgst=float(sg or 0))

def addv(a, b, s=1.0):
    for k in a: a[k] += s * b[k]
    return a

def tot(v): return v['igst'] + v['cgst'] + v['sgst']

def base_status(s):
    return s.replace('DUPLICATE booking - ', '')

# ------------------------------------------------------------------ data prep
bk = bk.copy()
bk['b_tax'] = bk['taxable'].fillna(0)
bk['b_ig']  = bk['IGST TAX'].fillna(0)
bk['b_cg']  = bk['CGST TAX'].fillna(0)
bk['b_sg']  = bk['SGST TAX'].fillna(0)
bk['b_gst'] = bk['gst_total'].fillna(0)
bk['b_gm']  = bk['g_idx'].map(lambda i: g.at[i, 'g2b_month'] if i >= 0 else '')
bk['b_gtax'] = bk['g_idx'].map(lambda i: g.at[i, 'taxable'] if i >= 0 else 0.0)
bk['b_gig']  = bk['g_idx'].map(lambda i: g.at[i, 'igst']  if i >= 0 else 0.0)
bk['b_gcg']  = bk['g_idx'].map(lambda i: g.at[i, 'cgst']  if i >= 0 else 0.0)
bk['b_gsg']  = bk['g_idx'].map(lambda i: g.at[i, 'sgst']  if i >= 0 else 0.0)
bk['b_ggst'] = bk['g_idx'].map(lambda i: g.at[i, 'gst_total'] if i >= 0 else 0.0)
bk['b_ginv'] = bk['g_idx'].map(lambda i: g.at[i, 'inv'] if i >= 0 else '')
bk['b_gsname'] = bk['g_idx'].map(lambda i: g.at[i, 'trade_name'] if i >= 0 else '')
bk['b_sdate'] = bk['Supplier Invoice Date']
bk['b_base'] = bk['status'].map(base_status)
bk['dup'] = bk['status'].str.startswith('DUPLICATE')

dn = dn.copy()
dn['d_ig']  = dn['IGST TAX'].fillna(0).abs()
dn['d_cg']  = dn['CGST TAX'].fillna(0).abs()
dn['d_sg']  = dn['SGST TAX'].fillna(0).abs()
dn['d_gst'] = dn['d_ig'] + dn['d_cg'] + dn['d_sg']
dn['d_tax'] = (dn['Gross Total'].fillna(0) - dn['d_gst'])
dn['c_gm']  = dn['c_idx'].map(lambda i: cdnr.at[i, 'g2b_month'] if i >= 0 else '')
dn['d_base'] = dn['status'].map(base_status)
dn['dup'] = dn['status'].str.startswith('DUPLICATE')

g = g.copy()
g['used'] = g.index.isin(set(bk[bk['g_idx'] >= 0]['g_idx'].tolist()))
cdnr = cdnr.copy()
cdnr['used'] = cdnr.index.isin(set(dn[dn['c_idx'] >= 0]['c_idx'].tolist()))

# double-assignment safety check
vc = bk[bk['g_idx'] >= 0]['g_idx'].value_counts()
assert (vc <= 1).all(), 'double assignment present: %s' % vc[vc > 1].to_dict()

# ------------------------------------------------------------------ monthly bridge
rows = {m: {k: dict(Z) for k in
             ['books','g2b_new','bk_only','dn_cn_adj','tm_out','tm_in','other','g2b']}
        for m in MONTHS}

def put(m, k, v): addv(rows[m][k], v)

# R1 ITC as per Books (invoices + DN)
for r in bk.itertuples():
    if r.book_month in rows:
        put(r.book_month, 'books', vec(r.b_tax, r.b_ig, r.b_cg, r.b_sg))
for r in dn.itertuples():
    if r.book_month in rows:
        put(r.book_month, 'books', vec(-r.d_tax, -r.d_ig, -r.d_cg, -r.d_sg))

# R8 ITC as per GSTR-2B (invoices + CN)
for r in g.itertuples():
    if r.g2b_month in rows:
        put(r.g2b_month, 'g2b', vec(r.taxable, r.igst, r.cgst, r.sgst))
for r in cdnr.itertuples():
    if r.g2b_month in rows:
        put(r.g2b_month, 'g2b', vec(-r.taxable, -r.igst, -r.cgst, -r.sgst))

def negv(v):
    return dict((k, -v[k]) for k in v)

def subv(a, b):
    return dict((k, a[k] - b[k]) for k in a)

# Pair-level bridge adjustments (guarantees exact reconciliation):
# adjustment = (2B side) - (Books side) for every object, placed in the months it affects.
# --- invoices
for r in bk.itertuples():
    if r.g_idx < 0:
        put(r.book_month, 'bk_only', vec(-r.b_tax, -r.b_ig, -r.b_cg, -r.b_sg))
        continue
    b = vec(r.b_tax, r.b_ig, r.b_cg, r.b_sg)
    gv = vec(r.b_gtax, r.b_gig, r.b_gcg, r.b_gsg)
    if r.b_gm == r.book_month:
        adj = subv(gv, b)
        if any(abs(adj[k]) > 1e-9 for k in Z):
            lab = 'g2b_new' if (r.b_gst == 0 and r.b_ggst > 0) else 'other'
            put(r.book_month, lab, adj)
    else:
        put(r.book_month, 'tm_out', negv(b))
        put(r.b_gm, 'tm_in', gv)
# --- unmatched GSTR-2B invoices (incl. RCM / Rejected)
for r in g.itertuples():
    if not r.used:
        put(r.g2b_month, 'g2b_new', vec(r.taxable, r.igst, r.cgst, r.sgst))
# --- debit notes (books) vs credit notes (GSTR-2B)
for r in dn.itertuples():
    d = vec(r.d_tax, r.d_ig, r.d_cg, r.d_sg)
    if r.c_idx < 0:
        put(r.book_month, 'bk_only', d)          # DN reduced books ITC; no CN in 2B -> add back
        continue
    ci = r.c_idx
    c = vec(cdnr.at[ci, 'taxable'], cdnr.at[ci, 'igst'],
            cdnr.at[ci, 'cgst'], cdnr.at[ci, 'sgst'])
    if r.c_gm == r.book_month:
        adj = subv(d, c)                          # book side -d, 2B side -c  =>  d - c
        if any(abs(adj[k]) > 1e-9 for k in Z):
            put(r.book_month, 'dn_cn_adj', adj)
    else:
        put(r.book_month, 'dn_cn_adj', d)
        put(r.c_gm, 'dn_cn_adj', negv(c))
# --- unmatched GSTR-2B credit notes (reduce 2B ITC; no DN in books)
for r in cdnr.itertuples():
    if not r.used:
        put(r.g2b_month, 'g2b_new', vec(-r.taxable, -r.igst, -r.cgst, -r.sgst))

BRIDGE_ROWS = [
    ('ITC as per Books',                                'books'),
    ('Add: ITC in GSTR-2B not booked in Books',         'g2b_new'),
    ('Less: ITC booked in Books not in GSTR-2B',        'bk_only'),
    ('Less/Add: Timing diff - ITC of current month received in subsequent GSTR-2B month', 'tm_out'),
    ('Add/Less: Timing diff - ITC of earlier month received in current GSTR-2B month', 'tm_in'),
    ('Add/Less: Debit/Credit Note adjustments',         'dn_cn_adj'),
    ('Other differences (GST value mismatch / rounding)', 'other'),
    ('ITC as per GSTR-2B',                              'g2b'),
]

def bridge_frame(m):
    out = []
    for label, key in BRIDGE_ROWS:
        v = rows[m][key]
        out.append([label, v['taxable'], v['igst'], v['cgst'], v['sgst'], tot(v)])
    lsum = {k: sum(x[i] for x in out[:7]) for i, k in
            enumerate(['taxable','igst','cgst','sgst'], 1)}
    g2b = out[7]
    fin = [lsum['taxable'] - g2b[1], lsum['igst'] - g2b[2],
           lsum['cgst'] - g2b[3], lsum['sgst'] - g2b[4],
           lsum['igst'] + lsum['cgst'] + lsum['sgst'] - g2b[5]]
    out.append(['Final Difference', *fin])
    return out

bridge_all = {m: bridge_frame(m) for m in MONTHS}
TOT = {key: dict(Z) for _, key in BRIDGE_ROWS + [('final', None)]}
for m in MONTHS:
    for i, (label, key) in enumerate(BRIDGE_ROWS):
        v = rows[m][key]
        addv(TOT[key], v)
LT_KEYS = [key for _, key in BRIDGE_ROWS[:7]]
ltot = {comp: sum(rows[m][key][comp] for key in LT_KEYS for m in MONTHS) for comp in Z}
g2btot = {comp: sum(rows[m]['g2b'][comp] for m in MONTHS) for comp in Z}
final = [ltot[k] - g2btot[k] for k in ['taxable','igst','cgst','sgst']] + \
        [ltot['igst']+ltot['cgst']+ltot['sgst'] - (g2btot['igst']+g2btot['cgst']+g2btot['sgst'])]

# ------------------------------------------------------------------ totals for explanation
bk_only_gst   = bk.loc[bk['g_idx'] < 0, 'b_gst'].sum()
g_only_gst    = g.loc[~g['used'], 'gst_total'].sum()
dn_only_gst   = dn.loc[dn['c_idx'] < 0, 'd_gst'].sum()
cn_only_gst   = cdnr.loc[~cdnr['used'], 'gst_total'].sum()
matched = bk[bk['g_idx'] >= 0]
notcl_gst     = matched.loc[matched['b_gst'].eq(0) & matched['b_ggst'].gt(0), 'b_ggst'].sum()
mm_dgst       = (matched['b_gst'] - matched['b_ggst']).sum()
tm_sub_gst    = matched.loc[matched['b_gm'] > matched['book_month'], 'b_gst'].sum()
tm_ear_gst    = matched.loc[matched['b_gm'] < matched['book_month'], 'b_ggst'].sum()
dn_tm_gst     = dn.loc[(dn['c_idx'] >= 0) & (dn['c_gm'] != '') &
                       (dn['c_gm'] != dn['book_month']), 'd_gst'].sum()
timing_gst    = tm_sub_gst + tm_ear_gst + dn_tm_gst

books_net_gst = bk['b_gst'].sum() - dn['d_gst'].sum()
g2b_net_gst   = g['gst_total'].sum() - cdnr['gst_total'].sum()
books_net_ig  = bk['b_ig'].sum() - dn['d_ig'].sum()
g2b_net_ig    = g['igst'].sum() - cdnr['igst'].sum()
books_net_cs  = bk['b_cg'].sum() + bk['b_sg'].sum() - dn['d_cg'].sum() - dn['d_sg'].sum()
g2b_net_cs    = g['cgst'].sum() + g['sgst'].sum() - cdnr['cgst'].sum() - cdnr['sgst'].sum()

print('=== BRIDGE CHECK (Final Difference should be 0) ===')
for m in MONTHS:
    f = bridge_all[m][-1]
    print('%-7s igst %+10.2f cgst %+10.2f sgst %+10.2f tot %+12.2f' %
          (MLBL[MONTHS.index(m)], f[2], f[3], f[4], f[5]))
print('TOTAL   igst %+10.2f cgst %+10.2f sgst %+10.2f tot %+12.2f' % tuple(final[1:]))
print('\nBooks net GST %.2f | 2B net GST %.2f | diff %.2f' % (books_net_gst, g2b_net_gst, books_net_gst - g2b_net_gst))
print('IGST  books %.2f | 2B %.2f | CS books %.2f | 2B %.2f' % (books_net_ig, g2b_net_ig, books_net_cs, g2b_net_cs))
print('book-only %.2f | 2B-only %.2f | dn-only %.2f | cn-only %.2f | not-claimed %.2f | mismatch-dgst %.2f' %
      (bk_only_gst, g_only_gst, dn_only_gst, cn_only_gst, notcl_gst, mm_dgst))
print('timing: subsequent %.2f earlier %.2f dn/cn %.2f total %.2f' % (tm_sub_gst, tm_ear_gst, dn_tm_gst, timing_gst))

# ------------------------------------------------------------------ save bridge
pickle.dump({'bridge': bridge_all, 'months': MONTHS, 'mlbl': MLBL,
             'rows_def': BRIDGE_ROWS, 'final': final,
             'stats': dict(books_net_gst=books_net_gst, g2b_net_gst=g2b_net_gst,
                           bk_only=bk_only_gst, g_only=g_only_gst, dn_only=dn_only_gst,
                           cn_only=cn_only_gst, notcl=notcl_gst, mm_dgst=mm_dgst,
                           tm_sub=tm_sub_gst, tm_ear=tm_ear_gst, dn_tm=dn_tm_gst,
                           timing=timing_gst,
                           books_ig=books_net_ig, g2b_ig=g2b_net_ig,
                           books_cs=books_net_cs, g2b_cs=g2b_net_cs)},
            open('/tmp/bridge.pkl', 'wb'))
print('\nbridge saved')
