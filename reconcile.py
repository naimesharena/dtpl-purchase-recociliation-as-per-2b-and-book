#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GST ITC Reconciliation: GSTR-2B vs Books of Accounts
Drive Trucking Pvt Ltd — FY 2025-26 (1-Apr-2025 to 31-Mar-2026)
GSTIN: 24AAJCD4457C1ZV (single GSTIN; purchase registers of 3 branches HR/PL/VL consolidated)

Method
------
1. Books = PURCHASE-HR/PL/VL.xls (purchase register, ITC additions) + DEBIT NOTE-HR/PL/VL.xls
   (supplier credit-note entries that reduce ITC, booked under "Debit Note" voucher type).
2. GSTR-2B = B2B invoices + B2BA amendments + B2B-CDNR credit notes + B2B-CDNRA
   + B2B(Rejected) [ITC rejected on IMS].
3. Matching (invoice level, per supplier GSTIN):
   A) exact (GSTIN, invoice no)
   B) amount+date (GSTIN, inv date, taxable, GST)
   C) amount (GSTIN, gross, GST)
   D) unique GST amount
   E) fuzzy invoice no (Levenshtein <=2) with amount confirmation
   F) optimal assignment per supplier with tolerance (scipy Hungarian) for
      rounding-level differences / missing dates
4. Book month = voucher "Date"; GSTR-2B month = GSTR-1 period of the invoice
   (month in which ITC became available in GSTR-2B).
5. Timing differences = matched invoices whose book month != GSTR-2B month.

No uploaded figure is altered; book 'taxable value' is derived as
Gross Total - SGST - CGST - IGST - Round Off.
"""
import pandas as pd, numpy as np, re, pickle, sys
from collections import defaultdict
from scipy.optimize import linear_sum_assignment

BASE = '/home/user/dtpl-purchase-recociliation-as-per-2b-and-book'
TOL = 0.05          # rupee tolerance for "amounts agree"
AMT_TOL = 1.00      # max total GST delta accepted by tolerance-assignment
DATE_TOL_DAYS = 3   # date consistency tolerance when both dates known
ROUND = 1.0           # rupee level considered rounding

# ---------------------------------------------------------------- helpers
def norm_inv(x):
    if pd.isna(x): return ''
    s = str(x).strip().upper().replace('\xa0', ' ')
    s = re.sub(r'\s+', ' ', s)
    # numeric invoice numbers: drop leading zeros so '69' == '0069'
    if s.isdigit():
        s = s.lstrip('0') or '0'
    return s

def norm_key(x):
    if pd.isna(x): return ''
    return re.sub(r'[^A-Z0-9]', '', str(x).upper())

def p2m(p):
    # GSTR-1 period is MMYYYY e.g. 042025 -> 2025-04
    p = str(p)
    if len(p) == 6 and p.isdigit():
        return f"{p[2:]}-{p[:2]}"
    return p

def lev(a, b):
    la, lb = len(a), len(b)
    if abs(la - lb) > 2: return 99
    if a == b: return 0
    prev = list(range(lb + 1))
    for i in range(1, la + 1):
        cur = [i] + [0] * lb
        for j in range(1, lb + 1):
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] != b[j - 1]))
        prev = cur
    return prev[lb]

# ---------------------------------------------------------------- load
def load_register(path):
    xl = pd.ExcelFile(path, engine='openpyxl')
    df = xl.parse(xl.sheet_names[0], header=None)
    hdr = None
    for i in range(df.shape[0]):
        rv = [str(v) for v in df.iloc[i].tolist()]
        if rv[0].strip() == 'Date' and 'Particulars' in rv:
            hdr = i; break
    headers = df.iloc[hdr].tolist()
    data = df.iloc[hdr + 1:].copy()
    data.columns = [str(h).strip() for h in headers]
    data = data.dropna(how='all').reset_index(drop=True)
    data = data[~data['Particulars'].astype(str).str.contains('Grand Total|^Total', na=False)].reset_index(drop=True)
    return data

import xlrd
wb = xlrd.open_workbook(f'{BASE}/gstr-2B.xls')
def xrows(name, start):
    sh = wb.sheet_by_name(name)
    out = []
    for r in range(start, sh.nrows):
        vals = [sh.cell_value(r, c) for c in range(sh.ncols)]
        if any(str(v).strip() for v in vals):
            out.append(vals)
    return out

# ----- GSTR-2B B2B
b2b_cols = ['gstin','trade_name','inv_no','inv_type','inv_date','inv_value','pos','rcm',
            'taxable','igst','cgst','sgst','cess','period','filing_date','itc_avail','reason',
            'app_pct','source','irn','irn_date']
g = pd.DataFrame(xrows('B2B', 6), columns=b2b_cols)
for c in ['taxable','igst','cgst','sgst','cess','inv_value']:
    g[c] = pd.to_numeric(g[c], errors='coerce').fillna(0.0)
g['gst'] = g['gstin'].map(norm_key)
g['inv'] = g['inv_no'].map(norm_inv)
g['g2b_month'] = g['period'].map(p2m)
g['inv_date_dt'] = pd.to_datetime(g['inv_date'], format='%d-%m-%Y', errors='coerce')
g['gst_total'] = g['igst'] + g['cgst'] + g['sgst']
g['gross'] = g['taxable'] + g['gst_total']
g['sheet'] = 'B2B'
g['status_2b'] = np.where(g['itc_avail'] == 'No', 'ITC Not Available', 'Available')
g['rcm'] = (g['rcm'] == 'Y')

# ----- B2BA amendment (single row)
b2ba = xrows('B2BA', 7)
g2ba = pd.DataFrame([b2ba[0]], columns=['orig_inv','orig_date','gstin','trade_name','inv_no','inv_type',
                                        'inv_date','inv_value','pos','rcm','taxable','igst','cgst','sgst',
                                        'cess','period','filing_date','itc_avail','reason','app_pct'])
g2ba = g2ba.iloc[0]
amend = pd.DataFrame([{
    'gstin': norm_key(g2ba['gstin']), 'trade_name': g2ba['trade_name'],
    'inv_no': g2ba['inv_no'], 'inv': norm_inv(g2ba['inv_no']),
    'inv_date_dt': pd.to_datetime(str(g2ba['inv_date'])[:10], format='%d-%m-%Y', errors='coerce'),
    'rcm': False, 'taxable': float(g2ba['taxable']), 'igst': float(g2ba['igst']),
    'cgst': float(g2ba['cgst']), 'sgst': float(g2ba['sgst']), 'cess': 0.0,
    'period': str(g2ba['period']), 'g2b_month': p2m(str(g2ba['period'])),
    'gst_total': float(g2ba['igst']) + float(g2ba['cgst']) + float(g2ba['sgst']),
    'gross': float(g2ba['taxable']) + float(g2ba['igst']) + float(g2ba['cgst']) + float(g2ba['sgst']),
    'sheet': 'B2BA', 'status_2b': 'Amendment (additional ITC)', 'itc_avail': 'Yes'}])
g = pd.concat([g, amend], ignore_index=True)
# B2BA rows can come with a blank GSTIN cell - inherit from the original B2B
# row of the same invoice number (same supplier) so amendments participate in matching
for gi in g.index[g['gst'].isna()]:
    same = g[(g['sheet'] != 'B2BA') & (g['inv'] == g.at[gi, 'inv']) & g['gst'].notna()]
    same = same[same['trade_name'] == g.at[gi, 'trade_name']] if len(same) > 1 else same
    if len(same):
        g.at[gi, 'gst'] = same.iloc[0]['gst']

# ----- B2B(Rejected): rows have inconsistent cell counts (one zero tax cell
# sometimes dropped). Parse positionally: find period cell (6 digits), cells
# before it = [IGST, CGST, SGST, (CESS)] possibly with one zero omitted.
# Interpretation rule (validated against the official 'ITC Rejected' summary
# totals: IGST 4,961.52 / CGST 16,326.23 / SGST 16,326.23):
#   4 tax cells  -> [IGST, CGST, SGST, CESS]
#   3 tax cells  -> intra-state supplier (GSTIN starts 24): [CGST, SGST, CESS], IGST=0
#                   inter-state supplier:                    [IGST, CGST, SGST], CESS=0
rej_rows = xrows('B2B(Rejected)', 6)
rej_list = []
for vals in rej_rows:
    gstin = str(vals[0]).strip()
    trade = str(vals[1]).strip()
    inv = str(vals[2]).strip()
    inv_type = str(vals[3]).strip()
    inv_date = str(vals[4]).strip()
    inv_value = float(str(vals[5]) or 0)
    pos = str(vals[6]).strip()
    # period cell = first 6-digit numeric after index 7
    pidx = None
    for k in range(7, len(vals)):
        if str(vals[k]).strip().isdigit() and len(str(vals[k]).strip()) == 6:
            pidx = k; break
    if pidx is None: continue
    taxcells = []
    for k in range(7, pidx):
        v = vals[k]
        if str(v).strip() == '':
            continue
        try: taxcells.append(float(v))
        except ValueError: pass
    taxable = None
    # taxable is the first taxcell only if cells before it... actually first cell after pos may be empty (RCM col)
    # taxable = taxcells[0] if len(taxcells)>=4 or intra else handle below
    period = str(vals[pidx]).strip()
    # determine split (see header comment for the validated rule)
    gst_key = norm_key(gstin)
    intra = gst_key[:2] == '24'
    taxable = taxcells[0]
    if len(taxcells) == 5:
        igst, cgst, sgst, cess = taxcells[1:5]
    elif len(taxcells) == 4:
        if intra:  # IGST (zero) omitted
            igst = 0.0; cgst, sgst, cess = taxcells[1:4]
        else:      # CESS (zero) omitted
            igst, cgst, sgst = taxcells[1:4]; cess = 0.0
    else:          # len 3
        if intra:
            igst = 0.0; cgst, sgst = taxcells[1:3]; cess = 0.0
        else:
            igst, cgst = taxcells[1:3]; sgst = 0.0; cess = 0.0
    rej_list.append(dict(gstin=gst_key, trade_name=trade, inv_no=inv, inv_type=inv_type,
                         inv_date=inv_date, inv_value=inv_value, pos=pos, rcm=False,
                         taxable=taxable, igst=igst, cgst=cgst, sgst=sgst, cess=cess,
                         period=period, itc_avail='No', reason=''))
rej = pd.DataFrame(rej_list)
if len(rej):
    rej['gst'] = rej['gstin']
    rej['inv'] = rej['inv_no'].map(norm_inv)
    rej['g2b_month'] = rej['period'].map(p2m)
    rej['inv_date_dt'] = pd.to_datetime(rej['inv_date'], format='%d-%m-%Y', errors='coerce')
    rej['gst_total'] = rej['igst'] + rej['cgst'] + rej['sgst']
    rej['gross'] = rej['taxable'] + rej['gst_total']
    rej['sheet'] = 'B2B(Rejected)'
    rej['status_2b'] = 'ITC Rejected (IMS)'
    g = pd.concat([g, rej], ignore_index=True)
g['r_tax'] = g['taxable'].round(2)
g['r_gst'] = g['gst_total'].round(2)
g['r_gross'] = g['gross'].round(2)

# ----- GSTR-2B credit notes (B2B-CDNR + B2B-CDNRA)
cd_cols = ['gstin','trade_name','note_no','note_type','note_supply_type','note_date','note_value',
           'pos','rcm','taxable','igst','cgst','sgst','cess','period','filing_date','itc_avail','reason',
           'app_pct','source','irn','irn_date']
cdnr = pd.DataFrame(xrows('B2B-CDNR', 6), columns=cd_cols)
for c in ['taxable','igst','cgst','sgst','cess','note_value']:
    cdnr[c] = pd.to_numeric(cdnr[c], errors='coerce').fillna(0.0)
cdnr['gst'] = cdnr['gstin'].map(norm_key)
cdnr['note'] = cdnr['note_no'].map(norm_inv)
cdnr['g2b_month'] = cdnr['period'].map(p2m)
cdnr['note_date_dt'] = pd.to_datetime(cdnr['note_date'], format='%d-%m-%Y', errors='coerce')
cdnr['gst_total'] = cdnr['igst'] + cdnr['cgst'] + cdnr['sgst']
cdnr['gross'] = cdnr['taxable'] + cdnr['gst_total']
cdnr['rcm'] = (cdnr['rcm'] == 'Y')
cdnr['sheet'] = 'B2B-CDNR'
cdnr['status_2b'] = np.where(cdnr['itc_avail'] == 'No', 'ITC Not Available', 'Available')
# CDNRA amendment
cra = xrows('B2B-CDNRA', 7)
if cra:
    r = cra[0]
    cra_row = pd.DataFrame([{
        'gstin': norm_key(r[3]), 'trade_name': r[4], 'note_no': r[5], 'note': norm_inv(r[5]),
        'note_type': 'C', 'note_date_dt': pd.to_datetime(str(r[8])[:10], format='%d-%m-%Y', errors='coerce'),
        'note_value': float(r[9]), 'pos': r[10], 'rcm': str(r[11]) == 'Y',
        'taxable': float(r[12]), 'igst': float(r[13]), 'cgst': float(r[14]), 'sgst': float(r[15]),
        'cess': float(r[16]), 'period': str(r[17]), 'g2b_month': p2m(str(r[17])),
        'gst_total': float(r[13]) + float(r[14]) + float(r[15]),
        'gross': float(r[12]) + float(r[13]) + float(r[14]) + float(r[15]),
        'sheet': 'B2B-CDNRA', 'status_2b': 'Amendment', 'itc_avail': str(r[18])}])
    cdnr = pd.concat([cdnr, cra_row], ignore_index=True)
cdnr = cdnr.reset_index(drop=True)

# ----- Books: purchase registers
regs = {}
for b in ['HR', 'PL', 'VL']:
    d = load_register(f'{BASE}/PURCHASE-{b}.xls')
    for c in ['Gross Total','SGST TAX','CGST TAX','IGST TAX','Round Off']:
        d[c] = pd.to_numeric(d[c], errors='coerce').fillna(0.0)
    d['Date'] = pd.to_datetime(d['Date'], errors='coerce')
    d['Supplier Invoice Date'] = pd.to_datetime(d['Supplier Invoice Date'], errors='coerce')
    d['branch'] = b
    regs[b] = d
book = pd.concat(list(regs.values()), ignore_index=True)
book['gst'] = book['GSTIN/UIN'].map(norm_key)
book['inv'] = book['Supplier Invoice No.'].map(norm_inv)
book['book_month'] = book['Date'].dt.strftime('%Y-%m')
book['taxable'] = book['Gross Total'] - book['SGST TAX'] - book['CGST TAX'] - book['IGST TAX'] - book['Round Off']
book['gst_total'] = book['SGST TAX'] + book['CGST TAX'] + book['IGST TAX']
book['r_tax'] = book['taxable'].round(2)
book['r_gst'] = book['gst_total'].round(2)
book['r_gross'] = book['Gross Total'].round(2)
book['kind'] = 'Invoice'

# ----- Books: debit-note register (supplier credit notes -> ITC reduction)
dn_all = []
for b in ['HR', 'PL', 'VL']:
    d = load_register(f'{BASE}/DEBIT NOTE-{b}.xls')
    for c in ['Gross Total','IGST TAX','CGST TAX','SGST TAX']:
        d[c] = pd.to_numeric(d[c], errors='coerce').fillna(0.0)
    d['Date'] = pd.to_datetime(d['Date'], errors='coerce')
    d['branch'] = b
    dn_all.append(d)
dn = pd.concat(dn_all, ignore_index=True)
dn['gst'] = dn['GSTIN/UIN'].map(norm_key)
dn['note'] = dn['Voucher No.'].map(norm_inv)
dn['book_month'] = dn['Date'].dt.strftime('%Y-%m')
dn['gst_total'] = dn['IGST TAX'] + dn['CGST TAX'] + dn['SGST TAX']
dn['kind'] = 'CreditNote'

# ================================================================ MATCH INVOICES
bk = book.copy()
match = np.full(len(bk), -1, dtype=int)
match_stage = {}
used_g = set()

def _gst_close(b, gg):
    return (abs(b['IGST TAX'] - gg['igst']) <= ROUND and
            abs(b['CGST TAX'] - gg['cgst']) <= ROUND and
            abs(b['SGST TAX'] - gg['sgst']) <= ROUND)

def stage(keyfunc, verify_gst=True):
    """match on exact key; if verify_gst, accept only when GST components agree
    (within ROUND) or one side has zero GST"""
    global match, used_g
    idx = defaultdict(list)
    for gi, row in g.iterrows():
        if gi in used_g: continue
        idx[keyfunc(row)].append(gi)
    for bi in bk.index:
        if match[bi] >= 0: continue
        k = keyfunc(bk.loc[bi])
        if k is None: continue
        cand = idx.get(k)
        if not cand: continue
        cand = [x for x in cand if x not in used_g]
        if not cand: continue
        if verify_gst:
            ok = [gi for gi in cand if _gst_close(bk.loc[bi], g.loc[gi])
                  or bk.at[bi, 'gst_total'] == 0 or g.at[gi, 'gst_total'] == 0]
            if not ok: continue
            cand = ok
        gi = cand[0]
        match[bi] = gi; used_g.add(gi); match_stage[bi] = 'key'

def stage_simple(keyfunc):
    global match, used_g
    idx = defaultdict(list)
    for gi, row in g.iterrows():
        if gi in used_g: continue
        idx[keyfunc(row)].append(gi)
    for bi in bk.index:
        if match[bi] >= 0: continue
        k = keyfunc(bk.loc[bi])
        if k is None: continue
        cand = idx.get(k)
        if not cand: continue
        cand = [x for x in cand if x not in used_g]
        if not cand: continue
        gi = cand[0]
        match[bi] = gi; used_g.add(gi); match_stage[bi] = 'amount'

def stage_unique(keyfunc):
    global match, used_g
    idx = defaultdict(list)
    for gi, row in g.iterrows():
        if gi in used_g: continue
        idx[keyfunc(row)].append(gi)
    for bi in bk.index:
        if match[bi] >= 0: continue
        k = keyfunc(bk.loc[bi])
        if k is None: continue
        cand = [x for x in idx.get(k, []) if x not in used_g]
        if cand and len(cand) == 1:
            match[bi] = cand[0]; used_g.add(cand[0]); match_stage[bi] = 'gst-unique'

def fA(r):
    return ('K', r['gst'], r['inv']) if r['inv'] else None
stage(fA)

def fB(r):
    d = r.get('Supplier Invoice Date')
    if d is None:
        d = r.get('inv_date_dt')
    if d is None or pd.isna(d): return None
    return ('D', r['gst'], d, r['r_tax'], r['r_gst'])
stage_simple(fB)

def fC(r): return ('G', r['gst'], r['r_gross'], r['r_gst'])
stage_simple(fC)

def fD(r): return ('T', r['gst'], r['r_gst'])
stage_unique(fD)

# Fuzzy stage (Levenshtein <= 2, with GST-amount confirmation)
def fuzzy_stage():
    global match, used_g
    bygst = defaultdict(list)
    for gi, row in g.iterrows():
        if gi not in used_g and row['inv']:
            bygst[row['gst']].append(gi)
    added = 0
    for bi in bk.index:
        if match[bi] >= 0: continue
        inv = bk.at[bi, 'inv']; gst = bk.at[bi, 'gst']
        if not inv or len(inv) < 4: continue
        L = len(inv)
        for gi in bygst.get(gst, []):
            if gi in used_g: continue
            ginv = g.at[gi, 'inv']
            if abs(len(ginv) - L) > 2: continue
            if lev(inv, ginv) > 2: continue
            if (_gst_close(bk.loc[bi], g.loc[gi])
                    or abs(bk.at[bi, 'r_gross'] - g.at[gi, 'r_gross']) <= TOL):
                match[bi] = gi; used_g.add(gi); match_stage[bi] = 'fuzzy'; added += 1
                break
    return added
nf = fuzzy_stage()

# Correction pass: key-matches whose GST differs significantly get re-pointed to a
# unique unused same-supplier row with agreeing GST (book invoice no likely mistyped)
ncorr = 0
for bi in bk.index:
    if match[bi] < 0 or match_stage.get(bi) != 'key': continue
    gi = match[bi]; b = bk.loc[bi]; gg = g.loc[gi]
    if (abs(b['IGST TAX'] - gg['igst']) <= ROUND and abs(b['CGST TAX'] - gg['cgst']) <= ROUND
            and abs(b['SGST TAX'] - gg['sgst']) <= ROUND):
        continue
    if b['gst_total'] == 0: continue
    cands = [x for x in g[g['gst'] == b['gst']].index
             if x not in used_g or x == gi]
    near = [x for x in cands if _gst_close(b, g.loc[x])]
    if len(near) == 1 and near[0] != gi:
        used_g.discard(gi)
        match[bi] = near[0]; used_g.add(near[0]); match_stage[bi] = 'key-corrected'
        ncorr += 1

# Optimal per-supplier assignment for the rest:
#  (i)  strict: |dTaxable| + |dGST| <= AMT_TOL
#  (ii) relaxed: |dGST| <= ROUND and uniquely best (handles same invoice booked
#      with different taxable value, e.g. Troops TS-478 vs 478)
rem_bi = [bi for bi in bk.index if match[bi] < 0]
rem_gi_all = [gi for gi in g.index if gi not in used_g]
bygst_b = defaultdict(list)
for bi in rem_bi: bygst_b[bk.at[bi, 'gst']].append(bi)
bygst_g = defaultdict(list)
for gi in rem_gi_all:
    bygst_g[g.at[gi, 'gst']].append(gi)
nassign = 0
for gst in set(bygst_b) & set(bygst_g):
    bis = bygst_b[gst]; gis = bygst_g[gst]
    if not bis or not gis: continue
    B = bk.loc[bis]; G = g.loc[gis]
    n, m = len(B), len(G)
    if n * m > 400000:
        cand = set()
        gmap2 = defaultdict(list)
        for gi in gis: gmap2[round(g.at[gi, 'r_gst'], 0)].append(gi)
        for bi in bis:
            for gi in gmap2.get(round(bk.at[bi, 'r_gst'], 0), []): cand.add(gi)
        if not cand: continue
        gis = [gi for gi in gis if gi in cand]
        G = g.loc[gis]
        m = len(G)
        if n * m > 400000: continue
    amt = np.zeros((n, m)); gstonly = np.zeros((n, m))
    for i in range(n):
        for j in range(m):
            b = B.iloc[i]; gg = G.iloc[j]
            gcost = (abs(b['IGST TAX'] - gg['igst']) + abs(b['CGST TAX'] - gg['cgst'])
                     + abs(b['SGST TAX'] - gg['sgst']))
            gstonly[i, j] = gcost
            amt[i, j] = abs(b['r_tax'] - gg['r_tax']) + gcost
    BIG = 1e9
    # (i) strict assignment
    costf = np.full((n, m + n), BIG)
    costf[:, :m] = amt
    costf[:, :m][amt > AMT_TOL] = BIG
    r_, c_ = linear_sum_assignment(costf)
    for i, j in zip(r_, c_):
        if j < m and costf[i, j] < BIG:
            match[bis[i]] = gis[j]
            used_g.add(gis[j]); match_stage[bis[i]] = 'assign-strict'; nassign += 1
    bis2 = [bi for bi in bis if match[bi] < 0]
    gis2 = [gi for gi in gis if gi not in used_g]
    if not bis2 or not gis2: continue
    B2 = bk.loc[bis2]; G2 = g.loc[gis2]
    n2, m2 = len(B2), len(G2)
    if n2 * m2 > 400000: continue
    # (ii) relaxed: per book row, unique nearest by |dGST| within ROUND
    gcost2 = np.zeros((n2, m2))
    for i in range(n2):
        for j in range(m2):
            b = B2.iloc[i]; gg = G2.iloc[j]
            gcost2[i, j] = (abs(b['IGST TAX'] - gg['igst']) + abs(b['CGST TAX'] - gg['cgst'])
                            + abs(b['SGST TAX'] - gg['sgst']))
    for i in range(n2):
        okj = [j for j in range(m2)
               if gis2[j] not in used_g and gcost2[i, j] <= ROUND]
        if len(okj) != 1: continue
        j = okj[0]
        # date guard when both dates known
        b = B2.iloc[i]; gg = G2.iloc[j]
        bd_ = b.get('Supplier Invoice Date'); gd = gg['inv_date_dt']
        if pd.notna(bd_) and pd.notna(gd) and abs((pd.Timestamp(bd_) - gd).days) > 15:
            continue
        match[bis2[i]] = gis2[j]
        used_g.add(gis2[j]); match_stage[bis2[i]] = 'assign-gstonly'; nassign += 1

bk['g_idx'] = match
pickle.dump(bk, open('/tmp/bk_final.pkl', 'wb'))
pickle.dump(g, open('/tmp/g_final.pkl', 'wb'))

rem_bk = bk[bk['g_idx'] < 0]
rem_g = g[~g.index.isin(used_g)]
print('Books rows matched:', int((match >= 0).sum()), '/', len(bk),
      '(fuzzy:', nf, 'corrected:', ncorr, 'assigned:', nassign, ')')
print('stage mix:', pd.Series(list(match_stage.values())).value_counts().to_dict())
print('Unmatched book rows:', len(rem_bk),
      '| with GST: %d sgst=%.2f cgst=%.2f igst=%.2f' % (
        (rem_bk['gst_total'] > 0).sum(), rem_bk['SGST TAX'].sum(), rem_bk['CGST TAX'].sum(), rem_bk['IGST TAX'].sum()))
print('Unmatched GSTR-2B rows:', len(rem_g),
      'taxable=%.2f igst=%.2f cgst=%.2f sgst=%.2f' % (rem_g['taxable'].sum(), rem_g['igst'].sum(), rem_g['cgst'].sum(), rem_g['sgst'].sum()))

# ================================================================ MATCH CREDIT NOTES
# Books: 'DEBIT NOTE' register entries = supplier credit notes (reduce ITC)
dnm = np.full(len(dn), -1, dtype=int)
used_c = set()
cidx = {}
for ci, row in cdnr.iterrows():
    cidx.setdefault((row['gst'], row['note']), []).append(ci)
for di in dn.index:
    k = (dn.at[di, 'gst'], dn.at[di, 'note'])
    if dn.at[di, 'gst'] and k in cidx and cidx[k] and cidx[k][0] not in used_c:
        dnm[di] = cidx[k][0]; used_c.add(cidx[k][0])
# fuzzy on note number within same gst
def cfuzzy():
    byg = defaultdict(list)
    for ci, row in cdnr.iterrows():
        if ci not in used_c and row['note']:
            byg[row['gst']].append(ci)
    n = 0
    for di in dn.index:
        if dnm[di] >= 0: continue
        note = dn.at[di, 'note']; gst = dn.at[di, 'gst']
        if not note or not gst: continue
        nk = norm_key(note); best = None; bd = 3
        for ci in byg.get(gst, []):
            cnk = norm_key(cdnr.at[ci, 'note'])
            if abs(len(cnk) - len(nk)) > 2: continue
            d = lev(nk, cnk)
            if d < bd: bd = d; best = ci
        if best is not None and bd <= 2:
            # amount confirmation on GST components
            if abs(dn.at[di, 'IGST TAX'] - cdnr.at[best, 'igst']) <= TOL or \
               abs(dn.at[di, 'CGST TAX'] - cdnr.at[best, 'cgst']) <= TOL or \
               abs(dn.at[di, 'SGST TAX'] - cdnr.at[best, 'sgst']) <= TOL:
                dnm[di] = best; used_c.add(best); n += 1
    return n
nfn = cfuzzy()
# cross-GSTIN: same note number, different supplier GSTIN (e.g. Daimler 9106250255
# booked under 33AABCF1590N1ZJ in books but reported under 06AABCF1590N1ZG in GSTR-1)
for di in dn.index:
    if dnm[di] >= 0: continue
    note = dn.at[di, 'note']
    if not note: continue
    nk = norm_key(note)
    cands = [ci for ci in cdnr.index
             if ci not in used_c and norm_key(cdnr.at[ci, 'note']) == nk
             and cdnr.at[ci, 'gst'] != dn.at[di, 'gst']]
    near = [ci for ci in cands
            if (abs(dn.at[di, 'IGST TAX'] - cdnr.at[ci, 'igst']) <= ROUND and
                abs(dn.at[di, 'CGST TAX'] - cdnr.at[ci, 'cgst']) <= ROUND and
                abs(dn.at[di, 'SGST TAX'] - cdnr.at[ci, 'sgst']) <= ROUND)]
    if len(near) == 1:
        dnm[di] = near[0]; used_c.add(near[0])
# unmatched notes: per-supplier amount assignment
rem_di = [di for di in dn.index if dnm[di] < 0 and dn.at[di, 'gst_total'] > 0]
rem_ci = [ci for ci in cdnr.index if ci not in used_c and cdnr.at[ci, 'gst_total'] > 0]
bg_ = defaultdict(list)
for di in rem_di: bg_[dn.at[di, 'gst']].append(di)
cg_ = defaultdict(list)
for ci in rem_ci: cg_[cdnr.at[ci, 'gst']].append(ci)
for gst in set(bg_) & set(cg_):
    dis = bg_[gst]; cis = cg_[gst]
    n, m = len(dis), len(cis)
    if not n or not m: continue
    cost = np.zeros((n, m))
    for i in range(n):
        for j in range(m):
            cost[i, j] = abs(dn.at[dis[i], 'IGST TAX'] - cdnr.at[cis[j], 'igst']) + \
                         abs(dn.at[dis[i], 'CGST TAX'] - cdnr.at[cis[j], 'cgst']) + \
                         abs(dn.at[dis[i], 'SGST TAX'] - cdnr.at[cis[j], 'sgst'])
    BIG = 1e9
    costf = np.full((n, m + n), BIG)
    costf[:, :m] = cost
    costf[:, :m][cost > AMT_TOL] = BIG
    r_, c_ = linear_sum_assignment(costf)
    for i, j in zip(r_, c_):
        if j < m and costf[i, j] < BIG:
            dnm[dis[i]] = cis[j]; used_c.add(cis[j])
dn['c_idx'] = dnm

# ================================================================ CLASSIFY
def r2(x): return round(float(x), 2)

# ROUND moved to constants

def inv_status(bi):
    """classify one book invoice row against its g match. Returns (status, reason, diffs dict)"""
    b = bk.loc[bi]
    diffs = dict(d_tax=0.0, d_igst=0.0, d_cgst=0.0, d_sgst=0.0, d_gst=0.0, d_gross=0.0)
    if b['g_idx'] < 0:
        if b['gst_total'] == 0:
            return ('Booked - no ITC claimed (no GST in books)', '', diffs)
        return ('In Books, not in GSTR-2B', 'To be verified', diffs)
    gi = b['g_idx']; gg = g.loc[gi]
    d_gst = b['gst_total'] - gg['gst_total']
    d_igst = b['IGST TAX'] - gg['igst']
    d_cg = b['CGST TAX'] - gg['cgst']
    d_sg = b['SGST TAX'] - gg['sgst']
    d_tax = b['taxable'] - gg['taxable']
    d_gross = b['Gross Total'] - gg['gross']
    diffs = dict(d_tax=r2(d_tax), d_igst=r2(d_igst), d_cgst=r2(d_cg), d_sgst=r2(d_sg),
                 d_gst=r2(d_gst), d_gross=r2(d_gross))
    # date mismatch check (invoice date vs book supplier invoice date)
    date_mis = ''
    if pd.notna(b['Supplier Invoice Date']) and pd.notna(gg['inv_date_dt']):
        if abs((b['Supplier Invoice Date'] - gg['inv_date_dt']).days) > 0:
            date_mis = 'Invoice date differs'
    bm = b['book_month']; gm = gg['g2b_month']
    rej_flag = gg['sheet'] == 'B2B(Rejected)'
    na_flag = (gg['itc_avail'] == 'No') or (gg['sheet'] == 'B2B(Rejected)')
    timing = ''
    if bm != gm:
        timing = ('Timing Difference - ITC reflected in subsequent GSTR-2B' if bm < gm
                  else 'Timing Difference - ITC reflected in earlier GSTR-2B') + \
                 f' (booked {bm}; in GSTR-2B {gm})'
    # (1) booked in books without any GST split, but supplier reported GST in GSTR-2B
    if b['gst_total'] == 0 and gg['gst_total'] > 0:
        s = 'ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split)'
        if rej_flag: s += ' [ITC Rejected on IMS]'
        elif na_flag: s += ' [ITC Not Available]'
        return (s, 'Verify whether ITC is claimable / should have been booked', diffs)
    # (2) genuine GST component difference
    if not (abs(d_igst) <= TOL and abs(d_cg) <= TOL and abs(d_sg) <= TOL):
        base = 'GST Mismatch'
        if rej_flag: base += ' + ITC Rejected (IMS)'
        elif na_flag: base += ' + ITC Not Available'
        reason = f'GST diff (Books-2B): IGST {diffs["d_igst"]} / CGST {diffs["d_cgst"]} / SGST {diffs["d_sgst"]} - to be verified'
        if timing: reason = timing + '; ' + reason
        return (base, reason, diffs)
    # (3) GST agrees (rounding <= Rs 1 treated as rounding); check rest
    if rej_flag:
        s = 'Matched - but ITC REJECTED on IMS' + (' (rounding diff)' if abs(d_gst) > TOL else '')
        return (s, 'Reversal of ITC required', diffs)
    if na_flag:
        s = 'Matched - but ITC NOT AVAILABLE in GSTR-2B' + (' (rounding diff)' if abs(d_gst) > TOL else '')
        return (s, 'Reversal of ITC required (e.g. sec 17(5))', diffs)
    if abs(d_gst) > TOL:
        base = 'Matched (GST rounding diff <= Rs 1)'
        reason = ''
        if abs(d_tax) > TOL: reason += f'Taxable value differs (Books-2B: {diffs["d_tax"]}) - no GST impact'
        if date_mis: reason += ('; ' if reason else '') + 'Invoice date differs'
        if timing: reason += ('; ' if reason else '') + timing
        return (base, reason, diffs)
    if timing:
        reason = timing
        if abs(d_tax) > TOL: reason += f'; Taxable value differs (Books-2B: {diffs["d_tax"]}) - no GST impact'
        return ('Timing Difference - ITC reflected in ' + ('subsequent' if bm < gm else 'earlier') + ' GSTR-2B', reason, diffs)
    if abs(d_tax) > TOL:
        reason = f'Taxable value differs (Books-2B: {diffs["d_tax"]}) - no GST impact'
        if date_mis: reason += '; ' + date_mis
        return ('Matched (taxable value differs - no GST impact)', reason, diffs)
    if date_mis:
        return ('Matched - invoice date differs', 'Verify invoice date', diffs)
    return ('Matched', '', diffs)

_st = bk.apply(lambda r: inv_status(r.name), axis=1)
bk['status'] = [x[0] for x in _st]
bk['reason'] = [x[1] for x in _st]
for dk in ['d_tax', 'd_igst', 'd_cgst', 'd_sgst', 'd_gst', 'd_gross']:
    bk[dk] = [x[2][dk] for x in _st]
# note invoice-number differences between books and GSTR-2B
for bi in bk.index:
    gi = bk.at[bi, 'g_idx']
    if gi >= 0 and bk.at[bi, 'inv']:
        gg_inv = g.at[gi, 'inv']
        if gg_inv and gg_inv != bk.at[bi, 'inv']:
            note = f'Invoice no differs (Books {bk.at[bi, "inv"]} vs GSTR-2B {gg_inv})'
            r = bk.at[bi, 'reason']
            bk.at[bi, 'reason'] = (r + '; ' + note) if r else note
# duplicate flag: same (gst, inv) booked >1 times in books (with GST)
dupmask = (bk['gst'] != '') & (bk['gst_total'] > 0) & bk.duplicated(subset=['gst', 'inv'], keep=False)
_s = bk.loc[dupmask, 'status']
bk.loc[dupmask, 'status'] = np.where(_s.str.contains('DUPLICATE'), _s, 'DUPLICATE booking - ' + _s)

# classify credit notes
def dn_status(di):
    d = dn.loc[di]
    if d['gst_total'] == 0:
        return ('Debit/Credit Note - no GST amount in books', 'Verify - RCM or unregistered', 0.0)
    if d['c_idx'] < 0:
        return ('In Books, not in GSTR-2B (Credit Note)', 'To be verified', 0.0)
    ci = d['c_idx']; c = cdnr.loc[ci]
    d_gst = d['gst_total'] - c['gst_total']
    gstin_note = ('Supplier GSTIN differs: Books %s vs GSTR-2B %s - to be verified'
                  % (d['gst'], c['gst'])) if d['gst'] != c['gst'] else ''
    if abs(d_gst) > ROUND:
        reason = f"Books {r2(d['gst_total'])} vs GSTR-2B {r2(c['gst_total'])} - to be verified"
        if gstin_note: reason = gstin_note + '; ' + reason
        return ('Credit Note - GST Mismatch', reason, r2(d_gst))
    bm = d['book_month']; gm = c['g2b_month']
    if bm != gm:
        base = f'Booked {bm}; in GSTR-2B {gm}'
        if gstin_note: base = gstin_note + '; ' + base
        if bm < gm:
            return ('Credit Note - Timing Difference (subsequent GSTR-2B)', base, r2(d_gst))
        return ('Credit Note - Timing Difference (earlier GSTR-2B)', base, r2(d_gst))
    if gstin_note:
        return ('Credit Note - Matched (GSTIN differs)', gstin_note, r2(d_gst))
    return ('Credit Note - Matched', '', r2(d_gst))

_st = dn.apply(lambda r: dn_status(r.name), axis=1)
dn['status'] = [x[0] for x in _st]
dn['reason'] = [x[1] for x in _st]
dn['d_gst'] = [x[2] for x in _st]
# duplicate DN
dupd = (dn['gst'] != '') & (dn['gst_total'] > 0) & dn.duplicated(subset=['gst', 'note'], keep=False)
_s = dn.loc[dupd, 'status']
dn.loc[dupd, 'status'] = np.where(_s.str.contains('DUPLICATE'), _s, 'DUPLICATE booking - ' + _s)

# unmatched GSTR-2B rows status
rem_gf = g[~g.index.isin(set(bk['g_idx'].tolist()))].copy()
rem_gf['status'] = np.where(rem_gf['sheet'] == 'B2B(Rejected)',
                            'In GSTR-2B (Rejected on IMS), not in Books',
                            np.where(rem_gf['rcm'], 'In GSTR-2B (RCM), not in Books',
                                     'In GSTR-2B, not in Books'))
rem_gf['reason'] = 'To be verified - check supplier GSTR-1 / posting in books'
# unmatched credit notes in GSTR-2B
rem_cf = cdnr[~cdnr.index.isin(set(dn['c_idx'].tolist()))].copy()
rem_cf['status'] = np.where(rem_cf['itc_avail'] == 'No',
                            'Credit Note in GSTR-2B (ITC not available), not in Books',
                            'Credit Note in GSTR-2B, not in Books')
rem_cf['reason'] = 'To be verified - check posting of credit note in books'

pickle.dump(bk, open('/tmp/bk_final.pkl', 'wb'))
pickle.dump(dn, open('/tmp/dn_final.pkl', 'wb'))
pickle.dump(rem_gf, open('/tmp/rem_gf.pkl', 'wb'))
pickle.dump(rem_cf, open('/tmp/rem_cf.pkl', 'wb'))

print('\n=== INVOICE STATUS MIX ===')
print(bk['status'].value_counts().to_string())
print('\n=== DN STATUS MIX ===')
print(dn['status'].value_counts().to_string())
print('\n=== REMAINING GSTR-2B (not in books) ===', len(rem_gf))
print(rem_gf.groupby('status').agg(n=('inv', 'size'), igst=('igst', 'sum'), cgst=('cgst', 'sum'), sgst=('sgst', 'sum')).to_string())
print('\n=== REMAINING GSTR-2B CREDIT NOTES (not in books) ===', len(rem_cf))
print(rem_cf[['gstin','trade_name','note','note_date','taxable','igst','cgst','sgst','g2b_month']].to_string(max_colwidth=24))
