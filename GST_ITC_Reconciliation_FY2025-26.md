# GSTR-2B vs Books ITC Reconciliation - Drive Trucking Pvt Ltd - FY 2025-26

**Period:** 01-Apr-2025 to 31-Mar-2026 | **Branches:** HR / PL / VL | **Sources:** GSTR-2B (B2B + B2BA + CDN), Purchase Register, Debit Note Register

**Matching basis:** supplier GSTIN + invoice number (fuzzy where mistyped), invoice date, taxable value, IGST/CGST/SGST and gross amount (rupee-level rounding tolerance Rs 1). Duplicate bookings in books are flagged, not removed.

**Note on debit/credit notes:** the "Debit Note" register in the books records the supplier credit notes (ITC reductions); these are matched against the credit notes (CDN) reported by the supplier in GSTR-2B. Updated DEBIT NOTE-HR/PL/VL files (with corrected Gross Total values) are used in this report.

## 1. Headline position (Rs)

| Particulars | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|
| ITC as per Books (net of debit notes) | 25,368,267.15 | 1,913,216.60 | 1,913,216.60 | 29,194,700.35 |
| ITC as per GSTR-2B (net of credit notes) | 25,066,835.11 | 2,176,295.98 | 2,176,295.98 | 29,419,427.07 |
| Difference (GSTR-2B minus Books) | -301,432.04 | 263,079.38 | 263,079.38 | 224,726.72 |

## 2. Section A - Month-wise Summary Reconciliation

Format: Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST. The bridge converts ITC as per Books into ITC as per GSTR-2B month by month; "Final Difference" is zero in every month by construction (all differences are explained).

**Line-item drill-down:** every rupee of each bridge line is traced to the underlying invoice/note in Excel sheet **A2. Bridge Line Details** (Month | Line | Type | Supplier | GSTIN | Invoice No (Books) | Invoice No (2B) | Book Month | 2B Month | Taxable | IGST | CGST | SGST | Total GST | Remarks). Filter by Month + Line to see exactly which invoices sit behind each bridge row - e.g. Line = "Add: ITC in GSTR-2B not booked in Books" + Month = Sep-25 lists the Parth credit notes and the 2B invoices not present in books for that month.

### Apr-25 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 19,701,627.43 | 3,450,613.81 | 251,889.43 | 251,889.43 | 3,954,392.67 |
| Add: ITC in GSTR-2B not booked in Books | 710,531.16 | 69,451.78 | 49,208.89 | 49,208.89 | 167,869.56 |
| Less: ITC booked in Books not in GSTR-2B | -1,491,171.89 | -69,451.79 | -5,352.77 | -5,352.77 | -80,157.33 |
| Add/Less: Timing difference - previous month ITC received in current month | 6,870.63 | 0.00 | 719.14 | 719.14 | 1,438.28 |
| Add/Less: Timing difference - current month ITC in subsequent month | -74,927.99 | -11,939.31 | -2,624.35 | -2,624.35 | -17,188.01 |
| Add/Less: Debit/Credit Note adjustments | -4,686.15 | -0.04 | -587.39 | -587.39 | -1,174.82 |
| Other differences | 36,940.30 | 0.13 | -0.09 | -0.09 | -0.05 |
| ITC as per GSTR-2B | 18,885,183.49 | 3,438,674.58 | 293,252.86 | 293,252.86 | 4,025,180.30 |
| Final Difference | -0.00 | 0.00 | -0.00 | -0.00 | 0.00 |

### May-25 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 12,546,246.44 | 1,787,047.73 | 257,231.81 | 257,231.81 | 2,301,511.35 |
| Add: ITC in GSTR-2B not booked in Books | 2,308.71 | 696.60 | 46,449.34 | 46,449.34 | 93,595.28 |
| Less: ITC booked in Books not in GSTR-2B | -339,477.00 | -673.90 | -4,319.96 | -4,319.96 | -9,313.82 |
| Add/Less: Timing difference - previous month ITC received in current month | 7,488.00 | 0.00 | 954.47 | 954.47 | 1,908.94 |
| Add/Less: Timing difference - current month ITC in subsequent month | -16,529.24 | -710.90 | -1,580.43 | -1,580.43 | -3,871.76 |
| Add/Less: Debit/Credit Note adjustments | 8,178.88 | 0.00 | 1,017.82 | 1,017.82 | 2,035.64 |
| Other differences | 41,817.93 | -0.01 | 0.09 | 0.09 | 0.17 |
| ITC as per GSTR-2B | 12,250,033.72 | 1,786,359.52 | 299,753.14 | 299,753.14 | 2,385,865.80 |
| Final Difference | -0.00 | -0.00 | -0.00 | -0.00 | -0.00 |

### Jun-25 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 12,921,132.11 | 1,600,787.87 | 222,751.82 | 222,751.82 | 2,046,291.51 |
| Add: ITC in GSTR-2B not booked in Books | 1,949,724.16 | 0.00 | 359,514.72 | 359,514.72 | 719,029.44 |
| Less: ITC booked in Books not in GSTR-2B | -495,271.00 | 0.00 | 0.00 | 0.00 | 0.00 |
| Add/Less: Timing difference - previous month ITC received in current month | 38,577.51 | 0.00 | 3,876.99 | 3,876.99 | 7,753.98 |
| Add/Less: Timing difference - current month ITC in subsequent month | -46,984.66 | 0.00 | -5,283.67 | -5,283.67 | -10,567.34 |
| Add/Less: Debit/Credit Note adjustments | -0.02 | 0.00 | 0.00 | 0.00 | 0.00 |
| Other differences | 43,137.19 | 0.14 | 0.02 | 0.02 | 0.18 |
| ITC as per GSTR-2B | 14,410,315.29 | 1,600,788.01 | 580,859.88 | 580,859.88 | 2,762,507.77 |
| Final Difference | -0.00 | -0.00 | 0.00 | 0.00 | -0.00 |

### Jul-25 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 9,462,660.61 | 1,305,964.67 | 226,480.75 | 226,480.75 | 1,758,926.17 |
| Add: ITC in GSTR-2B not booked in Books | 87,145.18 | 13,451.00 | 600.13 | 600.13 | 14,651.26 |
| Less: ITC booked in Books not in GSTR-2B | -869,222.04 | -13,412.13 | -67.50 | -67.50 | -13,547.13 |
| Add/Less: Timing difference - previous month ITC received in current month | 35,636.56 | 0.00 | 3,937.49 | 3,937.49 | 7,874.98 |
| Add/Less: Timing difference - current month ITC in subsequent month | -22,860.20 | 0.00 | -2,822.40 | -2,822.40 | -5,644.80 |
| Add/Less: Debit/Credit Note adjustments | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
| Other differences | 42,656.69 | -0.07 | 0.08 | 0.08 | 0.09 |
| ITC as per GSTR-2B | 8,736,016.80 | 1,306,003.47 | 228,128.55 | 228,128.55 | 1,762,260.57 |
| Final Difference | 0.00 | -0.00 | 0.00 | 0.00 | -0.00 |

### Aug-25 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 9,265,333.16 | 1,620,279.08 | 131,567.49 | 131,567.49 | 1,883,414.06 |
| Add: ITC in GSTR-2B not booked in Books | 54,444.13 | 14,123.38 | 10.71 | 10.71 | 14,144.80 |
| Less: ITC booked in Books not in GSTR-2B | -711,722.31 | -14,123.38 | 0.00 | 0.00 | -14,123.38 |
| Add/Less: Timing difference - previous month ITC received in current month | 1,000.00 | 0.00 | 90.00 | 90.00 | 180.00 |
| Add/Less: Timing difference - current month ITC in subsequent month | -6,385.24 | 0.00 | -646.88 | -646.88 | -1,293.76 |
| Add/Less: Debit/Credit Note adjustments | 0.08 | 0.00 | 0.00 | 0.00 | 0.00 |
| Other differences | 47,091.10 | -0.04 | 0.14 | 0.14 | 0.24 |
| ITC as per GSTR-2B | 8,649,760.92 | 1,620,279.04 | 131,021.46 | 131,021.46 | 1,882,321.96 |
| Final Difference | 0.00 | -0.00 | 0.00 | 0.00 | -0.00 |

### Sep-25 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 13,306,380.37 | 2,146,950.78 | 146,869.53 | 146,869.53 | 2,440,689.84 |
| Add: ITC in GSTR-2B not booked in Books | -2,164,563.01 | 3,601.52 | -195,083.41 | -195,083.41 | -386,565.30 |
| Less: ITC booked in Books not in GSTR-2B | -895,907.09 | -7,914.86 | -3,679.10 | -3,679.10 | -15,273.06 |
| Add/Less: Timing difference - previous month ITC received in current month | 76,364.50 | 13,170.11 | 2,857.28 | 2,857.28 | 18,884.67 |
| Add/Less: Timing difference - current month ITC in subsequent month | -882.00 | 0.00 | -90.00 | -90.00 | -180.00 |
| Add/Less: Debit/Credit Note adjustments | 0.38 | 0.00 | 0.00 | 0.00 | 0.00 |
| Other differences | 50,111.73 | 0.12 | -0.20 | -0.20 | -0.28 |
| ITC as per GSTR-2B | 10,371,504.88 | 2,155,807.67 | -49,125.90 | -49,125.90 | 2,057,555.87 |
| Final Difference | -0.00 | -0.00 | 0.00 | 0.00 | -0.00 |

### Oct-25 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 11,538,166.21 | 1,778,508.14 | 101,294.88 | 101,294.88 | 1,981,097.90 |
| Add: ITC in GSTR-2B not booked in Books | 7,276.96 | 860.13 | 224.86 | 224.86 | 1,309.85 |
| Less: ITC booked in Books not in GSTR-2B | -583,923.60 | 0.00 | 0.00 | 0.00 | 0.00 |
| Add/Less: Timing difference - previous month ITC received in current month | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
| Add/Less: Timing difference - current month ITC in subsequent month | -11,685.66 | 0.00 | -1,186.17 | -1,186.17 | -2,372.34 |
| Add/Less: Debit/Credit Note adjustments | -0.12 | 0.00 | 0.00 | 0.00 | 0.00 |
| Other differences | 63,749.48 | 0.16 | 0.03 | 0.03 | 0.22 |
| ITC as per GSTR-2B | 11,013,583.27 | 1,779,368.43 | 100,333.60 | 100,333.60 | 1,980,035.63 |
| Final Difference | 0.00 | -0.00 | -0.00 | -0.00 | -0.00 |

### Nov-25 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 13,451,060.80 | 2,089,293.68 | 124,662.82 | 124,662.82 | 2,338,619.32 |
| Add: ITC in GSTR-2B not booked in Books | 113,758.80 | 16,957.30 | 1,171.29 | 1,171.29 | 19,299.88 |
| Less: ITC booked in Books not in GSTR-2B | -602,175.97 | -13,501.28 | -2,083.50 | -2,083.50 | -17,668.28 |
| Add/Less: Timing difference - previous month ITC received in current month | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
| Add/Less: Timing difference - current month ITC in subsequent month | -5,648.32 | 0.00 | -508.34 | -508.34 | -1,016.68 |
| Add/Less: Debit/Credit Note adjustments | 0.18 | 0.00 | 0.00 | 0.00 | 0.00 |
| Other differences | 65,292.79 | 0.07 | 0.01 | 0.01 | 0.09 |
| ITC as per GSTR-2B | 13,022,288.28 | 2,092,749.77 | 123,242.28 | 123,242.28 | 2,339,234.33 |
| Final Difference | 0.00 | -0.00 | -0.00 | -0.00 | -0.00 |

### Dec-25 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 16,663,080.54 | 2,691,207.83 | 103,445.28 | 103,445.28 | 2,898,098.39 |
| Add: ITC in GSTR-2B not booked in Books | 35,483.91 | 519.89 | 2,933.60 | 2,933.60 | 6,387.09 |
| Less: ITC booked in Books not in GSTR-2B | -670,127.04 | 0.00 | -6,211.89 | -6,211.89 | -12,423.78 |
| Add/Less: Timing difference - previous month ITC received in current month | 7,157.78 | 0.00 | 644.21 | 644.21 | 1,288.42 |
| Add/Less: Timing difference - current month ITC in subsequent month | -3,799.31 | -519.90 | -82.00 | -82.00 | -683.90 |
| Add/Less: Debit/Credit Note adjustments | 1,007.49 | 181.33 | 0.00 | 0.00 | 181.33 |
| Other differences | 54,329.77 | -0.05 | 0.04 | 0.04 | 0.03 |
| ITC as per GSTR-2B | 16,087,133.14 | 2,691,389.10 | 100,729.24 | 100,729.24 | 2,892,847.58 |
| Final Difference | 0.00 | -0.00 | 0.00 | 0.00 | -0.00 |

### Jan-26 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 14,247,044.04 | 2,233,223.43 | 106,618.49 | 106,618.49 | 2,446,460.41 |
| Add: ITC in GSTR-2B not booked in Books | -7,052.00 | 0.00 | 3,526.00 | 3,526.00 | 7,052.00 |
| Less: ITC booked in Books not in GSTR-2B | -618,951.54 | -86.18 | -232.63 | -232.63 | -551.44 |
| Add/Less: Timing difference - previous month ITC received in current month | 6,483.36 | 0.00 | 583.50 | 583.50 | 1,167.00 |
| Add/Less: Timing difference - current month ITC in subsequent month | -11,278.74 | 0.00 | -1,015.13 | -1,015.13 | -2,030.26 |
| Add/Less: Debit/Credit Note adjustments | 0.28 | 0.00 | 0.00 | 0.00 | 0.00 |
| Other differences | 57,161.64 | 0.17 | 0.03 | 0.03 | 0.23 |
| ITC as per GSTR-2B | 13,673,407.04 | 2,233,137.42 | 109,480.26 | 109,480.26 | 2,452,097.94 |
| Final Difference | -0.00 | -0.00 | 0.00 | 0.00 | -0.00 |

### Feb-26 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 12,646,348.35 | 1,935,124.71 | 119,342.81 | 119,342.81 | 2,173,810.33 |
| Add: ITC in GSTR-2B not booked in Books | 157,866.16 | 14,510.83 | 6,075.00 | 6,075.00 | 26,660.83 |
| Less: ITC booked in Books not in GSTR-2B | -732,888.91 | -12,701.41 | -5,508.00 | -5,508.00 | -23,717.41 |
| Add/Less: Timing difference - previous month ITC received in current month | 18,470.00 | 0.00 | 1,662.30 | 1,662.30 | 3,324.60 |
| Add/Less: Timing difference - current month ITC in subsequent month | -25,232.24 | 0.00 | -1,582.38 | -1,582.38 | -3,164.76 |
| Add/Less: Debit/Credit Note adjustments | -1,007.71 | -181.41 | 0.00 | 0.00 | -181.41 |
| Other differences | 59,265.02 | -0.16 | 0.02 | 0.02 | -0.12 |
| ITC as per GSTR-2B | 12,122,820.67 | 1,936,752.56 | 119,989.75 | 119,989.75 | 2,176,732.06 |
| Final Difference | -0.00 | 0.00 | 0.00 | 0.00 | 0.00 |

### Mar-26 

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 17,474,561.80 | 2,729,265.42 | 121,061.49 | 121,061.49 | 2,971,388.40 |
| Add: ITC in GSTR-2B not booked in Books | 467,333.48 | 38,949.92 | 20,577.00 | 20,577.00 | 80,103.92 |
| Less: ITC booked in Books not in GSTR-2B | -2,979,690.63 | -342,689.92 | -5,687.14 | -5,687.14 | -354,064.20 |
| Add/Less: Timing difference - previous month ITC received in current month | 29,772.22 | 0.00 | 2,679.50 | 2,679.50 | 5,359.00 |
| Add/Less: Timing difference - current month ITC in subsequent month | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
| Add/Less: Debit/Credit Note adjustments | 0.64 | 0.00 | 0.00 | 0.00 | 0.00 |
| Other differences | 53,645.87 | 0.12 | 0.01 | 0.01 | 0.14 |
| ITC as per GSTR-2B | 15,045,623.38 | 2,425,525.54 | 138,630.86 | 138,630.86 | 2,702,787.26 |
| Final Difference | 0.00 | -0.00 | 0.00 | 0.00 | -0.00 |

### TOTAL (FY 2025-26)

| Particulars | Taxable | IGST | CGST | SGST/UTGST | Total GST |
|---|---|---|---|---|---|
| ITC as per Books | 163,223,641.86 | 25,368,267.15 | 1,913,216.60 | 1,913,216.60 | 29,194,700.35 |
| Add: ITC in GSTR-2B not booked in Books | 1,414,257.64 | 173,122.35 | 295,208.13 | 295,208.13 | 763,538.61 |
| Less: ITC booked in Books not in GSTR-2B | -10,990,529.02 | -474,554.85 | -33,142.49 | -33,142.49 | -540,839.83 |
| Add/Less: Timing difference - previous month ITC received in current month | 227,820.56 | 13,170.11 | 18,004.88 | 18,004.88 | 49,179.87 |
| Add/Less: Timing difference - current month ITC in subsequent month | -226,213.60 | -13,170.11 | -17,421.75 | -17,421.75 | -48,013.61 |
| Add/Less: Debit/Credit Note adjustments | 3,493.93 | -0.12 | 430.43 | 430.43 | 860.74 |
| Other differences | 615,199.51 | 0.58 | 0.18 | 0.18 | 0.94 |
| ITC as per GSTR-2B | 154,267,670.88 | 25,066,835.11 | 2,176,295.98 | 2,176,295.98 | 29,419,427.07 |
| Final Difference | -0.00 | -0.00 | 0.00 | 0.00 | -0.00 |

## 3. Section B - Invoice-wise Reconciliation (summary)

Complete invoice-wise listing (9,629 book rows + 139 unmatched GSTR-2B rows, with status, amounts, differences and remarks) is in the Excel workbook, sheet **B. Invoice Reconciliation**.

| Status | Rows |
|---|---|
| Matched | 7605 |
| Matched (taxable value differs - no GST impact) | 989 |
| Booked - no ITC claimed (no GST in books) | 664 |
| In Books, not in GSTR-2B | 73 |
| Matched - invoice date differs | 71 |
| DUPLICATE booking - Matched (taxable value differs - no GST impact) | 46 |
| ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | 45 |
| GST Mismatch | 40 |
| Timing Difference - ITC reflected in subsequent GSTR-2B | 34 |
| Matched (GST rounding diff <= Rs 1) | 27 |
| DUPLICATE booking - Matched | 11 |
| DUPLICATE booking - In Books, not in GSTR-2B | 10 |
| ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) [ITC Not Available] | 4 |
| Timing Difference - ITC reflected in earlier GSTR-2B | 4 |
| DUPLICATE booking - Timing Difference - ITC reflected in subsequent GSTR-2B | 3 |
| DUPLICATE booking - Matched (GST rounding diff <= Rs 1) | 2 |
| DUPLICATE booking - GST Mismatch | 1 |

Debit note register: 61 entries - Credit Note - Matched: 49; Debit/Credit Note - no GST amount in books: 3; In Books, not in GSTR-2B (Credit Note): 2; DUPLICATE booking - In Books, not in GSTR-2B (Credit Note): 2; Credit Note - Timing Difference (earlier GSTR-2B): 1; Credit Note - GST Mismatch: 1; DUPLICATE booking - Credit Note - Matched: 1; Credit Note - Matched (GSTIN differs): 1; DUPLICATE booking - Credit Note - Timing Difference (subsequent GSTR-2B): 1

## 4. Section C - Exception / Action Report

### 1. ITC booked in Books but NOT in GSTR-2B (supplier not reported / to follow up)  **(83 rows)**

| Branch | Book Date | Supplier | GSTIN | Invoice No | Book Month | Taxable | IGST | CGST | SGST | Total GST | Status | Remarks |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| HR | 2026-03-02 00:00:00 | Daimler India Commercial Veh.- Adblue and Oil Lub | 33AABCF1590N1ZJ | 1631356295 | 2026-03 | 1,883,320.00 | 338,997.60 | 0.00 | 0.00 | 338,997.60 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-28 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633190860 | 2025-04 | 135,457.44 | 24,382.34 | 0.00 | 0.00 | 24,382.34 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-11-25 00:00:00 | Daimler India Commercial Veh.- Adblue and Oil Lub | 33AABCF1590N1ZJ | 1633764135 | 2025-11 | 94,166.53 | 16,949.98 | 0.00 | 0.00 | 16,949.98 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-08-04 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633470328 | 2025-08 | 43,555.77 | 12,195.62 | 0.00 | 0.00 | 12,195.62 | In Books, not in GSTR-2B | To be verified |
| HR | 2026-03-01 00:00:00 | Urmi Equipments | 24ASRPS3707R1ZU |  | 2026-03 | 59,400.00 | 0.00 | 5,400.00 | 5,400.00 | 10,800.00 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2026-02-01 00:00:00 | Urmi Equipments | 24ASRPS3707R1ZU |  | 2026-02 | 59,400.00 | 0.00 | 5,400.00 | 5,400.00 | 10,800.00 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-01 00:00:00 | Troops11 Security Agency Private Limited | 24AALCT0561F1ZI | 423 | 2025-12 | 56,080.50 | 0.00 | 5,150.25 | 5,150.25 | 10,300.50 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-02-07 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633971996 | 2026-02 | 53,324.54 | 9,598.42 | 0.00 | 0.00 | 9,598.42 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-05-31 00:00:00 | Babaric Labour Service | 24AANFB7731Q1ZV | 47 | 2025-05 | 47,039.08 | 0.00 | 4,319.96 | 4,319.96 | 8,639.92 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-04-30 00:00:00 | Babaric Labour Service | 24AANFB7731Q1ZV | 17 | 2025-04 | 44,777.54 | 0.00 | 4,112.23 | 4,112.23 | 8,224.46 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-28 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633186807 | 2025-04 | 41,535.49 | 7,476.39 | 0.00 | 0.00 | 7,476.39 | In Books, not in GSTR-2B | To be verified |
| PL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633194851 | 2025-04 | 25,671.29 | 7,187.96 | 0.00 | 0.00 | 7,187.96 | In Books, not in GSTR-2B | To be verified |
| PL | 2025-07-01 00:00:00 | Csk Electronics and Automation Pvt Ltd | 33AAFCC7825B1Z1 | GST/25-26/14 | 2025-07 | 36,900.00 | 6,642.00 | 0.00 | 0.00 | 6,642.00 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-07-01 00:00:00 | Csk Electronics and Automation Pvt Ltd | 33AAFCC7825B1Z1 | GST/25-26/15 | 2025-07 | 36,900.00 | 6,642.00 | 0.00 | 0.00 | 6,642.00 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-09-02 00:00:00 | Akshar Auto Electricals | 24AFJPV2009J1ZD | JB-00033 | 2025-09 | 25,118.26 | 0.00 | 3,267.87 | 3,267.87 | 6,535.74 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-28 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633186799 | 2025-04 | 31,393.81 | 5,650.89 | 0.00 | 0.00 | 5,650.89 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-28 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633186802 | 2025-04 | 25,569.05 | 4,602.43 | 0.00 | 0.00 | 4,602.43 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-11-01 00:00:00 | Hare Krishna | 24AUOPP8762G1Z2 | GT/472/25-26 | 2025-11 | 22,000.00 | 0.00 | 1,980.00 | 1,980.00 | 3,960.00 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-02-28 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631354755 | 2026-02 | 18,246.64 | 3,284.40 | 0.00 | 0.00 | 3,284.40 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633191891 | 2025-04 | 10,392.18 | 2,909.81 | 0.00 | 0.00 | 2,909.81 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-04-09 00:00:00 | Kataria Motors Pvt Ltd -Rajkot(Pur) | 24AACCK0029Q1ZI | DDKTGJ112500001 | 2025-04 | 13,783.92 | 0.00 | 1,240.54 | 1,240.54 | 2,481.08 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-08-01 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633466246 | 2025-08 | 10,709.80 | 1,927.76 | 0.00 | 0.00 | 1,927.76 | In Books, not in GSTR-2B | To be verified |
| PL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633194527 | 2025-04 | 6,883.60 | 1,927.41 | 0.00 | 0.00 | 1,927.41 | In Books, not in GSTR-2B | To be verified |
| PL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633194183 | 2025-04 | 6,646.42 | 1,861.00 | 0.00 | 0.00 | 1,861.00 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-09-25 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1637701312 | 2025-09 | 9,040.76 | 1,848.75 | 0.00 | 0.00 | 1,848.75 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-09-06 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633559799 | 2025-09 | 9,040.76 | 1,848.75 | 0.00 | 0.00 | 1,848.75 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-09-17 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1637701286 | 2025-09 | 9,040.74 | 1,848.74 | 0.00 | 0.00 | 1,848.74 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-09-20 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633592519 | 2025-09 | 9,040.74 | 1,848.74 | 0.00 | 0.00 | 1,848.74 | In Books, not in GSTR-2B | To be verified |
| HR | 2026-03-01 00:00:00 | Right Angle Advert | 24ASWPR1164Q1ZT | 13 | 2026-03 | 10,200.00 | 0.00 | 918.00 | 918.00 | 1,836.00 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-28 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633186794 | 2025-04 | 9,923.05 | 1,786.15 | 0.00 | 0.00 | 1,786.15 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633194181 | 2025-04 | 6,891.22 | 1,721.21 | 0.00 | 0.00 | 1,721.21 | In Books, not in GSTR-2B | To be verified |
| PL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633193440 | 2025-04 | 8,595.96 | 1,547.27 | 0.00 | 0.00 | 1,547.27 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-28 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633187391 | 2025-04 | 5,257.80 | 1,472.18 | 0.00 | 0.00 | 1,472.18 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633192177 | 2025-04 | 5,202.36 | 1,456.66 | 0.00 | 0.00 | 1,456.66 | In Books, not in GSTR-2B | To be verified |
| PL | 2026-03-14 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631403348 | 2026-03 | 6,684.93 | 1,203.29 | 0.00 | 0.00 | 1,203.29 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-01 00:00:00 | Krishna Enterprise | 24ADOPP0118R1Z4 | 440 | 2025-12 | 6,000.00 | 0.00 | 540.00 | 540.00 | 1,080.00 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-28 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633186797 | 2025-04 | 5,447.92 | 980.63 | 0.00 | 0.00 | 980.63 | In Books, not in GSTR-2B | To be verified |
| PL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633194526 | 2025-04 | 3,441.80 | 963.70 | 0.00 | 0.00 | 963.70 | In Books, not in GSTR-2B | To be verified |
| PL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633194529 | 2025-04 | 5,345.67 | 962.22 | 0.00 | 0.00 | 962.22 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633193592 | 2025-04 | 4,298.00 | 773.64 | 0.00 | 0.00 | 773.64 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633192217 | 2025-04 | 3,946.00 | 710.28 | 0.00 | 0.00 | 710.28 | In Books, not in GSTR-2B | To be verified |
| PL | 2026-03-21 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631427846 | 2026-03 | 3,855.10 | 693.92 | 0.00 | 0.00 | 693.92 | In Books, not in GSTR-2B | To be verified |
| PL | 2025-05-22 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633254669 | 2025-05 | 3,743.90 | 673.90 | 0.00 | 0.00 | 673.90 | In Books, not in GSTR-2B | To be verified |
| PL | 2026-03-13 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631399892 | 2026-03 | 3,541.36 | 637.44 | 0.00 | 0.00 | 637.44 | In Books, not in GSTR-2B | To be verified |
| HR | 2026-03-24 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631437368 | 2026-03 | 3,319.26 | 597.47 | 0.00 | 0.00 | 597.47 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-09-15 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633579427 | 2025-09 | 2,738.47 | 519.88 | 0.00 | 0.00 | 519.88 | In Books, not in GSTR-2B | To be verified |
| PL | 2026-01-01 00:00:00 | Firoz Traders | 24ACDPB3131E1ZJ | 725 | 2026-01 | 2,584.74 | 0.00 | 232.63 | 232.63 | 465.26 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-10 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 124 | 2026-03 | 2,249.60 | 0.00 | 232.20 | 232.20 | 464.40 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-10 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 122 | 2026-03 | 2,249.60 | 0.00 | 232.20 | 232.20 | 464.40 | In Books, not in GSTR-2B | To be verified |
| PL | 2025-04-30 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633197116 | 2025-04 | 2,491.36 | 448.44 | 0.00 | 0.00 | 448.44 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-09-11 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/3861 | 2025-09 | 2,194.18 | 0.00 | 197.91 | 197.91 | 395.82 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633194319 | 2025-04 | 2,018.26 | 363.29 | 0.00 | 0.00 | 363.29 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-11 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 126 | 2026-03 | 1,500.40 | 0.00 | 154.80 | 154.80 | 309.60 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-10 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 123 | 2026-03 | 1,500.40 | 0.00 | 154.80 | 154.80 | 309.60 | In Books, not in GSTR-2B | To be verified |
| VL | 2025-04-28 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633186800 | 2025-04 | 956.75 | 267.89 | 0.00 | 0.00 | 267.89 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-02 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 1,312.76 | 0.00 | 118.62 | 118.62 | 237.24 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-20 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 1,312.76 | 0.00 | 118.62 | 118.62 | 237.24 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| PL | 2026-03-13 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631400641 | 2026-03 | 1,266.21 | 227.92 | 0.00 | 0.00 | 227.92 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-28 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 2476 | 2026-03 | 1,046.00 | 0.00 | 108.00 | 108.00 | 216.00 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-02-01 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 2340 | 2026-02 | 1,046.00 | 0.00 | 108.00 | 108.00 | 216.00 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-20 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 2453 | 2026-03 | 1,046.00 | 0.00 | 108.00 | 108.00 | 216.00 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-11-04 00:00:00 | Dineshbhai V. Tailor | 24AFHPT7182G2Z3 | 599 | 2025-11 | 1,149.00 | 0.00 | 103.50 | 103.50 | 207.00 | In Books, not in GSTR-2B | To be verified |
| PL | 2026-03-14 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631403466 | 2026-03 | 998.91 | 179.80 | 0.00 | 0.00 | 179.80 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-19 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 2452 | 2026-03 | 785.00 | 0.00 | 81.00 | 81.00 | 162.00 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-09-13 00:00:00 | Redbus [Gujarat] | 24AAHCP1178L1Z6 | RGJ25-A000993601 | 2025-09 | 3,240.00 | 0.00 | 81.00 | 81.00 | 162.00 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-16 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 2448 | 2026-03 | 749.20 | 0.00 | 77.40 | 77.40 | 154.80 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-31 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 2480 | 2026-03 | 749.20 | 0.00 | 77.40 | 77.40 | 154.80 | In Books, not in GSTR-2B | To be verified |
| PL | 2026-03-17 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631411070 | 2026-03 | 847.10 | 152.48 | 0.00 | 0.00 | 152.48 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-21 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 2461 | 2026-03 | 698.00 | 0.00 | 72.00 | 72.00 | 144.00 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-07-04 00:00:00 | Redbus [Gujarat] | 24AAHCP1178L1Z6 | RGJ25-A000546570 | 2025-07 | 2,700.00 | 0.00 | 67.50 | 67.50 | 135.00 | In Books, not in GSTR-2B | To be verified |
| PL | 2025-09-18 00:00:00 | Pramukh Auto | 24AVNPV2945D1ZC | 2661 | 2025-09 | 734.68 | 0.00 | 66.16 | 66.16 | 132.32 | In Books, not in GSTR-2B | To be verified |
| PL | 2025-09-11 00:00:00 | Pramukh Auto | 24AVNPV2945D1ZC | 2623 | 2025-09 | 734.68 | 0.00 | 66.16 | 66.16 | 132.32 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-09 00:00:00 | Diesel World (Pvt.) Ltd. | 24AAECD2472J1ZQ |  | 2025-12 | 592.40 | 0.00 | 64.80 | 64.80 | 129.60 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-02 00:00:00 | Diesel World (Pvt.) Ltd. | 24AAECD2472J1ZQ |  | 2025-12 | 592.40 | 0.00 | 64.80 | 64.80 | 129.60 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-04 00:00:00 | Diesel World (Pvt.) Ltd. | 24AAECD2472J1ZQ |  | 2025-12 | 592.40 | 0.00 | 64.80 | 64.80 | 129.60 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2026-03-17 00:00:00 | Diesel World (Pvt.) Ltd. | 24AAECD2472J1ZQ |  | 2026-03 | 592.40 | 0.00 | 64.80 | 64.80 | 129.60 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-07-05 00:00:00 | Redbus [Maharashtra] | 27AAHCP1178L1Z0 | RMH25-A002321799 | 2025-07 | 2,562.60 | 128.13 | 0.00 | 0.00 | 128.13 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-26 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 2469 | 2026-03 | 523.00 | 0.00 | 54.00 | 54.00 | 108.00 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-16 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 2447 | 2026-03 | 523.00 | 0.00 | 54.00 | 54.00 | 108.00 | In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-11 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 495.00 | 0.00 | 45.00 | 45.00 | 90.00 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-01 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 495.00 | 0.00 | 45.00 | 45.00 | 90.00 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2026-01-01 00:00:00 | Uniform Bucket (Jaf Enterprises Pvt Ltd) | 07AADCJ6419N1Z6 | JAF/25-26/4883 | 2026-01 | 1,006.00 | 86.18 | 0.00 | 0.00 | 86.18 | In Books, not in GSTR-2B | To be verified |
| VL | 2026-03-12 00:00:00 | Vishwahkarma Engineering Works | 24DNGPS4118R1ZD | 127 | 2026-03 | 130.00 | 0.00 | 13.50 | 13.50 | 27.00 | In Books, not in GSTR-2B | To be verified |

### 2. Booked in Books without GST split, and not present in GSTR-2B (no ITC claimed, not reported)  **(664 rows)**

| Branch | Book Date | Supplier | GSTIN | Invoice No | Book Month | Amount (no GST) | Status |
|---|---|---|---|---|---|---|---|
| HR | 2025-04-01 00:00:00 | Bobby Caman |  | 239 | 2025-04 | 15,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-02 00:00:00 | Gujarat Welding Works |  | 659 | 2025-04 | 1,200.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-03 00:00:00 | Overtech Business Solutions | 27ALBPP0585B1ZC | 002-04-01 | 2025-04 | 30,090.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-03 00:00:00 | Gayatri Computer and Printer |  | 7356 | 2025-04 | 250.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-05 00:00:00 | Gayatri Computer and Printer |  | 7360 | 2025-04 | 250.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-08 00:00:00 | Super Kisan Engineering Works |  | 487 | 2025-04 | 2,314.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-08 00:00:00 | Super Kisan Engineering Works |  | 488 | 2025-04 | 400.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-08 00:00:00 | Super Kisan Engineering Works |  | 490 | 2025-04 | 1,424.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-09 00:00:00 | Darpan Auto Glass |  | DAG1908 | 2025-04 | 2,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-10 00:00:00 | Gayatri Computer and Printer |  | 7374 | 2025-04 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-11 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI | DDKTGJ1B25000005 | 2025-04 | 3,831.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-14 00:00:00 | Prabal Motors Pvt Ltd (Purchase) | 27AAECP0166B1ZU | DDPVMH1K25000001 | 2025-04 | 2,646.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-14 00:00:00 | Super Kisan Engineering Works |  | 489 | 2025-04 | 1,424.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-14 00:00:00 | Super Kisan Engineering Works |  | 493 | 2025-04 | 356.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-14 00:00:00 | Super Kisan Engineering Works |  | 492 | 2025-04 | 890.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-15 00:00:00 | Super Kisan Engineering Works |  | 1801 | 2025-04 | 712.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-16 00:00:00 | Bal Krishan N. Gaddam |  | 831 | 2025-04 | 2,200.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-16 00:00:00 | Super Kisan Engineering Works |  | 1803 | 2025-04 | 1,068.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-18 00:00:00 | Kalptaru Super Market |  | 2050 | 2025-04 | 2,991.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-18 00:00:00 | Jatin M. Patel |  | 325 | 2025-04 | 1,854.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-19 00:00:00 | Spirited Motors Vehicle Limitd ( Pur)Raigadh Mumbai | 27ABBCS9610H1Z9 | DDERMH1A25000004 | 2025-04 | 34,295.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-22 00:00:00 | Patidar Steel & Traders |  | 1342 | 2025-04 | 16,450.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-22 00:00:00 | Chaturbhai Hirapara |  | 220425 | 2025-04 | 500,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-23 00:00:00 | Gayatri Computer and Printer |  | 7718 | 2025-04 | 250.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-25 00:00:00 | Motel Sunrise | 24AAFFM1071Q1Z3 | 104 | 2025-04 | 4,704.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-25 00:00:00 | Labheshwar Fabrication | 24BEQPV0376H1Z2 | GST /1/25-26 | 2025-04 | 396,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-25 00:00:00 | Super Kisan Engineering Works |  | 1804 | 2025-04 | 1,424.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-27 00:00:00 | Harkash Nrayab Singh |  | 301 | 2025-04 | 8,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-27 00:00:00 | Harkash Nrayab Singh |  | 302 | 2025-04 | 8,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-28 00:00:00 | Super Kisan Engineering Works |  | 1805 | 2025-04 | 712.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-29 00:00:00 | Super Kisan Engineering Works |  | 1807 | 2025-04 | 400.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-30 00:00:00 | Ajit Water Tanker | 24ARXPP4971F1Z6 | 4 | 2025-04 | 18,600.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-04-30 00:00:00 | Nirav and Company |  | 2025-26/9 | 2025-04 | 22,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-02 00:00:00 | Kataria Motors Pvt Ltd -Valsad(Pur) | 24AACCK0029Q1ZI | DDKTGJ1C25000004 | 2025-05 | 4,288.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-02 00:00:00 | Super Kisan Engineering Works |  | 1808 | 2025-05 | 712.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-06 00:00:00 | Super Kisan Engineering Works |  | 1809 | 2025-05 | 1,424.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-06 00:00:00 | Super Kisan Engineering Works |  | 1810 | 2025-05 | 1,424.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-07 00:00:00 | Super Kisan Engineering Works |  | 1811 | 2025-05 | 1,068.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-10 00:00:00 | Super Kisan Engineering Works |  | 1812 | 2025-05 | 2,714.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-14 00:00:00 | Gayatri Computer and Printer |  | 7768 | 2025-05 | 2,100.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-15 00:00:00 | Super Kisan Engineering Works |  | 1813 | 2025-05 | 890.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-18 00:00:00 | Super Kisan Engineering Works |  | 1816 | 2025-05 | 5,215.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-21 00:00:00 | Gayatri Computer and Printer |  | 7791 | 2025-05 | 250.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-21 00:00:00 | Super Kisan Engineering Works |  | 1815 | 2025-05 | 712.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-23 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI | DDKTGJ1B25000014 | 2025-05 | 8,087.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-26 00:00:00 | Super Kisan Engineering Works |  | 1817 | 2025-05 | 890.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-26 00:00:00 | Super Kisan Engineering Works |  | 1818 | 2025-05 | 712.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-27 00:00:00 | Super Kisan Engineering Works |  | 1819 | 2025-05 | 445.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-28 00:00:00 | Darpan Auto Glass |  | DAG2052 | 2025-05 | 1,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-29 00:00:00 | Gayatri Computer and Printer |  | 7912 | 2025-05 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-29 00:00:00 | Super Kisan Engineering Works |  | 1822 | 2025-05 | 11,125.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-29 00:00:00 | Super Kisan Engineering Works |  | 1820 | 2025-05 | 356.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-30 00:00:00 | Super Kisan Engineering Works |  | 1821 | 2025-05 | 1,424.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-31 00:00:00 | Nirav and Company |  | 2025-26/47 | 2025-05 | 22,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-05-31 00:00:00 | Ajit Water Tanker | 24ARXPP4971F1Z6 | 31 | 2025-05 | 25,200.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-02 00:00:00 | Giri Real Estate( Rajesh G. Goswami) |  | 1 | 2025-06 | 196,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-02 00:00:00 | Super Kisan Engineering Works |  | 1824 | 2025-06 | 445.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-02 00:00:00 | Super Kisan Engineering Works |  | 1825 | 2025-06 | 721.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-04 00:00:00 | Super Kisan Engineering Works |  | 1826 | 2025-06 | 801.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-04 00:00:00 | Gujarat Welding Works |  | 378 | 2025-06 | 400.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-04 00:00:00 | Super Kisan Engineering Works |  | 1827 | 2025-06 | 11,570.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-05 00:00:00 | Jay Khodiyar Enterprise [Pur] | 24AISPG2278G1Z4 | 1 | 2025-06 | 47,200.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-05 00:00:00 | Gujarat Welding Works |  | 1 | 2025-06 | 700.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-05 00:00:00 | Gujarat Welding Works |  | 2 | 2025-06 | 11,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-07 00:00:00 | Gujarat Welding Works |  | 3 | 2025-06 | 800.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-07 00:00:00 | Gujarat Welding Works |  | 4 | 2025-06 | 800.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-09 00:00:00 | Gujarat Welding Works |  | 5 | 2025-06 | 400.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-10 00:00:00 | Super Kisan Engineering Works |  | 1828 | 2025-06 | 2,314.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-11 00:00:00 | Gayatri Computer and Printer |  | 7919 | 2025-06 | 250.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-11 00:00:00 | Super Kisan Engineering Works |  | 1829 | 2025-06 | 2,136.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-11 00:00:00 | Super Kisan Engineering Works |  | 1830 | 2025-06 | 1,068.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-12 00:00:00 | Gujarat Welding Works |  | 7 | 2025-06 | 600.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-14 00:00:00 | Raghuvir Alang House |  | 968 | 2025-06 | 37,226.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-16 00:00:00 | Gujarat Welding Works |  | 9 | 2025-06 | 2,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-17 00:00:00 | Gurudev Gypsum & P.O.P. Traders | 24APZPP3911M1Z7 | 3859 | 2025-06 | 5,395.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-18 00:00:00 | Super Kisan Engineering Works |  | 1831 | 2025-06 | 890.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-18 00:00:00 | Super Kisan Engineering Works |  | 1832 | 2025-06 | 1,068.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-18 00:00:00 | Gujarat Welding Works |  | GWW-8-25-26 | 2025-06 | 11,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-19 00:00:00 | Gayatri Computer and Printer |  | 7977 | 2025-06 | 250.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-19 00:00:00 | Gayatri Computer and Printer |  | GCP/423/2025-26 | 2025-06 | 250.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-19 00:00:00 | Harkash Nrayab Singh |  | 168 | 2025-06 | 1,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-20 00:00:00 | Super Kisan Engineering Works |  |  | 2025-06 | 445.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-24 00:00:00 | Darpan Auto Glass |  | DAG2150 | 2025-06 | 2,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-25 00:00:00 | Gayatri Computer and Printer |  | GCP/441/2025-26 | 2025-06 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-27 00:00:00 | Super Kisan Engineering Works |  | 1836 | 2025-06 | 712.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-27 00:00:00 | Harkash Nrayab Singh |  | 167 | 2025-06 | 3,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-27 00:00:00 | Gujarat Welding Works |  | 10 | 2025-06 | 11,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-28 00:00:00 | Gayatri Computer and Printer |  | GCP/476/2025-26 | 2025-06 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-28 00:00:00 | Csk Electronics and Automation Pvt Ltd | 33AAFCC7825B1Z1 | 280625 | 2025-06 | 6,608.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-28 00:00:00 | Super Kisan Engineering Works |  |  | 2025-06 | 712.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-28 00:00:00 | Super Kisan Engineering Works |  |  | 2025-06 | 1,068.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-28 00:00:00 | Darpan Auto Glass |  | DAG2161 | 2025-06 | 2,950.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-30 00:00:00 | Ajit Water Tanker | 24ARXPP4971F1Z6 | 45 | 2025-06 | 28,200.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-30 00:00:00 | Radhey Techoserve |  | 16 | 2025-06 | 2,100.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-30 00:00:00 | Nirav and Company |  | 2025-26/49 | 2025-06 | 23,700.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-30 00:00:00 | Divyeshbhai M. Kamdi (Rent A/c) |  | DK-JUNE-2025 | 2025-06 | 2,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-06-30 00:00:00 | Super Kisan Engineering Works |  |  | 2025-06 | 445.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-01 00:00:00 | Radhey Techoserve |  | 18 | 2025-07 | 1,600.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-05 00:00:00 | Gujarat Welding Works |  | 11 | 2025-07 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-07 00:00:00 | Hotel Aarti Executive | 27AHKPB6123A2ZX |  | 2025-07 | 5,400.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-09 00:00:00 | Darpan Auto Glass |  |  | 2025-07 | 2,350.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-10 00:00:00 | Gayatri Computer and Printer |  |  | 2025-07 | 250.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-10 00:00:00 | Harkash Nrayab Singh |  | 123 | 2025-07 | 1,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-11 00:00:00 | Hotel President | 24AADFH7744R1ZS | 402 | 2025-07 | 10,920.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-11 00:00:00 | Gujarat Welding Works |  | 13 | 2025-07 | 10,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-11 00:00:00 | Gujarat Welding Works |  | 12 | 2025-07 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-12 00:00:00 | Super Kisan Engineering Works |  | 1844 | 2025-07 | 2,136.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-12 00:00:00 | Gujarat Welding Works |  | 14 | 2025-07 | 2,700.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-15 00:00:00 | Gayatri Computer and Printer |  |  | 2025-07 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-17 00:00:00 | Super Kisan Engineering Works |  | 1897 | 2025-07 | 890.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-17 00:00:00 | Super Kisan Engineering Works |  | 1896 | 2025-07 | 801.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-18 00:00:00 | D.G.V.C.L |  | HAZIRA JULY - 2025 | 2025-07 | 40,360.44 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-18 00:00:00 | Sharanam Hotels & Resorts Pvt. Ltd. |  | V1/FO/26/0001916 | 2025-07 | 11,202.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-18 00:00:00 | Gujarat Welding Works |  | 15 | 2025-07 | 450.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-22 00:00:00 | Gayatri Computer and Printer |  | 8876 | 2025-07 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-23 00:00:00 | Harkash Nrayab Singh |  | 124 | 2025-07 | 2,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-26 00:00:00 | Super Kisan Engineering Works |  | 1846 | 2025-07 | 1,246.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-31 00:00:00 | Nirav and Company |  | 2025-26/53 | 2025-07 | 22,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-31 00:00:00 | Ajit Water Tanker | 24ARXPP4971F1Z6 | 63 | 2025-07 | 20,400.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-31 00:00:00 | Mark Express Pvt. Ltd. (Courier Exp) |  | 31.07.2025 | 2025-07 | 2,715.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-31 00:00:00 | Mayur Parmar (Milk ) |  | 36/25-26 | 2025-07 | 6,820.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-07-31 00:00:00 | Divyeshbhai M. Kamdi (Rent A/c) |  | DK-JULY-2025 | 2025-07 | 2,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-01 00:00:00 | Gayatri Computer and Printer |  | 8903 | 2025-08 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-01 00:00:00 | Soul Communication Private Limited |  | 1082025 | 2025-08 | 49,900.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-01 00:00:00 | Gujarat Welding Works |  | 19 | 2025-08 | 2,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-01 00:00:00 | Gujarat Welding Works |  | 20 | 2025-08 | 2,400.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-07 00:00:00 | Gujarat Welding Works |  | 21 | 2025-08 | 800.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-08 00:00:00 | Gujarat Welding Works |  | 22 | 2025-08 | 400.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-08 00:00:00 | Gujarat Welding Works |  | 23 | 2025-08 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-14 00:00:00 | Gujarat Welding Works |  | 25 | 2025-08 | 800.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-14 00:00:00 | Gujarat Welding Works |  | 26 | 2025-08 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-18 00:00:00 | D.G.V.C.L |  | HAZIRA AUG - 2025 | 2025-08 | 40,513.80 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-18 00:00:00 | Gujarat Welding Works |  | 27 | 2025-08 | 800.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-19 00:00:00 | Gayatri Computer and Printer |  |  | 2025-08 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-19 00:00:00 | Super Kisan Engineering Works |  | 1850 | 2025-08 | 2,314.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-19 00:00:00 | Gujarat Welding Works |  | 28 | 2025-08 | 4,800.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-20 00:00:00 | Super Kisan Engineering Works |  | 1852 | 2025-08 | 495.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-20 00:00:00 | Darpan Auto Glass |  | DAG2395 | 2025-08 | 2,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-23 00:00:00 | Super Kisan Engineering Works |  | 1853 | 2025-08 | 1,068.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-26 00:00:00 | Super Kisan Engineering Works |  | 1854 | 2025-08 | 1,068.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-26 00:00:00 | Super Kisan Engineering Works |  | 1855 | 2025-08 | 712.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-26 00:00:00 | Darpan Auto Glass |  | DAG2373 | 2025-08 | 2,950.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-27 00:00:00 | Gujarat Welding Works |  | 29 | 2025-08 | 11,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-27 00:00:00 | Darpan Auto Glass |  | DAG2421 | 2025-08 | 2,950.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-29 00:00:00 | Om Sai Crane Serivce |  | 3016 | 2025-08 | 700.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-29 00:00:00 | Super Kisan Engineering Works |  | 1858 | 2025-08 | 1,424.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-30 00:00:00 | Om Sai Crane Serivce |  | 3209 | 2025-08 | 4,200.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-31 00:00:00 | Ajit Water Tanker | 24ARXPP4971F1Z6 | 76 | 2025-08 | 20,400.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-31 00:00:00 | Nirav and Company |  | 2025-26/59 | 2025-08 | 22,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-31 00:00:00 | Mayur Parmar (Milk ) |  | 60/25-26 | 2025-08 | 5,904.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-31 00:00:00 | Mark Express Pvt. Ltd. (Courier Exp) |  | 31.08.2025 | 2025-08 | 825.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-31 00:00:00 | Bikram Kariyana Store (Rent A/c) |  | BR-AUG-2025 | 2025-08 | 15,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-31 00:00:00 | Dimpleben Dinkarbhai Patel |  | 5 | 2025-08 | 8,800.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-08-31 00:00:00 | Divyeshbhai M. Kamdi (Rent A/c) |  | DK-AUG-2025 | 2025-08 | 2,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-02 00:00:00 | Super Kisan Engineering Works |  | 1847 | 2025-09 | 890.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-04 00:00:00 | Super Kisan Engineering Works |  | 1848 | 2025-09 | 445.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-04 00:00:00 | Super Kisan Engineering Works |  | 1859 | 2025-09 | 712.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-04 00:00:00 | Super Kisan Engineering Works |  | 1860 | 2025-09 | 2,136.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-08 00:00:00 | Super Kisan Engineering Works |  | 1861 | 2025-09 | 445.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-08 00:00:00 | Super Kisan Engineering Works |  | 1862 | 2025-09 | 178.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-08 00:00:00 | Super Kisan Engineering Works |  | 1864 | 2025-09 | 4,272.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-09 00:00:00 | Radhey Techoserve |  | 29 | 2025-09 | 1,700.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-09 00:00:00 | Gujarat Welding Works |  | 30 | 2025-09 | 800.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-09 00:00:00 | Darpan Auto Glass |  | DAG2465 | 2025-09 | 1,900.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-10 00:00:00 | Harkash Nrayab Singh |  | 307 | 2025-09 | 1,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-10 00:00:00 | Harkash Nrayab Singh |  | 308 | 2025-09 | 1,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-13 00:00:00 | Harkash Nrayab Singh |  | 310 | 2025-09 | 1,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-13 00:00:00 | Harkash Nrayab Singh |  | 315 | 2025-09 | 1,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-14 00:00:00 | Harkash Nrayab Singh |  | 311 | 2025-09 | 3,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-15 00:00:00 | Gayatri Computer and Printer |  | 8623 | 2025-09 | 750.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-16 00:00:00 | Darpan Auto Glass |  | DAG2490 | 2025-09 | 2,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-17 00:00:00 | D.G.V.C.L |  | 730 | 2025-09 | 36,664.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-18 00:00:00 | Labheshwar Fabrication | 24BEQPV0376H1Z2 | GST/4/25-26 | 2025-09 | 184,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-18 00:00:00 | Super Kisan Engineering Works |  | 1865 | 2025-09 | 979.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-22 00:00:00 | Gayatri Computer and Printer |  |  | 2025-09 | 250.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-22 00:00:00 | Super Kisan Engineering Works |  | 1866 | 2025-09 | 890.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-22 00:00:00 | Super Kisan Engineering Works |  | 1867 | 2025-09 | 712.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-24 00:00:00 | Gayatri Computer and Printer |  | GCP-01 | 2025-09 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-24 00:00:00 | Super Kisan Engineering Works |  | 1868 | 2025-09 | 445.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-25 00:00:00 | Super Kisan Engineering Works |  | 1869 | 2025-09 | 1,424.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-25 00:00:00 | Harkash Nrayab Singh |  | 314 | 2025-09 | 7,500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-30 00:00:00 | Ajit Water Tanker | 24ARXPP4971F1Z6 | 97 | 2025-09 | 19,200.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-30 00:00:00 | Nirav and Company |  | 2025-26/60 | 2025-09 | 28,800.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-30 00:00:00 | Mark Express Pvt. Ltd. (Courier Exp) |  | 30.09.2025 | 2025-09 | 2,370.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-30 00:00:00 | Sureshlal Harilal Kalal (Rent A/c) |  | SK-SEP-2025 | 2025-09 | 7,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-30 00:00:00 | Bikram Kariyana Store (Rent A/c) |  | BR-SEP-2025 | 2025-09 | 15,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-30 00:00:00 | Divyeshbhai M. Kamdi (Rent A/c) |  | DK-SEP-2025 | 2025-09 | 2,000.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-09-30 00:00:00 | Harkash Nrayab Singh |  | 312 | 2025-09 | 1,300.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-01 00:00:00 | Makadia Traders | 24AAEFM0858N1Z1 | 2661 | 2025-10 | 1,700.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-01 00:00:00 | Super Kisan Engineering Works |  | 1871 | 2025-10 | 1,068.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-01 00:00:00 | Super Kisan Engineering Works |  | 1857 | 2025-10 | 356.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-06 00:00:00 | Super Kisan Engineering Works |  | 1872 | 2025-10 | 890.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-06 00:00:00 | Super Kisan Engineering Works |  | 1873 | 2025-10 | 1,424.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-07 00:00:00 | Gayatri Computer and Printer |  | 8691 | 2025-10 | 500.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-07 00:00:00 | Om Sai Crane Serivce |  |  | 2025-10 | 700.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-09 00:00:00 | Airtel Ltd. |  | AIR-SEP-2025 | 2025-10 | 1,178.82 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-10 00:00:00 | Super Kisan Engineering Works |  | 1874 | 2025-10 | 2,670.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-11 00:00:00 | Super Kisan Engineering Works |  | 1875 | 2025-10 | 1,424.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-11 00:00:00 | Super Kisan Engineering Works |  | 1876 | 2025-10 | 712.00 | Booked - no ITC claimed (no GST in books) |
| HR | 2025-10-14 00:00:00 | Darpan Auto Glass |  | DAG2605 | 2025-10 | 2,500.00 | Booked - no ITC claimed (no GST in books) |

### 3. ITC in GSTR-2B but NOT booked in Books (incl. RCM / Rejected on IMS)  **(138 rows)**

| Supplier | GSTIN | Invoice No | Inv Date | Taxable | IGST | CGST | SGST | Total GST | 2B Month | Category |
|---|---|---|---|---|---|---|---|---|---|---|
| LABHESHWAR FABRICATION | 24BEQPV0376H1Z2 | GST/1/25-26 | 2025-04-25 00:00:00 | 338,983.43 | 0.00 | 30,508.51 | 30,508.51 | 61,017.02 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LIMITED | 33AABCF1590N1ZJ | 3325831988 | 2026-03-02 00:00:00 | 188,320.00 | 33,897.60 | 0.00 | 0.00 | 33,897.60 | 2026-03 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009181 | 2025-04-28 00:00:00 | 135,457.44 | 24,382.34 | 0.00 | 0.00 | 24,382.34 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LIMITED | 33AABCF1590N1ZJ | 3325575822 | 2025-11-25 00:00:00 | 94,160.00 | 16,948.80 | 0.00 | 0.00 | 16,948.80 | 2025-11 | In GSTR-2B, not in Books |
| CSK ELECTRONICS AND AUTOMATION PRIVATE LIMITED | 33AAFCC7825B1Z1 | GST/25-26/14 | 2025-07-05 00:00:00 | 73,800.00 | 13,284.00 | 0.00 | 0.00 | 13,284.00 | 2025-07 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025054146 | 2025-08-04 00:00:00 | 43,555.77 | 12,195.61 | 0.00 | 0.00 | 12,195.61 | 2025-08 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/343 | 2025-06-17 00:00:00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/350 | 2025-06-18 00:00:00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/341 | 2025-06-17 00:00:00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/376 | 2025-06-22 00:00:00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/367 | 2025-06-21 00:00:00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/377 | 2025-06-22 00:00:00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/374 | 2025-06-22 00:00:00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/356 | 2025-06-19 00:00:00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/361 | 2025-06-20 00:00:00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/364 | 2025-06-20 00:00:00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | 2025-06 | In GSTR-2B, not in Books |
| TROOPS11 SECURITY AGENCY PRIVATE LIMITED | 24AALCT0561F1ZI | 423 | 2025-12-01 00:00:00 | 65,100.00 | 0.00 | 5,859.00 | 5,859.00 | 11,718.00 | 2025-12 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/390 | 2025-06-24 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/355 | 2025-06-18 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/386 | 2025-06-24 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/391 | 2025-06-25 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/381 | 2025-06-22 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/380 | 2025-06-22 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/379 | 2025-06-22 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/378 | 2025-06-22 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/385 | 2025-06-24 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/353 | 2025-06-18 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/375 | 2025-06-22 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/370 | 2025-06-21 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/373 | 2025-06-21 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/365 | 2025-06-20 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/354 | 2025-06-18 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/362 | 2025-06-20 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/363 | 2025-06-20 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/342 | 2025-06-17 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/383 | 2025-06-23 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/382 | 2025-06-23 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/384 | 2025-06-23 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/344 | 2025-06-17 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/351 | 2025-06-18 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/392 | 2025-06-25 00:00:00 | 62,625.00 | 0.00 | 5,636.25 | 5,636.25 | 11,272.50 | 2025-06 | In GSTR-2B, not in Books |
| URMI EQUIPMENTS | 24ASRPS3707R1ZU | UE/2025-26/442 | 2026-02-01 00:00:00 | 60,000.00 | 0.00 | 5,400.00 | 5,400.00 | 10,800.00 | 2026-02 | In GSTR-2B, not in Books |
| URMI EQUIPMENTS | 24ASRPS3707R1ZU | UE/2025-26/478 | 2026-03-01 00:00:00 | 60,000.00 | 0.00 | 5,400.00 | 5,400.00 | 10,800.00 | 2026-03 | In GSTR-2B, not in Books |
| TROOPS11 SECURITY AGENCY PRIVATE LIMITED | 24AALCT0561F1ZI | 593 | 2026-03-01 00:00:00 | 58,500.00 | 0.00 | 5,265.00 | 5,265.00 | 10,530.00 | 2026-03 | In GSTR-2B (Rejected on IMS), not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/372 | 2025-06-21 00:00:00 | 58,450.00 | 0.00 | 5,260.50 | 5,260.50 | 10,521.00 | 2025-06 | In GSTR-2B, not in Books |
| TROOPS11 SECURITY AGENCY PRIVATE LIMITED | 24AALCT0561F1ZI | 594 | 2026-03-01 00:00:00 | 57,637.00 | 0.00 | 5,187.33 | 5,187.33 | 10,374.66 | 2026-03 | In GSTR-2B (Rejected on IMS), not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025133885 | 2026-02-07 00:00:00 | 53,324.54 | 9,598.41 | 0.00 | 0.00 | 9,598.41 | 2026-02 | In GSTR-2B, not in Books |
| BABARIC LABOUR SERVICE | 24AANFB7731Q1ZV | 47 | 2025-05-31 00:00:00 | 52,429.00 | 0.00 | 4,718.61 | 4,718.61 | 9,437.22 | 2025-05 | In GSTR-2B, not in Books |
| BABARIC LABOUR SERVICE | 24AANFB7731Q1ZV | 17 | 2025-04-30 00:00:00 | 47,537.59 | 0.00 | 4,278.38 | 4,278.38 | 8,556.76 | 2025-04 | In GSTR-2B, not in Books |
| TROOPS11 SECURITY AGENCY PRIVATE LIMITED | 24AALCT0561F1ZI | 595 | 2026-03-01 00:00:00 | 46,556.00 | 0.00 | 4,190.04 | 4,190.04 | 8,380.08 | 2026-03 | In GSTR-2B (Rejected on IMS), not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025008881 | 2025-04-28 00:00:00 | 41,535.49 | 7,476.39 | 0.00 | 0.00 | 7,476.39 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009799 | 2025-04-29 00:00:00 | 25,671.29 | 7,187.96 | 0.00 | 0.00 | 7,187.96 | 2025-04 | In GSTR-2B, not in Books |
| AKSHAR AUTO ELECTRICALS | 24AFJPV2009J1ZD | JB-00033 | 2025-09-02 00:00:00 | 23,498.00 | 0.00 | 3,122.07 | 3,122.07 | 6,244.14 | 2025-09 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025008875 | 2025-04-28 00:00:00 | 31,393.82 | 5,650.88 | 0.00 | 0.00 | 5,650.88 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025008878 | 2025-04-28 00:00:00 | 25,569.05 | 4,602.43 | 0.00 | 0.00 | 4,602.43 | 2025-04 | In GSTR-2B, not in Books |
| SPIRITED MOTOR VEHICLES LIMITED | 27ABBCS9610H1Z9 | DDERMH1E25000011 | 2025-05-02 00:00:00 | 20,008.45 | 3,601.52 | 0.00 | 0.00 | 3,601.52 | 2025-09 | In GSTR-2B (Rejected on IMS), not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025149294 | 2026-02-28 00:00:00 | 18,246.64 | 3,284.40 | 0.00 | 0.00 | 3,284.40 | 2026-02 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009269 | 2025-04-29 00:00:00 | 10,392.18 | 2,909.81 | 0.00 | 0.00 | 2,909.81 | 2025-04 | In GSTR-2B, not in Books |
| FORCE 11 SECURITY AND ALLIED SERVICES | 24BLCPS2141E1ZJ | 1038 | 2025-11-08 00:00:00 | 16,000.00 | 0.00 | 1,440.00 | 1,440.00 | 2,880.00 | 2025-11 | In GSTR-2B (RCM), not in Books |
| KIRAN HARDWARE AND SAFETY | 24BWPPP0170A1ZU | KHS-004991 | 2025-12-27 00:00:00 | 15,600.00 | 0.00 | 1,404.00 | 1,404.00 | 2,808.00 | 2025-12 | In GSTR-2B (Rejected on IMS), not in Books |
| CAMAL SECURITY AND ALLIED SERVICES PRIVATE LIMITED | 24AALCC4280G1ZP | PV/25-26/68 | 2025-06-01 00:00:00 | 11,871.00 | 0.00 | 1,068.39 | 1,068.39 | 2,136.78 | 2025-06 | In GSTR-2B, not in Books |
| VIRAL ENTERPRISE | 24AALFV6612Q3ZH | 25-26/L-0311 | 2025-06-19 00:00:00 | 11,840.00 | 0.00 | 1,065.60 | 1,065.60 | 2,131.20 | 2025-06 | In GSTR-2B, not in Books |
| TATA AIG GENERAL INSURANCE COMPANY LIMITED | 24AABCT3518Q1Z2 | GU25I00000147788 | 2025-05-13 00:00:00 | 10,909.00 | 0.00 | 981.00 | 981.00 | 1,962.00 | 2025-05 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025053537 | 2025-08-01 00:00:00 | 10,709.80 | 1,927.77 | 0.00 | 0.00 | 1,927.77 | 2025-08 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009775 | 2025-04-29 00:00:00 | 6,883.60 | 1,927.41 | 0.00 | 0.00 | 1,927.41 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009701 | 2025-04-29 00:00:00 | 6,646.42 | 1,861.00 | 0.00 | 0.00 | 1,861.00 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025008870 | 2025-04-28 00:00:00 | 9,923.05 | 1,786.15 | 0.00 | 0.00 | 1,786.15 | 2025-04 | In GSTR-2B, not in Books |
| KATARIA MOTORS PRIVATE LIMITED | 24AACCK0029Q1ZI | DDKTGJ1B25000058 | 2025-09-20 00:00:00 | 6,318.11 | 0.00 | 884.54 | 884.54 | 1,769.08 | 2025-09 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009699 | 2025-04-29 00:00:00 | 6,891.18 | 1,721.22 | 0.00 | 0.00 | 1,721.22 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009494 | 2025-04-29 00:00:00 | 8,595.96 | 1,547.27 | 0.00 | 0.00 | 1,547.27 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025008954 | 2025-04-28 00:00:00 | 5,257.80 | 1,472.18 | 0.00 | 0.00 | 1,472.18 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009441 | 2025-04-29 00:00:00 | 5,202.36 | 1,456.66 | 0.00 | 0.00 | 1,456.66 | 2025-04 | In GSTR-2B, not in Books |
| VIRAL ENTERPRISE | 24AALFV6612Q3ZH | 25-26/0673 | 2025-05-08 00:00:00 | 8,066.77 | 0.00 | 726.01 | 726.01 | 1,452.02 | 2025-05 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LIMITED | 27AABCF1590N2ZB | 2725000391 | 2026-03-14 00:00:00 | 6,684.93 | 1,203.29 | 0.00 | 0.00 | 1,203.29 | 2026-03 | In GSTR-2B, not in Books |
| SHARANAM HOTEL AND RESORTS PVT LTD | 27AAACG6153H1ZO | V1/FO/26/0001916 | 2025-07-18 00:00:00 | 10,002.18 | 0.00 | 600.13 | 600.13 | 1,200.26 | 2025-07 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025008873 | 2025-04-28 00:00:00 | 5,447.93 | 980.62 | 0.00 | 0.00 | 980.62 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009774 | 2025-04-29 00:00:00 | 3,441.80 | 963.70 | 0.00 | 0.00 | 963.70 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009777 | 2025-04-29 00:00:00 | 5,345.67 | 962.22 | 0.00 | 0.00 | 962.22 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LIMITED | 33AABCF1590N1ZJ | 3325505911 | 2025-10-30 00:00:00 | 4,778.52 | 860.13 | 0.00 | 0.00 | 860.13 | 2025-10 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009573 | 2025-04-29 00:00:00 | 4,298.00 | 773.64 | 0.00 | 0.00 | 773.64 | 2025-04 | In GSTR-2B, not in Books |
| PARTH CEMENT ARTICLES | 24AAJFP9360P1ZI | B2B/271 | 2025-06-03 00:00:00 | 4,000.00 | 0.00 | 360.00 | 360.00 | 720.00 | 2025-06 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009461 | 2025-04-29 00:00:00 | 3,946.00 | 710.28 | 0.00 | 0.00 | 710.28 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025158326 | 2026-03-21 00:00:00 | 3,855.10 | 693.92 | 0.00 | 0.00 | 693.92 | 2026-03 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025021020 | 2025-05-22 00:00:00 | 3,743.90 | 673.90 | 0.00 | 0.00 | 673.90 | 2025-05 | In GSTR-2B, not in Books |
| VIRAL ENTERPRISE | 24AALFV6612Q3ZH | 25-26/0454 | 2025-04-24 00:00:00 | 3,563.28 | 0.00 | 320.70 | 320.70 | 641.40 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LIMITED | 27AABCF1590N2ZB | 2725000279 | 2026-03-13 00:00:00 | 3,541.36 | 637.44 | 0.00 | 0.00 | 637.44 | 2026-03 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025160776 | 2026-03-24 00:00:00 | 3,319.26 | 597.47 | 0.00 | 0.00 | 597.47 | 2026-03 | In GSTR-2B, not in Books |
| KATARIA MOTORS PRIVATE LIMITED | 24AACCK0029Q1ZI | DDKTGJ1B25000055 | 2025-09-20 00:00:00 | 3,246.86 | 0.00 | 292.22 | 292.22 | 584.44 | 2025-09 | In GSTR-2B, not in Books |
| SHREE GANESH AUTO | 24AHOPC4366Q1ZT | 25-26/3861 | 2025-09-11 00:00:00 | 2,199.22 | 0.00 | 282.89 | 282.89 | 565.78 | 2025-09 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LIMITED | 33AABCF1590N1ZJ | 3325645702 | 2025-12-19 00:00:00 | 2,888.31 | 519.89 | 0.00 | 0.00 | 519.89 | 2025-12 | In GSTR-2B, not in Books |
| Motel Sunrise | 24AAFFM1071Q1Z3 | RR-104 | 2025-04-25 00:00:00 | 4,200.00 | 0.00 | 252.00 | 252.00 | 504.00 | 2025-06 | In GSTR-2B, not in Books |
| Maruti Auto Enterprise | 24AASPC6705G1ZT | A3920 | 2025-10-17 00:00:00 | 2,498.44 | 0.00 | 224.86 | 224.86 | 449.72 | 2025-10 | In GSTR-2B (Rejected on IMS), not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025010072 | 2025-04-30 00:00:00 | 2,491.36 | 448.44 | 0.00 | 0.00 | 448.44 | 2025-04 | In GSTR-2B, not in Books |
| HOTEL STONARC | 27AIFPK2609C2ZN | 670 | 2025-09-16 00:00:00 | 3,400.00 | 0.00 | 204.00 | 204.00 | 408.00 | 2025-09 | In GSTR-2B, not in Books |
| HOTEL STONARC | 27AIFPK2609C2ZN | 665 | 2025-09-15 00:00:00 | 3,400.00 | 0.00 | 204.00 | 204.00 | 408.00 | 2025-09 | In GSTR-2B, not in Books |
| Air India Express Limited | 29AABCA0522B1ZG | IBLR260200214612 | 2026-02-18 00:00:00 | 7,348.58 | 367.42 | 0.00 | 0.00 | 367.42 | 2026-02 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025009756 | 2025-04-29 00:00:00 | 2,018.26 | 363.29 | 0.00 | 0.00 | 363.29 | 2025-04 | In GSTR-2B, not in Books |
| CLEARTRIP PRIVATE LIMITED | 27AACCC6016B1Z8 | I/MH/43151953 | 2026-02-18 00:00:00 | 1,885.59 | 339.41 | 0.00 | 0.00 | 339.41 | 2026-02 | In GSTR-2B, not in Books |
| AKSHAR AUTO ELECTRICALS | 24AFJPV2009J1ZD | LB-00033 | 2025-09-02 00:00:00 | 1,620.00 | 0.00 | 145.80 | 145.80 | 291.60 | 2025-09 | In GSTR-2B, not in Books |
| INTERGLOBE AVIATION LIMITED | 33AABCI2726B1Z9 | TN1252602CL64504 | 2026-02-19 00:00:00 | 5,627.00 | 282.00 | 0.00 | 0.00 | 282.00 | 2026-02 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LTD | 06AABCF1590N1ZG | 6025008876 | 2025-04-28 00:00:00 | 956.75 | 267.89 | 0.00 | 0.00 | 267.89 | 2025-04 | In GSTR-2B, not in Books |
| HOTEL URBAN VISTA CHAKAN | 27AUDPN1919Q1Z3 | 45/147/2025-2026 | 2026-03-11 00:00:00 | 5,100.00 | 255.00 | 0.00 | 0.00 | 255.00 | 2026-03 | In GSTR-2B (Rejected on IMS), not in Books |
| HOTEL URBAN VISTA CHAKAN | 27AUDPN1919Q1Z3 | 14/57/2025-2026 | 2026-02-06 00:00:00 | 5,100.00 | 255.00 | 0.00 | 0.00 | 255.00 | 2026-02 | In GSTR-2B, not in Books |
| HOTEL URBAN VISTA CHAKAN | 27AUDPN1919Q1Z3 | 15/58/2025-2026 | 2026-02-06 00:00:00 | 5,100.00 | 255.00 | 0.00 | 0.00 | 255.00 | 2026-02 | In GSTR-2B, not in Books |
| HOTEL URBAN VISTA CHAKAN | 27AUDPN1919Q1Z3 | 42/141/2025-2026 | 2026-03-06 00:00:00 | 5,100.00 | 255.00 | 0.00 | 0.00 | 255.00 | 2026-03 | In GSTR-2B (Rejected on IMS), not in Books |
| HOTEL URBAN VISTA CHAKAN | 27AUDPN1919Q1Z3 | 44/146/2025-2026 | 2026-03-11 00:00:00 | 5,100.00 | 255.00 | 0.00 | 0.00 | 255.00 | 2026-03 | In GSTR-2B (Rejected on IMS), not in Books |
| HOTEL URBAN VISTA CHAKAN | 27AUDPN1919Q1Z3 | 41/140/2025-2026 | 2026-03-06 00:00:00 | 5,100.00 | 255.00 | 0.00 | 0.00 | 255.00 | 2026-03 | In GSTR-2B (Rejected on IMS), not in Books |
| NEST CORPORATE HOMES | 33AIBPD8987G1Z5 | 1531 | 2026-03-31 00:00:00 | 4,830.00 | 0.00 | 120.75 | 120.75 | 241.50 | 2026-03 | In GSTR-2B, not in Books |
| SHREE GANESH AUTO | 24AHOPC4366Q1ZT | 25-26/5902 | 2025-12-02 00:00:00 | 1,317.80 | 0.00 | 118.60 | 118.60 | 237.20 | 2025-12 | In GSTR-2B, not in Books |
| SHREE GANESH AUTO | 24AHOPC4366Q1ZT | 25-26/6446 | 2025-12-20 00:00:00 | 1,317.80 | 0.00 | 118.60 | 118.60 | 237.20 | 2025-12 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LIMITED | 27AABCF1590N2ZB | 2725000317 | 2026-03-13 00:00:00 | 1,266.21 | 227.92 | 0.00 | 0.00 | 227.92 | 2026-03 | In GSTR-2B, not in Books |
| Maruti Auto Enterprise | 24AASPC6705G1ZT | CN54 | 2026-03-31 00:00:00 | 1,225.00 | 0.00 | 110.25 | 110.25 | 220.50 | 2026-03 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LIMITED | 27AABCF1590N2ZB | 2725000412 | 2026-03-14 00:00:00 | 998.91 | 179.80 | 0.00 | 0.00 | 179.80 | 2026-03 | In GSTR-2B, not in Books |
| KATARIA AUTOMOBILES PRIVATE LTD | 24AAACK6221C1Z7 | 30/BR/25002057 | 2025-05-30 00:00:00 | 954.80 | 0.00 | 87.59 | 87.59 | 175.18 | 2025-05 | In GSTR-2B, not in Books |
| PRASAD ENTERPRISES | 27ACPPN8170R1ZH | 169/819/2025-26 | 2025-11-30 00:00:00 | 3,428.80 | 0.00 | 85.72 | 85.72 | 171.44 | 2025-11 | In GSTR-2B, not in Books |
| HOTEL URBAN VISTA CHAKAN | 27AUDPN1919Q1Z3 | 58/163/2025-2026 | 2026-03-13 00:00:00 | 3,400.00 | 170.00 | 0.00 | 0.00 | 170.00 | 2026-03 | In GSTR-2B (Rejected on IMS), not in Books |
| HOTEL URBAN VISTA CHAKAN | 27AUDPN1919Q1Z3 | 57/162/2025-2026 | 2026-03-13 00:00:00 | 3,400.00 | 170.00 | 0.00 | 0.00 | 170.00 | 2026-03 | In GSTR-2B (Rejected on IMS), not in Books |
| SHIV SHAKTI MOTORS | 24AKWPP5497L1Z1 | GSTT/128 | 2025-04-03 00:00:00 | 847.44 | 0.00 | 76.27 | 76.27 | 152.54 | 2025-04 | In GSTR-2B, not in Books |
| DAIMLER INDIA COMMERCIAL VEHICLES PRIVATE LIMITED | 27AABCF1590N2ZB | 2725000590 | 2026-03-17 00:00:00 | 847.10 | 152.48 | 0.00 | 0.00 | 152.48 | 2026-03 | In GSTR-2B, not in Books |
| DHYANI ENTERPRISE | 24ARSPP0690G2ZG | 454/25-26 | 2025-05-29 00:00:00 | 806.40 | 0.00 | 72.58 | 72.58 | 145.16 | 2025-05 | In GSTR-2B, not in Books |
| PRAMUKH AUTO | 24AVNPV2945D1ZC | 2661 | 2025-09-18 00:00:00 | 754.24 | 0.00 | 67.88 | 67.88 | 135.76 | 2025-09 | In GSTR-2B, not in Books |
| PRAMUKH AUTO | 24AVNPV2945D1ZC | 2623 | 2025-09-11 00:00:00 | 754.24 | 0.00 | 67.88 | 67.88 | 135.76 | 2025-09 | In GSTR-2B, not in Books |
| DIESEL WORD PVT LTD | 24AAECD2472J1ZQ | JB-00654/25-26 | 2025-12-02 00:00:00 | 720.00 | 0.00 | 64.80 | 64.80 | 129.60 | 2025-12 | In GSTR-2B, not in Books |
| DIESEL WORD PVT LTD | 24AAECD2472J1ZQ | JB-00996/25-26 | 2026-03-17 00:00:00 | 720.00 | 0.00 | 64.80 | 64.80 | 129.60 | 2026-03 | In GSTR-2B, not in Books |
| DIESEL WORD PVT LTD | 24AAECD2472J1ZQ | JB-00662/25-26 | 2025-12-04 00:00:00 | 720.00 | 0.00 | 64.80 | 64.80 | 129.60 | 2025-12 | In GSTR-2B, not in Books |
| DIESEL WORD PVT LTD | 24AAECD2472J1ZQ | JB-00679/25-26 | 2025-12-09 00:00:00 | 720.00 | 0.00 | 64.80 | 64.80 | 129.60 | 2025-12 | In GSTR-2B, not in Books |
| Air India Express Limited | 29AABCA0522B1ZG | IBLR260200212147 | 2026-02-18 00:00:00 | 2,583.81 | 129.19 | 0.00 | 0.00 | 129.19 | 2026-02 | In GSTR-2B, not in Books |
| HOTEL MAYURPARK RESIDENCY | 24BNUPD3394J1ZK | 1378 | 2026-01-23 00:00:00 | 2,200.00 | 0.00 | 55.00 | 55.00 | 110.00 | 2026-03 | In GSTR-2B (Rejected on IMS), not in Books |
| Turanth Logistic Private Limited | 27AAECT9480C1Z5 | 2092914 | 2025-07-04 00:00:00 | 1,857.00 | 93.00 | 0.00 | 0.00 | 93.00 | 2025-07 | In GSTR-2B, not in Books |
| SHREE GANESH AUTO | 24AHOPC4366Q1ZT | 25-26/5879 | 2025-12-01 00:00:00 | 500.00 | 0.00 | 45.00 | 45.00 | 90.00 | 2025-12 | In GSTR-2B, not in Books |
| SHREE GANESH AUTO | 24AHOPC4366Q1ZT | 25-26/6140 | 2025-12-11 00:00:00 | 500.00 | 0.00 | 45.00 | 45.00 | 90.00 | 2025-12 | In GSTR-2B, not in Books |
| Turanth Logistic Private Limited | 27AAECT9480C1Z5 | 2092915 | 2025-07-04 00:00:00 | 1,486.00 | 74.00 | 0.00 | 0.00 | 74.00 | 2025-07 | In GSTR-2B, not in Books |
| JR ENTERPRISE | 24AIAPM8938K1ZW | TI/0642/25-26 | 2025-06-05 00:00:00 | 535.68 | 0.00 | 32.14 | 32.14 | 64.28 | 2025-06 | In GSTR-2B, not in Books |
| JR ENTERPRISE | 24AIAPM8938K1ZW | TI/1192/25-26 | 2025-08-04 00:00:00 | 178.56 | 0.00 | 10.71 | 10.71 | 21.42 | 2025-08 | In GSTR-2B, not in Books |
| LE TRAVENUES TECHNOLOGY LIMITED | 06AABCL1932G1ZV | IXIFT00015965433 | 2025-05-22 00:00:00 | 117.79 | 21.21 | 0.00 | 0.00 | 21.21 | 2025-05 | In GSTR-2B, not in Books |
| FALCON CARGO | 27AAMPK5075B2ZP | FC/7670/25-26 | 2025-11-21 00:00:00 | 170.00 | 8.50 | 0.00 | 0.00 | 8.50 | 2025-11 | In GSTR-2B (RCM), not in Books |
| HDFC BANK LIMITED | 27AAACH2702H3ZY | UPI2615114863981 | 2025-05-31 00:00:00 | 8.26 | 1.49 | 0.00 | 0.00 | 1.49 | 2025-05 | In GSTR-2B, not in Books |
| FALCON CARGO | 27AAMPK5075B2ZP | A-7942-25/26 | 2025-04-20 00:00:00 | 270.00 | 0.00 | 0.00 | 0.00 | 0.00 | 2025-04 | In GSTR-2B (RCM), not in Books |

### 4. GST amount mismatch - same invoice, GST value differs (to be verified with supplier)  **(41 rows)**

| Branch | Book Date | Supplier | GSTIN | Inv (Books) | Inv (2B) | Book Mth | 2B Mth | Taxable (Bk) | Taxable (2B) | IGST (Bk) | IGST (2B) | CGST (Bk) | CGST (2B) | SGST (Bk) | SGST (2B) | GST (Bk) | GST (2B) | Status | Remarks |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| VL | 2025-04-28 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633189603 | 3325052309 | 2025-04 | 2025-04 | 285,340.97 | 285,340.97 | 66,354.68 | 66,354.76 | 0.00 | 0.00 | 0.00 | 0.00 | 66,354.68 | 66,354.76 | GST Mismatch | GST diff (Books-2B): IGST -0.08 / CGST 0.0 / SGST 0.0 - to be verified; Invoice no differs (Books 1633189603 vs GSTR-2B  |
| HR | 2025-08-06 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633474873 | 3325312387 | 2025-08 | 2025-08 | 105,085.60 | 105,085.20 | 18,915.41 | 18,915.34 | 0.00 | 0.00 | 0.00 | 0.00 | 18,915.41 | 18,915.34 | GST Mismatch | GST diff (Books-2B): IGST 0.07 / CGST 0.0 / SGST 0.0 - to be verified; Invoice no differs (Books 1633474873 vs GSTR-2B 3 |
| HR | 2025-11-21 00:00:00 | The New India Assurance Co Ltd | 24AAACN4165C2ZW | 23010025E0023299 | 23010025E0023299 | 2025-11 | 2025-11 | 99,588.16 | 99,588.00 | 0.00 | 0.00 | 8,962.92 | 8,963.00 | 8,962.92 | 8,963.00 | 17,925.84 | 17,926.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.08 / SGST -0.08 - to be verified |
| HR | 2025-05-23 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633258666 | 3325113624 | 2025-05 | 2025-05 | 21,201.29 | 21,201.00 | 3,816.23 | 3,816.18 | 0.00 | 0.00 | 0.00 | 0.00 | 3,816.23 | 3,816.18 | GST Mismatch | GST diff (Books-2B): IGST 0.05 / CGST 0.0 / SGST 0.0 - to be verified; Invoice no differs (Books 1633258666 vs GSTR-2B 3 |
| HR | 2025-04-30 00:00:00 | Bharat Hardware & Glass | 24AYCPP8314N1Z3 | 125 | 125 | 2025-04 | 2025-06 | 19,871.12 | 19,872.00 | 0.00 | 0.00 | 1,788.44 | 1,788.00 | 1,788.44 | 1,788.00 | 3,576.88 | 3,576.00 | GST Mismatch | Timing Difference - ITC reflected in subsequent GSTR-2B (booked 2025-04; in GSTR-2B 2025-06); GST diff (Books-2B): IGST  |
| VL | 2025-09-24 00:00:00 | The New India Assurance Co Ltd | 24AAACN4165C2ZW | 23010025P0017821 | 23010025P0017821 | 2025-09 | 2025-09 | 18,946.54 | 18,947.00 | 0.00 | 0.00 | 1,705.23 | 1,705.00 | 1,705.23 | 1,705.00 | 3,410.46 | 3,410.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.23 / SGST 0.23 - to be verified |
| VL | 2025-09-24 00:00:00 | The New India Assurance Co Ltd | 24AAACN4165C2ZW | 23010025P0017813 | 23010025P0017813 | 2025-09 | 2025-09 | 18,078.78 | 18,079.00 | 0.00 | 0.00 | 1,627.11 | 1,627.00 | 1,627.11 | 1,627.00 | 3,254.22 | 3,254.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.11 / SGST 0.11 - to be verified |
| HR | 2025-09-22 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633593572 | 3325419225 | 2025-09 | 2025-09 | 11,940.27 | 11,940.00 | 2,149.25 | 2,149.20 | 0.00 | 0.00 | 0.00 | 0.00 | 2,149.25 | 2,149.20 | GST Mismatch | GST diff (Books-2B): IGST 0.05 / CGST 0.0 / SGST 0.0 - to be verified; Invoice no differs (Books 1633593572 vs GSTR-2B 3 |
| HR | 2025-07-03 00:00:00 | Arihant Metal | 24DGXPS8290E1ZL | 430 | 430 | 2025-07 | 2025-07 | 11,412.74 | 11,412.50 | 0.00 | 0.00 | 1,027.13 | 1,027.00 | 1,027.13 | 1,027.00 | 2,054.26 | 2,054.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.13 / SGST 0.13 - to be verified |
| HR | 2025-06-04 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/1317 | 25-26/1317 | 2025-06 | 2025-06 | 6,607.50 | 6,613.00 | 0.00 | 0.00 | 900.75 | 900.82 | 900.75 | 900.82 | 1,801.50 | 1,801.64 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.07 / SGST -0.07 - to be verified |
| HR | 2025-07-14 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/2323 | 25-26/2323 | 2025-07 | 2025-07 | 5,674.88 | 5,680.00 | 0.00 | 0.00 | 770.06 | 770.20 | 770.06 | 770.20 | 1,540.12 | 1,540.40 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.14 / SGST -0.14 - to be verified |
| VL | 2025-10-14 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633646983 | 3325469694 | 2025-10 | 2025-10 | 7,922.92 | 7,922.67 | 1,426.13 | 1,426.08 | 0.00 | 0.00 | 0.00 | 0.00 | 1,426.13 | 1,426.08 | GST Mismatch | GST diff (Books-2B): IGST 0.05 / CGST 0.0 / SGST 0.0 - to be verified; Invoice no differs (Books 1633646983 vs GSTR-2B 3 |
| PL | 2025-09-15 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633578557 | 3325405513 | 2025-09 | 2025-09 | 7,049.03 | 7,050.03 | 1,268.83 | 1,269.01 | 0.00 | 0.00 | 0.00 | 0.00 | 1,268.83 | 1,269.01 | GST Mismatch | GST diff (Books-2B): IGST -0.18 / CGST 0.0 / SGST 0.0 - to be verified; Invoice no differs (Books 1633578557 vs GSTR-2B  |
| PL | 2025-10-11 00:00:00 | Lakhani Traders | 24AAAFL9494P1ZK | LT/1263/25-26 | LT/1263/25-26 | 2025-10 | 2025-10 | 6,910.20 | 6,910.80 | 0.00 | 0.00 | 621.90 | 621.97 | 621.90 | 621.97 | 1,243.80 | 1,243.94 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.07 / SGST -0.07 - to be verified |
| HR | 2025-09-01 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/3605 | 25-26/3605 | 2025-09 | 2025-09 | 4,475.10 | 4,490.21 | 0.00 | 0.00 | 504.95 | 504.90 | 504.95 | 504.90 | 1,009.90 | 1,009.80 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.05 / SGST 0.05 - to be verified |
| HR | 2025-04-15 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/324 | 25-26/324 | 2025-04 | 2025-04 | 3,935.84 | 3,941.07 | 0.00 | 0.00 | 489.58 | 489.46 | 489.58 | 489.46 | 979.16 | 978.92 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.12 / SGST 0.12 - to be verified |
| HR | 2025-04-14 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/308 | 25-26/308 | 2025-04 | 2025-04 | 3,133.36 | 3,137.50 | 0.00 | 0.00 | 414.32 | 414.25 | 414.32 | 414.25 | 828.64 | 828.50 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.07 / SGST 0.07 - to be verified |
| HR | 2025-06-30 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/1872 | 25-26/1872 | 2025-06 | 2025-06 | 3,132.50 | 3,138.00 | 0.00 | 0.00 | 414.25 | 414.32 | 414.25 | 414.32 | 828.50 | 828.64 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.07 / SGST -0.07 - to be verified |
| VL | 2025-08-26 00:00:00 | Kamal Commercial Vehicles Pvt. Ltd(Udaipur)(Pur) | 08AAECK7749K2ZW | DDKMRJ1B25000043 | DDKMRJ1B25000043 | 2025-08 | 2025-08 | 2,897.60 | 2,897.59 | 783.40 | 783.53 | 0.00 | 0.00 | 0.00 | 0.00 | 783.40 | 783.53 | GST Mismatch | GST diff (Books-2B): IGST -0.13 / CGST 0.0 / SGST 0.0 - to be verified |
| HR | 2025-04-10 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/205 | 25-26/205 | 2025-04 | 2025-04 | 2,812.52 | 2,816.41 | 0.00 | 0.00 | 381.74 | 381.80 | 381.74 | 381.80 | 763.48 | 763.60 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.06 / SGST -0.06 - to be verified |
| HR | 2025-05-30 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI |  | DDKTGJ1B25000017 | 2025-05 | 2025-06 | 2,600.68 | 2,600.68 | 0.00 | 0.00 | 364.16 | 364.10 | 364.16 | 364.10 | 728.32 | 728.20 | DUPLICATE booking - GST Mismatch | Timing Difference - ITC reflected in subsequent GSTR-2B (booked 2025-05; in GSTR-2B 2025-06); GST diff (Books-2B): IGST  |
| HR | 2025-07-13 00:00:00 | Mahalaxmi Enterprise | 24CDFPG9834G1ZI | FY25-26/232 | FY25-26/232 | 2025-07 | 2025-07 | 8,078.96 | 8,078.80 | 0.00 | 0.00 | 337.02 | 337.09 | 337.02 | 337.09 | 674.04 | 674.18 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.07 / SGST -0.07 - to be verified |
| HR | 2025-04-16 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/360 | 25-26/360 | 2025-04 | 2025-04 | 2,510.52 | 2,515.63 | 0.00 | 0.00 | 327.24 | 327.19 | 327.24 | 327.19 | 654.48 | 654.38 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.05 / SGST 0.05 - to be verified |
| HR | 2025-04-16 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/362 | 25-26/362 | 2025-04 | 2025-04 | 2,510.52 | 2,515.63 | 0.00 | 0.00 | 327.24 | 327.19 | 327.24 | 327.19 | 654.48 | 654.38 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.05 / SGST 0.05 - to be verified |
| HR | 2025-05-01 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/466 | 25-26/466 | 2025-05 | 2025-04 | 2,503.52 | 2,515.63 | 0.00 | 0.00 | 327.24 | 327.19 | 327.24 | 327.19 | 654.48 | 654.38 | GST Mismatch | Timing Difference - ITC reflected in earlier GSTR-2B (booked 2025-05; in GSTR-2B 2025-04); GST diff (Books-2B): IGST 0.0 |
| HR | 2025-09-10 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/3814 | 25-26/3814 | 2025-09 | 2025-09 | 2,510.52 | 2,515.63 | 0.00 | 0.00 | 327.24 | 327.19 | 327.24 | 327.19 | 654.48 | 654.38 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.05 / SGST 0.05 - to be verified |
| HR | 2025-08-07 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/3007 | 25-26/3007 | 2025-08 | 2025-08 | 2,221.58 | 2,227.00 | 0.00 | 0.00 | 286.71 | 286.78 | 286.71 | 286.78 | 573.42 | 573.56 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.07 / SGST -0.07 - to be verified |
| HR | 2025-11-21 00:00:00 | The New India Assurance Co Ltd | 24AAACN4165C2ZW | 23010025P0023317 | 23010025P0023317 | 2025-11 | 2025-11 | 2,985.52 | 2,986.00 | 0.00 | 0.00 | 268.74 | 269.00 | 268.74 | 269.00 | 537.48 | 538.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.26 / SGST -0.26 - to be verified |
| HR | 2025-12-31 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1633890649 | 6025118387 | 2025-12 | 2025-12 | 2,850.57 | 2,850.07 | 513.10 | 513.01 | 0.00 | 0.00 | 0.00 | 0.00 | 513.10 | 513.01 | GST Mismatch | GST diff (Books-2B): IGST 0.09 / CGST 0.0 / SGST 0.0 - to be verified; Invoice no differs (Books 1633890649 vs GSTR-2B 6 |
| HR | 2025-05-13 00:00:00 | National Enterprise | 24ACTPL0091R1ZX | 104 | 104 | 2025-05 | 2025-05 | 2,470.40 | 2,470.00 | 0.00 | 0.00 | 222.30 | 222.00 | 222.30 | 222.00 | 444.60 | 444.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.3 / SGST 0.3 - to be verified |
| PL | 2026-02-24 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1631336602 | 6025145924 | 2026-02 | 2026-02 | 2,233.90 | 2,232.90 | 402.10 | 401.92 | 0.00 | 0.00 | 0.00 | 0.00 | 402.10 | 401.92 | GST Mismatch | GST diff (Books-2B): IGST 0.18 / CGST 0.0 / SGST 0.0 - to be verified; Invoice no differs (Books 1631336602 vs GSTR-2B 6 |
| HR | 2025-11-21 00:00:00 | The New India Assurance Co Ltd | 24AAACN4165C2ZW | 23010025E0023310 | 23010025E0023310 | 2025-11 | 2025-11 | 2,225.50 | 2,225.00 | 0.00 | 0.00 | 200.25 | 200.00 | 200.25 | 200.00 | 400.50 | 400.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.25 / SGST 0.25 - to be verified |
| HR | 2025-06-30 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/1158 | 25-26/1158 | 2025-06 | 2025-05 | 1,560.72 | 1,566.00 | 0.00 | 0.00 | 194.14 | 194.24 | 194.14 | 194.24 | 388.28 | 388.48 | GST Mismatch | Timing Difference - ITC reflected in earlier GSTR-2B (booked 2025-06; in GSTR-2B 2025-05); GST diff (Books-2B): IGST 0.0 |
| HR | 2025-06-30 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/1043 | 25-26/1043 | 2025-06 | 2025-05 | 1,421.58 | 1,427.00 | 0.00 | 0.00 | 174.71 | 174.78 | 174.71 | 174.78 | 349.42 | 349.56 | GST Mismatch | Timing Difference - ITC reflected in earlier GSTR-2B (booked 2025-06; in GSTR-2B 2025-05); GST diff (Books-2B): IGST 0.0 |
| HR | 2025-06-02 00:00:00 | National Enterprise | 24ACTPL0091R1ZX | 149 | 149 | 2025-06 | 2025-06 | 1,920.40 | 1,920.00 | 0.00 | 0.00 | 172.80 | 173.00 | 172.80 | 173.00 | 345.60 | 346.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.2 / SGST -0.2 - to be verified |
| HR | 2025-04-07 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/126 | 25-26/126 | 2025-04 | 2025-04 | 1,251.60 | 1,255.47 | 0.00 | 0.00 | 163.20 | 163.27 | 163.20 | 163.27 | 326.40 | 326.54 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.07 / SGST -0.07 - to be verified |
| HR | 2025-06-11 00:00:00 | Rajaram Electricals | 24AHQPL8362B1Z9 | 960 | 960 | 2025-06 | 2025-06 | 1,767.32 | 1,753.00 | 0.00 | 0.00 | 157.84 | 157.77 | 157.84 | 157.77 | 315.68 | 315.54 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.07 / SGST 0.07 - to be verified |
| HR | 2025-05-22 00:00:00 | Interglobe Aviation Limited-Surat | 24AABCI2726B1Z8 | GJ1252605AF84889 | GJ1252605AF84889 | 2025-05 | 2025-05 | 4,203.80 | 4,204.00 | 0.00 | 0.00 | 105.10 | 105.00 | 105.10 | 105.00 | 210.20 | 210.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.1 / SGST 0.1 - to be verified |
| HR | 2025-05-13 00:00:00 | National Enterprise | 24ACTPL0091R1ZX | 103 | 103 | 2025-05 | 2025-05 | 1,016.94 | 1,017.00 | 0.00 | 0.00 | 91.53 | 92.00 | 91.53 | 92.00 | 183.06 | 184.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.47 / SGST -0.47 - to be verified |
| HR | 2025-06-14 00:00:00 | National Enterprise | 24ACTPL0091R1ZX | 168 | 168 | 2025-06 | 2025-06 | 1,016.12 | 1,016.00 | 0.00 | 0.00 | 91.44 | 91.00 | 91.44 | 91.00 | 182.88 | 182.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST 0.44 / SGST 0.44 - to be verified |
| HR | 2025-06-30 00:00:00 | National Enterprise | 24ACTPL0091R1ZX | 206 | 206 | 2025-06 | 2025-06 | 720.40 | 720.00 | 0.00 | 0.00 | 64.80 | 65.00 | 64.80 | 65.00 | 129.60 | 130.00 | GST Mismatch | GST diff (Books-2B): IGST 0.0 / CGST -0.2 / SGST -0.2 - to be verified |

### 5. ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split)  **(49 rows)**

| Branch | Book Date | Supplier | GSTIN | Invoice No | Book Mth | 2B Mth | Booked Amt (no GST) | Taxable (2B) | IGST (2B) | CGST (2B) | SGST (2B) | ITC in 2B not claimed | Status | Remarks |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| HR | 2025-05-05 00:00:00 | Labheshwar Fabrication | 24BEQPV0376H1Z2 | GST/3/25-26 | 2025-05 | 2025-05 | 495,000.00 | 423,728.80 | 0.00 | 38,135.59 | 38,135.59 | 76,271.18 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2025-04-30 00:00:00 | Shree Shivshakti Enteprises | 24ATFPR9208M3Z5 | 2025-26-0046 | 2025-04 | 2025-04 | 328,662.00 | 313,011.00 | 0.00 | 7,825.28 | 7,825.28 | 15,650.56 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| VL | 2025-06-11 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-312 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-312 vs GSTR-2B B2B/312) |
| VL | 2025-06-11 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-313 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-313 vs GSTR-2B B2B/313) |
| VL | 2025-06-11 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-314 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-314 vs GSTR-2B B2B/314) |
| VL | 2025-06-12 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-317 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-317 vs GSTR-2B B2B/317) |
| VL | 2025-06-12 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-319 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-319 vs GSTR-2B B2B/319) |
| VL | 2025-06-13 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-324 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-324 vs GSTR-2B B2B/324) |
| VL | 2025-06-14 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-325 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-325 vs GSTR-2B B2B/326) |
| VL | 2025-06-09 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-298 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-298 vs GSTR-2B B2B/298) |
| VL | 2025-06-10 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-299 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-299 vs GSTR-2B B2B/299) |
| VL | 2025-06-10 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-302 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-302 vs GSTR-2B B2B/302) |
| VL | 2025-06-10 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-304 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-304 vs GSTR-2B B2B/304) |
| VL | 2025-06-10 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-306 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-306 vs GSTR-2B B2B/306) |
| VL | 2025-06-10 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-305 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-305 vs GSTR-2B B2B/305) |
| VL | 2025-06-11 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-309 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-309 vs GSTR-2B B2B/309) |
| VL | 2025-06-05 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-281 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-281 vs GSTR-2B B2B/281) |
| VL | 2025-06-14 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-327 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-327 vs GSTR-2B B2B/327) |
| VL | 2025-06-14 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-329 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-329 vs GSTR-2B B2B/329) |
| VL | 2025-06-14 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-331 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-331 vs GSTR-2B B2B/331) |
| VL | 2025-06-14 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-326 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-326 vs GSTR-2B B2B/336) |
| VL | 2025-06-15 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-333 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-333 vs GSTR-2B B2B/333) |
| VL | 2025-06-15 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-335 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-335 vs GSTR-2B B2B/335) |
| VL | 2025-06-16 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-336 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-336 vs GSTR-2B B2B/337) |
| VL | 2025-06-11 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-310 | 2025-06 | 2025-06 | 78,824.00 | 66,800.00 | 0.00 | 6,012.00 | 6,012.00 | 12,024.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-310 vs GSTR-2B B2B/310) |
| VL | 2025-06-06 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-286 | 2025-06 | 2025-06 | 68,971.00 | 58,450.00 | 0.00 | 5,260.50 | 5,260.50 | 10,521.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-286 vs GSTR-2B B2B/286) |
| VL | 2025-06-03 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B/270 | 2025-06 | 2025-06 | 68,971.00 | 58,450.00 | 0.00 | 5,260.50 | 5,260.50 | 10,521.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2025-06-13 00:00:00 | Gurudev Gypsum & P.O.P. Traders | 24APZPP3911M1Z7 | G/32 | 2025-06 | 2025-06 | 51,148.00 | 43,345.48 | 0.00 | 3,901.09 | 3,901.09 | 7,802.18 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| VL | 2025-06-08 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-296 | 2025-06 | 2025-06 | 29,559.00 | 25,050.00 | 0.00 | 2,254.50 | 2,254.50 | 4,509.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-296 vs GSTR-2B B2B/296) |
| PL | 2026-01-05 00:00:00 | Magma Hdi General Insurance Ltd. | 24AAGCM1685C1ZP | POL2401260000292 | 2026-01 | 2026-01 | 34,684.00 | 31,158.00 | 0.00 | 1,763.00 | 1,763.00 | 3,526.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| VL | 2026-01-05 00:00:00 | Magma Hdi General Insurance Ltd. | 24AAGCM1685C1ZP | POL2401260000293 | 2026-01 | 2026-01 | 34,684.00 | 31,158.00 | 0.00 | 1,763.00 | 1,763.00 | 3,526.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2025-04-25 00:00:00 | Kamani Paints | 24AFUPP6416Q1ZJ | GT/215 | 2025-04 | 2025-04 | 22,500.00 | 19,067.82 | 0.00 | 1,716.10 | 1,716.10 | 3,432.20 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2025-04-04 00:00:00 | Viral Enterprise | 24AALFV6612Q3ZH | 25-26/0048 | 2025-04 | 2025-04 | 19,909.00 | 16,872.08 | 0.00 | 1,518.49 | 1,518.49 | 3,036.98 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2025-04-05 00:00:00 | Viral Enterprise | 24AALFV6612Q3ZH | 25-26/0072 | 2025-04 | 2025-04 | 16,805.00 | 14,241.38 | 0.00 | 1,281.73 | 1,281.73 | 2,563.46 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2025-04-01 00:00:00 | Viral Enterprise | 24AALFV6612Q3ZH | 25-26/0245 | 2025-04 | 2025-04 | 13,555.00 | 11,487.67 | 0.00 | 1,033.90 | 1,033.90 | 2,067.80 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| VL | 2025-06-06 00:00:00 | Parth Cement Articles | 24AAJFP9360P1ZI | B2B-285 | 2025-06 | 2025-06 | 9,912.00 | 8,400.00 | 0.00 | 756.00 | 756.00 | 1,512.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books B2B-285 vs GSTR-2B B2B/285) |
| HR | 2025-05-26 00:00:00 | Viral Enterprise | 24AALFV6612Q3ZH | 25-26/0933 | 2025-05 | 2025-05 | 8,692.00 | 7,366.30 | 0.00 | 662.96 | 662.96 | 1,325.92 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2025-05-26 00:00:00 | Viral Enterprise | 24AALFV6612Q3ZH | 25-26/L-0226 | 2025-05 | 2025-05 | 7,670.00 | 6,500.00 | 0.00 | 585.00 | 585.00 | 1,170.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2026-02-01 00:00:00 | Vinod Steel Center | 24ABOPJ2316H1ZV | G/25-26/7552 | 2026-02 | 2026-01 | 7,650.00 | 6,483.36 | 0.00 | 583.50 | 583.50 | 1,167.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2026-02-24 00:00:00 | Vinod Steel Center | 24ABOPJ2316H1ZV | G/25-26/8863 | 2026-02 | 2026-02 | 5,940.00 | 5,033.90 | 0.00 | 453.05 | 453.05 | 906.10 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2025-04-23 00:00:00 | Asheesh Sales Corporation | 24AATPJ3597Q1ZO | 96 | 2025-04 | 2025-04 | 5,942.00 | 5,053.06 | 0.00 | 451.53 | 451.53 | 903.06 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2025-06-20 00:00:00 | Nest Corporate Homes | 33AIBPD8987G1Z5 | 1401 | 2025-06 | 2025-06 | 7,455.00 | 6,700.00 | 0.00 | 377.50 | 377.50 | 755.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) [ITC Not Available] | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books 1401 vs GSTR-2B 1402) |
| HR | 2025-05-29 00:00:00 | Nest Corporate Homes | 33AIBPD8987G1Z5 | 1370 | 2025-05 | 2025-05 | 6,384.00 | 5,760.00 | 0.00 | 312.00 | 312.00 | 624.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) [ITC Not Available] | Verify whether ITC is claimable / should have been booked |
| HR | 2025-09-07 00:00:00 | Nest Corporate Homes | 33AIBPD8987G1Z5 | 1433 | 2025-09 | 2025-09 | 5,649.00 | 5,100.00 | 0.00 | 274.50 | 274.50 | 549.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) [ITC Not Available] | Verify whether ITC is claimable / should have been booked |
| HR | 2026-02-06 00:00:00 | Vinod Steel Center | 24ABOPJ2316H1ZV | G/25-26/7336 | 2026-02 | 2026-02 | 2,910.00 | 2,466.10 | 0.00 | 221.95 | 221.95 | 443.90 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books G/25-26/7336 vs GSTR-2B G/25-26/833 |
| HR | 2026-03-09 00:00:00 | Vinod Steel Center | 24ABOPJ2316H1ZV | G/25-26/9144 | 2026-03 | 2026-03 | 2,410.00 | 2,042.61 | 0.00 | 183.83 | 183.83 | 367.66 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2025-05-26 00:00:00 | Viral Enterprise | 24AALFV6612Q3ZH | 25-26/0932 | 2025-05 | 2025-05 | 2,203.00 | 1,866.69 | 0.00 | 168.00 | 168.00 | 336.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked |
| HR | 2025-04-18 00:00:00 | Hotel Aarti Executive | 27AHKPB6123A2ZX | 12 | 2025-04 | 2025-04 | 2,016.00 | 1,800.00 | 0.00 | 108.00 | 108.00 | 216.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) [ITC Not Available] | Verify whether ITC is claimable / should have been booked |
| VL | 2025-04-29 00:00:00 | Carewell Motors and Infra Private Limited | 24AAECC8137D1ZY | CMI/24-26/280 | 2025-04 | 2025-04 | 1,180.00 | 1,000.00 | 0.00 | 90.00 | 90.00 | 180.00 | ITC in GSTR-2B but NOT claimed in Books (invoice booked without GST split) | Verify whether ITC is claimable / should have been booked; Invoice no differs (Books CMI/24-26/280 vs GSTR-2B CMI/25-26/ |

### 6. Matched - invoice date differs (GST & values agree)  **(71 rows)**

| Branch | Book Date | Supplier | GSTIN | Inv (Books) | Inv (2B) | Inv Date (Books) | Book Mth | 2B Mth | Total GST | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| PL | 2025-05-19 00:00:00 | Kunj Enterprise | 24AAPPD6814Q1Z8 | KE/25-26/142 | KE/25-26/142 | 2025-05-19 00:00:00 | 2025-05 | 2025-05 | 19,800.00 | Matched - invoice date differs |
| HR | 2026-01-18 00:00:00 | Magma General Insurance Limited [MAHA] | 27AAGCM1685C1ZJ | POL2701260024569 | POL2701260024569 | 2026-01-18 00:00:00 | 2026-01 | 2026-01 | 15,528.96 | Matched - invoice date differs |
| VL | 2025-09-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633613970 | 3325438196 | 2025-09-28 00:00:00 | 2025-09 | 2025-09 | 8,127.45 | Matched - invoice date differs |
| PL | 2025-10-09 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633634225 | 3325457345 | 2025-10-07 00:00:00 | 2025-10 | 2025-10 | 6,764.74 | Matched - invoice date differs |
| PL | 2025-10-30 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633684710 | 3325504072 | 2025-10-30 00:00:00 | 2025-10 | 2025-10 | 6,371.25 | Matched - invoice date differs |
| VL | 2026-01-24 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633939945 | 3325736758 | 2026-01-22 00:00:00 | 2026-01 | 2026-01 | 5,581.42 | Matched - invoice date differs |
| HR | 2026-03-06 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631370680 | 3325850712 | 2026-03-05 00:00:00 | 2026-03 | 2026-03 | 3,903.38 | Matched - invoice date differs |
| HR | 2025-04-30 00:00:00 | D.I.C.V. Pvt Ltd(Amc/goodwill/ Warranty/sm/Rsa/cou) | 33AABCF1590N1ZJ | 3325056104 | 3325056104 | 2025-04-30 00:00:00 | 2025-04 | 2025-04 | 3,402.00 | Matched - invoice date differs |
| HR | 2025-09-30 00:00:00 | K. Pansuriya Llp | 24ABDFK8458L1ZU | KPL-204-2025-26 | KPL-204-2025-26 | 2025-09-30 00:00:00 | 2025-09 | 2025-09 | 3,240.00 | Matched - invoice date differs |
| HR | 2025-05-02 00:00:00 | Abound Erp Pvt. Ltd | 24AAOCA6538Q2ZX | 221/25-26 | 221/25-26 | 2025-05-02 00:00:00 | 2025-05 | 2025-05 | 2,430.00 | Matched - invoice date differs |
| VL | 2025-11-07 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633708485 | 3325526137 | 2025-11-06 00:00:00 | 2025-11 | 2025-11 | 2,271.66 | Matched - invoice date differs |
| HR | 2025-07-10 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633402871 | 3325244935 | 2025-07-09 00:00:00 | 2025-07 | 2025-07 | 2,012.58 | Matched - invoice date differs |
| HR | 2025-11-11 00:00:00 | Vodafone Idea Limited | 24AAACB2100P1Z3 | GJSO051125448537 | GJSO051125448537 | 2025-11-11 00:00:00 | 2025-11 | 2025-11 | 1,984.68 | Matched - invoice date differs |
| HR | 2025-11-11 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633717198 | 3325533779 | 2025-11-10 00:00:00 | 2025-11 | 2025-11 | 1,946.22 | Matched - invoice date differs |
| VL | 2025-10-31 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633690700 | 3325509906 | 2025-10-30 00:00:00 | 2025-10 | 2025-10 | 1,927.21 | Matched - invoice date differs |
| HR | 2025-07-10 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633402444 | 3325244508 | 2025-07-09 00:00:00 | 2025-07 | 2025-07 | 1,777.60 | Matched - invoice date differs |
| VL | 2025-12-10 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633807268 | 3325616948 | 2025-12-09 00:00:00 | 2025-12 | 2025-12 | 1,612.13 | Matched - invoice date differs |
| VL | 2026-03-04 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631363786 | 3325839432 | 2026-03-03 00:00:00 | 2026-03 | 2026-03 | 1,477.76 | Matched - invoice date differs |
| HR | 2025-04-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633195250 | 3325058154 | 2025-04-24 00:00:00 | 2025-04 | 2025-04 | 1,401.75 | Matched - invoice date differs |
| VL | 2025-12-12 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633815392 | 3325624516 | 2025-12-11 00:00:00 | 2025-12 | 2025-12 | 1,346.34 | Matched - invoice date differs |
| PL | 2026-03-23 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631433896 | 3325902368 | 2026-03-22 00:00:00 | 2026-03 | 2026-03 | 1,340.88 | Matched - invoice date differs |
| VL | 2025-12-24 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1633854453 | 6025112761 | 2025-12-23 00:00:00 | 2025-12 | 2025-12 | 1,325.27 | Matched - invoice date differs |
| VL | 2026-03-25 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1631443148 | 6025162111 | 2026-03-24 00:00:00 | 2026-03 | 2026-03 | 1,245.02 | Matched - invoice date differs |
| VL | 2026-03-16 00:00:00 | Daimler India Commercial Vehicles Pvt Ltd - MAHARA | 27AABCF1590N2ZB | 1631406568 | 2725000471 | 2026-03-14 00:00:00 | 2026-03 | 2026-03 | 1,157.44 | Matched - invoice date differs |
| VL | 2025-10-25 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633675548 | 3325495777 | 2025-10-24 00:00:00 | 2025-10 | 2025-10 | 1,032.22 | Matched - invoice date differs |
| HR | 2025-04-17 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633152119 | 3325017792 | 2025-04-16 00:00:00 | 2025-04 | 2025-04 | 998.28 | Matched - invoice date differs |
| HR | 2026-02-25 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631339201 | 3325815955 | 2026-02-24 00:00:00 | 2026-02 | 2026-02 | 943.63 | Matched - invoice date differs |
| PL | 2025-05-15 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633228735 | 3325088947 | 2025-05-14 00:00:00 | 2025-05 | 2025-05 | 893.58 | Matched - invoice date differs |
| PL | 2025-08-31 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633543032 | 3325374635 | 2025-08-30 00:00:00 | 2025-08 | 2025-08 | 665.71 | Matched - invoice date differs |
| VL | 2025-11-07 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633708989 | 3325526634 | 2025-11-06 00:00:00 | 2025-11 | 2025-11 | 633.56 | Matched - invoice date differs |
| VL | 2026-01-29 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633950234 | 3325746501 | 2026-01-28 00:00:00 | 2026-01 | 2026-01 | 613.85 | Matched - invoice date differs |
| HR | 2025-07-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP | C24E252600092721 | C24E252600092721 | 2025-07-01 00:00:00 | 2025-07 | 2025-07 | 608.94 | Matched - invoice date differs |
| HR | 2025-05-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP | C24E252600047731 | C24E252600047731 | 2025-05-01 00:00:00 | 2025-05 | 2025-05 | 608.94 | Matched - invoice date differs |
| HR | 2025-04-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP | C24E252600022548 | C24E252600022548 | 2025-04-01 00:00:00 | 2025-04 | 2025-04 | 608.94 | Matched - invoice date differs |
| HR | 2025-06-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP | C24E252600060988 | C24E252600060988 | 2025-06-01 00:00:00 | 2025-06 | 2025-06 | 608.94 | Matched - invoice date differs |
| HR | 2025-08-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP | C24E252600126040 | C24E252600126040 | 2025-08-01 00:00:00 | 2025-08 | 2025-08 | 606.62 | Matched - invoice date differs |
| PL | 2025-11-25 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633767638 | 3325579099 | 2025-11-25 00:00:00 | 2025-11 | 2025-11 | 604.69 | Matched - invoice date differs |
| PL | 2026-03-13 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631397560 | 3325870489 | 2026-03-12 00:00:00 | 2026-03 | 2026-03 | 603.11 | Matched - invoice date differs |
| VL | 2026-03-25 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1631442582 | 6025161809 | 2026-03-24 00:00:00 | 2026-03 | 2026-03 | 601.42 | Matched - invoice date differs |
| PL | 2026-02-21 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1631328839 | 6025144622 | 2026-02-20 00:00:00 | 2026-02 | 2026-02 | 600.63 | Matched - invoice date differs |
| HR | 2026-03-14 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631401674 | 3325874019 | 2026-02-14 00:00:00 | 2026-03 | 2026-03 | 584.60 | Matched - invoice date differs |
| HR | 2025-07-10 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633402869 | 3325244933 | 2025-07-09 00:00:00 | 2025-07 | 2025-07 | 583.83 | Matched - invoice date differs |
| HR | 2025-06-22 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633351300 | 3325198172 | 2025-06-21 00:00:00 | 2025-06 | 2025-06 | 578.29 | Matched - invoice date differs |
| HR | 2025-05-08 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1633214378 | 6025012134 | 2025-05-05 00:00:00 | 2025-05 | 2025-05 | 535.78 | Matched - invoice date differs |
| HR | 2025-12-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP | C24E252600246157 | C24E252600246157 | 2025-12-01 00:00:00 | 2025-12 | 2025-12 | 516.08 | Matched - invoice date differs |
| VL | 2025-10-16 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633657202 | 3325479384 | 2025-10-16 00:00:00 | 2025-10 | 2025-10 | 514.85 | Matched - invoice date differs |
| HR | 2025-09-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP | C24E252600146472 | C24E252600146472 | 2025-09-01 00:00:00 | 2025-09 | 2025-09 | 510.72 | Matched - invoice date differs |
| HR | 2025-11-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP | C24E252600199047 | C24E252600199047 | 2025-11-01 00:00:00 | 2025-11 | 2025-11 | 501.48 | Matched - invoice date differs |
| HR | 2025-10-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP | C24E252600181002 | C24E252600181002 | 2025-10-01 00:00:00 | 2025-10 | 2025-10 | 501.48 | Matched - invoice date differs |
| HR | 2026-02-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP | C24E252600300207 | C24E252600300207 | 2026-02-01 00:00:00 | 2026-02 | 2026-02 | 501.48 | Matched - invoice date differs |
| VL | 2026-03-25 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1631443283 | 6025162240 | 2026-03-24 00:00:00 | 2026-03 | 2026-03 | 474.13 | Matched - invoice date differs |
| VL | 2025-10-16 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633657203 | 3325479385 | 2025-10-16 00:00:00 | 2025-10 | 2025-10 | 448.87 | Matched - invoice date differs |
| HR | 2025-08-31 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633540746 | 3325372776 | 2025-08-30 00:00:00 | 2025-08 | 2025-08 | 446.36 | Matched - invoice date differs |
| HR | 2025-12-05 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1633798022 | 6025104848 | 2025-12-05 00:00:00 | 2025-12 | 2025-12 | 416.89 | Matched - invoice date differs |
| VL | 2026-01-20 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633924276 | 3325722122 | 2026-01-19 00:00:00 | 2026-01 | 2026-01 | 399.25 | Matched - invoice date differs |
| PL | 2025-11-25 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633768154 | 3325579562 | 2025-11-25 00:00:00 | 2025-11 | 2025-11 | 355.70 | Matched - invoice date differs |
| VL | 2026-01-24 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1633940285 | 6025127819 | 2026-01-22 00:00:00 | 2026-01 | 2026-01 | 341.08 | Matched - invoice date differs |
| VL | 2026-01-24 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1633940284 | 6025127818 | 2026-01-22 00:00:00 | 2026-01 | 2026-01 | 332.76 | Matched - invoice date differs |
| HR | 2025-05-17 00:00:00 | Sai Sarvin Corporation | 24AJGPP0067K1Z5 | SVC/00139/25-26 | SVC/00139/25-26 | 2025-05-17 00:00:00 | 2025-05 | 2025-05 | 270.00 | Matched - invoice date differs |
| PL | 2026-03-11 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631389152 | 3325862614 | 2026-03-10 00:00:00 | 2026-03 | 2026-03 | 262.03 | Matched - invoice date differs |
| HR | 2025-12-19 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633840528 | 3325647502 | 2025-11-19 00:00:00 | 2025-12 | 2025-12 | 238.09 | Matched - invoice date differs |
| PL | 2025-12-10 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633807205 | 3325616885 | 2025-12-09 00:00:00 | 2025-12 | 2025-12 | 237.22 | Matched - invoice date differs |
| VL | 2025-10-17 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633656861 | 3325479043 | 2025-10-16 00:00:00 | 2025-10 | 2025-10 | 235.22 | Matched - invoice date differs |
| VL | 2025-12-24 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633854062 | 3325659594 | 2025-12-23 00:00:00 | 2025-12 | 2025-12 | 227.88 | Matched - invoice date differs |
| VL | 2026-01-24 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - HARYANA | 06AABCF1590N1ZG | 1633940283 | 6025127817 | 2026-01-22 00:00:00 | 2026-01 | 2026-01 | 219.49 | Matched - invoice date differs |
| VL | 2025-06-19 00:00:00 | Carewell Motors and Infra Private Limited | 24AAECC8137D1ZY | CMI/25-26/684 | CMI/25-26/684 | 2025-04-29 00:00:00 | 2025-06 | 2025-06 | 180.00 | Matched - invoice date differs |
| PL | 2025-04-30 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633204191 | 3325067125 | 2025-04-28 00:00:00 | 2025-04 | 2025-04 | 111.53 | Matched - invoice date differs |
| PL | 2026-01-23 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633936031 | 3325733056 | 2026-01-22 00:00:00 | 2026-01 | 2026-01 | 72.68 | Matched - invoice date differs |
| HR | 2026-03-08 00:00:00 | Dineshbhai V. Tailor | 24AFHPT7182G2Z3 | A476 | A476 | 2026-03-08 00:00:00 | 2026-03 | 2026-03 | 54.00 | Matched - invoice date differs |
| VL | 2026-01-09 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1633900181 | 3325700982 | 2026-01-05 00:00:00 | 2026-01 | 2026-01 | 50.49 | Matched - invoice date differs |
| PL | 2026-03-03 00:00:00 | Daimler India Commercial Vehicle Pvt Ltd - TAMIL | 33AABCF1590N1ZJ | 1631360243 | 3325835646 | 2026-03-02 00:00:00 | 2026-03 | 2026-03 | 44.05 | Matched - invoice date differs |

### 7. Duplicate booking in Books (same invoice booked more than once)  **(73 rows)**

| Branch | Book Date | Supplier | GSTIN | Invoice No | Book Mth | Total GST | Status | Remarks |
|---|---|---|---|---|---|---|---|---|
| HR | 2025-07-01 00:00:00 | Camal Security and Allied Services Pvt Ltd | 24AALCC4280G1ZP | PV/25-26/69 | 2025-07 | 10,980.00 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -1220.0) - no GST impact; Invoice no differs (Books PV/25-26/69 vs GSTR-2B PV/25-26/108 |
| HR | 2026-03-01 00:00:00 | Urmi Equipments | 24ASRPS3707R1ZU |  | 2026-03 | 10,800.00 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2026-02-01 00:00:00 | Urmi Equipments | 24ASRPS3707R1ZU |  | 2026-02 | 10,800.00 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| PL | 2025-06-01 00:00:00 | Camal Security and Allied Services Pvt Ltd | 24AALCC4280G1ZP | PV/25-26/69 | 2025-06 | 5,400.00 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -600.0) - no GST impact |
| HR | 2025-08-21 00:00:00 | Kataria Motors Pvt Ltd-Kachchh(Pur) | 24AACCK0029Q1ZI |  | 2025-08 | 3,396.54 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.06) - no GST impact |
| HR | 2025-04-30 00:00:00 | Babaric Labour Service | 24AANFB7731Q1ZV | 8 | 2025-04 | 3,082.86 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -342.86) - no GST impact |
| PL | 2025-06-30 00:00:00 | Babaric Labour Service | 24AANFB7731Q1ZV | 8 | 2025-06 | 2,916.00 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -324.0) - no GST impact; Invoice no differs (Books 8 vs GSTR-2B 58) |
| HR | 2025-07-14 00:00:00 | Kataria Motors Pvt Ltd -Valsad(Pur) | 24AACCK0029Q1ZI |  | 2025-07 | 2,201.14 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.36) - no GST impact |
| HR | 2026-02-04 00:00:00 | Kataria Motors Pvt Ltd- Ankleshwar(Purchase) | 24AACCK0029Q1ZI |  | 2026-02 | 2,039.00 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.23) - no GST impact |
| HR | 2025-07-18 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI |  | 2025-07 | 1,563.98 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.24) - no GST impact |
| HR | 2025-07-09 00:00:00 | Kataria Motors Pvt Ltd -Valsad(Pur) | 24AACCK0029Q1ZI |  | 2025-07 | 1,547.86 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.28) - no GST impact |
| HR | 2025-09-20 00:00:00 | Kataria Motors Pvt Ltd-Chacharwadi (Pur) | 24AACCK0029Q1ZI |  | 2025-09 | 1,343.42 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.33) - no GST impact |
| HR | 2026-02-04 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI |  | 2026-02 | 1,205.04 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.32) - no GST impact |
| HR | 2026-02-06 00:00:00 | Spirited Motor Vehicles Limited (Purchase)Vasai | 27ABBCS9610H1Z9 |  | 2026-02 | 1,175.96 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.05) - no GST impact |
| HR | 2025-12-10 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 998.64 | DUPLICATE booking - Matched (GST rounding diff <= Rs 1) | Taxable value differs (Books-2B: 0.05) - no GST impact |
| HR | 2025-12-10 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 958.32 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -5.05) - no GST impact |
| HR | 2025-12-18 00:00:00 | Diesel World (Pvt.) Ltd. | 24AAECD2472J1ZQ |  | 2025-12 | 937.52 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -921.99) - no GST impact |
| HR | 2025-12-19 00:00:00 | Kataria Motors Pvt Ltd-Kachchh(Pur) | 24AACCK0029Q1ZI |  | 2025-12 | 867.44 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.48) - no GST impact |
| HR | 2025-05-31 00:00:00 | Shiv Enterprise (Drinking Water ) | 24AHAPP9298E1Z2 |  | 2025-05 | 839.66 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.17) - no GST impact |
| HR | 2025-08-02 00:00:00 | Kataria Motors Pvt Ltd -Vadodara(Pur) | 24AACCK0029Q1ZI |  | 2025-08 | 818.94 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.26) - no GST impact |
| HR | 2025-12-29 00:00:00 | Kataria Motors Pvt Ltd -Rajkot(Pur) | 24AACCK0029Q1ZI |  | 2025-12 | 738.24 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.39) - no GST impact |
| HR | 2025-05-30 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI |  | 2025-05 | 728.32 | DUPLICATE booking - GST Mismatch | Timing Difference - ITC reflected in subsequent GSTR-2B (booked 2025-05; in GSTR-2B 2025-06); GST diff (Books-2B): IGST  |
| HR | 2025-05-31 00:00:00 | Akp Sons Motors Pvt Ltd(Purcahse) | 24AAYCA5330J1ZE |  | 2025-05 | 728.32 | DUPLICATE booking - Timing Difference - ITC reflected in subsequent GSTR-2B | Timing Difference - ITC reflected in subsequent GSTR-2B (booked 2025-05; in GSTR-2B 2025-06); Taxable value differs (Boo |
| HR | 2025-05-19 00:00:00 | Spirited Motor Vehicles Limited (Purchase)Vasai | 27ABBCS9610H1Z9 |  | 2025-05 | 710.90 | DUPLICATE booking - Timing Difference - ITC reflected in subsequent GSTR-2B | Timing Difference - ITC reflected in subsequent GSTR-2B (booked 2025-05; in GSTR-2B 2025-09); Taxable value differs (Boo |
| HR | 2025-12-03 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 707.76 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -9.96) - no GST impact |
| HR | 2025-10-16 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI |  | 2025-10 | 703.84 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.1) - no GST impact |
| HR | 2025-07-05 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI |  | 2025-07 | 658.56 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.21) - no GST impact |
| HR | 2025-07-02 00:00:00 | Kataria Motors Pvt Ltd-Chacharwadi (Pur) | 24AACCK0029Q1ZI |  | 2025-07 | 636.04 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.43) - no GST impact |
| HR | 2025-08-27 00:00:00 | Kataria Motors Pvt Ltd-Chacharwadi (Pur) | 24AACCK0029Q1ZI |  | 2025-08 | 607.78 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.39) - no GST impact |
| HR | 2025-10-14 00:00:00 | Kataria Motors Pvt Ltd -Rajkot(Pur) | 24AACCK0029Q1ZI |  | 2025-10 | 579.74 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.48) - no GST impact |
| HR | 2025-12-25 00:00:00 | Kataria Motors Pvt Ltd-Kachchh(Pur) | 24AACCK0029Q1ZI |  | 2025-12 | 575.32 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.47) - no GST impact |
| HR | 2026-03-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP |  | 2026-03 | 501.48 | DUPLICATE booking - Matched |  |
| HR | 2026-01-01 00:00:00 | Reliance Jio Infocomm Limited | 24AABCI6363G1ZP |  | 2026-01 | 501.48 | DUPLICATE booking - Matched |  |
| HR | 2025-07-01 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI |  | 2025-07 | 463.36 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.26) - no GST impact |
| HR | 2025-12-29 00:00:00 | Akp Sons Motors Pvt Ltd(Purcahse) | 24AAYCA5330J1ZE |  | 2025-12 | 450.72 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.26) - no GST impact |
| HR | 2025-12-02 00:00:00 | Maruti Auto Enterprise | 24AASPC6705G1ZT |  | 2025-12 | 421.78 | DUPLICATE booking - Matched |  |
| HR | 2025-12-11 00:00:00 | Reliable Diesels | 24AXUPK0979K1ZR |  | 2025-12 | 391.28 | DUPLICATE booking - Matched |  |
| HR | 2025-12-09 00:00:00 | Reliable Diesels | 24AXUPK0979K1ZR |  | 2025-12 | 391.28 | DUPLICATE booking - Matched |  |
| HR | 2025-10-17 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI |  | 2025-10 | 378.66 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.28) - no GST impact |
| HR | 2025-12-10 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 370.98 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -5.0) - no GST impact |
| HR | 2025-12-29 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI |  | 2025-12 | 361.34 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.18) - no GST impact |
| HR | 2025-12-10 00:00:00 | Maruti Auto Enterprise | 24AASPC6705G1ZT |  | 2025-12 | 345.52 | DUPLICATE booking - Matched |  |
| HR | 2025-12-24 00:00:00 | Akp Sons Motors Pvt Ltd(Purcahse) | 24AAYCA5330J1ZE |  | 2025-12 | 315.78 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.06) - no GST impact |
| HR | 2025-12-21 00:00:00 | National Enterprise | 24ACTPL0091R1ZX |  | 2025-12 | 315.00 | DUPLICATE booking - Matched |  |
| HR | 2025-09-15 00:00:00 | Mahalaxmi Enterprise | 24CDFPG9834G1ZI |  | 2025-09 | 310.80 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.2) - no GST impact |
| HR | 2025-12-31 00:00:00 | Shiv Enterprise (Drinking Water ) | 24AHAPP9298E1Z2 |  | 2025-12 | 291.00 | DUPLICATE booking - Matched |  |
| HR | 2026-02-17 00:00:00 | Mahalaxmi Enterprise | 24CDFPG9834G1ZI |  | 2026-02 | 272.24 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.36) - no GST impact |
| HR | 2025-12-02 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 237.24 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-20 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 237.24 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2026-01-03 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI |  | 2026-01 | 233.00 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.44) - no GST impact |
| HR | 2025-07-18 00:00:00 | Akp Sons Motors Pvt Ltd(Purcahse) | 24AAYCA5330J1ZE |  | 2025-07 | 216.64 | DUPLICATE booking - Timing Difference - ITC reflected in subsequent GSTR-2B | Timing Difference - ITC reflected in subsequent GSTR-2B (booked 2025-07; in GSTR-2B 2025-09); Taxable value differs (Boo |
| HR | 2025-12-09 00:00:00 | Reliable Diesels | 24AXUPK0979K1ZR |  | 2025-12 | 202.50 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.5) - no GST impact |
| HR | 2025-12-11 00:00:00 | Reliable Diesels | 24AXUPK0979K1ZR |  | 2025-12 | 202.50 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.5) - no GST impact |
| HR | 2025-12-21 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 195.12 | DUPLICATE booking - Matched (GST rounding diff <= Rs 1) | Taxable value differs (Books-2B: -3.5) - no GST impact |
| HR | 2025-11-04 00:00:00 | Kataria Motors Pvt Ltd-Chacharwadi (Pur) | 24AACCK0029Q1ZI |  | 2025-11 | 175.34 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.47) - no GST impact |
| HR | 2025-10-14 00:00:00 | Kataria Motors Pvt Ltd -Valsad(Pur) | 24AACCK0029Q1ZI |  | 2025-10 | 172.50 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.18) - no GST impact |
| HR | 2025-12-01 00:00:00 | Maruti Auto Enterprise | 24AASPC6705G1ZT |  | 2025-12 | 170.68 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.09) - no GST impact |
| HR | 2025-12-30 00:00:00 | Maruti Auto Enterprise | 24AASPC6705G1ZT |  | 2025-12 | 170.68 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.09) - no GST impact |
| HR | 2025-12-02 00:00:00 | National Enterprise | 24ACTPL0091R1ZX |  | 2025-12 | 153.00 | DUPLICATE booking - Matched |  |
| HR | 2025-10-07 00:00:00 | Spirited Motor Vehicles Limited (Purchase)Vasai | 27ABBCS9610H1Z9 |  | 2025-10 | 142.05 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.22) - no GST impact |
| HR | 2025-12-30 00:00:00 | Kataria Motors Pvt Ltd-Chacharwadi (Pur) | 24AACCK0029Q1ZI |  | 2025-12 | 131.04 | DUPLICATE booking - Matched |  |
| HR | 2025-12-04 00:00:00 | Diesel World (Pvt.) Ltd. | 24AAECD2472J1ZQ |  | 2025-12 | 129.60 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-02 00:00:00 | Diesel World (Pvt.) Ltd. | 24AAECD2472J1ZQ |  | 2025-12 | 129.60 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-09 00:00:00 | Diesel World (Pvt.) Ltd. | 24AAECD2472J1ZQ |  | 2025-12 | 129.60 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-22 00:00:00 | National Enterprise | 24ACTPL0091R1ZX |  | 2025-12 | 129.60 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.4) - no GST impact |
| HR | 2025-12-11 00:00:00 | National Enterprise | 24ACTPL0091R1ZX |  | 2025-12 | 129.60 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.4) - no GST impact |
| HR | 2026-03-17 00:00:00 | Diesel World (Pvt.) Ltd. | 24AAECD2472J1ZQ |  | 2026-03 | 129.60 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-08-21 00:00:00 | Kataria Motors Pvt Ltd -Kamrej (Pur) | 24AACCK0029Q1ZI |  | 2025-08 | 101.96 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -0.43) - no GST impact |
| HR | 2025-12-01 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 90.00 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-11 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 90.00 | DUPLICATE booking - In Books, not in GSTR-2B | To be verified |
| HR | 2025-12-22 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 88.92 | DUPLICATE booking - Matched |  |
| HR | 2025-11-05 00:00:00 | Kataria Motors Pvt Ltd-Kachchh(Pur) | 24AACCK0029Q1ZI |  | 2025-11 | 78.70 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: 0.08) - no GST impact |
| HR | 2025-12-05 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT |  | 2025-12 | 54.00 | DUPLICATE booking - Matched (taxable value differs - no GST impact) | Taxable value differs (Books-2B: -1.0) - no GST impact |

### 8. Other / to be verified (small value, blank invoice numbers etc.)  **(0 rows)**

| Branch | Book Date | Supplier | GSTIN | Invoice No | Book Mth | Total GST | Status |
|---|---|---|---|---|---|---|---|

### Debit/Credit Note exceptions  **(10 rows)**

| Branch | Date | Supplier | GSTIN | Note No | Book Mth | CN 2B Mth | IGST | CGST | SGST | Total GST | Status | Remarks |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| HR | 2025-05-01 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/2 | 2025-05 | 2025-04 | 0.00 | 587.39 | 587.39 | 1,174.78 | Credit Note - Timing Difference (earlier GSTR-2B) | Booked 2025-05; in GSTR-2B 2025-04 |
| HR | 2025-05-31 00:00:00 | Shree Ganesh Auto | 24AHOPC4366Q1ZT | 25-26/5 | 2025-05 | 2025-05 | 0.00 | 632.17 | 632.17 | 1,264.34 | Credit Note - GST Mismatch | Books 1264.34 vs GSTR-2B 403.48 - to be verified |
| HR | 2025-11-30 00:00:00 | Maruti Auto Enterprise | 24AASPC6705G1ZT | CN27 | 2025-11 |  | 0.00 | 0.00 | 0.00 | 0.00 | Debit/Credit Note - no GST amount in books | Verify - RCM or unregistered |
| HR | 2026-03-31 00:00:00 | Viral Enterprise | 24AALFV6612Q3ZH | 310326 | 2026-03 |  | 0.00 | 0.00 | 0.00 | 0.00 | Debit/Credit Note - no GST amount in books | Verify - RCM or unregistered |
| PL | 2026-03-31 00:00:00 | Dhyani Enterprise | 24ARSPP0690G2ZG | 30.09.25 | 2026-03 |  | 0.00 | 1,769.81 | 1,769.81 | 3,539.62 | In Books, not in GSTR-2B (Credit Note) | To be verified |
| VL | 2025-11-10 00:00:00 | D.I.C.V. Pvt Ltd(Amc/goodwill/ Warranty/sm/Rsa/cou) | 33AABCF1590N1ZJ | 9132502849 | 2025-11 |  | 3,448.70 | 0.00 | 0.00 | 3,448.70 | DUPLICATE booking - In Books, not in GSTR-2B (Credit Note) | To be verified |
| VL | 2025-11-26 00:00:00 | Mona Anandkumar Shah [Rent] [Valsad] |  | 1/25-26 | 2025-11 |  | 0.00 | 0.00 | 0.00 | 0.00 | Debit/Credit Note - no GST amount in books | Verify - RCM or unregistered |
| VL | 2025-12-23 00:00:00 | D.I.C.V. Pvt Ltd(Amc/goodwill/ Warranty/sm/Rsa/cou) | 33AABCF1590N1ZJ | 9132503668 | 2025-12 | 2026-02 | 181.41 | 0.00 | 0.00 | 181.41 | DUPLICATE booking - Credit Note - Timing Difference (subsequent GSTR-2B) | Booked 2025-12; in GSTR-2B 2026-02 |
| VL | 2026-02-25 00:00:00 | D.I.C.V. Pvt Ltd(Amc/goodwill/ Warranty/sm/Rsa/cou) | 33AABCF1590N1ZJ | 9132503668 | 2026-02 |  | 181.41 | 0.00 | 0.00 | 181.41 | DUPLICATE booking - In Books, not in GSTR-2B (Credit Note) | To be verified |
| VL | 2026-03-31 00:00:00 | Dhyani Enterprise | 24ARSPP0690G2ZG | 71125 | 2026-03 |  | 0.00 | 345.15 | 345.15 | 690.30 | In Books, not in GSTR-2B (Credit Note) | To be verified |

## 5. Debit / Credit Note reconciliation

Full listing in Excel sheet **D. Debit/Credit Notes**.

| Status | Notes |
|---|---|
| Credit Note - Matched | 49 |
| Debit/Credit Note - no GST amount in books | 3 |
| In Books, not in GSTR-2B (Credit Note) | 2 |
| DUPLICATE booking - In Books, not in GSTR-2B (Credit Note) | 2 |
| Credit Note - Timing Difference (earlier GSTR-2B) | 1 |
| Credit Note - GST Mismatch | 1 |
| DUPLICATE booking - Credit Note - Matched | 1 |
| Credit Note - Matched (GSTIN differs) | 1 |
| DUPLICATE booking - Credit Note - Timing Difference (subsequent GSTR-2B) | 1 |

Debit notes as per books: IGST 88,339.98, CGST 13,919.10, SGST 13,919.10, Total 116,178.18. Credit notes as per GSTR-2B: IGST 84,709.99, CGST 217,208.33, SGST 217,208.33, Total 519,126.65.

## 6. Timing Differences (month-wise movement only - no impact on FY total)

Total timing movement: Rs 50,537.04 = current-month ITC appearing in subsequent GSTR-2B month Rs 43,966.73 + previous-month ITC received in current GSTR-2B month Rs 5,214.12 + debit/credit-note timing Rs 1,356.19. Full listing in Excel sheet **E. Timing Differences**.

## 7. Why Books and GSTR-2B differ - final explanation

ITC as per books (net) is Rs 29,194,700.35; ITC as per GSTR-2B (net) is Rs 29,419,427.07 - GSTR-2B is higher by Rs 224,726.72. The difference is fully explained:

| # | Component | Amount (Rs) | Nature |
|---|---|---|---|
| 1 | ITC in GSTR-2B not booked in Books (incl. RCM and invoices rejected on IMS) | 745,939.85 | Supplier reported but not in books - verify and book ITC where valid |
| 2 | ITC in GSTR-2B not claimed in Books (invoices booked without GST split) | 430,435.00 | Books carry no ITC on these invoices - claim/regularise ITC |
| 3 | Credit notes in GSTR-2B without matching debit note in Books | -411,669.24 | ITC reduction in 2B with no book entry - investigate |
| 4 | ITC booked in Books not in GSTR-2B (supplier not reported / RCM / to follow up) | -548,699.86 | ITC at risk - obtain supplier GSTR-1 report |
| 5 | Debit notes in Books without matching credit note in GSTR-2B | 7,860.03 | ITC reduction booked with no 2B credit note - verify |
| 6 | Net GST difference on matched invoices (value mismatches + rounding) | 0.20 | Gross mismatches listed in Exception Report item 4 - verify with suppliers |
| 7 | Net debit/credit-note value difference on matched pairs | 860.74 | Minor DN/CN value differences - see DN section |

**Timing vs other:** Of the total difference, **Rs 50,537.04 of invoice-level movement is timing only** (booked in one month, ITC in another - it cancels within the year and does not create a loss or excess). The balance of **Rs 224,726.72 is net of all timing effects** and is dominated by (a) ITC in GSTR-2B not booked/claimed (+Rs 1,176,374.85) largely offset by (b) ITC booked in books not appearing in GSTR-2B (−Rs 548,699.86, mostly Daimler India invoices not yet reported by the supplier and several small suppliers) and (c) credit notes in 2B without matching debit notes (−Rs 411,669.24).

**Action priorities:** (1) follow up suppliers for the book-side ITC not in GSTR-2B (largest: Daimler India ~Rs 4.8L IGST on PO-number invoices); (2) regularise the 49 invoices booked without GST split where 2B shows ITC (Rs 4.3L); (3) verify the 40 GST-value mismatches with suppliers; (4) match the 6 credit notes without book debit notes (Rs 4.1L, mostly Parth Cement CN/02-04 and Sunrise); (5) book/reverse the 13 invoices rejected on IMS as per final IMS position.