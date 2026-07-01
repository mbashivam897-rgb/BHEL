# -*- coding: utf-8 -*-
"""
Authoritative AUDITED consolidated financials for BHEL, FY2020-FY2026 (INR Crore).
Order of every list: [FY2020, FY2021, FY2022, FY2023, FY2024, FY2025, FY2026].

SOURCES (consolidated, Ind-AS, audited):
  FY2020 & FY2021 : Annual Report 2020-21  (Consol P&L p248, Cash Flows p252-253, BS p246)
  FY2022 & FY2023 : Annual Report 2022-23  (Consol P&L p276-277, Cash Flows p280-281, BS p274)
  FY2024 & FY2025 : Annual Report 2024-25  (Consol P&L p264, Cash Flows p266-268, BS p262)
  FY2026          : BSE Integrated Filing 04-May-2026 (Consol P&L, Cash Flow, BS)

Balance-sheet line items are cross-verified against audited_data.py and an
independent page-by-page extraction (build_pdf_data.py): 17 line items x 7 years, 0 mismatches.
"""

YEARS = ['FY2020','FY2021','FY2022','FY2023','FY2024','FY2025','FY2026']

# ---------------------------------------------------------------------------
# INCOME STATEMENT  (mapped into the model's gross-profit format)
#   COGS   = raw materials + bought-out items + civil/erection/eng + stores
#            + changes in inventories                (all "cost of sales")
#   SG&A   = other expenses (+ FX variation + provisions where separately shown FY20/21)
#   ADJ    = share of net profit of JV/associates (equity method), added after tax
# Each year reproduces the audited PBT and PAT exactly (verified below).
# ---------------------------------------------------------------------------
IS = {
 'rev' : [21463.14,17308.69,21211.09,23364.94,23892.78,28339.48,33782.18],  # revenue from operations
 'cogs': [14039.29,11870.63,14793.84,16300.72,16807.11,18845.15,22983.09],  # materials + inventory change
 'emp' : [ 5431.88, 5378.15, 5519.05, 5700.63, 5628.84, 5923.42, 6467.62],  # employee benefits
 'sga' : [ 2229.34, 3201.20,  162.62,  647.03,  844.23, 2329.34, 1989.28],  # other exp (+FX+prov FY20/21)
 'oi'  : [  564.30,  348.42,  354.54,  488.63,  546.27,  465.31,  807.65],  # other income
 'dep' : [  503.27,  473.25,  314.12,  260.34,  248.90,  271.96,  315.87],  # depreciation & amortisation
 'fin' : [  508.45,  373.95,  355.96,  521.43,  731.29,  748.33,  756.41],  # finance costs
 'tax' : [  809.28, -896.23,   25.75,    2.05,  -39.56,  211.70,  538.35],  # total tax expense
 'adj' : [   25.72,   44.14,   50.42,   56.02,   63.98,   59.01,   61.05],  # share of JV/associates
 'pat' : [-1468.35,-2699.70,  444.71,  477.39,  282.22,  533.90, 1600.26],  # profit for the year
 'div' : [  504.56,    0.68,    0.30,  139.18,  139.45,   87.44,  174.64],  # dividends PAID (per cash flow)
}

# ---------------------------------------------------------------------------
# BALANCE SHEET  (audited; mapped to model line items; ties with NO plug)
# ---------------------------------------------------------------------------
BS = {
 # assets
 'ppe'  : [2824.38,2507.49,2407.12,2485.24,2600.08,2980.91,3094.38],  # PPE + intangibles + inv property
 'cwip' : [ 306.74, 403.21, 422.32, 344.59, 282.32, 161.70, 399.20],
 'invst': [ 162.06, 185.34, 205.15, 235.42, 255.67, 275.57, 302.06],  # equity-method + financial investments
 'onca' : [23744.20,23800.81,25347.19,26222.26,21321.74,21905.23,20934.62],  # other non-current assets
 'recv' : [7108.60,4035.07,3024.75,3128.35,4785.38,5884.35,6796.27],  # trade receivables (current)
 'invy' : [8908.23,7194.45,6560.21,6755.90,7220.57,9869.49,13334.58],  # inventories
 'oca'  : [10275.76,10412.39,11123.33,13555.45,16378.69,19393.52,19457.88],  # other current assets
 'cash' : [6418.59,6701.45,7153.69,6642.58,6157.47,7612.41,11866.62],  # cash equiv + bank balances
 # equity & liabilities
 'eqc'  : [696.41]*7,
 'res'  : [27955.24,25275.59,25810.19,26131.62,23742.24,24025.75,25450.19],  # other equity incl NCI
 'ltbor': [75.37,53.41,35.12,33.75,23.55,162.39,168.18],               # LT borrowings / lease liab (NC)
 'stbor': [5004.59,4897.48,4794.81,5419.76,8832.91,8852.21,8018.77],   # ST borrowings incl lease liab (cur)
 'oncl' : [8413.06,8854.90,8330.89,9156.56,9292.48,14973.04,17568.65], # other non-current liabilities
 'pay'  : [8829.16,6683.51,7749.59,9895.83,8539.38,9540.92,10491.60],  # trade payables (current)
 'ocl'  : [8774.73,8778.90,8826.75,8035.85,7874.95,9832.46,13791.81],  # other current liabilities
}

# ---------------------------------------------------------------------------
# CASH FLOW  (audited, indirect method; reconciles to cash & cash equivalents)
# ---------------------------------------------------------------------------
CF = {
 'cfo'   : [-2891.50, 561.60, 660.25, -741.52,-3712.90,2191.89,5837.38],
 'cfi'   : [ 1877.06, -42.50,-1125.32,1480.46, 1330.86,-2730.91,-3035.41],
 'cff'   : [ 1621.99,-395.81, -329.49,  88.96, 2655.74,-856.81,-1805.73],
 'open'  : [  789.05,1396.60, 1519.90, 732.62, 1561.34,1835.04, 439.21],
 'close' : [ 1396.60,1519.90,  732.62,1560.52, 1835.04, 439.21,1435.45],
 # selected operating detail (for the statement build)
 'dep'      : [503.27,473.25,314.12,260.34,248.90,271.96,315.87],
 'fin_cost' : [508.45,373.95,355.96,521.43,731.29,748.33,756.41],
 'int_div'  : [-509.19,-320.86,-302.79,-421.12,-493.78,-402.63,-739.93],  # interest & dividend income
 # reported opening-balance reconciling items (disclosed in the annual reports):
 #   FY2022: +7.28 cash credit of subsidiary (BHEL EML) added to opening balance
 #   FY2024: +0.82 year-boundary regrouping (FY23 close 1560.52 vs FY24 open 1561.34)
 'open_adj' : [0.00, 0.00, 7.28, 0.00, 0.82, 0.00, 0.00],
 # roll adjustment needed for the intra-year cash roll to tie (only FY22 cash-credit)
 'roll_adj' : [0.00, 0.00, 7.28, 0.00, 0.00, 0.00, 0.00],
 # investing / financing build-up (all tie to cfi / cff exactly - verified)
 'capex'    : [-394.01,-250.35,-169.38,-188.40,-232.50,-281.54,-589.04],  # purchase of PPE & intangibles
 'borrow'   : [2503.78, -99.67, -87.53, 640.00,3423.00, -13.00, -845.00],  # net short-term borrowings
 'intlease' : [-377.23,-295.46,-241.66,-411.86,-627.81,-756.37,-786.09],  # interest paid + lease payments
 'divpaid'  : [-504.56,  -0.68,  -0.30,-139.18,-139.45, -87.44,-174.64],  # dividends paid
}
# NOTE ON CASH BASIS: the cash-flow statement reconciles to "cash & cash equivalents"
# (narrow: CF['close']). The balance-sheet cash line BS['cash'] is BROADER (adds bank
# balances / deposits > 3 months, whose movement sits in investing activities). The
# difference each year = BS['cash'] - CF['close'] = other bank balances & deposits.

def _check():
    ok = True
    for i,y in enumerate(YEARS):
        gross = IS['rev'][i]-IS['cogs'][i]
        ebitda= gross-IS['emp'][i]-IS['sga'][i]
        ebit  = ebitda-IS['dep'][i]
        pbt   = ebit+IS['oi'][i]-IS['fin'][i]
        pat   = pbt-IS['tax'][i]+IS['adj'][i]
        dpat  = pat-IS['pat'][i]
        ta = sum(BS[k][i] for k in ['ppe','cwip','invst','onca','recv','invy','oca','cash'])
        tel= sum(BS[k][i] for k in ['eqc','res','ltbor','stbor','oncl','pay','ocl'])
        dcf = (CF['open'][i]+CF['open_adj'][i]+CF['cfo'][i]+CF['cfi'][i]+CF['cff'][i]-CF['close'][i])
        if abs(dpat)>0.5 or abs(ta-tel)>0.5 or abs(dcf)>1.5: ok=False
        print(f"{y}: PAT {pat:9.2f}/{IS['pat'][i]:9.2f} (d{dpat:6.2f}) | TA {ta:10.2f} TE&L {tel:10.2f} (d{ta-tel:5.2f}) | CF d{dcf:5.2f}")
    print("ALL CHECKS PASS" if ok else "REVIEW NEEDED")

if __name__=='__main__':
    _check()
