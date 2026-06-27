# -*- coding: utf-8 -*-
"""
Institutional-grade integrated financial model for Bharat Heavy Electricals Ltd (BHEL)
Historical: FY2020-FY2026 (consolidated, sourced from BSE filings via Screener.in)
Forecast:   FY2027-FY2033 (order-book-driven, fully linked 3-statement model)

Conventions (Wall Street / IB):
  - BLACK font  = hardcoded inputs
  - BLUE  font  = in-sheet formulas
  - GREEN font  = links to other sheets
  - Grey fill   = input cells
All figures in INR Crore unless stated. Per-share in INR.
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.chart.axis import ChartLines

wb = Workbook()
wb.remove(wb.active)

# ----------------------------------------------------------------------------
# Column / year layout
# ----------------------------------------------------------------------------
HIST = ['FY2020','FY2021','FY2022','FY2023','FY2024','FY2025','FY2026']
FCST = ['FY2027','FY2028','FY2029','FY2030','FY2031','FY2032','FY2033']
YEARS = HIST + FCST
HCOLS = list(range(2, 9))     # B..H  (FY20..FY26)
FCOLS = list(range(9, 16))    # I..O  (FY27..FY33)
ACOLS = HCOLS + FCOLS         # B..O
def CL(c): return get_column_letter(c)

# ----------------------------------------------------------------------------
# Styles
# ----------------------------------------------------------------------------
NAVY   = '1F3864'
MIDBL  = '2E5496'
LTBL   = 'D9E1F2'
LTBL2  = 'EAF0FA'
GREY   = 'E7E6E6'
INPUTF = 'FFF2CC'   # input cells (light gold = grey-ish convention)
GREEN_F= 'E2EFDA'
RED_F  = 'F8CBAD'
WHITE  = 'FFFFFF'

F_BLACK  = Font(name='Calibri', size=10, color='000000')
F_BLUE   = Font(name='Calibri', size=10, color='0000CC')
F_GREEN  = Font(name='Calibri', size=10, color='008000')
F_LBL    = Font(name='Calibri', size=10, color='000000')
F_LBLB   = Font(name='Calibri', size=10, color='000000', bold=True)
F_HDRW   = Font(name='Calibri', size=10, color='FFFFFF', bold=True)
F_TITLE  = Font(name='Calibri', size=16, color='FFFFFF', bold=True)
F_SUB    = Font(name='Calibri', size=11, color='FFFFFF')
F_SECT   = Font(name='Calibri', size=11, color=NAVY, bold=True)
F_NOTE   = Font(name='Calibri', size=8,  color='808080', italic=True)
F_ITAL   = Font(name='Calibri', size=9,  color='595959', italic=True)

FILL_HDR  = PatternFill('solid', fgColor=NAVY)
FILL_SUB  = PatternFill('solid', fgColor=LTBL)
FILL_SUB2 = PatternFill('solid', fgColor=LTBL2)
FILL_IN   = PatternFill('solid', fgColor=INPUTF)
FILL_GREY = PatternFill('solid', fgColor=GREY)
FILL_TOT  = PatternFill('solid', fgColor='DDEBF7')
FILL_OK   = PatternFill('solid', fgColor=GREEN_F)
FILL_BAD  = PatternFill('solid', fgColor=RED_F)
FILL_TITLE= PatternFill('solid', fgColor=NAVY)

thin = Side(style='thin', color='BFBFBF')
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
B_TOP = Border(top=Side(style='thin', color='000000'))
B_TB  = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

FMT_CR   = '#,##0;(#,##0)'
FMT_CR1  = '#,##0.0;(#,##0.0)'
FMT_PCT  = '0.0%'
FMT_PCT2 = '0.00%'
FMT_PS   = '0.00'
FMT_X    = '0.0"x"'
FMT_DAYS = '0'

CEN = Alignment(horizontal='center', vertical='center')
RGT = Alignment(horizontal='right')
LFT = Alignment(horizontal='left')
WRP = Alignment(horizontal='left', vertical='top', wrap_text=True)

# ----------------------------------------------------------------------------
# Registry: REG[sheet_title][key] = row number  -> robust cross-sheet refs
# ----------------------------------------------------------------------------
REG = {}
def reg(ws, key, row):
    REG.setdefault(ws.title, {})[key] = row
def XR(title, key, c):
    return "'%s'!%s%d" % (title, CL(c), REG[title][key])

# ----------------------------------------------------------------------------
# Generic cell writer
# ----------------------------------------------------------------------------
def put(ws, r, c, v, font=F_BLACK, fmt=None, fill=None, align=None, border=None, bold=False):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name=font.name, size=font.size, color=font.color.rgb if font.color else '000000', bold=(bold or font.bold), italic=font.italic)
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    if align: cell.alignment = align
    if border: cell.border = border
    return cell

def section(ws, r, text, span=15):
    for c in range(1, span+1):
        ws.cell(row=r, column=c).fill = FILL_HDR
    put(ws, r, 1, text, font=F_HDRW, fill=FILL_HDR)
    return r

def year_header(ws, r, hist_label=True):
    put(ws, r, 1, '', fill=FILL_SUB)
    for i, c in enumerate(ACOLS):
        put(ws, r, c, YEARS[i], font=F_LBLB, fill=FILL_SUB, align=CEN, border=B_ALL)
    # period tag row
    put(ws, r, 1, 'INR Crore', font=F_ITAL, fill=FILL_SUB)
    return r

def ah_label(ws, r, text, bold=False, indent=0, italic=False):
    f = F_LBLB if bold else (F_ITAL if italic else F_LBL)
    cell = put(ws, r, 1, ('   '*indent)+text, font=f)
    return cell

def setw(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

def std_widths(ws, first=44):
    ws.column_dimensions['A'].width = first
    for c in ACOLS:
        ws.column_dimensions[CL(c)].width = 11.5

def freeze(ws, cell='B5'):
    ws.freeze_panes = cell

print("Framework loaded.")


# ============================================================================
# 1. COVER PAGE
# ============================================================================
cv = wb.create_sheet('Cover Page')
cv.sheet_view.showGridLines = False
setw(cv, {'A':3,'B':30,'C':24,'D':24,'E':24,'F':22,'G':20})
for r in range(1, 4):
    for c in range(1, 8):
        cv.cell(row=r, column=c).fill = FILL_TITLE
put(cv, 2, 2, 'BHARAT HEAVY ELECTRICALS LIMITED (BHEL)', font=F_TITLE, fill=FILL_TITLE)
put(cv, 3, 2, 'Integrated Financial Model & Equity Valuation  |  NSE: BHEL  |  BSE: 500103', font=F_SUB, fill=FILL_TITLE)

put(cv, 5, 2, 'Prepared on the basis of:', font=F_SECT)
meta = [
 ('Company','Bharat Heavy Electricals Ltd (Maharatna CPSE, Govt. of India)'),
 ('Sector','Capital Goods / Heavy Electrical Equipment'),
 ('Reporting basis','Consolidated, Indian Accounting Standards (Ind-AS)'),
 ('Currency / Units','INR Crore (per-share in INR)'),
 ('Historical period','FY2020 - FY2026 (7 years, audited / reported)'),
 ('Forecast period','FY2027 - FY2033 (7 years, analyst projections)'),
 ('Valuation date','27-Jun-2026'),
 ('Last traded price','INR 403 (25-Jun-2026)'),
 ('Shares outstanding','~348.0 crore (Face value INR 2)'),
 ('Promoter (Govt. of India)','58.17%'),
]
r = 6
for k, v in meta:
    put(cv, r, 2, k, font=F_LBLB); put(cv, r, 3, v, font=F_BLACK); r += 1

put(cv, r+1, 2, 'Workbook contents', font=F_SECT); r += 2
sheets_list = [
 '1  Cover Page','2  Assumptions','3  Historical Financial Statements','4  Historical Ratio Analysis',
 '5  Operational Drivers','6  Revenue Build-up','7  Cost Forecast','8  Working Capital Schedule',
 '9  Fixed Asset Schedule','10 Depreciation Schedule','11 Debt Schedule','12 Equity Schedule',
 '13 Other Assets & Liabilities','14 Cash Flow Statement','15 Forecast Financial Statements',
 '16 Ratio Analysis','17 DCF Valuation','18 Relative Valuation','19 Sensitivity Analysis',
 '20 Scenario Analysis','21 Dashboard','22 Error Checks']
rr = r
for i, s in enumerate(sheets_list):
    col = 2 + (i // 11)
    put(cv, rr + (i % 11), col, s, font=F_BLACK)
r = rr + 11 + 1

put(cv, r+1, 2, 'Formatting legend', font=F_SECT); r += 2
put(cv, r, 2, 'Black', font=F_BLACK); put(cv, r, 3, 'Hardcoded input / actual reported data', font=F_LBL)
put(cv, r+1, 2, 'Blue', font=F_BLUE); put(cv, r+1, 3, 'In-sheet calculation (formula)', font=F_LBL)
put(cv, r+2, 2, 'Green', font=F_GREEN); put(cv, r+2, 3, 'Link to another worksheet', font=F_LBL)
put(cv, r+3, 2, 'Gold fill', font=F_BLACK, fill=FILL_IN); put(cv, r+3, 3, 'Editable assumption / input cell', font=F_LBL)
r += 5

put(cv, r, 2, 'Data sources & provenance', font=F_SECT); r += 1
srcs = [
 'Primary headline financials (P&L, Balance Sheet, Cash Flow FY20-FY26): BHEL filings to BSE/NSE, aggregated via Screener.in.',
 'Order book, order inflow & segment data: BHEL results press releases, Directors\' Report and exchange filings (standalone basis).',
 'Market data (price, shares, promoter holding): exchange data as of 25-Jun-2026.',
 'CAVEAT: Detailed expense split (raw material / employee / subcontracting) and detailed balance-sheet sub-lines',
 '(inventory, receivables, payables, cash, contract assets/liabilities) are NOT separately disclosed in the aggregated',
 'sources used. These are shown as clearly-flagged analytical estimates anchored to the audited reported TOTALS.',
 'Verify against the company\'s audited Annual Report notes before any transactional use.',
]
for s in srcs:
    put(cv, r, 2, s, font=F_ITAL); r += 1
freeze(cv, 'A1')

# ============================================================================
# 2. ASSUMPTIONS
# ============================================================================
asm = wb.create_sheet('Assumptions')
asm.sheet_view.showGridLines = False
std_widths(asm, 46)
put(asm, 1, 1, 'BHEL  |  Key Modelling Assumptions  (editable inputs in gold)', font=F_TITLE if False else F_LBLB)
put(asm, 1, 1, 'KEY MODELLING ASSUMPTIONS', font=F_HDRW, fill=FILL_HDR)
for c in range(2, 16): asm.cell(row=1, column=c).fill = FILL_HDR
# year header for forecast block
put(asm, 3, 1, 'Operating & financial drivers', font=F_SECT)
hr = 4
put(asm, hr, 1, 'INR Crore / ratios', font=F_ITAL, fill=FILL_SUB)
for i, c in enumerate(ACOLS):
    put(asm, hr, c, YEARS[i], font=F_LBLB, fill=FILL_SUB, align=CEN, border=B_ALL)

def aset(key, label, vals7, fmt, note='', indent=0):
    global rowA
    r = rowA
    ah_label(asm, r, label, indent=indent)
    for j, c in enumerate(FCOLS):
        put(asm, r, c, vals7[j], font=F_BLACK, fmt=fmt, fill=FILL_IN, align=RGT, border=B_ALL)
    if note: put(asm, r, 16, note, font=F_NOTE)
    reg(asm, key, r)
    rowA += 1
    return r

rowA = 5
aset('exec_rate', 'Execution rate (revenue / opening order book)', [0.165,0.170,0.175,0.180,0.185,0.190,0.195], FMT_PCT)
aset('order_inflow', 'New order inflow', [78000,82000,86000,90000,92000,94000,96000], FMT_CR)
aset('ebitda_margin', 'EBITDA margin (% revenue)', [0.075,0.085,0.095,0.105,0.110,0.115,0.120], FMT_PCT)
aset('oth_inc_pct', 'Other income (% revenue)', [0.018]*7, FMT_PCT)
aset('emp_pct', 'Employee benefits expense (% revenue)', [0.135,0.128,0.122,0.116,0.112,0.108,0.105], FMT_PCT)
aset('subc_pct', 'Subcontracting & erection (% revenue)', [0.085]*7, FMT_PCT)
aset('oopex_pct', 'Other operating expenses (% revenue)', [0.075]*7, FMT_PCT)
aset('capex', 'Capital expenditure', [600,700,750,800,850,900,950], FMT_CR)
aset('dep_rate', 'Depreciation rate (% opening net block)', [0.095]*7, FMT_PCT)
aset('recv_days', 'Trade receivable days', [75]*7, FMT_DAYS)
aset('inv_days', 'Inventory days', [210,205,200,200,195,195,190], FMT_DAYS)
aset('pay_days', 'Trade payable days', [175]*7, FMT_DAYS)
aset('ca_pct', 'Contract assets (% revenue)', [0.235]*7, FMT_PCT)
aset('cl_pct', 'Contract liabilities / advances (% revenue)', [0.355]*7, FMT_PCT)
aset('oca_g', 'Other current/non-curr assets growth %', [0.03]*7, FMT_PCT)
aset('ol_g', 'Provisions & other liabilities growth %', [0.03]*7, FMT_PCT)
aset('cwip', 'Capital work-in-progress (closing)', [400]*7, FMT_CR)
aset('investments', 'Investments (closing)', [310,320,330,340,350,360,370], FMT_CR)
aset('tax_rate', 'Effective tax rate', [0.25]*7, FMT_PCT)
aset('div_payout', 'Dividend payout (% PAT)', [0.30]*7, FMT_PCT)
aset('new_borrow', 'New borrowings drawn', [0]*7, FMT_CR)
aset('debt_repay', 'Debt repayment', [1000,1000,1000,1000,1000,1000,1000], FMT_CR)
aset('int_rate', 'Interest rate on opening debt', [0.085]*7, FMT_PCT)

# WACC & valuation single-value inputs
wr = rowA + 1
put(asm, wr, 1, 'WACC & valuation inputs', font=F_SECT); wr += 1
def asingle(key, label, value, fmt, note=''):
    global wr
    ah_label(asm, wr, label)
    put(asm, wr, 3, value, font=F_BLACK, fmt=fmt, fill=FILL_IN, align=RGT, border=B_ALL)
    if note: put(asm, wr, 5, note, font=F_NOTE)
    reg(asm, key, wr); wr += 1

asingle('rf', 'Risk-free rate (10Y India G-Sec)', 0.068, FMT_PCT)
asingle('beta', 'Levered beta', 1.25, FMT_PCT2.replace('%','') if False else '0.00')
asingle('erp', 'Equity risk premium', 0.065, FMT_PCT)
asingle('kd', 'Pre-tax cost of debt', 0.085, FMT_PCT)
asingle('tax_w', 'Marginal tax rate (WACC)', 0.25, FMT_PCT)
asingle('wd', 'Target weight of debt', 0.15, FMT_PCT)
asingle('we', 'Target weight of equity', 0.85, FMT_PCT)
asingle('term_g', 'Terminal growth rate', 0.045, FMT_PCT)
# computed
ah_label(asm, wr, 'Cost of equity  (Ke = Rf + B x ERP)', bold=True)
put(asm, wr, 3, "=%s+%s*%s" % (XR('Assumptions','rf',3), XR('Assumptions','beta',3), XR('Assumptions','erp',3)), font=F_BLUE, fmt=FMT_PCT, align=RGT, border=B_ALL)
reg(asm, 'ke', wr); wr += 1
ah_label(asm, wr, 'WACC  (We x Ke + Wd x Kd x (1-t))', bold=True)
put(asm, wr, 3, "=%s*%s+%s*%s*(1-%s)" % (XR('Assumptions','we',3),XR('Assumptions','ke',3),XR('Assumptions','wd',3),XR('Assumptions','kd',3),XR('Assumptions','tax_w',3)), font=F_BLUE, fmt=FMT_PCT, fill=FILL_TOT, align=RGT, border=B_ALL)
reg(asm, 'wacc', wr); wr += 1

asingle('shares', 'Shares outstanding (crore)', 348.05, FMT_CR1)
asingle('price', 'Current market price (INR)', 403.0, FMT_PS)
asingle('fv', 'Face value (INR)', 2.0, FMT_PS)

put(asm, wr+1, 1, 'Note: Forecast drivers are analyst base-case assumptions. Rationale is documented on the "Operational Drivers" and "Revenue Build-up" sheets and in the delivery notes.', font=F_ITAL)
freeze(asm, 'B5')
print("Cover + Assumptions built.")


# ============================================================================
# 3. HISTORICAL FINANCIAL STATEMENTS  (consolidated, reported)
# ============================================================================
hf = wb.create_sheet('Historical Financial Statements')
hf.sheet_view.showGridLines = False
std_widths(hf, 46)
put(hf, 1, 1, 'HISTORICAL FINANCIAL STATEMENTS  (Consolidated, INR Crore)', font=F_HDRW, fill=FILL_HDR)
for c in range(2, 16): hf.cell(row=1, column=c).fill = FILL_HDR

def hyear(ws, r):
    put(ws, r, 1, 'INR Crore', font=F_ITAL, fill=FILL_SUB)
    for i, c in enumerate(HCOLS):
        put(ws, r, c, HIST[i], font=F_LBLB, fill=FILL_SUB, align=CEN, border=B_ALL)

def hrow(ws, r, label, vals, fmt=FMT_CR, font=F_BLACK, bold=False, indent=0, fill=None, key=None, border=None):
    ah_label(ws, r, label, bold=bold, indent=indent)
    for j, c in enumerate(HCOLS):
        put(ws, r, c, vals[j], font=font, fmt=fmt, fill=fill, align=RGT, border=border)
    if key: reg(ws, key, r)

def hformula(ws, r, label, fn, fmt=FMT_CR, bold=False, indent=0, fill=None, key=None, border=None):
    ah_label(ws, r, label, bold=bold, indent=indent)
    for c in HCOLS:
        put(ws, r, c, fn(c), font=F_BLUE, fmt=fmt, fill=fill, align=RGT, border=border)
    if key: reg(ws, key, r)

# ---- Income Statement ----
section(hf, 3, 'A.  INCOME STATEMENT'); hyear(hf, 4)
rev = [21463,17309,21211,23365,23893,28339,33782]
oi  = [590,393,405,544,608,524,869]
opex= [21596,20357,20383,22321,23182,26940,31440]
dep = [503,473,314,260,249,272,316]
fin = [613,467,448,612,828,906,756]
tax = [809,-896,25,62,-39,212,539]
hrow(hf, 5, 'Revenue from operations', rev, key='rev', bold=True)
hrow(hf, 6, 'Other income', oi, key='oi')
hrow(hf, 7, 'Total operating expenditure', opex, key='opex', font=F_BLACK)
hformula(hf, 8, 'EBITDA', lambda c: "=%s%d-%s%d" % (CL(c),5,CL(c),7), key='ebitda', bold=True, fill=FILL_TOT, border=B_TOP)
hrow(hf, 9, 'Depreciation & amortisation', dep, key='dep')
hformula(hf, 10, 'EBIT', lambda c: "=%s%d-%s%d" % (CL(c),8,CL(c),9), key='ebit', bold=True)
hrow(hf, 11, 'Finance costs', fin, key='fin')
hformula(hf, 12, 'Profit before tax', lambda c: "=%s%d+%s%d-%s%d" % (CL(c),10,CL(c),6,CL(c),11), key='pbt', bold=True, border=B_TOP)
hrow(hf, 13, 'Tax expense', tax, key='tax')
hformula(hf, 14, 'Profit after tax (PAT)', lambda c: "=%s%d-%s%d" % (CL(c),12,CL(c),13), key='pat', bold=True, fill=FILL_TOT, border=B_TB)
hformula(hf, 15, 'EPS (INR)', lambda c: "=%s%d/%s" % (CL(c),14, XR('Assumptions','shares',3)), fmt=FMT_PS, key='eps')
div = [0,0,138,137,87,176,487]
hrow(hf, 16, 'Dividend (paid on PAT)', div, key='div')
put(hf, 17, 1, 'Source: BHEL BSE/NSE filings (consolidated). EBITDA = Revenue less total operating expenditure (excl. other income).', font=F_NOTE)

# ---- Balance Sheet ----
section(hf, 19, 'B.  BALANCE SHEET'); hyear(hf, 20)
ppe = [2817,2491,2398,2476,2574,2947,3094]
cwip= [314,420,431,354,308,195,399]
inv = [162,185,205,235,256,276,302]
oth_a=[56998,52864,53956,54599,56646,65431,72390]
hrow(hf, 21, 'Net property, plant & equipment', ppe, key='ppe')
hrow(hf, 22, 'Capital work-in-progress', cwip, key='cwip')
hrow(hf, 23, 'Investments', inv, key='inv')
hrow(hf, 24, 'Other assets (inventory, receivables, cash, contract & other)', oth_a, key='otha')
hformula(hf, 25, 'TOTAL ASSETS', lambda c: "=SUM(%s21:%s24)" % (CL(c),CL(c)), key='ta', bold=True, fill=FILL_TOT, border=B_TB)
eqc = [696,696,696,696,696,696,696]
res = [27964,25287,25810,23682,23742,24026,25450]
bor = [5080,4951,4830,5454,8856,9015,8187]
othl= [26550,25025,25654,27832,26489,35112,41852]
hrow(hf, 26, 'Equity share capital', eqc, key='eqc')
hrow(hf, 27, 'Other equity (reserves & surplus)', res, key='res')
hrow(hf, 28, 'Borrowings (incl. lease liabilities)', bor, key='bor')
hrow(hf, 29, 'Other liabilities (payables, advances, provisions)', othl, key='othl')
hformula(hf, 30, 'TOTAL EQUITY & LIABILITIES', lambda c: "=SUM(%s26:%s29)" % (CL(c),CL(c)), key='tle', bold=True, fill=FILL_TOT, border=B_TB)
hformula(hf, 31, 'Balance check (TA - TE&L)', lambda c: "=%s25-%s30" % (CL(c),CL(c)), key='bschk')
put(hf, 32, 1, 'Reported aggregated structure (consolidated). Sub-line detail of "Other assets/liabilities" shown as estimates on schedules.', font=F_NOTE)

# ---- Cash Flow ----
section(hf, 34, 'C.  CASH FLOW STATEMENT (reported summary)'); hyear(hf, 35)
cfo = [-2892,560,660,-741,-3713,2192,5837]
cfi = [1877,-42,-1118,1480,1331,-2731,-3035]
cff = [1622,-394,-329,89,2656,-857,-1806]
hrow(hf, 36, 'Cash flow from operating activities', cfo, key='cfo')
hrow(hf, 37, 'Cash flow from investing activities', cfi, key='cfi')
hrow(hf, 38, 'Cash flow from financing activities', cff, key='cff')
hformula(hf, 39, 'Net increase / (decrease) in cash', lambda c: "=SUM(%s36:%s38)" % (CL(c),CL(c)), key='netcf', bold=True, fill=FILL_TOT, border=B_TB)
put(hf, 40, 1, 'Source: BHEL consolidated cash flow statements (BSE/NSE filings).', font=F_NOTE)
freeze(hf, 'B5')
print("Historical Financial Statements built.")


# ============================================================================
# 4. HISTORICAL RATIO ANALYSIS
# ============================================================================
HFS = 'Historical Financial Statements'
hr = wb.create_sheet('Historical Ratio Analysis')
hr.sheet_view.showGridLines = False
std_widths(hr, 46)
put(hr, 1, 1, 'HISTORICAL RATIO ANALYSIS  (Consolidated)', font=F_HDRW, fill=FILL_HDR)
for c in range(2, 16): hr.cell(row=1, column=c).fill = FILL_HDR
hyear(hr, 3)

def lnk(c, key, sheet=HFS):  # green link to a stored row
    return "%s!%s%d" % ("'"+sheet+"'", CL(c), REG[sheet][key])

def rrow(ws, r, label, fn, fmt, bold=False, key=None, start=2, fill=None):
    ah_label(ws, r, label, bold=bold)
    for c in HCOLS:
        if c < start:
            put(ws, r, c, '', align=RGT)
        else:
            put(ws, r, c, fn(c), font=F_GREEN, fmt=fmt, align=RGT, fill=fill)
    if key: reg(ws, key, r)

section(hr, 4, 'Growth & profitability')
rrow(hr, 5, 'Revenue growth %', lambda c: "=%s/%s-1" % (lnk(c,'rev'), lnk(c-1,'rev')), FMT_PCT, start=3)
rrow(hr, 6, 'EBITDA margin %', lambda c: "=%s/%s" % (lnk(c,'ebitda'), lnk(c,'rev')), FMT_PCT)
rrow(hr, 7, 'EBIT margin %', lambda c: "=%s/%s" % (lnk(c,'ebit'), lnk(c,'rev')), FMT_PCT)
rrow(hr, 8, 'PAT margin %', lambda c: "=%s/%s" % (lnk(c,'pat'), lnk(c,'rev')), FMT_PCT)
rrow(hr, 9, 'Effective tax rate %', lambda c: "=IF(%s=0,0,%s/%s)" % (lnk(c,'pbt'),lnk(c,'tax'), lnk(c,'pbt')), FMT_PCT)
section(hr, 11, 'Returns')
rrow(hr, 12, 'Return on equity (ROE) %', lambda c: "=%s/((%s+%s+%s+%s)/2)" % (lnk(c,'pat'), lnk(c,'eqc'),lnk(c,'res'),lnk(c-1,'eqc'),lnk(c-1,'res')), FMT_PCT, start=3)
rrow(hr, 13, 'Return on capital employed (ROCE) %', lambda c: "=%s/(%s+%s+%s)" % (lnk(c,'ebit'), lnk(c,'eqc'),lnk(c,'res'),lnk(c,'bor')), FMT_PCT)
rrow(hr, 14, 'Return on assets (ROA) %', lambda c: "=%s/%s" % (lnk(c,'pat'), lnk(c,'ta')), FMT_PCT)
section(hr, 16, 'Leverage & liquidity')
rrow(hr, 17, 'Debt / equity (x)', lambda c: "=%s/(%s+%s)" % (lnk(c,'bor'), lnk(c,'eqc'),lnk(c,'res')), FMT_X)
rrow(hr, 18, 'Interest coverage (EBIT/interest, x)', lambda c: "=%s/%s" % (lnk(c,'ebit'), lnk(c,'fin')), FMT_X)
rrow(hr, 19, 'Asset turnover (x)', lambda c: "=%s/%s" % (lnk(c,'rev'), lnk(c,'ta')), FMT_X)
section(hr, 21, 'Working capital (reported day ratios)')
def drow_h(ws, r, label, vals, fmt=FMT_DAYS, key=None):
    ah_label(ws, r, label)
    for j, c in enumerate(HCOLS):
        put(ws, r, c, vals[j], font=F_BLACK, fmt=fmt, align=RGT)
    if key: reg(ws, key, r)
drow_h(hr, 22, 'Debtor days', [121,85,52,49,73,76,73])
drow_h(hr, 23, 'Inventory days', [321,333,274,261,253,289,212])
drow_h(hr, 24, 'Payable days', [300,281,291,345,270,259,167])
drow_h(hr, 25, 'Cash conversion cycle (days)', [142,137,35,-35,56,106,119])
put(hr, 26, 1, 'Day ratios as reported by data aggregator from BHEL filings; used to seed forecast working-capital assumptions.', font=F_NOTE)
freeze(hr, 'B4')

# ============================================================================
# 5. OPERATIONAL DRIVERS
# ============================================================================
od = wb.create_sheet('Operational Drivers')
od.sheet_view.showGridLines = False
std_widths(od, 46)
put(od, 1, 1, 'OPERATIONAL DRIVERS & KPIs', font=F_HDRW, fill=FILL_HDR)
for c in range(2, 16): od.cell(row=1, column=c).fill = FILL_HDR
put(od, 2, 1, 'Order data on standalone basis (as disclosed); FY20-FY22 order book are estimates (flagged).', font=F_NOTE)
hyear(od, 4)

def odrow(ws, r, label, vals, fmt=FMT_CR, font=F_BLACK, bold=False, key=None, note=''):
    ah_label(ws, r, label, bold=bold)
    for j, c in enumerate(HCOLS):
        put(ws, r, c, vals[j], font=font, fmt=fmt, align=RGT)
    if key: reg(ws, key, r)
    if note: put(ws, r, 16, note, font=F_NOTE)

section(od, 5, 'Order book dynamics (INR Crore)')
odrow(od, 6, 'Opening order book', [115000,110000,102000,98000,91336,131598,196328], key='ob_open', note='FY20-22 est.')
odrow(od, 7, 'Add: new order inflow', [18000,15000,22000,23548,63000,92535,75000], key='ob_inflow', note='FY20-22,FY24 est.')
odrow(od, 8, 'Less: revenue executed (standalone)', [-21000,-17000,-20800,-22921,-22921,-27355,-33800], key='ob_exec')
odrow(od, 9, 'Closing order book', [110000,102000,98000,91336,131598,196328,240000], key='ob_close', bold=True)
def odf(ws, r, label, fn, fmt, key=None):
    ah_label(ws, r, label)
    for c in HCOLS:
        put(ws, r, c, fn(c), font=F_BLUE, fmt=fmt, align=RGT)
    if key: reg(ws, key, r)
odf(od, 10, 'Book-to-bill (closing OB / revenue, x)', lambda c: "=%s9/-%s8" % (CL(c),CL(c)), FMT_X)
odf(od, 11, 'Execution rate (revenue / opening OB)', lambda c: "=-%s8/%s6" % (CL(c),CL(c)), FMT_PCT)

section(od, 13, 'Segment revenue mix (consolidated, INR Crore)')
odrow(od, 14, 'Power segment', [16800,13500,16500,18100,18600,20937,25407], key='seg_pow', note='FY20-24 est.')
odrow(od, 15, 'Industry segment', [3900,3200,4000,4600,4700,6400,7400], key='seg_ind', note='estimates')
odrow(od, 16, 'Others / exports / renewables', [763,609,711,665,593,1002,975], key='seg_oth')
odf(od, 17, 'Total (check vs revenue)', lambda c: "=SUM(%s14:%s16)" % (CL(c),CL(c)), FMT_CR, key='seg_tot')
odf(od, 18, 'Power as % of revenue', lambda c: "=%s14/%s17" % (CL(c),CL(c)), FMT_PCT)

section(od, 20, 'Other operating KPIs')
odrow(od, 21, 'Approx. workforce (number)', [33500,31500,29800,28500,27000,26000,25500], fmt='#,##0', note='approx.')
odrow(od, 22, 'R&D expenditure', [1180,1090,1050,1100,1150,1300,1500], note='approx.')
odf(od, 23, 'R&D as % of revenue', lambda c: "=%s22/%s" % (CL(c), lnk(c,'rev')), FMT_PCT)
odrow(od, 24, 'Capacity utilisation (manufacturing)', [0.55,0.42,0.50,0.55,0.58,0.65,0.72], fmt=FMT_PCT, note='approx.')
freeze(od, 'B5')
print("Historical Ratios + Operational Drivers built.")


# ============================================================================
#  Generic helpers for forecast schedule sheets
# ============================================================================
ASM = 'Assumptions'
def yhdr(ws, r):
    put(ws, r, 1, 'INR Crore', font=F_ITAL, fill=FILL_SUB)
    for i, c in enumerate(ACOLS):
        put(ws, r, c, YEARS[i], font=F_LBLB, fill=FILL_SUB, align=CEN, border=B_ALL)

def newsheet(title):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    std_widths(ws, 46)
    put(ws, 1, 1, title.upper(), font=F_HDRW, fill=FILL_HDR)
    for c in range(2, 16): ws.cell(row=1, column=c).fill = FILL_HDR
    return ws

def frow(ws, r, label, hist=None, fcst=None, fmt=FMT_CR, bold=False, indent=0,
         fill=None, key=None, histfont=F_GREEN, fcstfont=F_BLUE, seed=None, seedcol=8):
    ah_label(ws, r, label, bold=bold, indent=indent)
    if hist is not None:
        for c in HCOLS:
            v = hist(c)
            if v is None: continue
            put(ws, r, c, v, font=histfont, fmt=fmt, fill=fill, align=RGT)
    if seed is not None:
        put(ws, r, seedcol, seed, font=F_BLACK, fmt=fmt, fill=FILL_IN, align=RGT)
    if fcst is not None:
        for c in FCOLS:
            put(ws, r, c, fcst(c), font=fcstfont, fmt=fmt, fill=fill, align=RGT)
    if key: reg(ws, key, r)

# ============================================================================
# 6. REVENUE BUILD-UP
# ============================================================================
RB = 'Revenue Build-up'
rb = newsheet(RB)
put(rb, 2, 1, 'Order-book-driven revenue: Revenue = Execution rate x Opening order book. Closing OB = Opening + Inflow - Revenue.', font=F_NOTE)
yhdr(rb, 4)
section(rb, 5, 'Order book roll-forward & revenue recognition')
ob_close_od = REG['Operational Drivers']['ob_close']
ob_open_od  = REG['Operational Drivers']['ob_open']
ob_inf_od   = REG['Operational Drivers']['ob_inflow']
frow(rb, 6, 'Opening order book',
     hist=lambda c: "='Operational Drivers'!%s%d" % (CL(c), ob_open_od),
     fcst=lambda c: ("='Operational Drivers'!H%d" % ob_close_od) if c==9 else ("=%s9" % CL(c-1)),
     key='open_ob', bold=True)
frow(rb, 7, 'Add: new order inflow',
     hist=lambda c: "='Operational Drivers'!%s%d" % (CL(c), ob_inf_od),
     fcst=lambda c: "=%s" % XR(ASM,'order_inflow',c), key='inflow')
frow(rb, 8, 'Less: revenue executed (recognised)',
     hist=lambda c: "=-%s" % lnk(c,'rev'),
     fcst=lambda c: "=-%s*%s6" % (XR(ASM,'exec_rate',c), CL(c)), key='exec_neg')
frow(rb, 9, 'Closing order book',
     hist=lambda c: "='Operational Drivers'!%s%d" % (CL(c), ob_close_od),
     fcst=lambda c: "=%s6+%s7+%s8" % (CL(c),CL(c),CL(c)), key='close_ob', bold=True, fill=FILL_TOT)
frow(rb, 11, 'Revenue from operations',
     hist=lambda c: "=%s" % lnk(c,'rev'),
     fcst=lambda c: "=-%s8" % CL(c), key='rev', bold=True, fill=FILL_TOT)
def rb_pct(r, label, fn, fmt, key=None, start=3):
    ah_label(rb, r, label)
    for c in ACOLS:
        if c < start: continue
        put(rb, r, c, fn(c), font=F_BLUE, fmt=fmt, align=RGT)
    if key: reg(rb, key, r)
rb_pct(12, 'Revenue growth %', lambda c: "=%s11/%s11-1" % (CL(c),CL(c-1)), FMT_PCT)
rb_pct(13, 'Execution rate (rev/opening OB) %', lambda c: "=%s11/%s6" % (CL(c),CL(c)), FMT_PCT, start=2)
rb_pct(14, 'Book-to-bill (closing OB/revenue, x)', lambda c: "=%s9/%s11" % (CL(c),CL(c)), FMT_X, start=2)
section(rb, 16, 'Indicative segment split of forecast revenue')
frow(rb, 17, 'Power segment (~75%)', fcst=lambda c: "=%s11*0.75" % CL(c), key='seg_pow')
frow(rb, 18, 'Industry segment (~22%)', fcst=lambda c: "=%s11*0.22" % CL(c))
frow(rb, 19, 'Others/exports/renewables (~3%)', fcst=lambda c: "=%s11*0.03" % CL(c))
freeze(rb, 'B5')

# ============================================================================
# 7. COST FORECAST
# ============================================================================
CF = 'Cost Forecast'
cf = newsheet(CF)
put(cf, 2, 1, 'Costs forecast as % of revenue; raw materials derived as residual so EBITDA = revenue x target margin.', font=F_NOTE)
yhdr(cf, 4)
section(cf, 5, 'Operating cost build-up')
frow(cf, 6, 'Revenue from operations',
     hist=lambda c: "=%s" % lnk(c,'rev'),
     fcst=lambda c: "=%s" % XR(RB,'rev',c), key='rev', bold=True)
frow(cf, 7, 'Raw materials, components & traded goods',
     fcst=lambda c: "=%s6*(1-%s-%s-%s-%s)" % (CL(c), XR(ASM,'ebitda_margin',c),XR(ASM,'emp_pct',c),XR(ASM,'subc_pct',c),XR(ASM,'oopex_pct',c)), key='rm', indent=1)
frow(cf, 8, 'Employee benefits expense',
     fcst=lambda c: "=%s6*%s" % (CL(c), XR(ASM,'emp_pct',c)), key='emp', indent=1)
frow(cf, 9, 'Subcontracting & erection',
     fcst=lambda c: "=%s6*%s" % (CL(c), XR(ASM,'subc_pct',c)), key='subc', indent=1)
frow(cf, 10, 'Other operating expenses',
     fcst=lambda c: "=%s6*%s" % (CL(c), XR(ASM,'oopex_pct',c)), key='oopex', indent=1)
frow(cf, 11, 'Total operating expenditure',
     hist=lambda c: "=%s" % lnk(c,'opex'),
     fcst=lambda c: "=SUM(%s7:%s10)" % (CL(c),CL(c)), key='opex', bold=True, fill=FILL_TOT, histfont=F_GREEN)
frow(cf, 12, 'EBITDA',
     hist=lambda c: "=%s" % lnk(c,'ebitda'),
     fcst=lambda c: "=%s6-%s11" % (CL(c),CL(c)), key='ebitda', bold=True, fill=FILL_TOT)
ah_label(cf, 13, 'EBITDA margin %')
for c in ACOLS:
    put(cf, 13, c, "=%s12/%s6" % (CL(c),CL(c)), font=F_BLUE, fmt=FMT_PCT, align=RGT)
freeze(cf, 'B5')
print("Revenue Build + Cost Forecast built.")


# ============================================================================
# 8. WORKING CAPITAL SCHEDULE
# ============================================================================
WC = 'Working Capital Schedule'
wc = newsheet(WC)
put(wc, 2, 1, 'FY2026 components are flagged analytical estimates that sum to the reported aggregate totals; used to seed the forecast.', font=F_NOTE)
yhdr(wc, 4)
section(wc, 5, 'Operating working capital')
frow(wc, 6, 'Revenue from operations',
     hist=lambda c: "=%s" % lnk(c,'rev'),
     fcst=lambda c: "=%s" % XR(CF,'rev',c), key='rev')
frow(wc, 7, 'Cost of sales (total opex, COGS proxy)',
     hist=lambda c: "=%s" % lnk(c,'opex'),
     fcst=lambda c: "=%s" % XR(CF,'opex',c), key='opex')
frow(wc, 8, 'Trade receivables', seed=6757,
     fcst=lambda c: "=%s/365*%s6" % (XR(ASM,'recv_days',c), CL(c)), key='recv', indent=1)
frow(wc, 9, 'Inventories', seed=18260,
     fcst=lambda c: "=%s/365*%s7" % (XR(ASM,'inv_days',c), CL(c)), key='inv', indent=1)
frow(wc, 10, 'Contract assets', seed=8000,
     fcst=lambda c: "=%s*%s6" % (XR(ASM,'ca_pct',c), CL(c)), key='ca', indent=1)
frow(wc, 11, 'Other current & non-current assets', seed=32373,
     fcst=lambda c: "=%s*(1+%s)" % (("H11" if c==9 else "%s11"%CL(c-1)), XR(ASM,'oca_g',c)), key='oca', indent=1)
ah_label(wc, 12, 'Total operating assets', bold=True)
for c in range(8, 16):
    put(wc, 12, c, "=SUM(%s8:%s11)" % (CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT, fill=FILL_TOT)
reg(wc, 'ta_op', 12)
frow(wc, 13, 'Trade payables', seed=14386,
     fcst=lambda c: "=%s/365*%s7" % (XR(ASM,'pay_days',c), CL(c)), key='pay', indent=1)
frow(wc, 14, 'Contract liabilities / advances', seed=12000,
     fcst=lambda c: "=%s*%s6" % (XR(ASM,'cl_pct',c), CL(c)), key='cl', indent=1)
frow(wc, 15, 'Provisions & other liabilities', seed=15466,
     fcst=lambda c: "=%s*(1+%s)" % (("H15" if c==9 else "%s15"%CL(c-1)), XR(ASM,'ol_g',c)), key='ol', indent=1)
ah_label(wc, 16, 'Total operating liabilities', bold=True)
for c in range(8, 16):
    put(wc, 16, c, "=SUM(%s13:%s15)" % (CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT, fill=FILL_TOT)
reg(wc, 'tl_op', 16)
ah_label(wc, 17, 'Net working capital', bold=True)
for c in range(8, 16):
    put(wc, 17, c, "=%s12-%s16" % (CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT, fill=FILL_TOT, border=B_TOP)
reg(wc, 'nwc', 17)
ah_label(wc, 18, '(Increase) / decrease in NWC')
for c in FCOLS:
    put(wc, 18, c, "=-(%s17-%s17)" % (CL(c),CL(c-1)), font=F_BLUE, fmt=FMT_CR, align=RGT)
reg(wc, 'dnwc', 18)
frow(wc, 19, 'Memo: cash & bank (FY26 estimate, seed)', seed=7000, fcst=None, key='cash_seed')
freeze(wc, 'B5')

# ============================================================================
# 9. FIXED ASSET SCHEDULE  + 10. DEPRECIATION SCHEDULE
# ============================================================================
FA = 'Fixed Asset Schedule'
fa = newsheet(FA)
put(fa, 2, 1, 'Net-block roll-forward. Opening FY2027 = reported FY2026 net PPE. Depreciation = rate x opening net block.', font=F_NOTE)
yhdr(fa, 4)
section(fa, 5, 'Net block roll-forward')
frow(fa, 6, 'Opening net block',
     hist=lambda c: ("=%s" % lnk(c,'ppe')) if c < 8 else None,
     fcst=lambda c: ("=%s" % lnk(8,'ppe')) if c==9 else ("=%s10" % CL(c-1)), key='open_nb', bold=True)
put(fa, 6, 8, "=%s" % lnk(8,'ppe'), font=F_GREEN, fmt=FMT_CR, align=RGT)
frow(fa, 7, 'Add: capital expenditure',
     fcst=lambda c: "=%s" % XR(ASM,'capex',c), key='capex', indent=1)
frow(fa, 8, 'Less: depreciation',
     fcst=lambda c: "=-%s*%s6" % (XR(ASM,'dep_rate',c), CL(c)), key='dep_neg', indent=1)
frow(fa, 9, 'Less: disposals / transfers', fcst=lambda c: 0, key='disp', indent=1)
frow(fa, 10, 'Closing net block',
     hist=lambda c: "=%s" % lnk(c,'ppe'),
     fcst=lambda c: "=%s6+%s7+%s8-%s9" % (CL(c),CL(c),CL(c),CL(c)), key='close_nb', bold=True, fill=FILL_TOT)
freeze(fa, 'B5')

DS = 'Depreciation Schedule'
ds = newsheet(DS)
yhdr(ds, 4)
section(ds, 5, 'Depreciation computation')
frow(ds, 6, 'Opening net block', fcst=lambda c: "=%s" % XR(FA,'open_nb',c), key='open_nb')
frow(ds, 7, 'Depreciation rate (%)', fcst=lambda c: "=%s" % XR(ASM,'dep_rate',c), fmt=FMT_PCT, key='rate')
frow(ds, 8, 'Depreciation expense',
     hist=lambda c: "=%s" % lnk(c,'dep'),
     fcst=lambda c: "=%s6*%s7" % (CL(c),CL(c)), key='dep', bold=True, fill=FILL_TOT)
freeze(ds, 'B5')

# ============================================================================
# 11. DEBT SCHEDULE
# ============================================================================
DBT = 'Debt Schedule'
db = newsheet(DBT)
put(db, 2, 1, 'Interest charged on opening debt to avoid circularity. Opening FY2027 = reported FY2026 borrowings.', font=F_NOTE)
yhdr(db, 4)
section(db, 5, 'Debt roll-forward')
frow(db, 6, 'Opening debt',
     hist=lambda c: ("=%s" % lnk(c,'bor')) if c < 8 else None,
     fcst=lambda c: ("=%s" % lnk(8,'bor')) if c==9 else ("=%s9" % CL(c-1)), key='open_d', bold=True)
put(db, 6, 8, "=%s" % lnk(8,'bor'), font=F_GREEN, fmt=FMT_CR, align=RGT)
frow(db, 7, 'Add: new borrowings', fcst=lambda c: "=%s" % XR(ASM,'new_borrow',c), key='new', indent=1)
frow(db, 8, 'Less: repayments', fcst=lambda c: "=-%s" % XR(ASM,'debt_repay',c), key='repay', indent=1)
frow(db, 9, 'Closing debt',
     hist=lambda c: "=%s" % lnk(c,'bor'),
     fcst=lambda c: "=%s6+%s7+%s8" % (CL(c),CL(c),CL(c)), key='close_d', bold=True, fill=FILL_TOT)
frow(db, 10, 'Average debt', fcst=lambda c: "=(%s6+%s9)/2" % (CL(c),CL(c)), key='avg_d')
frow(db, 11, 'Interest expense (on opening debt)',
     hist=lambda c: "=%s" % lnk(c,'fin'),
     fcst=lambda c: "=%s*%s6" % (XR(ASM,'int_rate',c), CL(c)), key='interest', bold=True)
freeze(db, 'B5')
print("WC + Fixed Asset + Depreciation + Debt schedules built.")


# ============================================================================
# 15. FORECAST FINANCIAL STATEMENTS  (integrated IS + BS, all years)
# ============================================================================
FFS = 'Forecast Financial Statements'
ff = newsheet(FFS)
put(ff, 2, 1, 'Fully integrated three-statement model. FY20-FY26 = reported; FY27-FY33 = forecast (linked to schedules). BS detail shown FY26 onward.', font=F_NOTE)

# ---- INCOME STATEMENT BLOCK ----
section(ff, 3, 'A.  INCOME STATEMENT'); yhdr(ff, 4)
frow(ff, 5, 'Revenue from operations',
     hist=lambda c: "=%s" % lnk(c,'rev'),
     fcst=lambda c: "=%s" % XR(RB,'rev',c), key='rev', bold=True)
frow(ff, 6, 'Other income',
     hist=lambda c: "=%s" % lnk(c,'oi'),
     fcst=lambda c: "=%s*%s5" % (XR(ASM,'oth_inc_pct',c), CL(c)), key='oi')
frow(ff, 7, 'Total operating expenditure',
     hist=lambda c: "=%s" % lnk(c,'opex'),
     fcst=lambda c: "=%s" % XR(CF,'opex',c), key='opex')
ah_label(ff, 8, 'EBITDA', bold=True)
for c in ACOLS: put(ff, 8, c, "=%s5-%s7" % (CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT, fill=FILL_TOT, border=B_TOP)
reg(ff, 'ebitda', 8)
frow(ff, 9, 'Depreciation & amortisation',
     hist=lambda c: "=%s" % lnk(c,'dep'),
     fcst=lambda c: "=%s" % XR(DS,'dep',c), key='dep')
ah_label(ff, 10, 'EBIT', bold=True)
for c in ACOLS: put(ff, 10, c, "=%s8-%s9" % (CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT)
reg(ff, 'ebit', 10)
frow(ff, 11, 'Finance costs',
     hist=lambda c: "=%s" % lnk(c,'fin'),
     fcst=lambda c: "=%s" % XR(DBT,'interest',c), key='fin')
ah_label(ff, 12, 'Profit before tax', bold=True)
for c in ACOLS: put(ff, 12, c, "=%s10+%s6-%s11" % (CL(c),CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT, border=B_TOP)
reg(ff, 'pbt', 12)
frow(ff, 13, 'Tax expense',
     hist=lambda c: "=%s" % lnk(c,'tax'),
     fcst=lambda c: "=%s*%s12" % (XR(ASM,'tax_rate',c), CL(c)), key='tax')
ah_label(ff, 14, 'Profit after tax (PAT)', bold=True)
for c in ACOLS: put(ff, 14, c, "=%s12-%s13" % (CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT, fill=FILL_TOT, border=B_TB)
reg(ff, 'pat', 14)
ah_label(ff, 15, 'EPS (INR)')
for c in ACOLS: put(ff, 15, c, "=%s14/%s" % (CL(c), XR(ASM,'shares',3)), font=F_BLUE, fmt=FMT_PS, align=RGT)
reg(ff, 'eps', 15)
frow(ff, 16, 'Dividend',
     hist=lambda c: "=%s" % lnk(c,'div'),
     fcst=lambda c: "=%s*%s14" % (XR(ASM,'div_payout',c), CL(c)), key='div')

# ============================================================================
# 12. EQUITY SCHEDULE
# ============================================================================
EQ = 'Equity Schedule'
eq = newsheet(EQ)
yhdr(eq, 4)
section(eq, 5, 'Other equity (reserves) roll-forward')
frow(eq, 6, 'Opening other equity',
     fcst=lambda c: ("=%s" % lnk(8,'res')) if c==9 else ("=%s10" % CL(c-1)), key='open_eq', bold=True)
frow(eq, 7, 'Add: profit after tax', fcst=lambda c: "=%s" % XR(FFS,'pat',c), key='pat', indent=1)
frow(eq, 8, 'Less: dividends', fcst=lambda c: "=-%s" % XR(FFS,'div',c), key='div', indent=1)
frow(eq, 9, 'Add: other comprehensive income', fcst=lambda c: 0, key='oci', indent=1)
frow(eq, 10, 'Closing other equity',
     fcst=lambda c: "=SUM(%s6:%s9)" % (CL(c),CL(c)), key='close_eq', bold=True, fill=FILL_TOT)
frow(eq, 12, 'Equity share capital (constant)', fcst=lambda c: 696, key='eqc')
frow(eq, 13, 'Total shareholders equity', fcst=lambda c: "=%s10+%s12" % (CL(c),CL(c)), key='tse', bold=True)
freeze(eq, 'B5')

# ============================================================================
# 13. OTHER ASSETS & LIABILITIES SCHEDULE
# ============================================================================
OAL = 'Other Assets & Liabilities'
oal = newsheet(OAL)
yhdr(oal, 4)
section(oal, 5, 'Other (non-working-capital) assets & liabilities')
frow(oal, 6, 'Capital work-in-progress',
     hist=lambda c: "=%s" % lnk(c,'cwip'),
     fcst=lambda c: "=%s" % XR(ASM,'cwip',c), key='cwip')
frow(oal, 7, 'Investments',
     hist=lambda c: "=%s" % lnk(c,'inv'),
     fcst=lambda c: "=%s" % XR(ASM,'investments',c), key='inv')
frow(oal, 8, 'Contract assets (memo, from WC)', fcst=lambda c: "=%s" % XR(WC,'ca',c), key='ca')
frow(oal, 9, 'Other operating assets (memo, from WC)', fcst=lambda c: "=%s" % XR(WC,'oca',c), key='oca')
frow(oal, 10, 'Contract liabilities (memo, from WC)', fcst=lambda c: "=%s" % XR(WC,'cl',c), key='cl')
frow(oal, 11, 'Provisions & other liabilities (memo, from WC)', fcst=lambda c: "=%s" % XR(WC,'ol',c), key='ol')
freeze(oal, 'B5')
print("Forecast IS + Equity + Other A&L built.")


# ============================================================================
# 14. CASH FLOW STATEMENT  (indirect, fully linked)
# ============================================================================
CFS = 'Cash Flow Statement'
cs = newsheet(CFS)
put(cs, 2, 1, 'Indirect method. FY20-26 = reported summary. FY27-33 built from PAT + D&A - increase in NWC; investing & financing from schedules. No plugs.', font=F_NOTE)
yhdr(cs, 4)
section(cs, 5, 'A.  Operating activities')
frow(cs, 6, 'Profit after tax', fcst=lambda c: "=%s" % XR(FFS,'pat',c), key='pat', indent=1)
frow(cs, 7, 'Add: depreciation & amortisation', fcst=lambda c: "=%s" % XR(DS,'dep',c), indent=1)
frow(cs, 8, '(Increase)/decrease in net working capital', fcst=lambda c: "=%s" % XR(WC,'dnwc',c), indent=1)
frow(cs, 9, 'Cash flow from operating activities',
     hist=lambda c: "=%s" % lnk(c,'cfo'),
     fcst=lambda c: "=SUM(%s6:%s8)" % (CL(c),CL(c)), key='cfo', bold=True, fill=FILL_TOT)
section(cs, 11, 'B.  Investing activities')
frow(cs, 12, 'Capital expenditure', fcst=lambda c: "=-%s" % XR(ASM,'capex',c), indent=1)
frow(cs, 13, 'Change in CWIP',
     fcst=lambda c: "=-(%s-%s)" % (XR(ASM,'cwip',c), (lnk(8,'cwip') if c==9 else XR(ASM,'cwip',c-1))), indent=1)
frow(cs, 14, 'Change in investments',
     fcst=lambda c: "=-(%s-%s)" % (XR(ASM,'investments',c), (lnk(8,'inv') if c==9 else XR(ASM,'investments',c-1))), indent=1)
frow(cs, 15, 'Cash flow from investing activities',
     hist=lambda c: "=%s" % lnk(c,'cfi'),
     fcst=lambda c: "=SUM(%s12:%s14)" % (CL(c),CL(c)), key='cfi', bold=True, fill=FILL_TOT)
section(cs, 17, 'C.  Financing activities')
frow(cs, 18, 'New borrowings', fcst=lambda c: "=%s" % XR(DBT,'new',c), indent=1)
frow(cs, 19, 'Debt repayment', fcst=lambda c: "=%s" % XR(DBT,'repay',c), indent=1)
frow(cs, 20, 'Dividends paid', fcst=lambda c: "=-%s" % XR(FFS,'div',c), indent=1)
frow(cs, 21, 'Cash flow from financing activities',
     hist=lambda c: "=%s" % lnk(c,'cff'),
     fcst=lambda c: "=SUM(%s18:%s20)" % (CL(c),CL(c)), key='cff', bold=True, fill=FILL_TOT)
ah_label(cs, 23, 'Net increase/(decrease) in cash', bold=True)
for c in HCOLS: put(cs, 23, c, "=%s" % lnk(c,'netcf'), font=F_GREEN, fmt=FMT_CR, align=RGT)
for c in FCOLS: put(cs, 23, c, "=%s9+%s15+%s21" % (CL(c),CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT, fill=FILL_TOT)
reg(cs, 'netcf', 23)
frow(cs, 24, 'Opening cash & bank',
     fcst=lambda c: ("='Working Capital Schedule'!H%d" % REG[WC]['cash_seed']) if c==9 else ("=%s25" % CL(c-1)), key='open_cash')
frow(cs, 25, 'Closing cash & bank',
     fcst=lambda c: "=%s24+%s23" % (CL(c),CL(c)), key='close_cash', bold=True, fill=FILL_TOT)
freeze(cs, 'B5')

# ============================================================================
#  FORECAST FINANCIAL STATEMENTS -- BALANCE SHEET BLOCK (FY26 onward)
# ============================================================================
section(ff, 18, 'B.  BALANCE SHEET  (FY2026 reported, FY2027-33 forecast)')
put(ff, 19, 1, 'INR Crore', font=F_ITAL, fill=FILL_SUB)
for i, c in enumerate(ACOLS):
    lab = YEARS[i] if c >= 8 else ''
    put(ff, 19, c, lab, font=F_LBLB, fill=FILL_SUB, align=CEN, border=B_ALL)
put(ff, 19, 1, 'FY20-25: see Historical Financial Statements', font=F_NOTE, fill=FILL_SUB)

def bsrow(r, label, col8, fcst, bold=False, fill=None, indent=0, key=None):
    ah_label(ff, r, label, bold=bold, indent=indent)
    put(ff, r, 8, col8, font=F_GREEN, fmt=FMT_CR, align=RGT, fill=fill)
    for c in FCOLS:
        put(ff, r, c, fcst(c), font=F_GREEN, fmt=FMT_CR, align=RGT, fill=fill)
    if key: reg(ff, key, r)

bsrow(20, 'Net property, plant & equipment', "=%s" % lnk(8,'ppe'), lambda c: "=%s" % XR(FA,'close_nb',c))
bsrow(21, 'Capital work-in-progress', "=%s" % lnk(8,'cwip'), lambda c: "=%s" % XR(ASM,'cwip',c))
bsrow(22, 'Investments', "=%s" % lnk(8,'inv'), lambda c: "=%s" % XR(ASM,'investments',c))
bsrow(23, 'Inventories', "=%s" % XR(WC,'inv',8), lambda c: "=%s" % XR(WC,'inv',c), indent=1)
bsrow(24, 'Trade receivables', "=%s" % XR(WC,'recv',8), lambda c: "=%s" % XR(WC,'recv',c), indent=1)
bsrow(25, 'Contract assets', "=%s" % XR(WC,'ca',8), lambda c: "=%s" % XR(WC,'ca',c), indent=1)
bsrow(26, 'Cash & bank balances', "=%s" % XR(WC,'cash_seed',8), lambda c: "=%s" % XR(CFS,'close_cash',c), indent=1)
bsrow(27, 'Other current & non-current assets', "=%s" % XR(WC,'oca',8), lambda c: "=%s" % XR(WC,'oca',c), indent=1)
ah_label(ff, 28, 'TOTAL ASSETS', bold=True)
for c in range(8, 16): put(ff, 28, c, "=SUM(%s20:%s27)" % (CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT, fill=FILL_TOT, border=B_TB)
reg(ff, 'ta', 28)
bsrow(29, 'Equity share capital', "=%s" % lnk(8,'eqc'), lambda c: "=%s" % XR(EQ,'eqc',c))
bsrow(30, 'Other equity (reserves & surplus)', "=%s" % lnk(8,'res'), lambda c: "=%s" % XR(EQ,'close_eq',c))
bsrow(31, 'Borrowings (incl. leases)', "=%s" % lnk(8,'bor'), lambda c: "=%s" % XR(DBT,'close_d',c))
bsrow(32, 'Trade payables', "=%s" % XR(WC,'pay',8), lambda c: "=%s" % XR(WC,'pay',c), indent=1)
bsrow(33, 'Contract liabilities / advances', "=%s" % XR(WC,'cl',8), lambda c: "=%s" % XR(WC,'cl',c), indent=1)
bsrow(34, 'Provisions & other liabilities', "=%s" % XR(WC,'ol',8), lambda c: "=%s" % XR(WC,'ol',c), indent=1)
ah_label(ff, 35, 'TOTAL EQUITY & LIABILITIES', bold=True)
for c in range(8, 16): put(ff, 35, c, "=SUM(%s29:%s34)" % (CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT, fill=FILL_TOT, border=B_TB)
reg(ff, 'tle', 35)
ah_label(ff, 36, 'Balance check (TA - TE&L)', bold=True)
for c in range(8, 16): put(ff, 36, c, "=%s28-%s35" % (CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT)
reg(ff, 'bschk', 36)
freeze(ff, 'B5')
print("Cash Flow + Forecast BS built.")


# ============================================================================
# 17. DCF VALUATION  (FCFF)
# ============================================================================
DCF = 'DCF Valuation'
dc = newsheet(DCF)
put(dc, 2, 1, 'Free Cash Flow to Firm (FCFF) valuation. WACC computed dynamically on the Assumptions sheet.', font=F_NOTE)
# FCFF build across forecast years
put(dc, 4, 1, 'INR Crore', font=F_ITAL, fill=FILL_SUB)
for i, c in enumerate(FCOLS):
    put(dc, 4, c, FCST[i], font=F_LBLB, fill=FILL_SUB, align=CEN, border=B_ALL)
section(dc, 5, 'A.  Unlevered free cash flow', span=15)
def dcf_row(r, label, fcst, fmt=FMT_CR, bold=False, key=None, fill=None, indent=0, font=F_GREEN):
    ah_label(dc, r, label, bold=bold, indent=indent)
    for c in FCOLS:
        put(dc, r, c, fcst(c), font=font, fmt=fmt, align=RGT, fill=fill)
    if key: reg(dc, key, r)
dcf_row(6, 'EBIT', lambda c: "=%s" % XR(FFS,'ebit',c), bold=True)
dcf_row(7, 'Less: cash taxes on EBIT', lambda c: "=-%s6*%s" % (CL(c), XR(ASM,'tax_rate',c)), font=F_BLUE, indent=1)
dcf_row(8, 'NOPAT', lambda c: "=%s6+%s7" % (CL(c),CL(c)), font=F_BLUE, bold=True)
dcf_row(9, 'Add: depreciation & amortisation', lambda c: "=%s" % XR(DS,'dep',c), indent=1)
dcf_row(10, 'Less: capital expenditure', lambda c: "=-%s" % XR(ASM,'capex',c), indent=1)
dcf_row(11, 'Add/(less): change in working capital', lambda c: "=%s" % XR(WC,'dnwc',c), indent=1)
dcf_row(12, 'Free cash flow to firm (FCFF)', lambda c: "=%s8+%s9+%s10+%s11" % (CL(c),CL(c),CL(c),CL(c)), font=F_BLUE, bold=True, fill=FILL_TOT, key='fcff')
# period & discounting
ah_label(dc, 13, 'Discount period (years)')
for i, c in enumerate(FCOLS): put(dc, 13, c, i+1, font=F_BLACK, fmt='0', align=RGT, fill=FILL_IN)
reg(dc, 'period', 13)
ah_label(dc, 14, 'Discount factor @ WACC')
for c in FCOLS: put(dc, 14, c, "=1/(1+%s)^%s13" % (XR(ASM,'wacc',3), CL(c)), font=F_BLUE, fmt='0.000', align=RGT)
reg(dc, 'df', 14)
ah_label(dc, 15, 'PV of FCFF', bold=True)
for c in FCOLS: put(dc, 15, c, "=%s12*%s14" % (CL(c),CL(c)), font=F_BLUE, fmt=FMT_CR, align=RGT, fill=FILL_TOT)
reg(dc, 'pvfcff', 15)

# Valuation summary (single column C)
section(dc, 17, 'B.  Enterprise & equity value', span=15)
def dsum(r, label, formula, fmt=FMT_CR, bold=False, key=None, fill=None, font=F_BLUE):
    ah_label(dc, r, label, bold=bold)
    put(dc, r, 3, formula, font=font, fmt=fmt, align=RGT, fill=fill, border=B_ALL)
    if key: reg(dc, key, r)
fc = lambda k: "%s%d" % (CL(FCOLS[-1]), REG[DCF][k])  # last forecast col (O) for given row key -> e.g. O12
dsum(18, 'Sum of PV of explicit FCFF (FY27-33)', "=SUM(I15:O15)", key='sumpv', bold=True)
dsum(19, 'WACC', "=%s" % XR(ASM,'wacc',3), fmt=FMT_PCT, font=F_GREEN, key='wacc')
dsum(20, 'Terminal growth rate (g)', "=%s" % XR(ASM,'term_g',3), fmt=FMT_PCT, font=F_GREEN, key='termg')
dsum(21, 'Terminal value (Gordon growth)', "=O12*(1+C20)/(C19-C20)", key='tv')
dsum(22, 'PV of terminal value', "=C21*O14", key='pvtv')
dsum(23, 'Enterprise value (EV)', "=C18+C22", bold=True, fill=FILL_TOT, key='ev')
dsum(24, 'Less: net debt (FY26 debt - cash)', "=-(%s-%s)" % ("'"+FFS+"'!H31", "'"+FFS+"'!H26"), key='netdebt', font=F_GREEN)
dsum(25, 'Equity value', "=C23+C24", bold=True, fill=FILL_TOT, key='eqval')
dsum(26, 'Shares outstanding (crore)', "=%s" % XR(ASM,'shares',3), fmt=FMT_CR1, font=F_GREEN, key='shares')
dsum(27, 'Intrinsic value per share (INR)', "=C25/C26", fmt=FMT_PS, bold=True, fill=FILL_OK, key='ivps')
dsum(28, 'Current market price (INR)', "=%s" % XR(ASM,'price',3), fmt=FMT_PS, font=F_GREEN, key='price')
dsum(29, 'Upside / (downside)', "=C27/C28-1", fmt=FMT_PCT, bold=True, key='upside')
dc.column_dimensions['C'].width = 16
freeze(dc, 'B5')
print("DCF built.")


# ============================================================================
# 18. RELATIVE VALUATION
# ============================================================================
RV = 'Relative Valuation'
rv = newsheet(RV)
put(rv, 2, 1, 'Peer trading multiples are INDICATIVE / illustrative and must be refreshed with live market data before use.', font=F_NOTE)
setw(rv, {'A':30,'B':14,'C':14,'D':14,'E':14,'F':14,'G':16,'H':16})
hdr = ['Peer company','EV/Revenue (x)','EV/EBITDA (x)','P/E (x)','P/B (x)','Div yield %']
section(rv, 4, 'A.  Peer trading multiples (indicative)', span=8)
for j, h in enumerate(hdr):
    put(rv, 5, 1+j, h, font=F_HDRW, fill=FILL_SUB if j else FILL_SUB, align=CEN, border=B_ALL)
    rv.cell(row=5, column=1+j).font = F_LBLB
peers = [
 ('Siemens India',        5.5, 42, 60,  9, 0.4),
 ('ABB India',            7.0, 50, 68, 14, 0.4),
 ('CG Power & Ind.',      6.5, 45, 62, 18, 0.2),
 ('GE Vernova T&D India', 8.0, 48, 65, 20, 0.1),
 ('Thermax',              4.0, 38, 55,  8, 0.3),
 ('Triveni Turbine',      7.5, 40, 52, 16, 0.5),
 ('Larsen & Toubro',      2.0, 22, 32,  5, 0.7),
 ('Kalpataru Projects',   0.9, 12, 22,  3, 0.4),
]
r = 6
for nm, evr, eve, pe, pb, dy in peers:
    put(rv, r, 1, nm, font=F_BLACK)
    for j, v in enumerate([evr, eve, pe, pb, dy]):
        put(rv, r, 2+j, v, font=F_BLACK, fmt=(FMT_X if j < 4 else FMT_PCT.replace('0.0%','0.0"%"')), fill=FILL_IN, align=RGT, border=B_ALL)
    r += 1
med = r
put(rv, med, 1, 'Median', font=F_LBLB, fill=FILL_TOT)
for j in range(5):
    put(rv, med, 2+j, "=MEDIAN(%s6:%s%d)" % (CL(2+j),CL(2+j),r-1), font=F_BLUE, fmt=FMT_X if j<4 else '0.0"%"', align=RGT, fill=FILL_TOT, border=B_ALL)
reg(rv, 'median', med)

section(rv, med+2, 'B.  Implied valuation for BHEL (on FY2027E)', span=8)
b = med+3
def rvline(rr, label, formula, fmt=FMT_CR, bold=False, key=None, fill=None, font=F_BLUE):
    ah_label(rv, rr, label, bold=bold)
    put(rv, rr, 3, formula, font=font, fmt=fmt, align=RGT, fill=fill, border=B_ALL)
    if key: reg(rv, key, rr)
rvline(b,   'BHEL FY2027E revenue', "='%s'!I5" % FFS, font=F_GREEN, key='rev')
rvline(b+1, 'BHEL FY2027E EBITDA', "='%s'!I8" % FFS, font=F_GREEN, key='ebitda')
rvline(b+2, 'BHEL FY2027E EPS (INR)', "='%s'!I15" % FFS, fmt=FMT_PS, font=F_GREEN, key='eps')
rvline(b+3, 'Net debt (FY26)', "='%s'!H31-'%s'!H26" % (FFS,FFS), font=F_GREEN, key='nd')
# implied values
rvline(b+5, 'Implied EV (EV/EBITDA method)', "=C%d*C%d" % (b+1, med), key='ev_e')
rvline(b+6, 'Implied equity value (EV/EBITDA)', "=C%d-C%d" % (b+5, b+3), key='eqv_e')
rvline(b+7, 'Implied price - EV/EBITDA (INR)', "=C%d/%s" % (b+6, XR(ASM,'shares',3)), fmt=FMT_PS, key='px_e', fill=FILL_OK, bold=True)
rvline(b+8, 'Implied price - P/E method (INR)', "=C%d*D%d" % (b+2, med), fmt=FMT_PS, key='px_pe', fill=FILL_OK, bold=True)
rvline(b+9, 'Implied EV (EV/Revenue method)', "=C%d*B%d" % (b, med), key='ev_r')
rvline(b+10,'Implied price - EV/Revenue (INR)', "=(C%d-C%d)/%s" % (b+9, b+3, XR(ASM,'shares',3)), fmt=FMT_PS, key='px_r', fill=FILL_OK, bold=True)
rvline(b+12,'Average implied target price (INR)', "=AVERAGE(C%d,C%d,C%d)" % (b+7,b+8,b+10), fmt=FMT_PS, bold=True, fill=FILL_TOT, key='px_avg')
freeze(rv, 'B5')

# ============================================================================
# 19. SENSITIVITY ANALYSIS  (WACC x terminal growth)
# ============================================================================
SEN = 'Sensitivity Analysis'
se = newsheet(SEN)
put(se, 2, 1, 'Two-way data table: intrinsic value per share (INR) vs WACC and terminal growth. Centre = base case.', font=F_NOTE)
fcff_rng = "'%s'!$I$12:$O$12" % DCF
per_rng  = "'%s'!$I$13:$O$13" % DCF
lastf    = "'%s'!$O$12" % DCF
nd_add   = "'%s'!$C$24" % DCF
shr      = "'%s'!$C$26" % DCF
put(se, 4, 2, 'Intrinsic value/share (INR)', font=F_LBLB)
put(se, 5, 2, 'WACC \\ g', font=F_LBLB, fill=FILL_SUB, align=CEN, border=B_ALL)
gdeltas = [-0.015, -0.0075, 0.0, 0.0075, 0.015]
wdeltas = [0.02, 0.01, 0.0, -0.01, -0.02]
# column headers (g) across cols C..G (3..7)
for j, gd in enumerate(gdeltas):
    put(se, 5, 3+j, "=%s+%s" % (XR(ASM,'term_g',3), gd), font=F_BLACK, fmt=FMT_PCT, fill=FILL_SUB, align=CEN, border=B_ALL)
# row headers (WACC) in col B rows 6..10
for i, wd in enumerate(wdeltas):
    put(se, 6+i, 2, "=%s+%s" % (XR(ASM,'wacc',3), wd), font=F_BLACK, fmt=FMT_PCT, fill=FILL_SUB, align=CEN, border=B_ALL)
for i in range(5):
    for j in range(5):
        w = "$B%d" % (6+i)
        g = "%s$5" % CL(3+j)
        f = ("=(SUMPRODUCT(%s,(1+%s)^(-%s))+(%s*(1+%s)/(%s-%s))*(1+%s)^(-7)+%s)/%s"
             % (fcff_rng, w, per_rng, lastf, g, w, g, w, nd_add, shr))
        fill = FILL_OK if (i == 2 and j == 2) else None
        put(se, 6+i, 3+j, f, font=F_BLUE, fmt=FMT_PS, align=RGT, border=B_ALL, fill=fill)
put(se, 12, 1, 'Other key sensitivities (order inflow, execution rate, EBITDA margin, commodity prices, capex, interest & tax rates, working-capital cycle) are modelled on the Scenario Analysis sheet and via the editable Assumptions sheet.', font=F_ITAL)
freeze(se, 'A1')

# ============================================================================
# 20. SCENARIO ANALYSIS
# ============================================================================
SC = 'Scenario Analysis'
sc = newsheet(SC)
put(sc, 2, 1, 'Bull / Base / Bear scenarios via exit EV/EBITDA on FY2033E, discounted to present. Base links to the live model.', font=F_NOTE)
setw(sc, {'A':36,'B':16,'C':16,'D':16})
heads = ['Driver / output','Bear case','Base case','Bull case']
for j, h in enumerate(heads):
    put(sc, 4, 1+j, h, font=F_LBLB, fill=FILL_SUB, align=CEN, border=B_ALL)
def scrow(rr, label, bear, base, bull, fmt=FMT_CR, bold=False, inputrow=True, key=None):
    ah_label(sc, rr, label, bold=bold)
    vals = [bear, base, bull]
    for j, v in enumerate(vals):
        isf = isinstance(v, str) and v.startswith('=')
        put(sc, rr, 2+j, v, font=(F_GREEN if isf else F_BLACK), fmt=fmt,
            fill=(None if isf else (FILL_IN if inputrow else None)), align=RGT, border=B_ALL)
    if key: reg(sc, key, rr)
section(sc, 3, 'Scenario drivers & valuation', span=4)
# place header again handled above (row4). Inputs from row5:
scrow(5, 'FY2033E revenue (INR Cr)', 58000, "='%s'!O5" % FFS, 92000, key='rev')
scrow(6, 'Terminal EBITDA margin', 0.09, "='%s'!O8/'%s'!O5" % (FFS,FFS), 0.15, fmt=FMT_PCT, key='mgn')
scrow(7, 'Exit EV/EBITDA multiple (x)', 14, 18, 24, fmt=FMT_X, key='mult')
scrow(8, 'WACC', 0.155, "=%s" % XR(ASM,'wacc',3), 0.12, fmt=FMT_PCT, key='wacc')
ah_label(sc, 9, 'FY2033E EBITDA (INR Cr)')
for j in range(3):
    put(sc, 9, 2+j, "=%s5*%s6" % (CL(2+j),CL(2+j)), font=F_BLUE, fmt=FMT_CR, align=RGT, border=B_ALL)
ah_label(sc, 10, 'Terminal EV (INR Cr)')
for j in range(3):
    put(sc, 10, 2+j, "=%s9*%s7" % (CL(2+j),CL(2+j)), font=F_BLUE, fmt=FMT_CR, align=RGT, border=B_ALL)
ah_label(sc, 11, 'PV of EV (discounted 7y)')
for j in range(3):
    put(sc, 11, 2+j, "=%s10/(1+%s8)^7" % (CL(2+j),CL(2+j)), font=F_BLUE, fmt=FMT_CR, align=RGT, border=B_ALL)
ah_label(sc, 12, 'Less: net debt (FY26)')
for j in range(3):
    put(sc, 12, 2+j, "='%s'!H31-'%s'!H26" % (FFS,FFS), font=F_GREEN, fmt=FMT_CR, align=RGT, border=B_ALL)
ah_label(sc, 13, 'Equity value (INR Cr)')
for j in range(3):
    put(sc, 13, 2+j, "=%s11-%s12" % (CL(2+j),CL(2+j)), font=F_BLUE, fmt=FMT_CR, align=RGT, border=B_ALL)
ah_label(sc, 14, 'Target price (INR)', bold=True)
for j in range(3):
    put(sc, 14, 2+j, "=%s13/%s" % (CL(2+j), XR(ASM,'shares',3)), font=F_BLUE, fmt=FMT_PS, align=RGT, border=B_ALL, fill=FILL_OK)
reg(sc, 'tp', 14)
ah_label(sc, 15, 'Scenario probability', bold=True)
probs = [0.25, 0.50, 0.25]
for j, p in enumerate(probs):
    put(sc, 15, 2+j, p, font=F_BLACK, fmt=FMT_PCT, fill=FILL_IN, align=RGT, border=B_ALL)
ah_label(sc, 17, 'Probability-weighted target price (INR)', bold=True)
put(sc, 17, 2, "=SUMPRODUCT(B14:D14,B15:D15)", font=F_BLUE, fmt=FMT_PS, fill=FILL_TOT, align=RGT, border=B_ALL)
reg(sc, 'wtp', 17)
freeze(sc, 'A1')
print("Relative + Sensitivity + Scenario built.")


# ============================================================================
# 16. RATIO ANALYSIS  (full period, FY20-FY33)
# ============================================================================
RA = 'Ratio Analysis'
ra = newsheet(RA)
put(ra, 2, 1, 'FY20-25 sourced from Historical Financial Statements; FY26-33 from the integrated Forecast statements.', font=F_NOTE)
yhdr(ra, 4)
def S(c, key):
    return lnk(c, key, HFS) if c < 8 else "'%s'!%s%d" % (FFS, CL(c), REG[FFS][key])
def EQc(c):
    return "(%s+%s)" % (lnk(c,'eqc',HFS), lnk(c,'res',HFS)) if c < 8 else "('%s'!%s29+'%s'!%s30)" % (FFS,CL(c),FFS,CL(c))
def DBc(c):
    return lnk(c,'bor',HFS) if c < 8 else "'%s'!%s31" % (FFS,CL(c))
def rax(r, label, fn, fmt, start=2, bold=False, fill=None):
    ah_label(ra, r, label, bold=bold)
    for c in ACOLS:
        if c < start: continue
        put(ra, r, c, fn(c), font=F_GREEN, fmt=fmt, align=RGT, fill=fill)

section(ra, 5, 'Growth & margins')
rax(6, 'Revenue growth %', lambda c: "=%s/%s-1" % (S(c,'rev'),S(c-1,'rev')), FMT_PCT, start=3)
rax(7, 'EBITDA margin %', lambda c: "=%s/%s" % (S(c,'ebitda'),S(c,'rev')), FMT_PCT)
rax(8, 'EBIT margin %', lambda c: "=%s/%s" % (S(c,'ebit'),S(c,'rev')), FMT_PCT)
rax(9, 'PAT margin %', lambda c: "=%s/%s" % (S(c,'pat'),S(c,'rev')), FMT_PCT)
section(ra, 11, 'Returns')
rax(12, 'ROE %', lambda c: "=%s/((%s+%s)/2)" % (S(c,'pat'),EQc(c),EQc(c-1)), FMT_PCT, start=3)
rax(13, 'ROCE %', lambda c: "=%s/(%s+%s)" % (S(c,'ebit'),EQc(c),DBc(c)), FMT_PCT)
rax(14, 'ROA %', lambda c: "=%s/%s" % (S(c,'pat'),S(c,'ta')), FMT_PCT)
rax(15, 'Asset turnover (x)', lambda c: "=%s/%s" % (S(c,'rev'),S(c,'ta')), FMT_X)
section(ra, 17, 'Leverage & coverage')
rax(18, 'Debt / equity (x)', lambda c: "=%s/%s" % (DBc(c),EQc(c)), FMT_X)
rax(19, 'Interest coverage (x)', lambda c: "=%s/%s" % (S(c,'ebit'),S(c,'fin')), FMT_X)
rax(20, 'Net debt / EBITDA (x)', lambda c: "=(%s-'%s'!%s26)/%s" % (DBc(c),FFS,CL(c),S(c,'ebitda')), FMT_X, start=8)
section(ra, 22, 'Per share & payout')
rax(23, 'EPS (INR)', lambda c: "=%s" % S(c,'eps'), FMT_PS)
rax(24, 'DPS (INR)', lambda c: "=%s/%s" % (S(c,'div'), XR(ASM,'shares',3)), FMT_PS)
rax(25, 'Dividend payout %', lambda c: "=IF(%s=0,0,%s/%s)" % (S(c,'pat'),S(c,'div'),S(c,'pat')), FMT_PCT)
rax(26, 'Book value per share (INR)', lambda c: "=%s/%s" % (EQc(c), XR(ASM,'shares',3)), FMT_PS)
section(ra, 28, 'Working capital & cash flow (forecast)')
rax(29, 'Receivable days', lambda c: "=%s/'%s'!%s5*365" % ("'"+FFS+"'!"+CL(c)+"24", FFS, CL(c)), FMT_DAYS, start=8)
rax(30, 'Inventory days', lambda c: "='%s'!%s23/'%s'!%s7*365" % (FFS,CL(c),FFS,CL(c)), FMT_DAYS, start=9)
rax(31, 'Payable days', lambda c: "='%s'!%s32/'%s'!%s7*365" % (FFS,CL(c),FFS,CL(c)), FMT_DAYS, start=9)
rax(32, 'FCFF (INR Cr)', lambda c: "='%s'!%s12" % (DCF,CL(c)), FMT_CR, start=9)
rax(33, 'Capex / sales %', lambda c: "=%s/'%s'!%s5" % (XR(ASM,'capex',c),FFS,CL(c)), FMT_PCT, start=9)
rax(34, 'FCFE (INR Cr)', lambda c: "='%s'!%s12-'%s'!%s11*(1-%s)+%s+%s" % (DCF,CL(c),FFS,CL(c),XR(ASM,'tax_rate',c),XR(DBT,'new',c),XR(DBT,'repay',c)), FMT_CR, start=9)
freeze(ra, 'B5')
print("Ratio Analysis built.")


# ============================================================================
# 22. ERROR CHECKS
# ============================================================================
EC = 'Error Checks'
ec = newsheet(EC)
put(ec, 2, 1, 'Automated integrity checks. "OK" = pass (difference < INR 0.5 Cr). Conditional formatting flags any "REVIEW".', font=F_NOTE)
put(ec, 4, 1, 'Check', font=F_LBLB, fill=FILL_SUB, border=B_ALL)
for i, c in enumerate(ACOLS):
    lab = YEARS[i] if c >= 8 else ''
    put(ec, 4, c, lab, font=F_LBLB, fill=FILL_SUB, align=CEN, border=B_ALL)
def echk(r, label, a, b, start=9):
    ah_label(ec, r, label)
    for c in range(start, 16):
        put(ec, r, c, '=IF(ABS(%s-%s)<0.5,"OK","REVIEW")' % (a(c), b(c)), font=F_BLUE, align=CEN, border=B_ALL)
FF = lambda c, row: "'%s'!%s%d" % (FFS, CL(c), row)
echk(5, 'Balance sheet balances (TA = TE&L)', lambda c: FF(c,28), lambda c: FF(c,35), start=8)
echk(6, 'Cash: Balance sheet = Cash flow stmt', lambda c: FF(c,26), lambda c: "'%s'!%s25" % (CFS,CL(c)))
echk(7, 'Retained earnings roll-forward ties', lambda c: FF(c,30), lambda c: "'%s'!%s10" % (EQ,CL(c)))
echk(8, 'Debt roll-forward ties', lambda c: FF(c,31), lambda c: "'%s'!%s9" % (DBT,CL(c)))
echk(9, 'Depreciation schedule ties to P&L', lambda c: FF(c,9), lambda c: "'%s'!%s8" % (DS,CL(c)))
echk(10, 'Revenue ties to Revenue Build-up', lambda c: FF(c,5), lambda c: "'%s'!%s11" % (RB,CL(c)))
echk(11, 'Net fixed assets tie to FA schedule', lambda c: FF(c,20), lambda c: "'%s'!%s10" % (FA,CL(c)))
# single-value checks
ah_label(ec, 13, 'Historical balance check max |diff| (FY20-26)')
put(ec, 13, 3, "=MAX(ABS('%s'!B31),ABS('%s'!C31),ABS('%s'!D31),ABS('%s'!E31),ABS('%s'!F31),ABS('%s'!G31),ABS('%s'!H31))" % tuple([HFS]*7), font=F_BLUE, fmt=FMT_CR1, align=RGT, border=B_ALL)
ah_label(ec, 14, 'Forecast balance check max |diff|')
put(ec, 14, 3, "=MAX(ABS('%s'!I36),ABS('%s'!J36),ABS('%s'!K36),ABS('%s'!L36),ABS('%s'!M36),ABS('%s'!N36),ABS('%s'!O36))" % tuple([FFS]*7), font=F_BLUE, fmt=FMT_CR1, align=RGT, border=B_ALL)
ah_label(ec, 15, 'Circular references')
put(ec, 15, 3, 'None (interest charged on opening debt)', font=F_BLACK)
ah_label(ec, 16, 'Overall model status', bold=True)
put(ec, 16, 3, '=IF(COUNTIF(B5:O11,"REVIEW")=0,"ALL CHECKS PASS","REVIEW REQUIRED")', font=F_BLUE, bold=True, align=CEN, border=B_ALL, fill=FILL_TOT)
from openpyxl.formatting.rule import CellIsRule
grn = PatternFill('solid', fgColor='C6EFCE'); red = PatternFill('solid', fgColor='FFC7CE')
ec.conditional_formatting.add('B5:O11', CellIsRule(operator='equal', formula=['"OK"'], fill=grn))
ec.conditional_formatting.add('B5:O11', CellIsRule(operator='equal', formula=['"REVIEW"'], fill=red))
ec.conditional_formatting.add('C16:C16', CellIsRule(operator='equal', formula=['"ALL CHECKS PASS"'], fill=grn))
freeze(ec, 'B5')

# ============================================================================
# 21. DASHBOARD
# ============================================================================
DB = 'Dashboard'
db = newsheet(DB)
db.sheet_view.showGridLines = False
setw(db, {'A':26,'B':16,'C':4,'D':26,'E':16,'F':4,'G':26,'H':16})
put(db, 1, 1, 'BHEL  |  EXECUTIVE DASHBOARD', font=F_HDRW, fill=FILL_HDR)
for c in range(2, 9): db.cell(row=1, column=c).fill = FILL_HDR
def tile(r, c, label, formula, fmt, font=F_GREEN):
    put(db, r, c, label, font=F_LBLB, fill=FILL_SUB, border=B_ALL)
    put(db, r+1, c, formula, font=font, fmt=fmt, fill=FILL_TOT, border=B_ALL, align=CEN, bold=True)
    db.cell(row=r+1, column=c).font = Font(name='Calibri', size=14, color='1F3864', bold=True)
tile(3, 1, 'DCF value / share (INR)', "='%s'!C27" % DCF, FMT_PS)
tile(3, 2, 'Current price (INR)', "='%s'!C28" % DCF, FMT_PS)
tile(3, 4, 'Upside / (downside)', "='%s'!C29" % DCF, FMT_PCT)
tile(3, 5, 'Prob-weighted TP (INR)', "='%s'!B17" % SC, FMT_PS)
tile(3, 7, 'Avg relative TP (INR)', "='%s'!C%d" % (RV, REG[RV]['px_avg']), FMT_PS)
tile(3, 8, 'WACC', "=%s" % XR(ASM,'wacc',3), FMT_PCT)
tile(6, 1, 'FY27E revenue (INR Cr)', "='%s'!I5" % FFS, FMT_CR)
tile(6, 2, 'FY33E revenue (INR Cr)', "='%s'!O5" % FFS, FMT_CR)
tile(6, 4, 'FY27E EBITDA margin', "='%s'!I8/'%s'!I5" % (FFS,FFS), FMT_PCT)
tile(6, 5, 'FY33E EBITDA margin', "='%s'!O8/'%s'!O5" % (FFS,FFS), FMT_PCT)
tile(6, 7, 'FY26 order book (INR Cr)', "='%s'!H9" % RB, FMT_CR)
tile(6, 8, 'FY33E PAT (INR Cr)', "='%s'!O14" % FFS, FMT_CR)

# DCF bridge mini-table
put(db, 9, 1, 'DCF bridge (INR Cr)', font=F_SECT)
bridge = [('PV of explicit FCFF',"='%s'!C18"%DCF),('PV of terminal value',"='%s'!C22"%DCF),
          ('Enterprise value',"='%s'!C23"%DCF),('Less: net debt',"='%s'!C24"%DCF),('Equity value',"='%s'!C25"%DCF)]
for i,(lab,frm) in enumerate(bridge):
    put(db, 10+i, 1, lab, font=F_LBL, border=B_ALL)
    put(db, 10+i, 2, frm, font=F_GREEN, fmt=FMT_CR, align=RGT, border=B_ALL)

# Charts
c1 = LineChart(); c1.title = "Revenue & EBITDA (INR Cr)"; c1.height = 7.5; c1.width = 15; c1.style = 10
c1.add_data(Reference(ff, min_col=2, max_col=15, min_row=5, max_row=5), from_rows=True, titles_from_data=False)
c1.add_data(Reference(ff, min_col=2, max_col=15, min_row=8, max_row=8), from_rows=True, titles_from_data=False)
c1.set_categories(Reference(ff, min_col=2, max_col=15, min_row=4, max_row=4))
c1.series[0].graphicalProperties.line.width = 28000
db.add_chart(c1, "A17")

c2 = BarChart(); c2.title = "PAT (INR Cr)"; c2.height = 7.5; c2.width = 15; c2.style = 12
c2.add_data(Reference(ff, min_col=2, max_col=15, min_row=14, max_row=14), from_rows=True, titles_from_data=False)
c2.set_categories(Reference(ff, min_col=2, max_col=15, min_row=4, max_row=4))
db.add_chart(c2, "E17")

c3 = LineChart(); c3.title = "Closing order book (INR Cr)"; c3.height = 7.5; c3.width = 15; c3.style = 11
c3.add_data(Reference(rb, min_col=2, max_col=15, min_row=9, max_row=9), from_rows=True, titles_from_data=False)
c3.set_categories(Reference(rb, min_col=2, max_col=15, min_row=4, max_row=4))
db.add_chart(c3, "A33")

c4 = BarChart(); c4.title = "FCFF (INR Cr, FY27-33)"; c4.height = 7.5; c4.width = 15; c4.style = 13
c4.add_data(Reference(dc, min_col=9, max_col=15, min_row=12, max_row=12), from_rows=True, titles_from_data=False)
c4.set_categories(Reference(dc, min_col=9, max_col=15, min_row=4, max_row=4))
db.add_chart(c4, "E33")
print("Dashboard + Error Checks built.")

# ============================================================================
#  FINALISE: tab order, colours, save
# ============================================================================
order = ['Cover Page','Assumptions','Historical Financial Statements','Historical Ratio Analysis',
 'Operational Drivers','Revenue Build-up','Cost Forecast','Working Capital Schedule',
 'Fixed Asset Schedule','Depreciation Schedule','Debt Schedule','Equity Schedule',
 'Other Assets & Liabilities','Cash Flow Statement','Forecast Financial Statements',
 'Ratio Analysis','DCF Valuation','Relative Valuation','Sensitivity Analysis',
 'Scenario Analysis','Dashboard','Error Checks']
wb._sheets.sort(key=lambda s: order.index(s.title))
for t in ['Cover Page','Dashboard']:
    wb[t].sheet_properties.tabColor = NAVY
for t in ['DCF Valuation','Relative Valuation','Scenario Analysis']:
    wb[t].sheet_properties.tabColor = '548235'
wb['Error Checks'].sheet_properties.tabColor = 'C00000'
wb['Assumptions'].sheet_properties.tabColor = 'BF8F00'
wb.active = wb._sheets.index(wb['Dashboard'])

out = '/projects/sandbox/BHEL_Financial_Model.xlsx'
wb.save(out)
print("SAVED:", out)
