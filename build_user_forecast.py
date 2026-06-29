# -*- coding: utf-8 -*-
"""Fill FY2027-FY2033 forecast into the user's BHEL FM.xlsx (PL & BS),
deriving assumptions from their historical actuals. Adds Assumptions,
Forecast Engine and Discrepancies sheets. Fully formula-driven & balancing."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL

PATH='/projects/sandbox/BHEL/BHEL FM.xlsx'
wb=openpyxl.load_workbook(PATH)

# forecast years in cols J..P (10..16); 2026 seed = col I (9)
FC=list(range(10,17))           # 2027..2033
YR=[2027,2028,2029,2030,2031,2032,2033]
BLUE=Font(color='0000CC'); BLACK=Font(color='000000'); GRN=Font(color='008000')
BOLD=Font(bold=True); HDRW=Font(color='FFFFFF',bold=True)
NAVY=PatternFill('solid',fgColor='1F3864'); GOLD=PatternFill('solid',fgColor='FFF2CC')
SUB=PatternFill('solid',fgColor='D9E1F2'); TOT=PatternFill('solid',fgColor='DDEBF7')
RGT=Alignment(horizontal='right'); CEN=Alignment(horizontal='center')
WRAP=Alignment(wrap_text=True,vertical='top')
thin=Side(style='thin',color='BFBFBF'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
NF='#,##0.0;(#,##0.0)'; PF='0.0%'; DF='0'

# -------------------------------------------------------------------
# 1. ASSUMPTIONS sheet (drivers derived from the company's own history)
# -------------------------------------------------------------------
if 'Assumptions' in wb.sheetnames: del wb['Assumptions']
asm=wb.create_sheet('Assumptions')
asm.sheet_view.showGridLines=False
asm.column_dimensions['A'].width=2; asm.column_dimensions['B'].width=42
for c in FC: asm.column_dimensions[CL(c)].width=10
asm.column_dimensions[CL(17)].width=70
asm.cell(1,2,'FORECAST ASSUMPTIONS  -  derived from BHEL FY2020-FY2026 actuals').font=Font(bold=True,size=13,color='1F3864')
asm.cell(3,2,'Driver').font=BOLD
for j,c in enumerate(FC):
    cell=asm.cell(3,c,YR[j]); cell.font=BOLD; cell.fill=SUB; cell.alignment=CEN; cell.border=BOX
asm.cell(3,17,'Rationale (historical basis | driver | sensitivity)').font=BOLD
A={}
rr=[4]
def arow(key,label,vals,fmt,rat):
    r=rr[0]; asm.cell(r,2,label).font=Font(size=10)
    for j,c in enumerate(FC):
        cell=asm.cell(r,c,vals[j]); cell.font=BLACK; cell.fill=GOLD; cell.number_format=fmt; cell.alignment=RGT; cell.border=BOX
    rc=asm.cell(r,17,rat); rc.font=Font(size=9,italic=True); rc.alignment=WRAP
    asm.row_dimensions[r].height=42
    A[key]=r; rr[0]+=1
arow('growth','Revenue growth %',[0.15,0.14,0.13,0.12,0.11,0.10,0.09],PF,
 'FY24-26 growth 2.3%/18.6%/19.2%. Record order book ~Rs2.4 lakh cr (~7x sales) underpins double-digit growth; tapered as base rises.')
arow('gross','Gross margin %',[0.325,0.330,0.335,0.340,0.340,0.345,0.350],PF,
 'Actual gross margin rose to 32.0% (FY26) from ~27%. Mix shift to higher-value execution & easing input costs; gradual expansion.')
arow('emp','Employee cost % of sales',[0.185,0.180,0.175,0.170,0.165,0.160,0.155],PF,
 'Employee cost ~flat in absolute terms (Rs5.4-6.5k cr) but falls as % as sales scale (FY20 25.3% -> FY26 19.1%); workforce declining via attrition.')
arow('ebitda','EBITDA margin %',[0.080,0.090,0.100,0.110,0.115,0.120,0.125],PF,
 'Actual EBITDA margin FY24 3.0% -> FY26 6.9%. Operating leverage on fixed overheads lifts margin toward low-teens (mgmt guidance).')
arow('oth_inc','Other income % of sales',[0.022]*7,PF,
 'FY26 other income Rs868.7 cr = 2.6% of sales (treasury income on large cash). Held ~2.2% conservatively.')
arow('recv_days','Receivable days',[73,73,73,72,72,72,72],DF,
 'Actual FY26 73 days (FY20 121 -> improving). Large PSU/SEB customers; collections stable.')
arow('inv_days','Inventory days (on COGS)',[210,205,202,200,198,195,193],DF,
 'Actual FY26 ~212 days; long manufacturing/project cycle. Gradual improvement from better planning/execution.')
arow('pay_days','Payable days (on COGS)',[167]*7,DF,
 'Actual FY26 ~167 days supplier credit; held flat.')
arow('oca_pct','Other current assets % of sales',[0.57,0.56,0.55,0.54,0.53,0.52,0.51],PF,
 'FY26 other current assets Rs19,458 cr = 57.6% of sales (unbilled/contract assets, advances). Held slightly declining.')
arow('ocl_pct','Other current liabilities % of sales',[0.40,0.40,0.39,0.39,0.38,0.38,0.37],PF,
 'FY26 other current liabilities Rs13,792 cr = 40.8% of sales (customer advances, provisions). Key WC funding source.')
arow('onca_g','Other non-current assets growth %',[0.03]*7,PF,
 'FY26 Rs20,935 cr (deferred tax assets, long-term/legacy receivables). Structural; grows slowly, not 1:1 with sales.')
arow('oncl_g','Other non-current liabilities growth %',[0.04]*7,PF,
 'FY26 Rs15,214 cr (long-term advances/deferred). Grew strongly historically; moderated to 4%.')
arow('ltprov_g','LT provisions growth %',[0.03]*7,PF,
 'FY26 Rs2,355 cr (warranty & employee-benefit provisions). Grows ~with activity.')
arow('capex','Capital expenditure (Rs cr)',[600,700,750,800,850,900,950],NF,
 'Modernisation + capacity for renewables/defence/electrolysers. Historical capex Rs300-900 cr; internally funded.')
arow('dep_rate','Depreciation % of opening net block',[0.095]*7,PF,
 'Consistent with historical D&A / net block (~9-10%).')
arow('cwip','CWIP closing (Rs cr)',[400]*7,NF,'Steady-state low; capex flows quickly to assets.')
arow('invst','Investments closing (Rs cr)',[310,320,330,340,350,360,370],NF,'JV/associate stakes; modest growth.')
arow('lt_pct','LT borrowings % of total debt',[0.02]*7,PF,
 'Actual FY26 LT borrowings only Rs168 cr (~2%); BHEL debt is almost entirely short-term working-capital lines.')
arow('repay','Debt repayment (Rs cr)',[500]*7,NF,
 'Strong cash & FCF allow gradual reduction of working-capital borrowings.')
arow('int_rate','Interest rate on opening debt %',[0.085]*7,PF,
 'Blended ~ FY26 finance cost / debt; current rate environment.')
arow('tax','Effective tax rate %',[0.25]*7,PF,'New corporate-tax regime (~25.17%); normalised 25%.')
arow('payout','Dividend payout % of PAT',[0.30]*7,PF,'CPSE/DIPAM policy ~30%; historical 30-33%.')
# WACC block
wr=rr[0]+1
asm.cell(wr,2,'WACC & valuation').font=Font(bold=True,color='1F3864'); wr+=1
def single(key,label,val,fmt,rat=''):
    global wr; asm.cell(wr,2,label).font=Font(size=10)
    cell=asm.cell(wr,3,val); cell.font=BLACK; cell.fill=GOLD; cell.number_format=fmt; cell.alignment=RGT; cell.border=BOX
    if rat: c=asm.cell(wr,17,rat); c.font=Font(size=9,italic=True); c.alignment=WRAP
    A[key]=('s',wr); wr+=1
asm.column_dimensions['C'].width=10
single('rf','Risk-free rate (10Y G-Sec)',0.068,PF,'~India 10-year G-Sec yield.')
single('beta','Levered beta',1.25,'0.00','High-beta cyclical capital-goods PSU.')
single('erp','Equity risk premium',0.065,PF,'India ERP ~6-7%.')
single('kd','Pre-tax cost of debt',0.085,PF)
single('wd','Weight of debt',0.10,PF,'Low leverage / net cash; small debt weight.')
single('we','Weight of equity',0.90,PF)
single('tg','Terminal growth',0.045,PF,'Below long-run nominal GDP; conservative perpetuity.')
asm.cell(wr,2,'Cost of equity (Rf+B*ERP)').font=Font(size=10,bold=True)
asm.cell(wr,3,"=C%d+C%d*C%d"%(A['rf'][1],A['beta'][1],A['erp'][1])).number_format=PF; A['ke']=('s',wr); asm.cell(wr,3).font=BLUE; wr+=1
asm.cell(wr,2,'WACC').font=Font(size=10,bold=True)
asm.cell(wr,3,"=C%d*C%d+C%d*C%d*(1-C%d)"%(A['we'][1],A['ke'][1],A['wd'][1],A['kd'][1],A['tax'] if False else A['rf'][1])).number_format=PF
# fix WACC tax ref to tax rate single? use 0.25 constant
asm.cell(wr,3,"=C%d*C%d+C%d*C%d*0.75"%(A['we'][1],A['ke'][1],A['wd'][1],A['kd'][1])).number_format=PF
A['wacc']=('s',wr); asm.cell(wr,3).font=BLUE; wr+=1
single('shares','Shares outstanding (cr)',348.2,NF)
single('price','Current price (Rs)',402.7,'0.00')
print("Assumptions sheet built. rows:",A)
wb.save(PATH)
print("saved stage1")


# -------------------------------------------------------------------
# 2. FORECAST ENGINE sheet (schedules; cols I..P = 2026 seed..2033)
# -------------------------------------------------------------------
if 'Forecast Engine' in wb.sheetnames: del wb['Forecast Engine']
fe=wb.create_sheet('Forecast Engine'); fe.sheet_view.showGridLines=False
fe.column_dimensions['A'].width=2; fe.column_dimensions['B'].width=34
for c in range(9,17): fe.column_dimensions[CL(c)].width=11
fe.cell(1,2,'FORECAST ENGINE  (2026 seed linked to actuals; 2027-2033 driven by Assumptions)').font=Font(bold=True,size=12,color='1F3864')
fe.cell(3,2,'INR Crore').font=Font(italic=True)
for j,c in enumerate(range(9,17)):
    cell=fe.cell(3,c,2026+j); cell.font=BOLD; cell.fill=SUB; cell.alignment=CEN; cell.border=BOX
def ad(key,c): return "'Assumptions'!%s%d"%(CL(c),A[key])
def E(r,c,f):
    cell=fe.cell(r,c,f); cell.font=BLUE; cell.number_format=NF; cell.alignment=RGT
def lbl(r,t,bold=False): fe.cell(r,2,t).font=Font(bold=bold,size=10)
ALL=range(9,17); FCc=range(10,17)
rows={'Revenue':4,'COGS':5,'Gross':6,'Employee':7,'EBITDA':8,'SGA':9,'NBopen':11,'Capex':12,'Dep':13,
 'NBclose':14,'EBIT':15,'Dopen':17,'Repay':18,'Dclose':19,'LTbor':20,'STbor':21,'Int':22,'OthInc':23,
 'PBT':24,'Tax':25,'PAT':26,'Div':27,'ResOpen':29,'ResClose':30,'Recv':32,'Inv':33,'Pay':34,'OCA':35,
 'OCL':36,'ONCA':37,'ONCL':38,'LTprov':39,'CWIP':40,'Invst':41,'NWC':42,'dNWC':43,'CFO':45,'CFI':46,
 'CFF':47,'NetCF':48,'CashOpen':49,'CashClose':50}
for k,r in rows.items(): lbl(r,k,bold=k in('Revenue','Gross','EBITDA','EBIT','PBT','PAT','NetCF','CashClose'))
fe.cell(10,2,'-- Fixed assets --').font=Font(italic=True,size=9)
fe.cell(16,2,'-- Debt / P&L below EBIT --').font=Font(italic=True,size=9)
fe.cell(28,2,'-- Equity --').font=Font(italic=True,size=9)
fe.cell(31,2,'-- Working capital --').font=Font(italic=True,size=9)
fe.cell(44,2,'-- Cash flow --').font=Font(italic=True,size=9)
for c in ALL:
    p=CL(c-1); cc=CL(c)
    # Revenue
    E(rows['Revenue'],c, "='PL'!I5" if c==9 else "=%s4*(1+%s)"%(p,ad('growth',c)))
    E(rows['COGS'],c, "=-'PL'!I6" if c==9 else "=%s4*(1-%s)"%(cc,ad('gross',c)))
    E(rows['Gross'],c, "=%s4-%s5"%(cc,cc))
    E(rows['Employee'],c, "=-'PL'!I9" if c==9 else "=%s4*%s"%(cc,ad('emp',c)))
    E(rows['EBITDA'],c, "='PL'!I12" if c==9 else "=%s4*%s"%(cc,ad('ebitda',c)))
    E(rows['SGA'],c, "=%s6-%s7-%s8"%(cc,cc,cc))
    E(rows['NBopen'],c, "='BS'!I30" if c==9 else "=%s14"%p)
    E(rows['Capex'],c, 0 if c==9 else "=%s"%ad('capex',c))
    E(rows['Dep'],c, "=-'PL'!I13" if c==9 else "=%s11*%s"%(cc,ad('dep_rate',c)))
    E(rows['NBclose'],c, "='BS'!I30" if c==9 else "=%s11+%s12-%s13"%(cc,cc,cc))
    E(rows['EBIT'],c, "=%s8-%s13"%(cc,cc))
    E(rows['Dopen'],c, "='BS'!I14+'BS'!I19" if c==9 else "=%s19"%p)
    E(rows['Repay'],c, 0 if c==9 else "=%s"%ad('repay',c))
    E(rows['Dclose'],c, "=%s17-%s18"%(cc,cc))
    E(rows['LTbor'],c, "='BS'!I14" if c==9 else "=%s19*%s"%(cc,ad('lt_pct',c)))
    E(rows['STbor'],c, "=%s19-%s20"%(cc,cc))
    E(rows['Int'],c, "=-'PL'!I16" if c==9 else "=%s17*%s"%(cc,ad('int_rate',c)))
    E(rows['OthInc'],c, "='PL'!I15" if c==9 else "=%s4*%s"%(cc,ad('oth_inc',c)))
    E(rows['PBT'],c, "=%s15+%s23-%s22"%(cc,cc,cc))
    E(rows['Tax'],c, "=-'PL'!I18" if c==9 else "=%s24*%s"%(cc,ad('tax',c)))
    E(rows['PAT'],c, "=%s24-%s25"%(cc,cc))
    E(rows['Div'],c, 0 if c==9 else "=%s26*%s"%(cc,ad('payout',c)))
    E(rows['ResOpen'],c, "='BS'!I9" if c==9 else "=%s30"%p)
    E(rows['ResClose'],c, "='BS'!I9" if c==9 else "=%s29+%s26-%s27"%(cc,cc,cc))
    E(rows['Recv'],c, "='BS'!I38" if c==9 else "=%s/365*%s4"%(ad('recv_days',c),cc))
    E(rows['Inv'],c, "='BS'!I39" if c==9 else "=%s/365*%s5"%(ad('inv_days',c),cc))
    E(rows['Pay'],c, "='BS'!I20" if c==9 else "=%s/365*%s5"%(ad('pay_days',c),cc))
    E(rows['OCA'],c, "='BS'!I40" if c==9 else "=%s4*%s"%(cc,ad('oca_pct',c)))
    E(rows['OCL'],c, "='BS'!I21" if c==9 else "=%s4*%s"%(cc,ad('ocl_pct',c)))
    E(rows['ONCA'],c, "='BS'!I34" if c==9 else "=%s37*(1+%s)"%(p,ad('onca_g',c)))
    E(rows['ONCL'],c, "='BS'!I16" if c==9 else "=%s38*(1+%s)"%(p,ad('oncl_g',c)))
    E(rows['LTprov'],c, "='BS'!I15" if c==9 else "=%s39*(1+%s)"%(p,ad('ltprov_g',c)))
    E(rows['CWIP'],c, "='BS'!I31" if c==9 else "=%s"%ad('cwip',c))
    E(rows['Invst'],c, "='BS'!I32" if c==9 else "=%s"%ad('invst',c))
    E(rows['NWC'],c, "=(%s32+%s33+%s35+%s37)-(%s34+%s36+%s38+%s39)"%(cc,cc,cc,cc,cc,cc,cc,cc))
    if c>=10:
        E(rows['dNWC'],c, "=-(%s42-%s42)"%(cc,p))
        E(rows['CFO'],c, "=%s26+%s13+%s43"%(cc,cc,cc))
        E(rows['CFI'],c, "=-%s12-(%s40-%s40)-(%s41-%s41)"%(cc,cc,p,cc,p))
        E(rows['CFF'],c, "=-%s18-%s27"%(cc,cc))
        E(rows['NetCF'],c, "=%s45+%s46+%s47"%(cc,cc,cc))
    E(rows['CashOpen'],c, "='BS'!I41" if c==9 else "=%s50"%p)
    E(rows['CashClose'],c, "='BS'!I41" if c==9 else "=%s49+%s48"%(cc,cc))
fe.sheet_view.tabSelected=False

# -------------------------------------------------------------------
# 3. Fill PL & BS forecast columns J..P with formulas -> engine
# -------------------------------------------------------------------
pl=wb['PL']; bs=wb['BS']
def setf(ws,r,c,f,bold=False,fill=None):
    cell=ws.cell(r,c,f); cell.font=Font(color='0000CC',bold=bold); cell.number_format=NF; cell.alignment=RGT
    if fill: cell.fill=fill
for c in FCc:
    cc=CL(c); g="='Forecast Engine'!%s"%cc
    # PL (their sign convention)
    setf(pl,5,c, g+"4", bold=True)
    setf(pl,6,c, "=-'Forecast Engine'!%s5"%cc)
    setf(pl,7,c, "=%s5+%s6"%(cc,cc), bold=True)
    setf(pl,9,c, "=-'Forecast Engine'!%s7"%cc)
    setf(pl,10,c, "=-'Forecast Engine'!%s9"%cc)
    setf(pl,11,c, "=%s9+%s10"%(cc,cc), bold=True)
    setf(pl,12,c, "=%s7+%s11"%(cc,cc), bold=True, fill=TOT)
    setf(pl,13,c, "=-'Forecast Engine'!%s13"%cc)
    setf(pl,14,c, "=%s12+%s13"%(cc,cc), bold=True)
    setf(pl,15,c, g+"23")
    setf(pl,16,c, "=-'Forecast Engine'!%s22"%cc)
    setf(pl,17,c, "=%s14+%s15+%s16"%(cc,cc,cc), bold=True)
    setf(pl,18,c, "=-'Forecast Engine'!%s25"%cc)
    setf(pl,19,c, 0)
    setf(pl,20,c, "=%s17+%s18+%s19"%(cc,cc,cc), bold=True, fill=TOT)
    # BS
    setf(bs,8,c, "=%s"%asg('shares') if False else 696.4)
    setf(bs,9,c, g+"30")
    setf(bs,10,c, "=%s8+%s9"%(cc,cc), bold=True)
    setf(bs,14,c, g+"20")
    setf(bs,15,c, g+"39")
    setf(bs,16,c, g+"38")
    setf(bs,17,c, "=%s14+%s15+%s16"%(cc,cc,cc), bold=True)
    setf(bs,19,c, g+"21")
    setf(bs,20,c, g+"34")
    setf(bs,21,c, g+"36")
    setf(bs,22,c, "=%s19+%s20+%s21"%(cc,cc,cc), bold=True)
    setf(bs,23,c, "=%s17+%s22"%(cc,cc), bold=True)
    setf(bs,25,c, "=%s10+%s23"%(cc,cc), bold=True, fill=TOT)
    setf(bs,30,c, g+"14")
    setf(bs,31,c, g+"40")
    setf(bs,32,c, g+"41")
    setf(bs,33,c, "=%s30+%s31+%s32"%(cc,cc,cc), bold=True)
    setf(bs,34,c, g+"37")
    setf(bs,35,c, "=%s33+%s34"%(cc,cc), bold=True)
    setf(bs,38,c, g+"32")
    setf(bs,39,c, g+"33")
    setf(bs,40,c, g+"35")
    setf(bs,41,c, g+"50")
    setf(bs,42,c, "=%s38+%s39+%s40+%s41"%(cc,cc,cc,cc), bold=True)
    setf(bs,44,c, "=%s35+%s42"%(cc,cc), bold=True, fill=TOT)
    setf(bs,47,c, "=%s44-%s25"%(cc,cc))  # balance check (row 47 'Test')
bs.cell(47,2,'Balance check (TA - TE&L)').font=Font(italic=True,size=9)

# -------------------------------------------------------------------
# 4. DISCREPANCIES sheet
# -------------------------------------------------------------------
if 'Discrepancies' in wb.sheetnames: del wb['Discrepancies']
dq=wb.create_sheet('Discrepancies')
dq.sheet_view.showGridLines=False
dq.column_dimensions['A'].width=2
dq.column_dimensions['B'].width=30
for c in range(3,11): dq.column_dimensions[CL(c)].width=11
dq.column_dimensions[CL(11)].width=60
dq.cell(1,2,'DATA DISCREPANCIES  (BS sheet vs Data Sheet/Screener vs Moneycontrol)').font=Font(bold=True,size=12,color='C00000')
hdr=['Item','FY20','FY21','FY22','FY23','FY24','FY25','FY26','Comment']
for j,h in enumerate(hdr):
    cell=dq.cell(3,2+j,h); cell.font=BOLD; cell.fill=SUB; cell.alignment=CEN; cell.border=BOX
def drow(r,item,vals,comment):
    dq.cell(r,2,item).font=Font(size=10)
    for j,v in enumerate(vals):
        cell=dq.cell(r,3+j,v); cell.number_format=NF; cell.alignment=RGT; cell.border=BOX
    dq.cell(r,11,comment).font=Font(size=9,italic=True); dq.cell(r,11).alignment=WRAP; dq.row_dimensions[r].height=30
drow(4,'BS-sheet total',[59757.6,55251.9,56243.8,56920.0,59005.5,68083.2,76185.6],'Sum of your Moneycontrol BS line items.')
drow(5,'Data Sheet total (Screener)',[60291.0,55959.7,56990.8,57663.8,59784.1,68848.9,76185.6],'Screener aggregated total.')
drow(6,'Difference',[533.4,707.8,747.0,743.8,778.6,765.7,0.0],'FY20-25 gap ~Rs530-780 cr; FY26 reconciles.')
drow(7,'Inventory (BS sheet)',[8908.2,7194.4,6560.2,6755.9,7220.6,9869.5,13334.6],'Moneycontrol inventory.')
drow(8,'Inventory (Screener)',[9450.7,7914.0,7307.3,7499.7,8002.7,10635.3,13334.6],'Screener inventory.')
drow(9,'Inventory difference',[542.5,719.6,747.1,743.8,782.1,765.8,0.0],'ROOT CAUSE: equals the total gap. Screener likely includes stores/spares Moneycontrol nets elsewhere.')
drow(10,'Borrowings (BS LT+ST)',[4947.9,4849.3,4745.0,5385.0,8808.0,8795.0,8187.0],'Moneycontrol.')
drow(11,'Borrowings (Screener)',[5080.0,4950.9,4829.9,5453.5,8856.5,9014.6,8186.9],'Screener; Rs50-220 cr higher FY20-25; FY26 matches.')
dq.cell(13,2,'Other findings').font=Font(bold=True,color='C00000')
notes=[
 "PL 'Selling and admin' is a residual (Gross - Employee - EBITDA): it turns POSITIVE in FY22 (+510.6) and FY23 (+351.1) because it absorbs Screener's volatile 'Other Expenses' (provision write-backs). EBITDA/EBIT/PBT/PAT are unaffected and correct.",
 "FY23 Cash differs by Rs55 cr (BS sheet 6,642.6 vs Data Sheet 6,698.1).",
 "FY26 cash is Rs11,866.6 cr against borrowings of Rs8,187 cr => BHEL is NET CASH ~Rs3,680 cr. Forecast/valuation uses this (an earlier estimate of Rs7,000 cr was too low).",
 "RESOLUTION: forecast is built on your BS-sheet (Moneycontrol) line items, which are internally consistent (Assets = Liabilities each year) and reconcile to Screener in FY26 - the seed year for the forecast.",
]
rn=14
for t in notes:
    c=dq.cell(rn,2,t); c.font=Font(size=9,italic=True); c.alignment=WRAP
    dq.merge_cells(start_row=rn,start_column=2,end_row=rn,end_column=11); dq.row_dimensions[rn].height=42; rn+=1

# order sheets: put Assumptions & Forecast Engine & Discrepancies sensibly
order=['Data Sheet','Sheet1','Assumptions','Forecast Engine','PL','BS','Discrepancies']
wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)
wb.save(PATH)
print("SAVED forecast into BHEL FM.xlsx")
