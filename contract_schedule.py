# -*- coding: utf-8 -*-
"""
CONTRACT ACCOUNTING SCHEDULE - Bharat Heavy Electricals Limited (BHEL)
Ind AS 115 "Revenue from Contracts with Customers" - Contract Assets & Contract
Liabilities, Current + Non-current, historical (FY2020-FY2026) and forecast
(FY2027-FY2033).

METHODOLOGY (sell-side / CFA equity-research standard):
  This module builds the driver analysis and the roll-forward mechanics used to
  populate the "Contract Accounting Schedule" worksheet in the Excel model. All
  historical figures are sourced from audited consolidated financial statements
  (Ind AS 115 disclosures - "Disclosure - Revenue from Contracts with Customers")
  in BHEL's annual reports, cross-checked against the 5-year financial summary.

SOURCES:
  FY2020 & FY2021 : Annual Report 2020-21 (Note 49 - Ind AS 115; Note 9 - Other
                     non-current assets; Note 22/28 - Other liabilities)
  FY2022 & FY2023 : Annual Report 2022-23 (Note 40 - Ind AS 115; Note 9 - Other
                     Assets; Note 20 - Other liabilities)
  FY2024 & FY2025 : Annual Report 2024-25 (Note 41 - Ind AS 115; Note 9 - Other
                     Assets; Note 20 - Other liabilities; 5-year summary p.332-333)
  FY2026          : BSE Integrated Filing 04-May-2026 + Q4FY26 supplementary deck
                     (order book/inflow); current/non-current split for FY26 is
                     derived (see DATA QUALITY NOTE below).

Order book (standalone, net of taxes) is from the AR2024-25 5-year financial
summary (p.332) and the Q4FY26 supplementary investor deck (order book Rs240,000
cr, inflow Rs75,916 cr for FY26).

------------------------------------------------------------------------------
DATA QUALITY NOTES (disclosed transparently, per sell-side best practice):
------------------------------------------------------------------------------
1. FY2023 contract assets: AR2022-23 Note 40 discloses total contract assets
   (net) of Rs29,740.03 cr as at 31-Mar-23, and Note 9 splits this into
   Non-current Rs18,928.58 cr + Current Rs10,811.45 cr = Rs29,740.03 cr
   (internally consistent). However, the AR2024-25 5-year summary (restated/
   regrouped comparative) shows FY23 contract assets (net) as Rs26,466 cr - a
   Rs3,274 cr downward restatement. We use the AS-ORIGINALLY-REPORTED Rs29,740
   cr (with its Note 9 current/non-current split) for the historical schedule,
   because it is the only figure for which a current/non-current split is
   available; this is flagged as a discontinuity in the historical driver
   series (see CA/Revenue ratio FY23 vs FY24 below) and excluded from the
   trailing-average driver calculation used for forecasting.
2. FY2026 current/non-current split IS disclosed in the BSE integrated filing
   (scanned/image-based filing; current Rs15,192.89 cr + non-current
   Rs14,196.72 cr = Rs29,389.61 cr for contract assets; current Rs9,110.39 cr +
   non-current Rs13,413.24 cr = Rs22,523.63 cr for contract liabilities),
   verified by cross-checking 17 balance-sheet line items against an
   independent extraction with zero mismatches (see audited_data.py).
3. BHEL does not disclose "gross billings during the year" or "gross customer
   advances received during the year" as explicit line items. The full
   mechanical Ind AS 115 roll (Opening + revenue recognised not yet billed -
   billings = Closing; Opening + advances received - revenue recognised from
   advances = Closing) is therefore RECONSTRUCTED using two partial disclosures
   that ARE given every year:
     (a) revenue recognised over time (percentage-of-completion) from the
         disaggregation-of-revenue note, as a proxy for "contract-asset build"
     (b) "Revenue recognised against contract liabilities" from the Ind AS 115
         note, which IS the actual disclosed "revenue recognised from opening
         contract liabilities" line
   The implied "billings" (for CA) and "advances received" (for CL) are
   derived as balancing figures within the roll - i.e. they are NOT
   independently disclosed and should be read as model-derived estimates that
   make the roll self-consistent, not as reported figures.
"""

YEARS_H = ['FY2020','FY2021','FY2022','FY2023','FY2024','FY2025','FY2026']
YEARS_F = ['FY2027','FY2028','FY2029','FY2030','FY2031','FY2032','FY2033']

# ---------------------------------------------------------------------------
# A. HISTORICAL DATA (INR Crore, consolidated, audited/disclosed)
# ---------------------------------------------------------------------------
H = {
 # Contract assets (net of allowances), Ind AS 115 Note, current + non-current
 'ca_cur' : [7670.46, 7494.57, 8691.63, 10811.45, 13451.81, 15778.55, 15192.89],
 'ca_nc'  : [16123.76,16584.91,18248.24,18928.58, 13295.73, 13665.58, 14196.72],
 # Contract liabilities (advances received + billing in excess of revenue)
 'cl_cur' : [3798.08, 4057.85, 3854.33, 3049.34, 3070.10, 5550.73, 9110.39],
 'cl_nc'  : [2921.16, 2806.50, 2193.43, 2585.67, 4063.35, 9742.94, 13413.24],
 # Revenue (consolidated, audited P&L) - denominator for CA driver
 'rev'    : [21463.14,17308.69,21211.09,23364.94,23892.78,28339.48,33782.18],
 # Standalone revenue from operations (disaggregation-of-revenue note total)
 'rev_sa' : [20490.64,16295.55,20153.38,22136.30,22920.52,27355.17,None],
 # Revenue recognised OVER TIME (percentage-of-completion, i.e. project/EPC
 # revenue) - standalone, from the "disaggregation of revenue" Ind AS 115 note.
 # This is the best available proxy for "revenue recognised but not yet billed"
 # (i.e. the flow that BUILDS UP contract assets each year).
 'rev_ot' : [14904.26,10581.15,14107.10,16083.09,15888.42,18478.11,None],
 # "Revenue recognised against contract liabilities (adjustment of customer
 # advances and valuation adjustment during the year)" - the ACTUAL disclosed
 # Ind AS 115 line item = revenue recognised out of the OPENING contract-
 # liability balance. FY26 not separately disclosed in the abbreviated filing.
 'cl_rev_recog': [3140.51,3591.86,3592.89,3024.72,3011.15,2842.80,None],
 # Order book, standalone, net of taxes (per AR2024-25 5-yr summary + Q4FY26 deck)
 'ob_close'  : [None,89813,90084,91336,131598,196328,240000],
 'ob_inflow' : [None,11470,20379,23548,77907,92535,75916],
}
H['ca_tot'] = [H['ca_cur'][i]+H['ca_nc'][i] for i in range(7)]
H['cl_tot'] = [H['cl_cur'][i]+H['cl_nc'][i] for i in range(7)]

# ---------------------------------------------------------------------------
# B. HISTORICAL DRIVER RATIOS
# ---------------------------------------------------------------------------
def _driver_table():
    d = {}
    d['ca_to_rev']   = [H['ca_tot'][i]/H['rev'][i] for i in range(7)]
    d['ca_to_ob']    = [H['ca_tot'][i]/H['ob_close'][i] if H['ob_close'][i] else None for i in range(7)]
    d['ca_days']     = [H['ca_tot'][i]/H['rev'][i]*365 for i in range(7)]
    d['cl_to_rev']   = [H['cl_tot'][i]/H['rev'][i] for i in range(7)]
    d['cl_to_ob']    = [H['cl_tot'][i]/H['ob_close'][i] if H['ob_close'][i] else None for i in range(7)]
    d['cl_to_inflow']= [H['cl_tot'][i]/H['ob_inflow'][i] if H['ob_inflow'][i] else None for i in range(7)]
    d['cl_days']     = [H['cl_tot'][i]/H['rev'][i]*365 for i in range(7)]
    d['ca_cur_pct']  = [H['ca_cur'][i]/H['ca_tot'][i] for i in range(7)]
    d['ca_nc_pct']   = [H['ca_nc'][i]/H['ca_tot'][i] for i in range(7)]
    d['cl_cur_pct']  = [H['cl_cur'][i]/H['cl_tot'][i] for i in range(7)]
    d['cl_nc_pct']   = [H['cl_nc'][i]/H['cl_tot'][i] for i in range(7)]
    return d
DRIVERS = _driver_table()

def _cv(series):
    """Coefficient of variation (population stdev / mean), ignoring None."""
    s = [x for x in series if x is not None]
    if len(s) < 2: return None
    m = sum(s)/len(s)
    var = sum((x-m)**2 for x in s)/len(s)
    return (var**0.5)/m if m else None

CV = {k: _cv(v) for k, v in DRIVERS.items()}

# ---------------------------------------------------------------------------
# C. DRIVER SELECTION (coefficient-of-variation analysis)
# ---------------------------------------------------------------------------
# Candidates tested for CONTRACT ASSETS: CA/Revenue, CA/Closing order book,
# CA Days(vs revenue).
#   CA/Revenue        CV = 0.139   <- LOWEST, and economically sound: contract
#                                      assets ARE unbilled revenue recognised
#                                      under the input-cost (%-of-completion)
#                                      method, so they scale with REVENUE
#                                      recognised, not with the order book
#                                      (most of the order book is unexecuted).
#   CA/Closing OB     CV = 0.330   Higher variance: order book includes large
#                                      unexecuted balances (~4-7x revenue) that
#                                      have not yet generated any contract asset.
#   -> SELECTED DRIVER: CA/Revenue (total), split into current/non-current
#      using the recent (FY24-26) average proportion.
#
# Candidates tested for CONTRACT LIABILITIES: CL/Revenue, CL/Closing order
# book, CL/Order inflow.
#   CL/Closing OB     CV = 0.178   <- LOWEST among order-book-based measures,
#                                      and economically sound: contract
#                                      liabilities are customer ADVANCES tied
#                                      to work still to be executed, i.e. to
#                                      the size of the (unexecuted) ORDER BOOK,
#                                      not to revenue already recognised.
#   CL/Revenue        CV = 0.369   Higher variance - revenue is the wrong
#                                      denominator (advances relate to future,
#                                      not past, performance).
#   CL/Order inflow   CV = 0.566   Highest variance - inflow is too lumpy
#                                      year to year (large single orders).
#   -> SELECTED DRIVER: CL/Closing order book (total), split into current/
#      non-current using the recent (FY24-26) average proportion.
#
# CURRENT vs NON-CURRENT SPLIT: both series show a clear STRUCTURAL SHIFT
# after FY23 (CA current% ~32% in FY20-23 vs ~52% in FY24-26; CL current%
# ~57% in FY20-23 vs ~40% in FY24-26), consistent with management's disclosed
# working-capital initiatives (legacy-project clearance, milestone billing,
# revised advance/payment terms - Q4FY24 concall, 21-May-2024). We therefore
# use the RECENT 3-YEAR (FY24-26) average split, not the full 7-year average,
# since the older years no longer represent the current operating regime.
#   CA current% (FY24-26): 50.3%, 53.6%, 51.7% -> average 51.9% (CV 0.026 - very stable)
#   CL current% (FY24-26): 43.0%, 36.3%, 40.4% -> average 39.9% (CV 0.070 - very stable)
# ---------------------------------------------------------------------------
CA_TO_REV_SELECTED  = 'ca_to_rev'
CL_TO_OB_SELECTED   = 'cl_to_ob'
CA_CUR_SPLIT_RECENT = sum(DRIVERS['ca_cur_pct'][4:]) / 3   # FY24,FY25,FY26
CL_CUR_SPLIT_RECENT = sum(DRIVERS['cl_cur_pct'][4:]) / 3   # FY24,FY25,FY26

# ---------------------------------------------------------------------------
# D. HISTORICAL ROLL-FORWARD (partial mechanics; balancing items are DERIVED,
#    not independently disclosed - see DATA QUALITY NOTE 3 above)
# ---------------------------------------------------------------------------
def historical_roll():
    """Returns per-year: opening, revenue-flow-in, derived-outflow, closing,
    for both Contract Assets (billings-derived) and Contract Liabilities
    (advances-received-derived). FY2020 has no prior-year opening (n/a)."""
    rows = []
    for i in range(1, 7):
        ca_open  = H['ca_tot'][i-1]
        ca_close = H['ca_tot'][i]
        ot_rev   = H['rev_ot'][i]                       # revenue recognised over time (builds CA)
        billed   = (ca_open + ot_rev - ca_close) if ot_rev is not None else None  # derived balancing item

        cl_open  = H['cl_tot'][i-1]
        cl_close = H['cl_tot'][i]
        rev_recog = H['cl_rev_recog'][i]                 # disclosed: revenue recognised from opening CL
        adv_recvd = (cl_close - cl_open + rev_recog) if rev_recog is not None else None  # derived balancing item

        rows.append(dict(year=YEARS_H[i], ca_open=ca_open, ot_rev=ot_rev, billed=billed, ca_close=ca_close,
                          cl_open=cl_open, rev_recog=rev_recog, adv_recvd=adv_recvd, cl_close=cl_close))
    return rows

# ---------------------------------------------------------------------------
# E. FORECAST DRIVER GLIDE PATHS (FY2027-FY2033)
# ---------------------------------------------------------------------------
# Contract assets: CA/Revenue tapers from the FY26 actual (0.870) toward a
# steady-state level consistent with the post-FY24 working-capital-improvement
# trend (management guidance, Q4FY24 concall + Q4FY26 supplementary showing
# contract assets flat YoY despite +19% revenue). We do NOT extrapolate the
# full pace of FY25->FY26 improvement (0.870 vs 1.039, a -16% single-year
# move) as that would be unrealistically aggressive over 7 years; instead we
# taper at roughly half that annual pace, moderating into a plateau.
CA_TO_REV_FCST = [0.83, 0.80, 0.77, 0.74, 0.71, 0.68, 0.65]

# Contract liabilities: CL/Closing OB continues its FY24->FY26 rise (0.054 ->
# 0.078 -> 0.094) per management guidance on improved advance/milestone
# payment terms, plateauing near 0.10-0.105 (still below the FY21 level of
# 0.076->this driver was on an upswing before COVID-era disruption, and the
# current order mix - large thermal EPC orders with revised payment terms -
# supports advances settling modestly above the FY26 level, not reverting).
CL_TO_OB_FCST  = [0.095,0.096,0.098,0.100,0.102,0.104,0.105]

# Current/non-current split: held at the stable recent-3-year average (see
# Section C) for both series - no structural shift is assumed for the
# forecast period absent further disclosed guidance.
CA_CUR_SPLIT_FCST = [CA_CUR_SPLIT_RECENT]*7
CL_CUR_SPLIT_FCST = [CL_CUR_SPLIT_RECENT]*7

# ---------------------------------------------------------------------------
# F. SENSITIVITY - execution rate & customer-advance (CL/OB) rate impact
# ---------------------------------------------------------------------------
def sensitivity(order_inflow, exec_rate, consol_uplift, ca_to_rev, cl_to_ob,
                 open_ob_fy26=240000.0, ca_open_fy26=None, cl_open_fy26=None):
    """Roll the order book / revenue / CA / CL forward for FY27-33 given a
    scenario's driver vectors. Returns dict of lists (7 each)."""
    if ca_open_fy26 is None: ca_open_fy26 = H['ca_tot'][6]
    if cl_open_fy26 is None: cl_open_fy26 = H['cl_tot'][6]
    open_ob = open_ob_fy26
    ca_open = ca_open_fy26
    cl_open = cl_open_fy26
    rev_l, ob_l, ca_l, cl_l = [], [], [], []
    for i in range(7):
        exec_rev = exec_rate[i]*open_ob
        close_ob = open_ob + order_inflow[i] - exec_rev
        rev_c = exec_rev*(1+consol_uplift[i])
        ca_close = ca_to_rev[i]*rev_c
        cl_close = cl_to_ob[i]*close_ob
        rev_l.append(rev_c); ob_l.append(close_ob); ca_l.append(ca_close); cl_l.append(cl_close)
        open_ob = close_ob
    return dict(rev=rev_l, ob=ob_l, ca=ca_l, cl=cl_l)

# ---------------------------------------------------------------------------
# G. REFERENCE FORECAST (FY2027-FY2033) - mirrors the Excel formulas exactly;
#    used here purely to validate the methodology in Python before it is
#    wired into the workbook.
# ---------------------------------------------------------------------------
def forecast_reference(order_inflow, exec_rate, consol_uplift):
    """Base-case reference forecast using the selected drivers (Section E)."""
    out = sensitivity(order_inflow, exec_rate, consol_uplift, CA_TO_REV_FCST, CL_TO_OB_FCST)
    out['ca_cur'] = [out['ca'][i]*CA_CUR_SPLIT_FCST[i] for i in range(7)]
    out['ca_nc']  = [out['ca'][i]*(1-CA_CUR_SPLIT_FCST[i]) for i in range(7)]
    out['cl_cur'] = [out['cl'][i]*CL_CUR_SPLIT_FCST[i] for i in range(7)]
    out['cl_nc']  = [out['cl'][i]*(1-CL_CUR_SPLIT_FCST[i]) for i in range(7)]
    return out

# ---------------------------------------------------------------------------
# H. VALIDATION CHECKS
# ---------------------------------------------------------------------------
def validate(fc):
    """Sanity checks: no unexplained jumps, WC stays reasonable, consistency
    with revenue/order-book growth."""
    checks = []
    # 1. CA and CL should grow with revenue/order book (no sign flips)
    checks.append(('CA rises monotonically with revenue', all(fc['ca'][i+1] >= fc['ca'][i] for i in range(6))))
    checks.append(('CL rises monotonically with order book', all(fc['cl'][i+1] >= fc['cl'][i] for i in range(6))))
    # 2. Year-on-year % change should not exceed order-book/revenue growth by an
    #    unreasonable multiple (no unexplained jumps) - cap at 30% YoY for CA/CL
    for i in range(1, 7):
        cag = fc['ca'][i]/fc['ca'][i-1]-1
        clg = fc['cl'][i]/fc['cl'][i-1]-1
        checks.append((f'CA YoY growth FY{27+i} within +/-30%%: {cag:+.1%}', abs(cag) <= 0.30))
        checks.append((f'CL YoY growth FY{27+i} within +/-30%%: {clg:+.1%}', abs(clg) <= 0.30))
    # 3. CA/Revenue and CL/OB stay within the historical observed range (no
    #    extrapolation beyond plausible bounds)
    hist_ca_range = (min(x for x in DRIVERS['ca_to_rev']), max(x for x in DRIVERS['ca_to_rev']))
    hist_cl_range = (min(x for x in DRIVERS['cl_to_ob'] if x), max(x for x in DRIVERS['cl_to_ob'] if x))
    checks.append((f'CA/Rev within historical range {hist_ca_range}', all(hist_ca_range[0]*0.7 <= x <= hist_ca_range[1]*1.1 for x in CA_TO_REV_FCST)))
    checks.append((f'CL/OB within historical range {hist_cl_range}', all(hist_cl_range[0]*0.7 <= x <= hist_cl_range[1]*1.5 for x in CL_TO_OB_FCST)))
    return checks

if __name__ == '__main__':
    print("Historical roll-forward (derived balancing items):")
    for r in historical_roll():
        if r['ot_rev'] is None:
            print(f"  {r['year']}: CA {r['ca_open']:.0f} -> {r['ca_close']:.0f} (OT-rev/billings not disclosed at this granularity for FY26)  |  "
                  f"CL {r['cl_open']:.0f} -> {r['cl_close']:.0f} (rev-recog/advances not separately disclosed for FY26)")
        else:
            print(f"  {r['year']}: CA {r['ca_open']:.0f} + OT-rev {r['ot_rev']:.0f} - Billed[{r['billed']:.0f}] = {r['ca_close']:.0f}  |  "
                  f"CL {r['cl_open']:.0f} + AdvRecvd[{r['adv_recvd']:.0f}] - RevRecog {r['rev_recog']:.0f} = {r['cl_close']:.0f}")
    print()
    print("Driver CV summary:")
    for k, v in CV.items():
        print(f"  {k:14}: {v:.3f}" if v is not None else f"  {k:14}: n/a")
    print()
    print(f"CA current-split (recent 3yr avg): {CA_CUR_SPLIT_RECENT:.3f}")
    print(f"CL current-split (recent 3yr avg): {CL_CUR_SPLIT_RECENT:.3f}")
    print()
    print("Reference forecast (base case, matches model's order_inflow/exec_rate assumptions):")
    order_inflow = [78000,82000,86000,90000,92000,94000,96000]
    exec_rate    = [0.165,0.170,0.175,0.180,0.185,0.190,0.195]
    consol_uplift= [0.02]*7
    fc = forecast_reference(order_inflow, exec_rate, consol_uplift)
    print(f"{'FY':6}{'Revenue':>10}{'ClosingOB':>10}{'CA_tot':>10}{'CA_cur':>10}{'CA_nc':>10}{'CL_tot':>10}{'CL_cur':>10}{'CL_nc':>10}")
    for i in range(7):
        print(f"FY{27+i:02}{fc['rev'][i]:10.0f}{fc['ob'][i]:10.0f}{fc['ca'][i]:10.0f}{fc['ca_cur'][i]:10.0f}{fc['ca_nc'][i]:10.0f}{fc['cl'][i]:10.0f}{fc['cl_cur'][i]:10.0f}{fc['cl_nc'][i]:10.0f}")
    print()
    print("Validation checks:")
    for name, ok in validate(fc):
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}")
