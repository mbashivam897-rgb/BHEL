"""Audited consolidated balance-sheet data FY2020-FY2026 (INR Cr) from BHEL annual
reports (FY2021, FY2023, FY2025) + FY2026 BSE filing. Order: FY20,FY21,FY22,FY23,FY24,FY25,FY26."""

D = {
 # --- assets ---
 'ppe_grp':   [2824.38,2507.49,2407.12,2485.24,2600.08,2980.91,3094.38],  # PPE+intangibles+inv property
 'cwip':      [306.74,403.21,422.32,344.59,282.32,161.70,399.20],
 'invest':    [162.06,185.34,205.15,235.42,255.67,275.57,302.06],         # equity method + financial
 'trnc':      [4533.50,3179.74,3203.84,3415.54,3224.69,3046.58,2426.93],
 'dta':       [2765.87,3671.24,3530.08,3422.62,4201.26,4067.72,3532.80],
 'note9_nc':  [16361.66,16852.44,18526.54,19300.14,13689.69,14074.98,14716.11],  # other NC assets (incl contract)
 'ofa_nc':    [83.17,97.39,86.73,83.96,206.10,715.95,258.78],             # other financial assets NC
 'invtry':    [8908.23,7194.45,6560.21,6755.90,7220.57,9869.49,13334.58],
 'trcur':     [7108.60,4035.07,3024.75,3128.35,4785.38,5884.35,6796.27],
 'note9_cur': [9783.95,9776.14,10792.53,13050.84,15909.80,18955.39,18923.03],   # other cur assets (incl contract)
 'ofa_cur':   [262.74,232.65,211.56,278.23,239.82,300.76,332.13],
 'curtax':    [229.07,403.60,119.24,226.38,229.07,137.37,202.72],
 'cash_grp':  [6418.59,6701.45,7153.69,6642.58,6157.47,7612.41,11866.62],  # cash equiv + bank balances
 # --- equity ---
 'sharecap':  [696.41]*7,
 'reserve':   [27955.24,25275.59,25810.19,26131.62,23742.24,24025.75,25450.19],  # other equity + NCI
 # --- liabilities ---
 'leaseNC':   [75.37,53.41,35.12,33.75,23.55,162.39,168.18],               # LT borrowings = lease liab NC
 'tpnc':      [1076.23,1881.08,2131.93,2194.03,2292.76,2170.79,1343.44],
 'ofl_nc':    [159.02,216.72,215.10,255.70,407.87,422.79,410.96],
 'provnc':    [4225.16,3925.56,3771.21,4101.02,2489.08,2585.56,2354.62],
 'note20_nc': [2952.65,2831.54,2212.65,2605.81,4102.77,9793.90,13459.63],   # other NC liab (incl contract)
 'stbor_grp': [5004.59,4897.48,4794.81,5419.76,8832.91,8852.21,8018.77],    # ST borrow + lease liab cur
 'tpcur':     [8829.16,6683.51,7749.59,9895.83,8539.38,9540.92,10491.60],
 'ofl_cur':   [1430.62,929.58,1124.09,1276.93,1493.32,1240.52,1313.30],
 'provcur':   [3085.76,3168.52,3066.70,2796.63,2318.27,1815.31,1918.91],
 'note20_cur':[4258.35,4680.80,4635.96,3962.29,4063.36,6776.63,10559.60],   # other cur liab (incl contract)
}
# contract split ratios from FY26 audited (BSE filing notes)
R_CAnc  = 14196.72/14716.11
R_CAcur = 15192.89/18923.03
R_CLnc  = 13413.24/13459.63
R_CLcur = 9110.39/10559.60

def derived(i):
    CAnc = D['note9_nc'][i]*R_CAnc
    CAcur= D['note9_cur'][i]*R_CAcur
    CLnc = D['note20_nc'][i]*R_CLnc
    CLcur= D['note20_cur'][i]*R_CLcur
    othncasset = D['note9_nc'][i]-CAnc + D['ofa_nc'][i]
    othcurasset= D['note9_cur'][i]-CAcur + D['ofa_cur'][i] + D['curtax'][i]
    othncliab  = D['note20_nc'][i]-CLnc + D['ofl_nc'][i]
    othcurliab = D['note20_cur'][i]-CLcur + D['ofl_cur'][i]
    return dict(CAnc=CAnc,CAcur=CAcur,CLnc=CLnc,CLcur=CLcur,
                othncasset=othncasset,othcurasset=othcurasset,othncliab=othncliab,othcurliab=othcurliab)

if __name__=='__main__':
    yrs=['FY20','FY21','FY22','FY23','FY24','FY25','FY26']
    print(f"{'Yr':5}{'TotAssets':>12}{'TotE&L':>12}{'Diff':>8}")
    for i,y in enumerate(yrs):
        d=derived(i)
        ta = (D['ppe_grp'][i]+D['cwip'][i]+D['invest'][i]+D['trnc'][i]+d['CAnc']+D['dta'][i]+d['othncasset']
              + D['invtry'][i]+D['trcur'][i]+d['CAcur']+d['othcurasset']+D['cash_grp'][i])
        tel= (D['sharecap'][i]+D['reserve'][i]
              + D['leaseNC'][i]+D['tpnc'][i]+d['CLnc']+D['provnc'][i]+d['othncliab']
              + D['stbor_grp'][i]+D['tpcur'][i]+d['CLcur']+D['provcur'][i]+d['othcurliab'])
        print(f"{y:5}{ta:12.2f}{tel:12.2f}{ta-tel:8.2f}")
