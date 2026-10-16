"""
Retirement Portfolio Stress Test Model
Builds a formula-driven Excel model (FMVA-style) with openpyxl.

Sheets
  Assumptions - client profile, scenario levers, capital-market assumptions, covariance
  Scenarios   - 3 model portfolios x 3 inflation-protection (TIPS) sleeves: return, risk, outcomes
  Projection  - year-by-year cash-flow projection for the selected scenario
  Dashboard   - key outputs + charts

All assumptions are illustrative for education/portfolio purposes, not forecasts or advice.
Run:  python build_model.py   ->  retirement_stress_test_model.xlsx
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.comments import Comment

OUT = "retirement_stress_test_model.xlsx"

F = "Arial"
BLUE = Font(name=F, color="0000FF")
BLACK = Font(name=F, color="000000")
GREEN = Font(name=F, color="008000")
BOLD = Font(name=F, bold=True)
TITLE = Font(name=F, bold=True, size=14)
HDR = Font(name=F, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F3864")
YELLOW = PatternFill("solid", fgColor="FFFF00")
THIN = Border(bottom=Side(style="thin", color="999999"))

USD = '$#,##0;($#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'

wb = Workbook()

# ---------------------------------------------------------------- Assumptions
a = wb.active
a.title = "Assumptions"
a["A1"] = "Retirement Portfolio Stress Test Model"; a["A1"].font = TITLE
a["A2"] = ("Blue = input you can change · Black = formula · Green = link to another sheet · "
           "Yellow = key lever. Illustrative assumptions only, not investment advice.")
a["A2"].font = Font(name=F, italic=True, size=9)

def header(ws, row, labels, col=1):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=t)
        c.font = HDR; c.fill = HDR_FILL; c.alignment = Alignment(horizontal="center", wrap_text=True)

header(a, 3, ["Client profile", "Value", "Note"])
profile = [
    ("Current age", 35, None, "Example client"),
    ("Retirement age", 65, None, ""),
    ("Plan-to age", 92, None, "Projection sheet rows sized for ages 36-92"),
    ("Current savings ($)", 50000, USD, ""),
    ("Annual contribution, year 1 ($)", 12000, USD, "e.g. 401(k) + IRA"),
    ("Contribution growth (%/yr)", 0.03, PCT, "Raises over time"),
    ("Inflation (%/yr)", 0.025, PCT, "Long-run assumption"),
    ("Retirement spending need (today's $/yr)", 50000, USD, ""),
    ("Social Security estimate (today's $/yr)", 24000, USD, "Assumed to start at retirement"),
    ("Advisory fee (%/yr)", 0.01, PCT, "Typical AUM fee"),
]
for i, (lab, val, fmt, note) in enumerate(profile):
    r = 4 + i
    a.cell(row=r, column=1, value=lab).font = BLACK
    c = a.cell(row=r, column=2, value=val); c.font = BLUE
    if fmt: c.number_format = fmt
    a.cell(row=r, column=3, value=note).font = Font(name=F, size=9, italic=True)
# rows: 4 age,5 ret,6 planto,7 savings,8 contrib,9 g,10 infl,11 spend,12 SS,13 fee

header(a, 15, ["Scenario levers", "Value", "Note"])
a["A16"] = "Model portfolio (1=Conservative, 2=Balanced, 3=Growth)"; a["B16"] = 2
a["A17"] = "Inflation-protection (TIPS) sleeve (%)"; a["B17"] = 0.05; a["B17"].number_format = PCT
a["A18"] = "Selected portfolio name"; a["B18"] = '=INDEX($E$21:$G$21,1,$B$16)'
for ref in ("B16", "B17"):
    a[ref].font = BLUE; a[ref].fill = YELLOW
a["A16"].font = a["A17"].font = a["A18"].font = BLACK
a["C16"] = "Change to see a different portfolio"; a["C17"] = "Carved pro-rata from the core mix"
for ref in ("C16", "C17"):
    a[ref].font = Font(name=F, size=9, italic=True)

header(a, 20, ["Asset class", "Example ETF", "Expected return", "Volatility", "Conservative", "Balanced", "Growth"])
a.cell(row=21, column=1, value="Base weights ->").font = BOLD
for j, n in enumerate(["Conservative", "Balanced", "Growth"]):
    a.cell(row=21, column=5 + j, value=n).font = BOLD
assets = [
    ("US stocks", "VTI", 0.070, 0.16, 0.30, 0.45, 0.60),
    ("International stocks", "VXUS", 0.065, 0.18, 0.10, 0.15, 0.25),
    ("US bonds", "BND", 0.040, 0.06, 0.60, 0.40, 0.15),
    ("Inflation-protected bonds (TIPS)", "SCHP", 0.038, 0.055, 0.00, 0.00, 0.00),
]
for i, row in enumerate(assets):
    r = 22 + i
    for j, v in enumerate(row):
        c = a.cell(row=r, column=1 + j, value=v)
        c.font = BLUE if j >= 2 else BLACK
        if j >= 2: c.number_format = PCT
a["A26"] = "Total"; a["A26"].font = BOLD
for col in "EFG":
    a[f"{col}26"] = f"=SUM({col}22:{col}25)"; a[f"{col}26"].number_format = PCT; a[f"{col}26"].font = BOLD
a["C22"].comment = Comment("Illustrative long-run nominal assumptions for teaching purposes, "
                           "not forecasts. TIPS principal adjusts with inflation (CPI).",
                           "Model")

header(a, 28, ["Correlation", "US stocks", "Intl stocks", "US bonds", "TIPS"])
names = ["US stocks", "Intl stocks", "US bonds", "TIPS"]
corr = [[1.00, 0.80, 0.10, 0.10],
        [0.80, 1.00, 0.10, 0.10],
        [0.10, 0.10, 1.00, 0.70],
        [0.10, 0.10, 0.70, 1.00]]
for i in range(4):
    a.cell(row=29 + i, column=1, value=names[i]).font = BOLD
    for j in range(4):
        c = a.cell(row=29 + i, column=2 + j, value=corr[i][j])
        c.font = BLUE; c.number_format = "0.00"

header(a, 34, ["Covariance", "US stocks", "Intl stocks", "US bonds", "TIPS"])
for i in range(4):
    a.cell(row=35 + i, column=1, value=names[i]).font = BOLD
    for j in range(4):
        rc = a.cell(row=29 + i, column=2 + j).coordinate
        c = a.cell(row=35 + i, column=2 + j, value=f"=$D${22+i}*$D${22+j}*{rc}")
        c.font = BLACK; c.number_format = "0.0000"
a["A39"] = "Covariance = vol(i) x vol(j) x correlation(i,j)"; a["A39"].font = Font(name=F, size=9, italic=True)

a.column_dimensions["A"].width = 48
for col in "BCDEFG":
    a.column_dimensions[col].width = 15
a.column_dimensions["C"].width = 34

# ----------------------------------------------------------------- Scenarios
s = wb.create_sheet("Scenarios")
s["A1"] = "Scenario Comparison: 3 Portfolios x 3 TIPS Sleeves"; s["A1"].font = TITLE
s["A2"] = ("Closed-form outcomes; the Projection sheet shows the selected case year by year. "
           "Bad-luck case = approx. 10th-percentile annualized return: net return - 1.28 x volatility / SQRT(years to retirement).")
s["A2"].font = Font(name=F, italic=True, size=9)
cols = ["Portfolio", "TIPS %", "US stocks", "Intl stocks", "US bonds", "TIPS",
        "Expected return", "Volatility", "Net return (after fee)", "Balance at retirement ($)",
        "Ending balance at plan-to age ($)", "Funded?", "Bad-luck net return (~10th pct)", "Bad-luck ending balance ($)"]
header(s, 4, cols)
# helper inputs
s["Q4"] = "Shared inputs"; s["Q4"].font = BOLD
helpers = [
    ("Years to retirement (n)", "=Assumptions!B5-Assumptions!B4"),
    ("Years in retirement (m)", "=Assumptions!B6-Assumptions!B5"),
    ("Net need, first retirement year ($)", "=(Assumptions!B11-Assumptions!B12)*(1+Assumptions!B10)^R5"),
]
for i, (lab, f) in enumerate(helpers):
    s.cell(row=5 + i, column=17, value=lab).font = BLACK
    c = s.cell(row=5 + i, column=18, value=f); c.font = GREEN
s["R7"].number_format = USD
# O5 n, O6 m, O7 W

digital = [0.0, 0.05, 0.10]
r = 5
for p in range(3):
    pcol = "EFG"[p]
    for d in digital:
        s.cell(row=r, column=1, value=f"=Assumptions!{pcol}21").font = GREEN
        c = s.cell(row=r, column=2, value=d); c.font = BLUE; c.number_format = PCT
        for k in range(3):  # core assets
            c = s.cell(row=r, column=3 + k, value=f"=Assumptions!{pcol}{22+k}*(1-$B{r})")
            c.font = BLACK; c.number_format = PCT
        c = s.cell(row=r, column=6, value=f"=$B{r}"); c.number_format = PCT
        s.cell(row=r, column=7, value=(f"=C{r}*Assumptions!$C$22+D{r}*Assumptions!$C$23"
                                               f"+E{r}*Assumptions!$C$24+F{r}*Assumptions!$C$25"))
        # variance w' * Cov * w, explicit SUMPRODUCT per row of covariance
        terms = "+".join(
            f"{'CDEF'[i]}{r}*SUMPRODUCT(C{r}:F{r},Assumptions!$B${35+i}:$E${35+i})" for i in range(4))
        s.cell(row=r, column=8, value=f"=SQRT({terms})")
        s.cell(row=r, column=9, value=f"=G{r}-Assumptions!$B$13")
        # FV of savings + growing contributions
        s.cell(row=r, column=10, value=(
            f"=Assumptions!$B$7*(1+I{r})^$R$5+IF(ABS(I{r}-Assumptions!$B$9)<0.000001,"
            f"Assumptions!$B$8*$R$5*(1+I{r})^($R$5-1),"
            f"Assumptions!$B$8*((1+I{r})^$R$5-(1+Assumptions!$B$9)^$R$5)/(I{r}-Assumptions!$B$9))"))
        # balance after inflation-growing withdrawals
        s.cell(row=r, column=11, value=(
            f"=J{r}*(1+I{r})^$R$6-IF(ABS(I{r}-Assumptions!$B$10)<0.000001,"
            f"$R$7*$R$6*(1+I{r})^($R$6-1),"
            f"$R$7*((1+I{r})^$R$6-(1+Assumptions!$B$10)^$R$6)/(I{r}-Assumptions!$B$10))"))
        s.cell(row=r, column=12, value=f'=IF(K{r}>=0,"Yes","Shortfall")')
        s.cell(row=r, column=13, value=f"=I{r}-1.28*H{r}/SQRT($R$5)")
        fv = (f"(Assumptions!$B$7*(1+M{r})^$R$5+IF(ABS(M{r}-Assumptions!$B$9)<0.000001,"
              f"Assumptions!$B$8*$R$5*(1+M{r})^($R$5-1),"
              f"Assumptions!$B$8*((1+M{r})^$R$5-(1+Assumptions!$B$9)^$R$5)/(M{r}-Assumptions!$B$9)))")
        s.cell(row=r, column=14, value=(
            f"={fv}*(1+M{r})^$R$6-IF(ABS(M{r}-Assumptions!$B$10)<0.000001,"
            f"$R$7*$R$6*(1+M{r})^($R$6-1),"
            f"$R$7*((1+M{r})^$R$6-(1+Assumptions!$B$10)^$R$6)/(M{r}-Assumptions!$B$10))"))
        s.cell(row=r, column=13).number_format = PCT
        s.cell(row=r, column=14).number_format = USD
        for col in (7, 8, 9):
            s.cell(row=r, column=col).number_format = PCT
        for col in (10, 11):
            s.cell(row=r, column=col).number_format = USD
        for col in range(7, 15):
            if s.cell(row=r, column=col).font != GREEN:
                s.cell(row=r, column=col).font = BLACK
        r += 1
# r now 14
s["A15"] = "Selected scenario row"; s["A15"].font = BOLD
s["B15"] = "=(Assumptions!B16-1)*3+MATCH(Assumptions!B17,B5:B7,0)"
s["C15"] = "TIPS % lever must be 0%, 5% or 10% to match a row here; Projection works for any %."
s["C15"].font = Font(name=F, size=9, italic=True)
s.column_dimensions["A"].width = 16
for col in "BCDEFGHIJKLMN":
    s.column_dimensions[col].width = 14
s.column_dimensions["Q"].width = 34
s.column_dimensions["R"].width = 14

# ---------------------------------------------------------------- Projection
pj = wb.create_sheet("Projection")
pj["A1"] = "Year-by-Year Projection (selected scenario)"; pj["A1"].font = TITLE
pj["A2"] = "Selected portfolio:"; pj["B2"] = "=Assumptions!B18"; pj["B2"].font = GREEN
pj["D2"] = "TIPS %:"; pj["E2"] = "=Assumptions!B17"; pj["E2"].font = GREEN; pj["E2"].number_format = PCT
# selected weights & returns computed here for any TIPS %
pj["G2"] = "Net return:"
pj["H2"] = ("=SUMPRODUCT(INDEX(Assumptions!$E$22:$G$24,0,Assumptions!$B$16),Assumptions!$C$22:$C$24)"
            "*(1-Assumptions!$B$17)+Assumptions!$B$17*Assumptions!$C$25-Assumptions!$B$13")
pj["H2"].number_format = PCT
pj["G3"] = "Net need yr 1 of retirement:"; pj["H3"] = "=Scenarios!R7"; pj["H3"].font = GREEN
pj["H3"].number_format = USD
header(pj, 5, ["Year", "Age", "Phase", "Start balance ($)", "Contribution ($)",
               "Withdrawal ($)", "Investment growth ($)", "End balance ($)"])
N = 57
for t in range(1, N + 1):
    r = 5 + t
    pj.cell(row=r, column=1, value=1 if t == 1 else f"=A{r-1}+1")
    pj.cell(row=r, column=2, value=f"=Assumptions!$B$4+A{r}")
    pj.cell(row=r, column=3, value=f'=IF(B{r}<=Assumptions!$B$5,"Accumulation","Retirement")')
    pj.cell(row=r, column=4, value="=Assumptions!B7" if t == 1 else f"=H{r-1}")
    pj.cell(row=r, column=5, value=f"=IF(B{r}<=Assumptions!$B$5,Assumptions!$B$8*(1+Assumptions!$B$9)^(A{r}-1),0)")
    pj.cell(row=r, column=6, value=f"=IF(B{r}>Assumptions!$B$5,$H$3*(1+Assumptions!$B$10)^(B{r}-Assumptions!$B$5-1),0)")
    pj.cell(row=r, column=7, value=f"=D{r}*$H$2")
    pj.cell(row=r, column=8, value=f"=D{r}+G{r}+E{r}-F{r}")
    for col in range(1, 9):
        c = pj.cell(row=r, column=col); c.font = BLACK
        if col >= 4: c.number_format = USD
    if t == 1:
        pj.cell(row=r, column=4).font = GREEN
last = 5 + N
for col, w in zip("ABCDEFGH", [7, 7, 14, 16, 16, 16, 18, 16]):
    pj.column_dimensions[col].width = w

# ----------------------------------------------------------------- Dashboard
d = wb.create_sheet("Dashboard", 0)
d["A1"] = "Client Plan Dashboard"; d["A1"].font = TITLE
d["A2"] = "Change the yellow levers on the Assumptions sheet; everything updates."
d["A2"].font = Font(name=F, italic=True, size=9)
header(d, 4, ["Key output", "Value"])
kpis = [
    ("Selected portfolio", "=Assumptions!B18", None),
    ("Inflation-protection (TIPS) sleeve", "=Assumptions!B17", PCT),
    ("Net expected return (after fee)", "=Projection!H2", PCT),
    ("Portfolio volatility", "=INDEX(Scenarios!H5:H13,Scenarios!B15)", PCT),
    ("Balance at retirement ($)", f"=INDEX(Projection!H6:H{last},Assumptions!B5-Assumptions!B4)", USD),
    ("Ending balance at plan-to age ($)", f"=Projection!H{last}", USD),
    ("Money lasts to age",
     f'=IF(COUNTIF(Projection!H6:H{last},"<0")=0,"Beyond "&Assumptions!B6,'
     f'Assumptions!B6-COUNTIF(Projection!H6:H{last},"<0")+1)', None),
]
for i, (lab, f, fmt) in enumerate(kpis):
    r = 5 + i
    d.cell(row=r, column=1, value=lab).font = BLACK
    c = d.cell(row=r, column=2, value=f); c.font = GREEN
    if fmt: c.number_format = fmt
    c.alignment = Alignment(horizontal="right")
d.column_dimensions["A"].width = 36; d.column_dimensions["B"].width = 18

lc = LineChart(); lc.title = "Portfolio balance by age (selected scenario)"
lc.y_axis.title = "End balance ($)"; lc.x_axis.title = "Age"; lc.height = 8; lc.width = 18
lc.add_data(Reference(pj, min_col=8, min_row=5, max_row=last), titles_from_data=True)
lc.set_categories(Reference(pj, min_col=2, min_row=6, max_row=last))
lc.y_axis.numFmt = '$#,##0'; lc.legend = None
d.add_chart(lc, "D4")

bc = BarChart(); bc.title = "Ending balance at plan-to age: expected vs bad-luck case"
bc.y_axis.title = "Ending balance ($)"; bc.height = 8; bc.width = 18
bc.add_data(Reference(s, min_col=11, min_row=4, max_row=13), titles_from_data=True)
bc.add_data(Reference(s, min_col=14, min_row=4, max_row=13), titles_from_data=True)
bc.set_categories(Reference(s, min_col=1, min_row=5, max_col=2, max_row=13))
bc.y_axis.numFmt = '$#,##0'; bc.legend.position = 'b'
d.add_chart(bc, "D21")

# ------------------------------------------------------- Account Comparison
ac = wb.create_sheet("Account Comparison")
ac["A1"] = "Account Comparison: Where Should the Next Dollar Go?"; ac["A1"].font = TITLE
ac["A2"] = ("Same pre-tax income goes into each account type each year. Compares after-tax value at retirement. "
            "Texas has no state income tax, so federal rates only.")
ac["A2"].font = Font(name=F, italic=True, size=9)

header(ac, 4, ["Tax & account inputs", "Value", "Note"])
acc_inputs = [
    ("Annual pre-tax income set aside, year 1 ($)", 7500, USD,
     "Equals the 2026 IRA limit ($7,500); 401(k) limit is $24,500 (IRS, Nov 2025)"),
    ("Current marginal tax rate (%)", 0.22, PCT, "Federal bracket while working"),
    ("Retirement tax rate on withdrawals (%)", 0.12, PCT, "Key lever: lower income in retirement"),
    ("Long-term capital gains rate (%)", 0.15, PCT, "Applies to taxable account gains"),
    ("Taxable account tax drag (%/yr)", 0.005, PCT, "Yearly tax on dividends and distributions"),
    ("Salary, year 1 ($)", 75000, USD, "Needed for employer match"),
    ("Employer match rate (%)", 0.50, PCT, "50 cents per dollar..."),
    ("Match cap (% of salary)", 0.06, PCT, "...on the first 6% of salary"),
]
for i, (lab, val, fmt, note) in enumerate(acc_inputs):
    r = 5 + i
    ac.cell(row=r, column=1, value=lab).font = BLACK
    c = ac.cell(row=r, column=2, value=val); c.font = BLUE; c.number_format = fmt
    ac.cell(row=r, column=3, value=note).font = Font(name=F, size=9, italic=True)
ac["B7"].fill = YELLOW
# B5 amount, B6 t_now, B7 t_ret, B8 cg, B9 drag, B10 salary, B11 match rate, B12 cap

header(ac, 14, ["Linked from model", "Value", "Note"])
links = [
    ("Years of saving (n)", "=Assumptions!B5-Assumptions!B4", "0", "Retirement age - current age"),
    ("Net return, selected portfolio (r)", "=Projection!H2", PCT, "From the stress-test model"),
    ("Contribution growth (g)", "=Assumptions!B9", PCT, "Raises; limits also rise with inflation"),
]
for i, (lab, f, fmt, note) in enumerate(links):
    r = 15 + i
    ac.cell(row=r, column=1, value=lab).font = BLACK
    c = ac.cell(row=r, column=2, value=f); c.font = GREEN; c.number_format = fmt
    ac.cell(row=r, column=3, value=note).font = Font(name=F, size=9, italic=True)
# B15 n, B16 r, B17 g

header(ac, 19, ["Growth factors", "Value", "Note"])
ac["A20"] = "Tax-free growth factor (FV of $1/yr growing at g)"
ac["B20"] = "=IF(ABS(B16-B17)<0.000001,B15*(1+B16)^(B15-1),((1+B16)^B15-(1+B17)^B15)/(B16-B17))"
ac["A21"] = "Taxable growth factor (return less tax drag)"
ac["B21"] = ("=IF(ABS((B16-B9)-B17)<0.000001,B15*(1+B16-B9)^(B15-1),"
             "((1+B16-B9)^B15-(1+B17)^B15)/((B16-B9)-B17))")
ac["A22"] = "Cost-basis factor (sum of contributions)"
ac["B22"] = "=IF(B17=0,B15,((1+B17)^B15-1)/B17)"
ac["A23"] = "Employer match, year 1 ($)"
ac["B23"] = "=B11*MIN(B5,B12*B10)"; ac["B23"].number_format = USD
for rr in range(20, 24):
    ac.cell(row=rr, column=1).font = BLACK; ac.cell(row=rr, column=2).font = BLACK
for rr in (20, 21, 22):
    ac.cell(row=rr, column=2).number_format = "0.00"
ac["C20"] = "Growing-annuity formula, end-of-year contributions"
ac["C20"].font = Font(name=F, size=9, italic=True)

header(ac, 25, ["Account", "Tax treatment", "Invested, year 1 ($)", "Balance at retirement ($)",
                "Tax due on withdrawal ($)", "After-tax value ($)", "vs taxable ($)"])
rows = [
    ("Taxable brokerage", "Taxed now; gains taxed yearly and at sale",
     "=B5*(1-B6)", "=C26*B21", "=B8*(D26-C26*B22)"),
    ("Traditional IRA", "Tax break now; withdrawals taxed",
     "=B5", "=C27*B20", "=D27*B7"),
    ("Roth IRA", "Taxed now; withdrawals tax-free",
     "=B5*(1-B6)", "=C28*B20", "=0"),
    ("401(k) with employer match", "Tax break now; match is free money; withdrawals taxed",
     "=B5+B23", "=C29*B20", "=D29*B7"),
]
for i, (name, treat, inv, bal, tax) in enumerate(rows):
    r = 26 + i
    ac.cell(row=r, column=1, value=name).font = BOLD
    ac.cell(row=r, column=2, value=treat).font = Font(name=F, size=9)
    ac.cell(row=r, column=3, value=inv)
    ac.cell(row=r, column=4, value=bal)
    ac.cell(row=r, column=5, value=tax)
    ac.cell(row=r, column=6, value=f"=D{r}-E{r}")
    ac.cell(row=r, column=7, value=f"=F{r}-$F$26")
    for col in range(3, 8):
        c = ac.cell(row=r, column=col); c.font = BLACK; c.number_format = USD
ac["A31"] = "Best choice at these inputs"; ac["A31"].font = BOLD
ac["B31"] = "=INDEX(A26:A29,MATCH(MAX(F26:F29),F26:F29,0))"
ac["A32"] = "Best IRA (Traditional vs Roth)"; ac["A32"].font = BOLD
ac["B32"] = '=IF(F27>F28,"Traditional IRA",IF(F28>F27,"Roth IRA","Tie"))'

header(ac, 34, ["Retirement tax rate", "Traditional IRA after-tax ($)", "Roth IRA after-tax ($)",
                "Better IRA", "Roth advantage ($)"])
for i, rate in enumerate([0.10, 0.12, 0.22, 0.24, 0.32]):
    r = 35 + i
    c = ac.cell(row=r, column=1, value=rate); c.font = BLUE; c.number_format = PCT
    ac.cell(row=r, column=2, value=f"=$D$27*(1-A{r})")
    ac.cell(row=r, column=3, value="=$F$28")
    ac.cell(row=r, column=4, value=f'=IF(B{r}>C{r}+0.5,"Traditional",IF(C{r}>B{r}+0.5,"Roth","Tie"))')
    ac.cell(row=r, column=5, value=f"=C{r}-B{r}")
    for col in (2, 3, 5):
        cc = ac.cell(row=r, column=col); cc.font = BLACK; cc.number_format = USD
    ac.cell(row=r, column=4).font = BLACK
ac["A40"] = ("Rule of thumb the table proves: Traditional wins if your tax rate drops in retirement, "
             "Roth wins if it rises, and they tie when it stays the same.")
ac["A40"].font = Font(name=F, size=9, italic=True)

ac.column_dimensions["A"].width = 44
ac.column_dimensions["B"].width = 26
for col in "CDEFG":
    ac.column_dimensions[col].width = 20

abc = BarChart(); abc.title = "After-tax value at retirement by account"
abc.y_axis.title = "After-tax value ($)"; abc.height = 8; abc.width = 16
abc.add_data(Reference(ac, min_col=6, min_row=25, max_row=29), titles_from_data=True)
abc.set_categories(Reference(ac, min_col=1, min_row=26, max_row=29))
abc.y_axis.numFmt = '$#,##0'; abc.legend = None
ac.add_chart(abc, "I4")

# Dashboard: link the account answer
d["A13"] = "Best place for the next dollar"; d["A13"].font = BLACK
d["B13"] = "='Account Comparison'!B31"; d["B13"].font = GREEN
d["B13"].alignment = Alignment(horizontal="right")

wb.save(OUT)
print("saved", OUT)
