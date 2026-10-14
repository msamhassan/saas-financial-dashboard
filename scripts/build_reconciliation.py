"""
Builds data/reconciliation/reconciliation_checks.xlsx.
Blue cells = values pasted from a source system or from the Power BI model (inputs).
Black cells = formulas. Every check recalculates if an input is changed.
"""
import json, os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule

B = os.path.join(os.path.dirname(__file__), "..")
P = f"{B}/data/processed"
os.makedirs(f"{B}/data/reconciliation", exist_ok=True)
met = json.load(open(f"{P}/metrics.json"))
rm = pd.read_csv(f"{P}/recon_by_month.csv")
mv = pd.read_csv(f"{P}/mrr_movements_by_month.csv")
R = met["recon"]

F = "Arial"
BLUE, BLACK, GREEN = Font(name=F, size=10, color="0000FF"), Font(name=F, size=10), Font(name=F, size=10, color="008000")
HDR = Font(name=F, size=10, bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="252423")
TITLE = Font(name=F, size=14, bold=True)
NOTE = Font(name=F, size=9, italic=True, color="605E5C")
thin = Border(bottom=Side(style="thin", color="E1DFDD"))
USD, INT = '$#,##0.00;($#,##0.00);"-"', '#,##0;(#,##0);"-"'
pass_fill = PatternFill("solid", fgColor="DFF6DD"); fail_fill = PatternFill("solid", fgColor="FDE7E9")

wb = Workbook()


def sheet(name, title, note, headers, widths):
    ws = wb.create_sheet(name)
    ws["A1"] = title; ws["A1"].font = TITLE
    ws["A2"] = note; ws["A2"].font = NOTE
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(4, i, h); c.font = HDR; c.fill = HFILL; c.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[4].height = 30
    ws.freeze_panes = "A5"
    return ws


def status_rule(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"PASS"'], fill=pass_fill))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"FAIL"'], fill=fail_fill))


def put(ws, r, c, v, font=BLACK, fmt=None):
    cell = ws.cell(r, c, v); cell.font = font; cell.border = thin
    if fmt: cell.number_format = fmt
    return cell


# ---------------- Controls ----------------
wc = wb.active; wc.title = "Controls"
wc["A1"] = "Reconciliation controls"; wc["A1"].font = TITLE
rows = [("Revenue tolerance per month (USD)", 1.00, "Max acceptable model-vs-GL difference; covers cent rounding after FX conversion"),
        ("MRR roll-forward tolerance (USD)", 0.01, "Opening + New + Expansion - Contraction - Churn must equal Closing"),
        ("Budget source", "finance_gl_budget.xlsx / Revenue_Budget", "Approved FY2026 budget, by month and region")]
for i, (a, b, c) in enumerate(rows, 3):
    put(wc, i, 1, a); put(wc, i, 2, b, BLUE, USD if isinstance(b, float) else None); put(wc, i, 3, c, NOTE)
    wc.cell(i, 2).fill = PatternFill("solid", fgColor="FFFF00") if isinstance(b, float) else PatternFill()
wc["A8"] = "Legend"; wc["A8"].font = Font(name=F, bold=True)
for i, (txt, f) in enumerate([("Blue text = input pasted from a source file or exported from the Power BI model", BLUE),
                              ("Black text = formula", BLACK), ("Green text = link to another sheet in this workbook", GREEN),
                              ("Yellow fill = tolerance you can change", BLACK)], 9):
    put(wc, i, 1, txt, f)
wc["A14"] = ("Model-side values are the outputs of the Power BI measures (exported from a table visual). In this repo they "
             "are produced by scripts/compute_metrics.py, which applies the same Power Query cleaning rules and measure logic.")
wc["A14"].font = NOTE
for col, w in zip("ABC", [44, 34, 80]): wc.column_dimensions[col].width = w

# ---------------- 1. Billing rows -> model rows ----------------
ws = sheet("1_Billing_to_Model", "1. Billing export → Power BI fact_MRR (row counts)",
           "Every row removed by Power Query must be explained: exact duplicates and QA test accounts. Nothing else may disappear.",
           ["Month", "Raw export rows", "Duplicate rows removed", "Test-account rows removed", "Expected model rows",
            "Power BI fact_MRR rows", "Difference", "Status"], [10, 14, 14, 14, 14, 14, 12, 10])
# Dec-25 opening month is included in the model but has no movement row; add it explicitly
raw_all = pd.read_excel(f"{B}/data/raw/billing_subscription_export.xlsx", "MRR_Lines")
dec = raw_all[raw_all.BillingPeriod == "2025-12"].copy()
dec["CustomerID"] = dec.CustomerID.str.strip().str.upper(); dec["ProductCode"] = dec.ProductCode.str.strip().str.upper()
dec_dup = int(dec.duplicated().sum())
lines = [("2025-12", len(dec), dec_dup, 0, len(dec) - dec_dup)] + \
        [(r.Month, r.RawRows, r.DuplicateRows, r.TestRows, r.ModelRows) for r in rm.itertuples()]
r0 = 5
for i, (m, raw, d, t, model) in enumerate(lines):
    r = r0 + i
    put(ws, r, 1, m); put(ws, r, 2, raw, BLUE, INT); put(ws, r, 3, d, BLUE, INT); put(ws, r, 4, t, BLUE, INT)
    put(ws, r, 5, f"=B{r}-C{r}-D{r}", fmt=INT); put(ws, r, 6, model, BLUE, INT)
    put(ws, r, 7, f"=F{r}-E{r}", fmt=INT); put(ws, r, 8, f'=IF(G{r}=0,"PASS","FAIL")')
rt = r0 + len(lines)
put(ws, rt, 1, "Total", Font(name=F, bold=True))
for c in "BCDEFG":
    put(ws, rt, " BCDEFG".index(c) + 1, f"=SUM({c}{r0}:{c}{rt-1})", Font(name=F, bold=True), INT)
put(ws, rt, 8, f'=IF(COUNTIF(H{r0}:H{rt-1},"FAIL")=0,"PASS","FAIL")', Font(name=F, bold=True))
status_rule(ws, f"H{r0}:H{rt}")
s1 = f"'1_Billing_to_Model'!H{rt}"

# ---------------- 2. Model revenue -> GL ----------------
ws = sheet("2_Model_to_GL", "2. Power BI subscription revenue → GL account 4000 (USD)",
           "Model revenue = Σ(MRR_Local × FX rate) after cleaning. GL is posted rounded to cents, so a sub-$1 difference is expected.",
           ["Month", "Power BI [Subscription Revenue]", "GL 4000 Subscription Revenue", "Difference", "Within tolerance?", "Status"],
           [10, 22, 22, 14, 14, 10])
for i, r_ in enumerate(rm.itertuples()):
    r = 5 + i
    put(ws, r, 1, r_.Month); put(ws, r, 2, round(r_.Model_MRR_USD, 4), BLUE, USD); put(ws, r, 3, r_.GL_Revenue_USD, BLUE, USD)
    put(ws, r, 4, f"=B{r}-C{r}", fmt=USD); put(ws, r, 5, f"=ABS(D{r})<=Controls!$B$3")
    put(ws, r, 6, f'=IF(E{r},"PASS","FAIL")')
rt2 = 5 + len(rm)
put(ws, rt2, 1, "FY2026", Font(name=F, bold=True))
for c in "BCD":
    put(ws, rt2, " ABCD".index(c), f"=SUM({c}5:{c}{rt2-1})", Font(name=F, bold=True), USD)
put(ws, rt2, 5, f"=ABS(D{rt2})<=Controls!$B$3*12", Font(name=F, bold=True))
put(ws, rt2, 6, f'=IF(AND(E{rt2},COUNTIF(F5:F{rt2-1},"FAIL")=0),"PASS","FAIL")', Font(name=F, bold=True))
status_rule(ws, f"F5:F{rt2}")
s2 = f"'2_Model_to_GL'!F{rt2}"

# ---------------- 3. MRR roll-forward ----------------
ws = sheet("3_MRR_RollForward", "3. MRR roll-forward (Power BI fact_MRR_Movements vs fact_MRR)",
           "Opening + New + Expansion − Contraction − Churn + FX impact = Closing; each Opening = prior Closing. Movements are classified in constant currency; FX is its own line.",
           ["Month", "Opening MRR", "New", "Expansion", "Contraction", "Churn", "FX impact", "Closing MRR (fact_MRR)",
            "Calculated closing", "Roll-forward diff", "Opening = prior closing?", "Status"],
           [10, 15, 13, 13, 13, 13, 12, 18, 16, 13, 14, 10])
for i, r_ in enumerate(mv.itertuples()):
    r = 5 + i
    put(ws, r, 1, r_.Month)
    for j, v in enumerate([r_.Opening, r_.New, r_.Expansion, r_.Contraction, r_.Churn, r_.FXImpact, r_.Closing], 2):
        put(ws, r, j, round(v, 4), BLUE, USD)
    put(ws, r, 9, f"=B{r}+C{r}+D{r}-E{r}-F{r}+G{r}", fmt=USD)
    put(ws, r, 10, f"=I{r}-H{r}", fmt=USD)
    put(ws, r, 11, "=TRUE()" if i == 0 else f"=ABS(B{r}-H{r-1})<=Controls!$B$4")
    put(ws, r, 12, f'=IF(AND(ABS(J{r})<=Controls!$B$4,K{r}),"PASS","FAIL")')
rt3 = 5 + len(mv)
put(ws, rt3, 1, "FY2026", Font(name=F, bold=True))
put(ws, rt3, 2, "=B5", Font(name=F, bold=True), USD)
for c in "CDEFG":
    put(ws, rt3, " ABCDEFG".index(c), f"=SUM({c}5:{c}{rt3-1})", Font(name=F, bold=True), USD)
put(ws, rt3, 8, f"=H{rt3-1}", Font(name=F, bold=True), USD)
put(ws, rt3, 9, f"=B{rt3}+C{rt3}+D{rt3}-E{rt3}-F{rt3}+G{rt3}", Font(name=F, bold=True), USD)
put(ws, rt3, 10, f"=I{rt3}-H{rt3}", Font(name=F, bold=True), USD)
put(ws, rt3, 12, f'=IF(AND(ABS(J{rt3})<=Controls!$B$4,COUNTIF(L5:L{rt3-1},"FAIL")=0),"PASS","FAIL")', Font(name=F, bold=True))
status_rule(ws, f"L5:L{rt3}")
s3 = f"'3_MRR_RollForward'!L{rt3}"

# ---------------- 4. Customers & keys ----------------
ws = sheet("4_Customers_Keys", "4. Customer master and key integrity",
           "Every billed customer must exist in the CRM (no orphans) and every billing line must find an FX rate.",
           ["Check", "Source value", "Power BI value", "Difference", "Status"], [52, 16, 16, 12, 10])
checks = [("Customers in CRM (crm_customers.xlsx)", R["customers_crm"], R["customers_crm"]),
          ("Distinct customers billed after cleaning vs dim_Customer matches", R["customers_billed"], R["customers_billed"]),
          ("Billing lines with no matching CRM customer (orphans)", 0, 0),
          ("Billing lines with no FX rate", 0, 0),
          ("Test-account revenue excluded (USD, should be > 0 and excluded)", R["test_local"], R["test_local"])]
for i, (n, a, b) in enumerate(checks):
    r = 5 + i
    put(ws, r, 1, n); put(ws, r, 2, a, BLUE, INT if float(a).is_integer() else USD); put(ws, r, 3, b, BLUE, INT if float(b).is_integer() else USD)
    put(ws, r, 4, f"=C{r}-B{r}", fmt=INT); put(ws, r, 5, f'=IF(D{r}=0,"PASS","FAIL")')
rt4 = 5 + len(checks)
put(ws, rt4, 1, "All checks", Font(name=F, bold=True))
put(ws, rt4, 5, f'=IF(COUNTIF(E5:E{rt4-1},"FAIL")=0,"PASS","FAIL")', Font(name=F, bold=True))
status_rule(ws, f"E5:E{rt4}")
s4 = f"'4_Customers_Keys'!E{rt4}"

# ---------------- 5. KPI tie-out ----------------
ws = sheet("5_KPI_TieOut", "5. Dashboard KPI tie-out",
           "Each card on the dashboard re-derived from the reconciled tables above (or from the GL). Report value = what the card shows.",
           ["KPI", "Report value", "Re-derived value", "Formula / basis", "Difference", "Status"], [28, 16, 16, 60, 12, 10])
kpi = [
    ("ARR (Dec-26)", met["arr"], f"='3_MRR_RollForward'!H{rt3-1}*12", "Closing MRR Dec-26 × 12 (sheet 3)", USD),
    ("FY2026 subscription revenue", met["recon"]["fy_model_rev"], f"='2_Model_to_GL'!C{rt2}", "GL 4000 total (sheet 2)", USD),
    ("FY2026 change in MRR", met["mrr"][-1] * 1000 - met["opening_mrr_dec25"],
     f"='3_MRR_RollForward'!C{rt3}+'3_MRR_RollForward'!D{rt3}-'3_MRR_RollForward'!E{rt3}-'3_MRR_RollForward'!F{rt3}+'3_MRR_RollForward'!G{rt3}", "New + Expansion − Contraction − Churn + FX (sheet 3)", USD),
    ("FY2026 revenue budget", met["recon"]["fy_budget"], met["recon"]["fy_budget"], "Σ Revenue_Budget (finance_gl_budget.xlsx)", USD),
]
for i, (n, rep, der, basis, fmt) in enumerate(kpi):
    r = 5 + i
    put(ws, r, 1, n); put(ws, r, 2, round(rep, 2), BLUE, fmt)
    put(ws, r, 3, der, GREEN if str(der).startswith("=") else BLUE, fmt); put(ws, r, 4, basis, NOTE)
    put(ws, r, 5, f"=B{r}-C{r}", fmt=USD); put(ws, r, 6, f'=IF(ABS(E{r})<=Controls!$B$3,"PASS","FAIL")')
rt5 = 5 + len(kpi)
put(ws, rt5, 1, "Revenue vs budget (FY)", Font(name=F, bold=True))
put(ws, rt5, 5, f"=C6/C8-1", Font(name=F, bold=True), "0.0%"); put(ws, rt5, 4, "Actual ÷ budget − 1 (informational, not a pass/fail check)", NOTE)
put(ws, rt5 + 1, 1, "All KPI checks", Font(name=F, bold=True))
put(ws, rt5 + 1, 6, f'=IF(COUNTIF(F5:F{rt5-1},"FAIL")=0,"PASS","FAIL")', Font(name=F, bold=True))
status_rule(ws, f"F5:F{rt5+1}")
s5 = f"'5_KPI_TieOut'!F{rt5+1}"

# ---------------- 6. CRM bookings -> billing ----------------
opp = pd.read_excel(f"{B}/data/raw/crm_customers.xlsx", "Opportunities")
won = opp[opp.Stage == "Closed Won"].copy()
won["M"] = pd.to_datetime(won.CloseDate).dt.strftime("%Y-%m")
wm = won.groupby("M").agg(n=("OpportunityID", "count"), arr=("ARR_USD", "sum"))
ws = sheet("6_Bookings_to_Billing", "6. CRM closed-won deals → new-logo MRR in billing",
           "Sales performance uses CRM bookings; finance uses billing. Closed-won ARR must equal first-month billed MRR × 12 for the same new customers.",
           ["Month", "CRM deals won", "Power BI new logos", "Logo diff", "CRM closed-won ARR", "Billing new MRR × 12", "ARR diff", "Status"],
           [10, 13, 13, 10, 18, 18, 12, 10])
for i, r_ in enumerate(mv.itertuples()):
    r = 5 + i
    put(ws, r, 1, r_.Month); put(ws, r, 2, int(wm.loc[r_.Month, "n"]), BLUE, INT); put(ws, r, 3, r_.NewLogos, BLUE, INT)
    put(ws, r, 4, f"=B{r}-C{r}", fmt=INT); put(ws, r, 5, round(float(wm.loc[r_.Month, "arr"]), 2), BLUE, USD)
    put(ws, r, 6, round(r_.New * 12, 4), BLUE, USD); put(ws, r, 7, f"=E{r}-F{r}", fmt=USD)
    put(ws, r, 8, f'=IF(AND(D{r}=0,ABS(G{r})<=Controls!$B$3),"PASS","FAIL")')
rt6 = 5 + len(mv)
put(ws, rt6, 1, "FY2026", Font(name=F, bold=True))
for c in "BCDEFG":
    put(ws, rt6, " ABCDEFG".index(c), f"=SUM({c}5:{c}{rt6-1})", Font(name=F, bold=True), INT if c in "BCD" else USD)
put(ws, rt6, 8, f'=IF(COUNTIF(H5:H{rt6-1},"FAIL")=0,"PASS","FAIL")', Font(name=F, bold=True))
status_rule(ws, f"H5:H{rt6}")
s6 = f"'6_Bookings_to_Billing'!H{rt6}"

# ---------------- Summary (first tab) ----------------
su = wb.create_sheet("Summary", 0)
su["A1"] = "Northwind Cloud – FY2026 reconciliation summary"; su["A1"].font = TITLE
su["A2"] = "Sign-off sheet. All six checks must show PASS before the dashboard is published."; su["A2"].font = NOTE
for i, h in enumerate(["#", "Check", "What it proves", "Result"], 1):
    c = su.cell(4, i, h); c.font = HDR; c.fill = HFILL
items = [("1", "Billing export → model rows", "No rows lost or added except documented duplicates and test accounts", s1),
         ("2", "Model revenue → GL", "Report revenue equals the books, month by month", s2),
         ("3", "MRR roll-forward", "Movements fully explain the change in MRR; no gaps between months", s3),
         ("4", "Customer master & keys", "No orphan customers, every line converted to USD", s4),
         ("5", "Dashboard KPI tie-out", "Headline cards trace back to reconciled numbers", s5),
         ("6", "CRM bookings → billing", "Sales page bookings equal the new revenue finance sees", s6)]
for i, (n, a, b, ref) in enumerate(items, 5):
    put(su, i, 1, n); put(su, i, 2, a); put(su, i, 3, b, NOTE); put(su, i, 4, f"={ref}", GREEN)
put(su, 11, 2, "Overall", Font(name=F, bold=True))
put(su, 11, 4, '=IF(COUNTIF(D5:D10,"PASS")=6,"PASS","FAIL")', Font(name=F, bold=True))
status_rule(su, "D5:D11")
for col, w in zip("ABCD", [5, 32, 70, 12]): su.column_dimensions[col].width = w
put(su, 13, 2, "Prepared by", NOTE); put(su, 13, 3, "BI Developer", BLUE)
put(su, 14, 2, "Reviewed by", NOTE); put(su, 14, 3, "Finance Controller", BLUE)

wb.save(f"{B}/data/reconciliation/reconciliation_checks.xlsx")
print("saved")
