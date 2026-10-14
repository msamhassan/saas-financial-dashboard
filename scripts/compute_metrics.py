"""
Independent re-computation of every dashboard figure straight from the Excel source files,
applying the same cleaning rules as the Power Query (M) layer. Used to:
  1) drive the dashboard mockup image, and
  2) provide the "expected" side of the reconciliation workbook.
Writes data/processed/metrics.json and data/processed/*.csv.
"""
import json, os
import pandas as pd

BASE = os.path.join(os.path.dirname(__file__), "..")
RAW, PROC = os.path.join(BASE, "data/raw"), os.path.join(BASE, "data/processed")
os.makedirs(PROC, exist_ok=True)

bill_raw = pd.read_excel(f"{RAW}/billing_subscription_export.xlsx", "MRR_Lines")
cust = pd.read_excel(f"{RAW}/crm_customers.xlsx", "Customers")
prod = pd.read_excel(f"{RAW}/crm_customers.xlsx", "Products")
gl = pd.read_excel(f"{RAW}/finance_gl_budget.xlsx", "GL_Actuals")
bud = pd.read_excel(f"{RAW}/finance_gl_budget.xlsx", "Revenue_Budget")
fx = pd.read_excel(f"{RAW}/finance_gl_budget.xlsx", "FX_Rates")
cash = pd.read_excel(f"{RAW}/finance_gl_budget.xlsx", "Cash_Balance")

# ---- same steps as stg_Billing in Power Query ----
b = bill_raw.copy()
raw_rows, raw_local_by_month = len(b), None
b["CustomerID"] = b["CustomerID"].str.strip().str.upper()
b["ProductCode"] = b["ProductCode"].str.strip().str.upper()
b["MonthStart"] = pd.to_datetime(b["BillingPeriod"] + "-01")
n_dupes = int(b.duplicated().sum())
b = b.drop_duplicates()
n_test = int(b["CustomerID"].str.startswith("TEST-").sum())
test_local = float(b.loc[b["CustomerID"].str.startswith("TEST-"), "MRR_Local"].sum())
b = b[~b["CustomerID"].str.startswith("TEST-")]
fx["Month"] = pd.to_datetime(fx["Month"])
b = b.merge(fx, left_on=["MonthStart", "Currency"], right_on=["Month", "Currency"], how="left")
assert b["USD_per_Unit"].notna().all(), "missing FX rate"
b["MRR_USD"] = b["MRR_Local"] * b["USD_per_Unit"]
assert b["CustomerID"].isin(cust["CustomerID"]).all(), "orphan customer"
b = b.merge(cust[["CustomerID", "Region", "Segment"]], on="CustomerID").merge(prod[["ProductCode", "PlanTier", "ProductType"]], on="ProductCode")

months = sorted(b["MonthStart"].unique())
fy = months[1:]

# ---- movements (same as fact_MRR_Movements) ----
cm = b.groupby(["CustomerID", "MonthStart"])["MRR_USD"].sum().unstack(fill_value=0).reindex(columns=months, fill_value=0)
# Movements are classified in CONSTANT CURRENCY (the customer's billing currency), so an exchange-rate move
# is never mistaken for an upsell or a downgrade. The FX effect is reported as its own bridge line.
cmL = b.groupby(["CustomerID", "MonthStart"])["MRR_Local"].sum().unstack(fill_value=0).reindex(index=cm.index, columns=months, fill_value=0)
ccy = cust.set_index("CustomerID")["BillingCurrency"].reindex(cm.index)
rate = pd.DataFrame({m: ccy.map(fx[fx.Month == m].set_index("Currency")["USD_per_Unit"]) for m in months})
EPS = 0.005
mv = []
for i, m in enumerate(fy, start=1):
    p = months[i - 1]
    pL, cL, pU, cU, r = cmL[p], cmL[m], cm[p], cm[m], rate[m]
    cont = (pL > EPS) & (cL > EPS)
    dL = (cL - pL).where(cont, 0)
    new_ = (pL <= EPS) & (cL > EPS); churn_ = (pL > EPS) & (cL <= EPS)
    exp_ = cont & (dL > EPS); con_ = cont & (dL < -EPS)
    mv.append(dict(Month=pd.Timestamp(m).strftime("%Y-%m"),
                   Opening=pU.sum(),
                   New=cU[new_].sum(),
                   Expansion=(dL * r)[exp_].sum(),
                   Contraction=-(dL * r)[con_].sum(),
                   Churn=pU[churn_].sum(),
                   FXImpact=(cU - pU - dL * r)[cont].sum(),
                   Closing=cU.sum(),
                   OpeningLogos=int((pL > EPS).sum()),
                   NewLogos=int(new_.sum()),
                   ChurnedLogos=int(churn_.sum()),
                   ExpandedLogos=int(exp_.sum()), ContractedLogos=int(con_.sum()),
                   ClosingLogos=int((cL > EPS).sum())))
mv = pd.DataFrame(mv)
mv["RollForwardCheck"] = mv.Opening + mv.New + mv.Expansion - mv.Contraction - mv.Churn + mv.FXImpact - mv.Closing
mv.to_csv(f"{PROC}/mrr_movements_by_month.csv", index=False)

dec25, dec26 = months[0], months[-1]
coh = cm[cm[dec25] > 0]
nrr = coh[dec26].sum() / coh[dec25].sum()
grr = coh[[dec25, dec26]].min(axis=1).sum() / coh[dec25].sum()
logo_churn = (mv.ChurnedLogos / mv.OpeningLogos).mean()

# ---- P&L ----
gl["PostingMonth"] = pd.to_datetime(gl["PostingMonth"])
pl = gl.pivot_table(index="PostingMonth", columns="Category", values="Amount_USD", aggfunc="sum")
pl["GM%"] = 1 - pl["COGS"] / pl["Revenue"]
pl["NetBurn"] = pl["COGS"] + pl[["OpEx - S&M", "OpEx - R&D", "OpEx - G&A"]].sum(axis=1) - pl["Revenue"]
cash["Month"] = pd.to_datetime(cash["Month"])
cash = cash.set_index("Month")["ClosingCash_USD"]
runway = lambda i: cash.iloc[i] / pl["NetBurn"].iloc[i - 3:i].mean()  # cash at month i / avg burn of last 3 months
runway_dec, runway_jun = runway(12), runway(6)

# ---- unit economics by segment (FY2026) ----
seg_rows = []
sm = gl[gl.Category == "OpEx - S&M"].groupby("Segment")["Amount_USD"].sum()
gm_fy = 1 - pl["COGS"].sum() / pl["Revenue"].sum()
seg_of = cust.set_index("CustomerID")["Segment"]
for s in ["SMB", "Mid-Market", "Enterprise"]:
    ids = seg_of[seg_of == s].index.intersection(cm.index)
    c = cm.loc[ids]
    new_mrr = sum(c.loc[(c[months[i - 1]] == 0) & (c[m] > 0), m].sum() for i, m in enumerate(fy, 1))
    new_logos = sum(int(((c[months[i - 1]] == 0) & (c[m] > 0)).sum()) for i, m in enumerate(fy, 1))
    churn_r = sum(int(((c[months[i - 1]] > 0) & (c[m] == 0)).sum()) for i, m in enumerate(fy, 1)) / \
              sum(int((c[months[i - 1]] > 0).sum()) for i, m in enumerate(fy, 1))
    cac = sm[s] / new_logos
    arpa = new_mrr / new_logos
    payback = cac / (arpa * gm_fy)
    lifetime = min(1 / churn_r, 60)  # cap at 60 months
    ltv = arpa * gm_fy * lifetime
    seg_rows.append(dict(Segment=s, SM_Spend=sm[s], NewLogos=new_logos, CAC=cac, NewARPA=arpa,
                         MonthlyLogoChurn=churn_r, LifetimeMonths=lifetime, LTV=ltv,
                         CAC_Payback_Months=payback, LTV_to_CAC=ltv / cac))
unit = pd.DataFrame(seg_rows)
unit.to_csv(f"{PROC}/unit_economics_by_segment.csv", index=False)

# ---- tier mix (Dec-26 ARR) ----
d26 = b[b.MonthStart == dec26]
tier = d26.groupby("PlanTier")["MRR_USD"].sum() / d26["MRR_USD"].sum()

# ---- budget vs actual by quarter ----
bud["Month"] = pd.to_datetime(bud["Month"])
act_q = b[b.MonthStart.isin(fy)].groupby(b.MonthStart.dt.quarter)["MRR_USD"].sum()
bud_q = bud.groupby(bud.Month.dt.quarter)["BudgetRevenue_USD"].sum()

# ---- regions ----
reg = []
for r in ["North America", "EMEA", "APAC", "LATAM"]:
    x = b[b.Region == r]
    a26, a25 = x[x.MonthStart == dec26].MRR_USD.sum() * 12, x[x.MonthStart == dec25].MRR_USD.sum() * 12
    ids = cust.loc[cust.Region == r, "CustomerID"]
    c = cm.loc[cm.index.intersection(ids)]
    ch = [((c[months[i - 1]] > 0) & (c[m] == 0)).sum() / (c[months[i - 1]] > 0).sum() for i, m in enumerate(fy, 1)]
    reg.append(dict(Region=r, ARR=a26, YoY=a26 / a25 - 1, LogoChurn=sum(ch) / len(ch)))

# ---- reconciliation control totals ----
gl_rev = gl[gl.AccountCode == 4000].set_index("PostingMonth")["Amount_USD"]
recon_month = pd.DataFrame({
    "Month": [pd.Timestamp(m).strftime("%Y-%m") for m in fy],
    "RawRows": [int((bill_raw.BillingPeriod == pd.Timestamp(m).strftime("%Y-%m")).sum()) for m in fy],
    "ModelRows": [int((b.MonthStart == m).sum()) for m in fy],
    "Model_MRR_USD": [b.loc[b.MonthStart == m, "MRR_USD"].sum() for m in fy],
    "GL_Revenue_USD": [gl_rev.loc[m] for m in fy],
})
# rows removed per month (duplicates + test) for the bridge
bb = bill_raw.copy()
bb["CustomerID"] = bb["CustomerID"].str.strip().str.upper(); bb["ProductCode"] = bb["ProductCode"].str.strip().str.upper()
bb["dup"] = bb.duplicated()
bb["test"] = bb.CustomerID.str.startswith("TEST-") & ~bb.dup
recon_month["DuplicateRows"] = [int(bb.loc[bb.BillingPeriod == mm, "dup"].sum()) for mm in recon_month.Month]
recon_month["TestRows"] = [int(bb.loc[bb.BillingPeriod == mm, "test"].sum()) for mm in recon_month.Month]
recon_month.to_csv(f"{PROC}/recon_by_month.csv", index=False)

metrics = dict(
    months=[pd.Timestamp(m).strftime("%b") for m in fy],
    mrr=list(mv.Closing / 1000), new=list(mv.New / 1000), exp=list(mv.Expansion / 1000),
    churn_contr=list(-(mv.Churn + mv.Contraction) / 1000), fx=list(mv.FXImpact / 1000),
    opening_mrr_dec25=float(mv.Opening.iloc[0]), arr=float(mv.Closing.iloc[-1] * 12),
    arr_dec25=float(mv.Opening.iloc[0] * 12), mrr_mom=float((mv.Closing.iloc[-1] - mv.Closing.iloc[-2])),
    nrr=float(nrr), grr=float(grr), logo_churn=float(logo_churn),
    logo_churn_h1=float((mv.ChurnedLogos / mv.OpeningLogos)[:6].mean()),
    logo_churn_h2=float((mv.ChurnedLogos / mv.OpeningLogos)[6:].mean()),
    gm=list(pl["GM%"]), gm_fy=float(gm_fy), burn=list(pl["NetBurn"] / 1000),
    runway_dec=float(runway_dec), runway_jun=float(runway_jun), cash_dec=float(cash.iloc[-1]),
    tier={k: float(v) for k, v in tier.items()},
    unit=unit.to_dict("records"),
    bva=dict(budget=list(bud_q / 1000), actual=list(act_q / 1000)),
    regions=reg,
    recon=dict(raw_rows=raw_rows, dup_rows=n_dupes, test_rows=n_test, test_local=test_local,
               model_rows=len(b), customers_crm=len(cust), customers_billed=int(b.CustomerID.nunique()),
               fy_model_rev=float(b[b.MonthStart.isin(fy)].MRR_USD.sum()), fy_gl_rev=float(gl_rev.sum()),
               rollforward_max_abs=float(mv.RollForwardCheck.abs().max()),
               fy_budget=float(bud.BudgetRevenue_USD.sum())),
)
# =====================================================================
# Page 2 – Decomposition trees
# =====================================================================
def tree(df, value, levels, top=7):
    """Power BI decomposition-tree style: at each level, split the selected node and select the largest child."""
    out, cur = [], df
    for lvl in levels:
        g = cur.groupby(lvl)[value].sum().sort_values(ascending=False)
        nodes = [dict(label=str(k), value=float(v)) for k, v in g.items()]
        if len(nodes) > top:
            rest = sum(n["value"] for n in nodes[top - 1:])
            nodes = nodes[:top - 1] + [dict(label=f"Other ({len(g) - top + 1})", value=rest)]
        sel = str(g.index[0])
        out.append(dict(level=lvl, nodes=nodes, selected=sel))
        cur = cur[cur[lvl].astype(str) == sel]
    return dict(total=float(df[value].sum()), levels=out)


custx = cust.set_index("CustomerID")
d26c = d26.merge(cust[["CustomerID", "Country", "Industry", "AcquisitionChannel"]], on="CustomerID")
d26c["ARR"] = d26c["MRR_USD"] * 12
arr_tree = tree(d26c, "ARR", ["Region", "Country", "Segment", "Industry"])

churn_rows = []
for i, m in enumerate(fy, start=1):
    prev, cur = cm[months[i - 1]], cm[m]
    for cid, v in prev[(prev > 0) & (cur == 0)].items():
        churn_rows.append(dict(CustomerID=cid, ChurnMRR=v))
chdf = pd.DataFrame(churn_rows).merge(cust, on="CustomerID")
churn_tree = tree(chdf, "ChurnMRR", ["Segment", "Region", "Industry", "AcquisitionChannel"])

# =====================================================================
# Page 3 – Sales rep performance
# =====================================================================
reps = pd.read_excel(f"{RAW}/crm_customers.xlsx", "Sales_Reps")
opp = pd.read_excel(f"{RAW}/crm_customers.xlsx", "Opportunities")
opp["CloseDate"] = pd.to_datetime(opp["CloseDate"]); opp["CreatedDate"] = pd.to_datetime(opp["CreatedDate"])
opp["CycleDays"] = (opp.CloseDate - opp.CreatedDate).dt.days
won, lost = opp[opp.Stage == "Closed Won"], opp[opp.Stage == "Closed Lost"]
own = cust.set_index("CustomerID")["AccountOwnerID"]
rep_rows = []
for r in reps.itertuples():
    w, l = won[won.AccountOwnerID == r.SalesRepID], lost[lost.AccountOwnerID == r.SalesRepID]
    book = cm.loc[cm.index.intersection(own[own == r.SalesRepID].index)]
    coh_r = book[book[dec25] > 0]
    rep_rows.append(dict(id=r.SalesRepID, name=r.SalesRepName, team=r.Team, segs=r.SegmentsCovered,
                         new_arr=float(w.ARR_USD.sum()), quota=float(r.AnnualQuota_NewARR_USD),
                         attain=float(w.ARR_USD.sum() / r.AnnualQuota_NewARR_USD), won=len(w), lost=len(l),
                         win_rate=len(w) / (len(w) + len(l)), avg_deal=float(w.ARR_USD.mean()),
                         cycle=float(w.CycleDays.mean()),
                         book_nrr=float(coh_r[dec26].sum() / coh_r[dec25].sum()) if len(coh_r) else None,
                         book_arr=float(book[dec26].sum() * 12)))
rep_rows.sort(key=lambda x: -x["new_arr"])
team_month = won.assign(M=won.CloseDate.dt.month).pivot_table(index="M", columns="Region", values="ARR_USD", aggfunc="sum").reindex(range(1, 13)).fillna(0)
sales = dict(
    reps=rep_rows,
    new_arr=float(won.ARR_USD.sum()), quota=float(reps.AnnualQuota_NewARR_USD.sum()),
    win_rate=len(won) / len(opp), avg_deal=float(won.ARR_USD.mean()), cycle=float(won.CycleDays.mean()),
    deals_won=len(won), deals_lost=len(lost), reps_at_quota=sum(1 for x in rep_rows if x["attain"] >= 1),
    team_month={c: list(team_month[c] / 1000) for c in ["North America", "EMEA", "APAC", "LATAM"]},
    loss=lost.groupby("LossReason").agg(n=("OpportunityID", "count"), arr=("ARR_USD", "sum")).sort_values("n", ascending=False).reset_index().to_dict("records"),
    seg_win={s: len(won[won.Segment == s]) / len(opp[opp.Segment == s]) for s in ["SMB", "Mid-Market", "Enterprise"]},
    booked_vs_model_new_arr=float(won.ARR_USD.sum() - mv.New.sum() * 12),
)

# =====================================================================
# Page 4 – Customer insights
# =====================================================================
rep_name = reps.set_index("SalesRepID")["SalesRepName"]
prod_n = d26.groupby("CustomerID")["ProductCode"].nunique()
top = (cm[dec26] * 12).sort_values(ascending=False).head(10)
top10 = []
for rank, (cid, arr) in enumerate(top.items(), 1):
    c = custx.loc[cid]; a25 = cm.loc[cid, dec25] * 12
    top10.append(dict(rank=rank, name=c.CustomerName, seg=c.Segment, region=c.Region, industry=c.Industry,
                      arr=float(arr), yoy=float(arr / a25 - 1) if a25 > 0 else None, products=int(prod_n.get(cid, 0)),
                      tenure=int((pd.Timestamp(dec26) - pd.Timestamp(c.SignupDate)).days // 30.44),
                      owner=rep_name[c.AccountOwnerID]))
m_sep = months[-4]   # Sep-26: at-risk = MRR down vs 3 months ago and still active
# measured in constant (billing) currency so FX moves don't flag healthy accounts
risk_ids = cmL.index[(cmL[dec26] > EPS) & (cmL[dec26] < cmL[m_sep] - EPS)]
risk = cm.loc[risk_ids].copy()
risk["drop"] = (cmL.loc[risk_ids, m_sep] - cmL.loc[risk_ids, dec26]) * rate.loc[risk_ids, dec26]
risk_list = []
for cid, r_ in risk.sort_values("drop", ascending=False).head(6).iterrows():
    c = custx.loc[cid]
    risk_list.append(dict(name=c.CustomerName, seg=c.Segment, arr=float(r_[dec26] * 12),
                          change=float(cmL.loc[cid, dec26] / cmL.loc[cid, m_sep] - 1), owner=rep_name[c.AccountOwnerID]))
# cohort revenue retention: customers acquired in each FY26 month, MRR in month k / MRR in month 0
coh_rows = []
for i, m0 in enumerate(fy):
    prev = months[i]  # month before m0
    ids = cm.index[(cm[prev] == 0) & (cm[m0] > 0)]
    base = cm.loc[ids, m0].sum()
    coh_rows.append(dict(cohort=pd.Timestamp(m0).strftime("%b"), n=len(ids),
                         ret=[float(cm.loc[ids, fy[j]].sum() / base) for j in range(i, len(fy))]))
ind = (d26c.groupby("Industry")["ARR"].sum().sort_values(ascending=False) / 1e6)
active = [int((cm[m] > 0).sum()) for m in fy]
custpage = dict(
    active=active[-1], active_dec25=int((cm[dec25] > 0).sum()), arpa=float(cm[dec26].sum() / active[-1]),
    top10=top10, top10_share=float(top.sum() / (cm[dec26].sum() * 12)), risk=risk_list,
    risk_n=len(risk), risk_arr=float(risk[dec26].sum() * 12), cohorts=coh_rows,
    industry=[dict(label=k, value=float(v)) for k, v in ind.items()],
    active_series=active, new_logos=list(mv.NewLogos), churned_logos=list(mv.ChurnedLogos),
    median_tenure=float(((pd.Timestamp(dec26) - pd.to_datetime(cust.loc[cust.CustomerID.isin(cm.index[cm[dec26] > 0]), "SignupDate"])).dt.days / 30.44).median()),
)
metrics.update(arr_tree=arr_tree, churn_tree=churn_tree, sales=sales, cust=custpage)
pd.DataFrame(rep_rows).to_csv(f"{PROC}/sales_rep_performance.csv", index=False)
pd.DataFrame(top10).to_csv(f"{PROC}/top10_customers.csv", index=False)

json.dump(metrics, open(f"{PROC}/metrics.json", "w"), indent=1, default=float)
print(mv[["Month", "Opening", "New", "Expansion", "Contraction", "Churn", "Closing", "NewLogos", "ChurnedLogos"]].round(0).to_string())
print(unit.round(2).to_string())
print({k: v for k, v in metrics.items() if k not in ("mrr", "new", "exp", "churn_contr", "gm", "burn", "unit", "months")})
