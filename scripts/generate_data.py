"""
Generates the synthetic source files for the Northwind Cloud SaaS financial dashboard.

Outputs (data/raw/):
  billing_subscription_export.xlsx  - monthly billed MRR lines (local currency), with realistic data-quality issues
  crm_customers.xlsx                - customer master + product catalogue
  finance_gl_budget.xlsx            - GL actuals, revenue budget, FX rates, cash balances

All data is fictional. Seeded, so every run produces identical files.
"""
import random, math, os
from datetime import date
import pandas as pd
from openpyxl.styles import Font

random.seed(42)
OUT = os.path.join(os.path.dirname(__file__), "..", "data", "raw")
os.makedirs(OUT, exist_ok=True)

MONTHS = [date(2025, 12, 1)] + [date(2026, m, 1) for m in range(1, 13)]  # Dec-25 opening + FY2026

# ---------------- reference data ----------------
REGIONS = {
    "North America": [("United States", "USD", .78), ("Canada", "CAD", .22)],
    "EMEA": [("United Kingdom", "GBP", .40), ("Germany", "EUR", .35), ("France", "EUR", .25)],
    "APAC": [("Australia", "AUD", .55), ("Singapore", "USD", .45)],
    "LATAM": [("Brazil", "BRL", .60), ("Mexico", "USD", .40)],
}
REGION_W = {"North America": .50, "EMEA": .29, "APAC": .15, "LATAM": .06}
FX_BASE = {"USD": 1.0, "CAD": 0.73, "EUR": 1.09, "GBP": 1.27, "AUD": 0.66, "BRL": 0.19}

TIERS = {  # tier: (segment, core product, MRR median, sigma, monthly churn prob)
    "Starter": ("SMB", "STR-CORE", 450, .35, .024),
    "Growth": ("Mid-Market", "GRW-CORE", 2800, .30, .012),
    "Enterprise": ("Enterprise", "ENT-CORE", 15000, .30, .005),
}
PRODUCTS = [
    ("STR-CORE", "Starter Plan", "Starter", "Core Plan", 49),
    ("GRW-CORE", "Growth Plan", "Growth", "Core Plan", 79),
    ("ENT-CORE", "Enterprise Plan", "Enterprise", "Core Plan", 119),
    ("ADD-ANL", "Advanced Analytics Add-on", "Add-on", "Add-on", 15),
    ("ADD-SEC", "Security & SSO Add-on", "Add-on", "Add-on", 12),
    ("ADD-API", "API Volume Pack", "Add-on", "Add-on", 0.02),
]
INDUSTRIES = ["Logistics", "Retail", "Healthcare", "Financial Services", "Manufacturing",
              "Education", "Media", "Professional Services", "Energy", "Hospitality"]
CHANNELS = ["Inbound", "Outbound SDR", "Partner", "Self-serve"]
W1 = ["Blue", "North", "Iron", "Silver", "Maple", "Harbor", "Summit", "Cedar", "Bright", "Atlas",
      "Pioneer", "Lumen", "Granite", "Coral", "Vertex", "Willow", "Falcon", "Orbit", "Prairie", "Nova"]
W2 = ["peak", "line", "field", "stone", "bridge", "gate", "wave", "point", "ridge", "works",
      "path", "light", "core", "forge", "dale", "port", "crest", "brook", "view", "spring"]
SUF = ["Logistics", "Labs", "Health", "Retail Group", "Systems", "Partners", "Foods", "Energy",
       "Analytics", "Studios", "Manufacturing", "Capital", "Supply Co.", "Learning", "Hotels"]


def fx_rate(cur, mi):
    if cur == "USD":
        return 1.0
    drift = 1 + 0.012 * math.sin(mi / 2.0 + len(cur)) + random.uniform(-0.004, 0.004)
    return round(FX_BASE[cur] * drift, 6)


FX = {(m, c): fx_rate(c, i) for i, m in enumerate(MONTHS) for c in FX_BASE}

used_names = set()


def company_name():
    while True:
        n = f"{random.choice(W1)}{random.choice(W2)} {random.choice(SUF)}"
        if n not in used_names:
            used_names.add(n)
            return n


customers = {}  # id -> dict
state = {}      # id -> {product: usd_mrr}
next_id = [1]


def new_customer(tier, signup):
    cid = f"C-{next_id[0]:05d}"
    next_id[0] += 1
    region = random.choices(list(REGION_W), weights=list(REGION_W.values()))[0]
    country, cur, _ = random.choices(REGIONS[region], weights=[x[2] for x in REGIONS[region]])[0]
    seg, core, med, sig, _ = TIERS[tier]
    customers[cid] = dict(CustomerID=cid, CustomerName=company_name(), Industry=random.choice(INDUSTRIES),
                          Country=country, Region=region, Currency=cur, Segment=seg, PlanTier=tier,
                          AcquisitionChannel=random.choice(CHANNELS), SignupDate=signup,
                          Status="Active", ChurnDate=None)
    lines = {core: round(med * math.exp(random.gauss(0, sig)), 0)}
    if random.random() < 0.25:
        lines[random.choice(["ADD-ANL", "ADD-SEC", "ADD-API"])] = round(lines[core] * random.uniform(.10, .22), 0)
    state[cid] = lines
    return cid


# opening book (Dec-2025)
for tier, n in [("Starter", 600), ("Growth", 186), ("Enterprise", 50)]:
    for _ in range(n):
        y = random.choice([2022, 2023, 2024, 2025])
        new_customer(tier, date(y, random.randint(1, 11), random.randint(1, 28)))

# scale opening book to exactly $1.80M MRR so the story has a clean starting point
tot = sum(sum(l.values()) for l in state.values())
for l in state.values():
    for p in l:
        l[p] = round(l[p] * 1_800_000 / tot, 0)

snap = {MONTHS[0]: {c: dict(l) for c, l in state.items()}}

for mi, m in enumerate(MONTHS[1:], start=1):
    g = 1 + 0.03 * mi
    for cid in list(state):
        tier = customers[cid]["PlanTier"]
        r = random.random()
        if r < TIERS[tier][4]:                      # churn
            del state[cid]
            customers[cid]["Status"] = "Churned"
            customers[cid]["ChurnDate"] = m
            continue
        if r < TIERS[tier][4] + 0.010:              # contraction (seat reduction / downgrade)
            core = TIERS[tier][1]
            state[cid][core] = round(state[cid][core] * random.uniform(.65, .9), 0)
        elif r < TIERS[tier][4] + 0.010 + 0.090:    # expansion (seats) or add-on attach
            core = TIERS[tier][1]
            if random.random() < .35:
                addon = random.choice(["ADD-ANL", "ADD-SEC", "ADD-API"])
                state[cid][addon] = state[cid].get(addon, 0) + round(state[cid][core] * random.uniform(.10, .2), 0)
            else:
                state[cid][core] = round(state[cid][core] * random.uniform(1.15, 1.45), 0)
    for tier, base in [("Starter", 32), ("Growth", 9), ("Enterprise", 1.6)]:
        n = int(base * g + random.uniform(-.5, 1.0) * base * .25 + .5)
        for _ in range(n):
            new_customer(tier, date(m.year, m.month, random.randint(1, 28)))
    snap[m] = {c: dict(l) for c, l in state.items()}

# ---------------- billing export (local currency) ----------------
rows = []
for m in MONTHS:
    for cid, lines in snap[m].items():
        cur = customers[cid]["Currency"]
        for prod, usd in lines.items():
            local = round(usd / FX_BASE[cur], 2)  # price list fixed in local currency
            rows.append(dict(BillingPeriod=m.strftime("%Y-%m"), InvoiceID=None, CustomerID=cid,
                             ProductCode=prod, Currency=cur, MRR_Local=local))
random.shuffle(rows)
rows.sort(key=lambda r: r["BillingPeriod"])
for i, r in enumerate(rows, 1):
    r["InvoiceID"] = f"INV-{r['BillingPeriod'].replace('-', '')}-{i:06d}"

# data-quality issues a real billing export would have
dq = []
for r in random.sample(rows, 140):           # inconsistent ID / code formatting
    r["CustomerID"] = "  " + r["CustomerID"].lower() + " "
    r["ProductCode"] = r["ProductCode"].lower()
dups = [dict(r) for r in random.sample(rows, 57)]   # re-exported invoice lines (exact duplicates)
rows += dups
for m in MONTHS[4:]:                          # QA test accounts billed $1
    for t in ["TEST-0001", "TEST-0002"]:
        rows.append(dict(BillingPeriod=m.strftime("%Y-%m"), InvoiceID=f"INV-TEST-{m:%Y%m}-{t[-1]}",
                         CustomerID=t, ProductCode="STR-CORE", Currency="USD", MRR_Local=1.00))
rows.sort(key=lambda r: (r["BillingPeriod"], r["InvoiceID"]))
billing = pd.DataFrame(rows)[["InvoiceID", "BillingPeriod", "CustomerID", "ProductCode", "Currency", "MRR_Local"]]

# ---------------- CRM ----------------
crm = pd.DataFrame(customers.values())
crm = crm.drop(columns=["PlanTier"]).rename(columns={"Currency": "BillingCurrency"})
products = pd.DataFrame(PRODUCTS, columns=["ProductCode", "ProductName", "PlanTier", "ProductType", "ListPricePerSeat_USD"])

# ---------------- Sales team, account ownership, opportunities ----------------
# Separate RNG so adding sales data does not change any billing / GL number above.
rng = random.Random(2026)
REPS = [  # id, name, team(region), segments covered, manager
    ("R01", "Olivia Bennett", "North America", ["SMB"], "Daniel Brooks"),
    ("R02", "Ethan Walker", "North America", ["SMB"], "Daniel Brooks"),
    ("R03", "Priya Nair", "North America", ["SMB", "Mid-Market"], "Daniel Brooks"),
    ("R04", "Marcus Chen", "North America", ["Mid-Market"], "Daniel Brooks"),
    ("R05", "Sofia Alvarez", "North America", ["Mid-Market", "Enterprise"], "Daniel Brooks"),
    ("R06", "James O'Connor", "North America", ["Enterprise"], "Daniel Brooks"),
    ("R07", "Hannah Schmidt", "EMEA", ["SMB"], "Claire Dubois"),
    ("R08", "Lucas Moreau", "EMEA", ["SMB", "Mid-Market"], "Claire Dubois"),
    ("R09", "Amara Okafor", "EMEA", ["Mid-Market", "Enterprise"], "Claire Dubois"),
    ("R10", "Kenji Tanaka", "APAC", ["SMB", "Mid-Market"], "Grace Lim"),
    ("R11", "Chloe Wright", "APAC", ["Mid-Market", "Enterprise"], "Grace Lim"),
    ("R12", "Rafael Costa", "LATAM", ["SMB", "Mid-Market", "Enterprise"], "Grace Lim"),
]
ROLE = {"SMB": "Account Executive", "Mid-Market": "Account Executive", "Enterprise": "Senior Account Executive"}


def owner_for(region, seg):
    pool = [r[0] for r in REPS if r[2] == region and seg in r[3]]
    return rng.choice(pool)


crm["AccountOwnerID"] = [owner_for(r, s) for r, s in zip(crm.Region, crm.Segment)]


def first_month_usd(cid, m):
    cur = customers[cid]["Currency"]
    return sum(round(u / FX_BASE[cur], 2) * FX[(m, cur)] for u in snap[m][cid].values())


CYCLE = {"SMB": (14, 45), "Mid-Market": (35, 90), "Enterprise": (75, 180)}
LOSS = ["Price", "Lost to competitor", "No decision", "Timing / budget", "Feature gap"]
LOSS_W = [.27, .30, .21, .14, .08]
rep_skill = {r[0]: rng.uniform(.18, .36) for r in REPS}   # win rate tendency per rep
owner = dict(zip(crm.CustomerID, crm.AccountOwnerID))
opps, won_by_rep = [], {}
oid = 1
for mi, m in enumerate(MONTHS[1:], start=1):
    for cid in snap[m]:
        if cid in snap[MONTHS[mi - 1]]:
            continue
        seg, rep = customers[cid]["Segment"], owner[cid]
        arr = round(first_month_usd(cid, m) * 12, 2)
        close = customers[cid]["SignupDate"]
        cyc = rng.randint(*CYCLE[seg])
        opps.append(dict(OpportunityID=f"OPP-{oid:05d}", CustomerID=cid, AccountOwnerID=rep, Segment=seg,
                         Region=customers[cid]["Region"], CreatedDate=pd.Timestamp(close) - pd.Timedelta(days=cyc),
                         CloseDate=pd.Timestamp(close), Stage="Closed Won", ARR_USD=arr, LossReason=None))
        oid += 1
        won_by_rep[rep] = won_by_rep.get(rep, 0) + arr
# closed-lost deals so each rep's win rate is realistic
for rep_id, name, team, segs, mgr in REPS:
    won = [o for o in opps if o["AccountOwnerID"] == rep_id and o["Stage"] == "Closed Won"]
    wr = rep_skill[rep_id]
    n_lost = int(round(len(won) * (1 - wr) / wr))
    for _ in range(n_lost):
        seg = rng.choice(segs)
        med = {"SMB": 450, "Mid-Market": 2800, "Enterprise": 15000}[seg] * 12
        close = pd.Timestamp(2026, rng.randint(1, 12), rng.randint(1, 28))
        cyc = rng.randint(*CYCLE[seg])
        opps.append(dict(OpportunityID=f"OPP-{oid:05d}", CustomerID=None, AccountOwnerID=rep_id, Segment=seg,
                         Region=team, CreatedDate=close - pd.Timedelta(days=cyc), CloseDate=close,
                         Stage="Closed Lost", ARR_USD=round(med * math.exp(rng.gauss(0, .35)), -2),
                         LossReason=rng.choices(LOSS, weights=LOSS_W)[0]))
        oid += 1
opps_df = pd.DataFrame(opps).sort_values(["CloseDate", "OpportunityID"]).reset_index(drop=True)
def prospect_name():
    while True:
        n = f"{rng.choice(W1)}{rng.choice(W2)} {rng.choice(SUF)}"
        if n not in used_names:
            used_names.add(n)
            return n


opps_df["ProspectName"] = [prospect_name() if pd.isna(c) else None for c in opps_df.CustomerID]
reps_df = pd.DataFrame([dict(SalesRepID=r[0], SalesRepName=r[1], Team=r[2], SegmentsCovered=", ".join(r[3]),
                             Role=ROLE[r[3][-1]], Manager=r[4],
                             AnnualQuota_NewARR_USD=round(won_by_rep.get(r[0], 0) * rng.uniform(.78, 1.22), -4))
                        for r in REPS])

# ---------------- finance: GL, budget, FX, cash ----------------
def usd_rev(m):
    return sum(sum(l.values()) for l in snap[m].values())


def seg_new_mrr(m, prev):
    out = {"SMB": 0, "Mid-Market": 0, "Enterprise": 0}
    for cid, l in snap[m].items():
        if cid not in snap[prev] and not cid.startswith("TEST"):
            out[customers[cid]["Segment"]] += sum(l.values())
    return out


gl, budget, cash = [], [], []
payback = {"SMB": 9, "Mid-Market": 14, "Enterprise": 19}
closing = 12_150_000.0
for mi, m in enumerate(MONTHS[1:], start=1):
    # revenue in GL = billed lines converted at month FX, rounded to cents
    rev = round(sum(round(usd / FX_BASE[customers[c]["Currency"]], 2) * FX[(m, customers[c]["Currency"])]
                    for c, l in snap[m].items() for usd in l.values()), 2)
    cogs_pct = 0.240 - 0.0026 * (mi - 1) + random.uniform(-.0015, .0015)
    hosting = round(rev * cogs_pct * .62, 2)
    support = round(rev * cogs_pct * .38, 2)
    gm = 1 - (hosting + support) / rev
    nm = seg_new_mrr(m, MONTHS[mi - 1])
    sm = {s: round(payback[s] * nm[s] * gm * random.uniform(.94, 1.06), 2) for s in nm}
    burn_target = 425_000 - 11_500 * (mi - 1) + random.uniform(-6000, 6000)
    other = burn_target + rev - hosting - support - sum(sm.values())
    rnd, ga = round(other * .64, 2), round(other * .36, 2)
    lines = [("4000", "Subscription Revenue", "Revenue", "All", rev),
             ("5000", "Cloud Hosting", "COGS", "All", hosting),
             ("5100", "Customer Support", "COGS", "All", support),
             ("6000", "Sales & Marketing", "OpEx - S&M", "SMB", sm["SMB"]),
             ("6000", "Sales & Marketing", "OpEx - S&M", "Mid-Market", sm["Mid-Market"]),
             ("6000", "Sales & Marketing", "OpEx - S&M", "Enterprise", sm["Enterprise"]),
             ("7000", "Research & Development", "OpEx - R&D", "All", rnd),
             ("8000", "General & Administrative", "OpEx - G&A", "All", ga)]
    for a, n, cat, seg, amt in lines:
        gl.append(dict(PostingMonth=m, AccountCode=a, AccountName=n, Category=cat, Segment=seg, Amount_USD=amt))
    net_burn = hosting + support + sum(sm.values()) + rnd + ga - rev
    closing = round(closing - net_burn, 2)
    cash.append(dict(Month=m, ClosingCash_USD=closing))

# revenue budget by region (set at start of year; quarterly over/under-run baked in)
qvar = {1: .016, 2: -.009, 3: .019, 4: .022}
for mi, m in enumerate(MONTHS[1:], start=1):
    q = (m.month - 1) // 3 + 1
    by_reg = {}
    for c, l in snap[m].items():
        cur = customers[c]["Currency"]
        by_reg[customers[c]["Region"]] = by_reg.get(customers[c]["Region"], 0) + sum(
            round(u / FX_BASE[cur], 2) * FX[(m, cur)] for u in l.values())
    for reg, act in by_reg.items():
        b = round(act / (1 + qvar[q] + random.uniform(-.006, .006)) / 1000, 0) * 1000
        budget.append(dict(Month=m, Region=reg, BudgetRevenue_USD=b))

fx = pd.DataFrame([dict(Month=m, Currency=c, USD_per_Unit=r) for (m, c), r in FX.items()])
cash.insert(0, dict(Month=MONTHS[0], ClosingCash_USD=12_150_000.0))


def save(path, sheets):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            for col in ws.columns:
                w = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(w + 2, 10), 40)
                for c in col:
                    c.font = Font(name="Arial", size=10, bold=(c.row == 1))
                    if isinstance(c.value, (date,)) and c.row > 1:
                        c.number_format = "yyyy-mm-dd"
                    if c.row > 1 and ws.cell(1, c.column).value and str(ws.cell(1, c.column).value).endswith(("USD", "Local")):
                        c.number_format = "#,##0.00"
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions


save(os.path.join(OUT, "billing_subscription_export.xlsx"), {"MRR_Lines": billing})
save(os.path.join(OUT, "crm_customers.xlsx"), {"Customers": crm, "Products": products,
                                                "Sales_Reps": reps_df, "Opportunities": opps_df})
save(os.path.join(OUT, "finance_gl_budget.xlsx"), {
    "GL_Actuals": pd.DataFrame(gl), "Revenue_Budget": pd.DataFrame(budget),
    "FX_Rates": fx, "Cash_Balance": pd.DataFrame(cash)})
print(f"billing rows={len(billing)}  customers={len(crm)}  gl rows={len(gl)}  budget rows={len(budget)}")
