// =====================================================================
// Northwind Cloud – SaaS Financial Dashboard
// Power Query (M) – one block per query. Paste each into
// Power Query Editor > New Source > Blank Query > Advanced Editor.
// =====================================================================


// ---------------------------------------------------------------------
// pDataFolder  (Parameter, Text) – folder that holds the three source workbooks
// ---------------------------------------------------------------------
"C:\Projects\saas-financial-dashboard\data\raw\" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]


// ---------------------------------------------------------------------
// fnReadSheet  (Function) – reads one sheet from one workbook and promotes headers
// ---------------------------------------------------------------------
(fileName as text, sheetName as text) as table =>
let
    Source   = Excel.Workbook(File.Contents(pDataFolder & fileName), null, true),
    Sheet    = Source{[Item = sheetName, Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true])
in
    Promoted


// ---------------------------------------------------------------------
// stg_Billing  (staging, Enable Load = OFF)
// Cleans the billing export: ID formatting, duplicates, QA test accounts
// ---------------------------------------------------------------------
let
    Source     = fnReadSheet("billing_subscription_export.xlsx", "MRR_Lines"),
    Typed      = Table.TransformColumnTypes(Source, {
                    {"InvoiceID", type text}, {"BillingPeriod", type text}, {"CustomerID", type text},
                    {"ProductCode", type text}, {"Currency", type text}, {"MRR_Local", type number}}),
    // 1. normalise keys – the export has values like "  c-00412 " and "grw-core"
    CleanKeys  = Table.TransformColumns(Typed, {
                    {"CustomerID",  each Text.Upper(Text.Trim(_)), type text},
                    {"ProductCode", each Text.Upper(Text.Trim(_)), type text}}),
    // 2. remove exact duplicate invoice lines produced by re-exports
    Deduped    = Table.Distinct(CleanKeys),
    // 3. drop QA test accounts (billed $1, not real revenue, not in GL)
    NoTest     = Table.SelectRows(Deduped, each not Text.StartsWith([CustomerID], "TEST-")),
    // 4. "2026-03" -> 2026-03-01
    MonthStart = Table.AddColumn(NoTest, "MonthStart", each Date.FromText([BillingPeriod] & "-01"), type date)
in
    MonthStart


// ---------------------------------------------------------------------
// stg_FX  (staging, Enable Load = OFF)
// ---------------------------------------------------------------------
let
    Source = fnReadSheet("finance_gl_budget.xlsx", "FX_Rates"),
    Typed  = Table.TransformColumnTypes(Source, {{"Month", type date}, {"Currency", type text}, {"USD_per_Unit", type number}})
in
    Typed


// ---------------------------------------------------------------------
// fact_MRR  – grain: invoice line (customer x product x month), USD
// ---------------------------------------------------------------------
let
    Source    = stg_Billing,
    MergeFX   = Table.NestedJoin(Source, {"MonthStart", "Currency"}, stg_FX, {"Month", "Currency"}, "FX", JoinKind.LeftOuter),
    ExpandFX  = Table.ExpandTableColumn(MergeFX, "FX", {"USD_per_Unit"}),
    // guard: every line must find a rate, otherwise fail the refresh loudly
    CheckFX   = if Table.RowCount(Table.SelectRows(ExpandFX, each [USD_per_Unit] = null)) > 0
                then error "Missing FX rate for one or more billing lines"
                else ExpandFX,
    AddUSD    = Table.AddColumn(CheckFX, "MRR_USD", each [MRR_Local] * [USD_per_Unit], type number),
    Selected  = Table.SelectColumns(AddUSD, {"InvoiceID", "MonthStart", "CustomerID", "ProductCode", "Currency", "MRR_Local", "USD_per_Unit", "MRR_USD"})
in
    Selected


// ---------------------------------------------------------------------
// fact_MRR_Movements – grain: customer x month
// Builds a full customer-month grid so churn (MRR going to 0) is visible, then classifies
// each month in CONSTANT CURRENCY (the customer's billing currency): an exchange-rate move
// is not an upsell or a downgrade, so it is reported separately as FXImpactMRR.
//   Opening + New + Expansion - Contraction - Churn + FX Impact = Closing   (all USD)
// ---------------------------------------------------------------------
let
    Eps         = 0.005,
    CustMonth   = Table.Buffer(Table.Group(fact_MRR, {"CustomerID", "MonthStart"}, {
                    {"MRR",      each List.Sum([MRR_USD]),       type number},
                    {"MRRLocal", each List.Sum([MRR_Local]),     type number},
                    {"Rate",     each List.Max([USD_per_Unit]),  type number}})),
    MinMonth    = List.Min(CustMonth[MonthStart]),
    MaxMonth    = List.Max(CustMonth[MonthStart]),
    Months      = List.Generate(() => MinMonth, each _ <= MaxMonth, each Date.AddMonths(_, 1)),
    Customers   = Table.Distinct(Table.SelectColumns(CustMonth, {"CustomerID"})),
    Grid        = Table.ExpandListColumn(Table.AddColumn(Customers, "MonthStart", each Months), "MonthStart"),
    GridTyped   = Table.TransformColumnTypes(Grid, {{"MonthStart", type date}}),
    AddPrev     = Table.AddColumn(GridTyped, "PrevMonth", each Date.AddMonths([MonthStart], -1), type date),
    JoinCurr    = Table.NestedJoin(AddPrev, {"CustomerID", "MonthStart"}, CustMonth, {"CustomerID", "MonthStart"}, "Curr", JoinKind.LeftOuter),
    ExpCurr     = Table.ExpandTableColumn(JoinCurr, "Curr", {"MRR", "MRRLocal", "Rate"}, {"ClosingMRR", "ClosingLocal", "ClosingRate"}),
    JoinPrev    = Table.NestedJoin(ExpCurr, {"CustomerID", "PrevMonth"}, CustMonth, {"CustomerID", "MonthStart"}, "Prev", JoinKind.LeftOuter),
    ExpPrev     = Table.ExpandTableColumn(JoinPrev, "Prev", {"MRR", "MRRLocal"}, {"OpeningMRR", "OpeningLocal"}),
    ZeroFill    = Table.ReplaceValue(ExpPrev, null, 0, Replacer.ReplaceValue, {"OpeningMRR", "ClosingMRR", "OpeningLocal", "ClosingLocal", "ClosingRate"}),
    // first month (Dec-2025) is the opening balance only; drop months with no MRR either side
    InPeriod    = Table.SelectRows(ZeroFill, each [MonthStart] > MinMonth and ([OpeningLocal] > Eps or [ClosingLocal] > Eps)),
    AddDelta    = Table.AddColumn(InPeriod, "DeltaLocal", each
                    if [OpeningLocal] > Eps and [ClosingLocal] > Eps then [ClosingLocal] - [OpeningLocal] else 0, type number),
    AddType     = Table.AddColumn(AddDelta, "MovementType", each
                    if [OpeningLocal] <= Eps then "New"
                    else if [ClosingLocal] <= Eps then "Churn"
                    else if [DeltaLocal] > Eps then "Expansion"
                    else if [DeltaLocal] < -Eps then "Contraction"
                    else "Flat", type text),
    AddNew      = Table.AddColumn(AddType, "NewMRR",         each if [MovementType] = "New"         then [ClosingMRR] else 0, type number),
    AddExp      = Table.AddColumn(AddNew,  "ExpansionMRR",   each if [MovementType] = "Expansion"   then [DeltaLocal] * [ClosingRate] else 0, type number),
    AddCon      = Table.AddColumn(AddExp,  "ContractionMRR", each if [MovementType] = "Contraction" then -[DeltaLocal] * [ClosingRate] else 0, type number),
    AddChurn    = Table.AddColumn(AddCon,  "ChurnMRR",       each if [MovementType] = "Churn"       then [OpeningMRR] else 0, type number),
    AddFX       = Table.AddColumn(AddChurn, "FXImpactMRR",   each
                    if [MovementType] = "New" or [MovementType] = "Churn" then 0
                    else [ClosingMRR] - [OpeningMRR] - [DeltaLocal] * [ClosingRate], type number),
    AddFlags    = Table.AddColumn(Table.AddColumn(Table.AddColumn(AddFX,
                    "IsOpeningLogo", each if [OpeningLocal] > Eps then 1 else 0, Int64.Type),
                    "IsNewLogo",     each if [MovementType] = "New" then 1 else 0, Int64.Type),
                    "IsChurnLogo",   each if [MovementType] = "Churn" then 1 else 0, Int64.Type),
    Result      = Table.RemoveColumns(AddFlags, {"PrevMonth"})
in
    Result


// ---------------------------------------------------------------------
// dim_Customer
// ---------------------------------------------------------------------
let
    Source  = fnReadSheet("crm_customers.xlsx", "Customers"),
    Typed   = Table.TransformColumnTypes(Source, {
                {"CustomerID", type text}, {"CustomerName", type text}, {"Industry", type text},
                {"Country", type text}, {"Region", type text}, {"BillingCurrency", type text},
                {"Segment", type text}, {"AcquisitionChannel", type text}, {"SignupDate", type date},
                {"Status", type text}, {"ChurnDate", type date}}),
    Cohort  = Table.AddColumn(Typed, "SignupCohort", each Date.StartOfMonth([SignupDate]), type date),
    SegSort = Table.AddColumn(Cohort, "SegmentSort", each
                if [Segment] = "SMB" then 1 else if [Segment] = "Mid-Market" then 2 else 3, Int64.Type),
    Owner   = Table.TransformColumnTypes(SegSort, {{"AccountOwnerID", type text}})
in
    Owner


// ---------------------------------------------------------------------
// dim_SalesRep
// ---------------------------------------------------------------------
let
    Source = fnReadSheet("crm_customers.xlsx", "Sales_Reps"),
    Typed  = Table.TransformColumnTypes(Source, {
               {"SalesRepID", type text}, {"SalesRepName", type text}, {"Team", type text},
               {"SegmentsCovered", type text}, {"Role", type text}, {"Manager", type text},
               {"AnnualQuota_NewARR_USD", type number}})
in
    Typed


// ---------------------------------------------------------------------
// fact_Opportunities – closed won / lost deals from the CRM
// ---------------------------------------------------------------------
let
    Source  = fnReadSheet("crm_customers.xlsx", "Opportunities"),
    Typed   = Table.TransformColumnTypes(Source, {
                {"OpportunityID", type text}, {"CustomerID", type text}, {"AccountOwnerID", type text},
                {"Segment", type text}, {"Region", type text}, {"CreatedDate", type date}, {"CloseDate", type date},
                {"Stage", type text}, {"ARR_USD", type number}, {"LossReason", type text}, {"ProspectName", type text}}),
    Account = Table.AddColumn(Typed, "AccountName", each if [CustomerID] = null then [ProspectName] else null, type text),
    Cycle   = Table.AddColumn(Account, "SalesCycleDays", each Duration.Days([CloseDate] - [CreatedDate]), Int64.Type),
    IsWon   = Table.AddColumn(Cycle, "IsWon", each if [Stage] = "Closed Won" then 1 else 0, Int64.Type),
    CloseM  = Table.AddColumn(IsWon, "CloseMonth", each Date.StartOfMonth([CloseDate]), type date)
in
    CloseM


// ---------------------------------------------------------------------
// dim_Product
// ---------------------------------------------------------------------
let
    Source = fnReadSheet("crm_customers.xlsx", "Products"),
    Typed  = Table.TransformColumnTypes(Source, {
               {"ProductCode", type text}, {"ProductName", type text}, {"PlanTier", type text},
               {"ProductType", type text}, {"ListPricePerSeat_USD", type number}})
in
    Typed


// ---------------------------------------------------------------------
// fact_GL – monthly GL actuals (revenue, COGS, OpEx)
// ---------------------------------------------------------------------
let
    Source = fnReadSheet("finance_gl_budget.xlsx", "GL_Actuals"),
    Typed  = Table.TransformColumnTypes(Source, {
               {"PostingMonth", type date}, {"AccountCode", type text}, {"AccountName", type text},
               {"Category", type text}, {"Segment", type text}, {"Amount_USD", type number}}),
    Group  = Table.AddColumn(Typed, "PnLGroup", each
               if [Category] = "Revenue" then "Revenue"
               else if [Category] = "COGS" then "COGS" else "OpEx", type text)
in
    Group


// ---------------------------------------------------------------------
// fact_Budget
// ---------------------------------------------------------------------
let
    Source = fnReadSheet("finance_gl_budget.xlsx", "Revenue_Budget"),
    Typed  = Table.TransformColumnTypes(Source, {{"Month", type date}, {"Region", type text}, {"BudgetRevenue_USD", type number}})
in
    Typed


// ---------------------------------------------------------------------
// fact_Cash
// ---------------------------------------------------------------------
let
    Source = fnReadSheet("finance_gl_budget.xlsx", "Cash_Balance"),
    Typed  = Table.TransformColumnTypes(Source, {{"Month", type date}, {"ClosingCash_USD", type number}})
in
    Typed


// ---------------------------------------------------------------------
// dim_Date – daily calendar covering the data; mark as Date table on [Date]
// ---------------------------------------------------------------------
let
    StartDate = #date(2025, 12, 1),
    EndDate   = #date(2026, 12, 31),
    Dates     = List.Dates(StartDate, Duration.Days(EndDate - StartDate) + 1, #duration(1, 0, 0, 0)),
    ToTable   = Table.FromList(Dates, Splitter.SplitByNothing(), {"Date"}),
    Typed     = Table.TransformColumnTypes(ToTable, {{"Date", type date}}),
    AddCols   = Table.AddColumn(Table.AddColumn(Table.AddColumn(Table.AddColumn(Table.AddColumn(Typed,
                  "MonthStart",  each Date.StartOfMonth([Date]), type date),
                  "Year",        each Date.Year([Date]), Int64.Type),
                  "Quarter",     each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),
                  "MonthName",   each Date.ToText([Date], "MMM"), type text),
                  "MonthNumber", each Date.Month([Date]), Int64.Type),
    FiscalYr  = Table.AddColumn(AddCols, "FiscalYear", each "FY" & Text.From(Date.Year([Date])), type text)
in
    FiscalYr
