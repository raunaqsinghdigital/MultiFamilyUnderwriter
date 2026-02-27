const modelUrl = "/api/model";
const calculateUrl = "/api/calculate";
const adminMortgageUrl = "/api/admin/mortgage";
const adminMortgageOverrideUrl = "/api/admin/mortgage-overrides";
const PDF_REPORT_SHEET = "__pdf_report";
const isAdminMode = new URLSearchParams(window.location.search).get("admin") === "1";
const ROW_FORM_SHEETS = new Set();
const RENT_ROLL_REQUIRED_ROW_FIELDS = [
  "regular_rent",
];
const RENT_ROLL_CONFIG_REQUIRED_IDS = ["#rr-unit-count"];
const PERCENT_WHOLE_INPUT_KEYS = new Set([
  "Valuation!D14", // Vacancy
  "Valuation!D24", // Management
  "Valuation!D25", // Other/Advertising
  "Valuation!E29", // Market Cap Rate
  "Valuation!E32", // Market Interest Rate
  "Valuation!E34", // Max LTV
  "Returns!D13", // CMHC Premium
  "Returns!C26", // Repairs & Maintenance (% of income)
  "Returns!C27", // Property Manager (% of income)
  "Returns!C28", // Asset Manager (% of income)
  "Returns!C55", // GST Rebate
  "Returns!C56", // Acquisition Fee
  "Returns!C57", // Finance Fee
  "Returns!C79", // Property Appreciation (%)
  "Returns!D61", // Percent Ownership
  "Returns!H61", // Estimated Appreciation / Inflation
]);
const VALUATION_RENTROLL_AUTOFILL_KEYS = new Map([
  ["Valuation!E11", { rentRollField: "parking", label: "Total Parking" }],
  ["Valuation!E12", { rentRollField: "pet_fee", label: "Total Pet Fee" }],
]);
const VALUATION_INPUT_HINTS = new Map([
  ["Valuation!D6", "Enter purchase price as dollar amount."],
  ["Valuation!E10", "Enter annual laundry income in $/year."],
  ["Valuation!E11", "Enter annual parking income in $/year (or auto-fill from Rent Roll)."],
  ["Valuation!E12", "Enter annual other income in $/year (or auto-fill from Rent Roll pet fee total)."],
  ["Valuation!D14", "Enter vacancy as percent (example: 3 for 3%)."],
  ["Valuation!E18", "Enter annual property taxes in $/year."],
  ["Valuation!E19", "Enter annual insurance in $/year."],
  ["Valuation!F20", "Enter utilities in $/unit/year."],
  ["Valuation!F21", "Enter repairs and maintenance in $/unit/year."],
  ["Valuation!F22", "Enter appliance reserve in $/unit/year."],
  ["Valuation!F23", "Enter wages/on-site manager in $/unit/year."],
  ["Valuation!D24", "Enter management fee as percent of EGI (example: 4 for 4%)."],
  ["Valuation!D25", "Enter other/advertising fee as percent of EGI (example: 2 for 2%)."],
  ["Valuation!E29", "Enter market cap rate as percent (example: 7 for 7%)."],
  ["Valuation!E32", "Enter lender interest rate as percent (example: 5.25 for 5.25%)."],
  ["Valuation!E33", "Enter amortization period in years."],
  ["Valuation!E34", "Enter max lender LTV as percent (example: 75 for 75%)."],
  ["Valuation!E35", "Enter required DSCR ratio (example: 1.25)."],
]);
const VALUATION_RESULT_FIELDS = [
  { key: "Valuation!E15", label: "Effective Gross Income (EGI)", format: "money" },
  {
    key: "Valuation!E26",
    label: "Operating Expenses (OE)",
    format: "money",
    secondaryKey: "Valuation!H26",
    secondaryLabel: "% EGI",
    secondaryFormat: "percent",
  },
  { key: "Valuation!E28", label: "Net Operating Income (NOI)", format: "money" },
  {
    key: "Valuation!E30",
    label: "Value Based on Area Cap Rate",
    format: "money",
    secondaryKey: "Valuation!E29",
    secondaryLabel: "Area Cap Rate",
    secondaryFormat: "percent",
  },
  { key: "Valuation!E47", label: "Actual Property Cap Rate (NOI/Price)", format: "percent" },
  { key: "Valuation!E43", label: "Lesser Of The Three Loans", format: "money" },
  { key: "Valuation!E44", label: "Actual Loan To Cost (LTC)", format: "percent" },
  { key: "Valuation!E46", label: "Actual DSCR", format: "number" },
];
const VALUATION_TABLE_COLUMNS = [
  { id: "pct", label: "%", format: "percent" },
  { id: "year", label: "$/year", format: "money" },
  { id: "unit_year", label: "$/unit/yr", format: "money" },
  { id: "month", label: "$/month", format: "money" },
  { id: "pct_egi", label: "% EGI", format: "percent" },
];
const VALUATION_SECTIONS = [
  {
    title: "Rental Revenue",
    rows: [
      { label: "Annual Rent (actual, not projected)", cells: { year: "Valuation!E9", unit_year: "Valuation!F9", month: "Valuation!G9" } },
      {
        label: "Laundry",
        description: "12-15/unit/month (144-180 PUPA).",
        cells: { year: "Valuation!E10", unit_year: "Valuation!F10", month: "Valuation!G10" },
      },
      {
        label: "Parking",
        description: "If not included on Rent Roll tab.",
        cells: { year: "Valuation!E11", unit_year: "Valuation!F11", month: "Valuation!G11" },
      },
      {
        label: "Other",
        description: "If not included on Rent Roll tab.",
        cells: { year: "Valuation!E12", unit_year: "Valuation!F12", month: "Valuation!G12" },
      },
      { label: "Potential Gross Income (PGI)", emphasis: true, cells: { year: "Valuation!E13", unit_year: "Valuation!F13", month: "Valuation!G13" } },
      {
        label: "Vacancy",
        description: "Use CMHC values for the specific neighborhood.",
        cells: { pct: "Valuation!D14", year: "Valuation!E14", unit_year: "Valuation!F14", month: "Valuation!G14" },
      },
      { label: "Effective Gross Income (EGI)", emphasis: true, cells: { year: "Valuation!E15", unit_year: "Valuation!F15", month: "Valuation!G15" } },
    ],
  },
  {
    title: "Operating Expenses",
    rows: [
      {
        label: "Property Taxes",
        description: "0.7 - 1% of sale price.",
        descriptionSupplementals: [{ label: "Current", key: "Valuation!L18", format: "percent" }],
        cells: { year: "Valuation!E18", unit_year: "Valuation!F18", month: "Valuation!G18", pct_egi: "Valuation!H18" },
      },
      {
        label: "Insurance",
        description: "Best to get quote. PUPA ~600-800 /unit/yr.",
        cells: { year: "Valuation!E19", unit_year: "Valuation!F19", month: "Valuation!G19", pct_egi: "Valuation!H19" },
      },
      {
        label: "Utilities",
        description: "All utilities: ~3000 PUPA; water/gas only: ~1000-1350 by unit mix.",
        cells: { year: "Valuation!E20", unit_year: "Valuation!F20", month: "Valuation!G20", pct_egi: "Valuation!H20" },
      },
      {
        label: "Repairs and Maintenance",
        description: "PUPA ~750-830 newer, 850-900 older.",
        cells: { year: "Valuation!E21", unit_year: "Valuation!F21", month: "Valuation!G21", pct_egi: "Valuation!H21" },
      },
      {
        label: "Appliances",
        description: "PUPA ~60 per appliance.",
        cells: { year: "Valuation!E22", unit_year: "Valuation!F22", month: "Valuation!G22", pct_egi: "Valuation!H22" },
      },
      {
        label: "Wages /on-site manager",
        description: "40-45/unit/month.",
        descriptionSupplementals: [{ label: "Current", key: "Valuation!L23", format: "money", suffix: "/unit/month" }],
        cells: { year: "Valuation!E23", unit_year: "Valuation!F23", month: "Valuation!G23", pct_egi: "Valuation!H23" },
      },
      {
        label: "Management",
        description: "4.5-6% for <24 units, or 8-10% including onsite.",
        cells: { pct: "Valuation!D24", year: "Valuation!E24", unit_year: "Valuation!F24", month: "Valuation!G24", pct_egi: "Valuation!H24" },
      },
      {
        label: "Other/Advertising",
        description: "~1-2%.",
        cells: { pct: "Valuation!D25", year: "Valuation!E25", unit_year: "Valuation!F25", month: "Valuation!G25", pct_egi: "Valuation!H25" },
      },
      {
        label: "Total Operating Expense (OE)",
        emphasis: true,
        description: "Newer assets should be ~35-45%; older assets ~45-55% of EGI.",
        cells: { year: "Valuation!E26", unit_year: "Valuation!F26", month: "Valuation!G26", pct_egi: "Valuation!H26" },
      },
      {
        label: "Net Operating Income (NOI)",
        emphasis: true,
        cells: { year: "Valuation!E28", unit_year: "Valuation!F28", month: "Valuation!G28" },
      },
    ],
  },
  {
    title: "Financing & Value",
    rows: [
      { label: "Value Based on Cap Rate", cells: { year: "Valuation!E30", unit_year: "Valuation!F30" } },
      {
        label: "Max Annual Debt Service",
        cells: { year: "Valuation!E36", unit_year: "Valuation!F36", month: "Valuation!G36" },
      },
      {
        label: "Debt Coverage Ratio (NOI/DS)",
        cells: { pct: "Valuation!E37" },
        formats: { pct: "number" },
      },
      { label: "Value Based on DSCR", cells: { year: "Valuation!E38" } },
      { label: "Max Loan Based on LTV", cells: { year: "Valuation!E39" } },
      { label: "Max Loan Based on Purchase Price (LTC)", cells: { year: "Valuation!E40" } },
      { label: "Max Loan Based on DSCR", cells: { year: "Valuation!E41" } },
      { label: "Lesser Of The Three Loans", cells: { year: "Valuation!E43" }, emphasis: true },
      { label: "Actual LTC", cells: { pct: "Valuation!E44" }, formats: { pct: "percent" } },
      { label: "Actual Debt Service (P&I)", cells: { year: "Valuation!E45" } },
      { label: "Actual Debt Coverage Ratio", cells: { pct: "Valuation!E46" }, formats: { pct: "number" } },
      { label: "Property Cap Rate", cells: { pct: "Valuation!E47" }, formats: { pct: "percent" } },
    ],
  },
];
const VALUATION_SECONDARY_ROWS = [
  {
    label: "Value Based on Cap Rate",
    key: "Valuation!E30",
    format: "money",
    supplementalKey: "Valuation!F30",
    supplementalFormat: "money",
    supplementalSuffix: "/door",
  },
  { label: "Max Loan Based on LTV", key: "Valuation!E39", format: "money" },
  { label: "Max Loan Based on Purchase Price (LTC)", key: "Valuation!E40", format: "money" },
  { label: "Max Loan Based on DSCR", key: "Valuation!E41", format: "money" },
];
const VALUATION_ASSUMPTION_NOTES = new Map([
  ["Valuation!E29", "Average of comparables."],
  ["Valuation!E32", "GOC + (1.25-1.5) for CMHC, +1% conventional."],
]);
const RETURNS_KEY_RESULTS = [
  { key: "Returns!D34", label: "Monthly Cash Flow", format: "money" },
  { key: "Returns!D33", label: "Annual Cash Flow", format: "money" },
  { key: "Returns!D58", label: "Total Investment", format: "money" },
  { key: "Returns!H58", label: "Investor Annual ROI", format: "percent" },
  { key: "Returns!D14", label: "Mortgage Amount (incl. CMHC)", format: "money" },
  { key: "Returns!D37", label: "Down Payment", format: "money" },
];
const RETURNS_TABLE_COLUMNS = [
  { id: "pct", label: "%", format: "percent" },
  { id: "value", label: "Value", format: "money" },
];
const RETURNS_SECTIONS = [
  {
    title: "Purchase Summary",
    rows: [
      {
        label: "Property",
        cells: { value: "Returns!B1" },
        formats: { value: "text" },
      },
      { label: "Appraised Value", cells: { value: "Returns!D5" }, formats: { value: "money" } },
      { label: "Purchase Price", cells: { value: "Returns!D6" }, formats: { value: "money" } },
    ],
  },
  {
    title: "Commercial Financing",
    rows: [
      { label: "Interest Rate", cells: { value: "Returns!D9" }, formats: { value: "percent" } },
      { label: "Amortization Years", cells: { value: "Returns!D10" }, formats: { value: "number" } },
      { label: "Loan To Value", cells: { value: "Returns!D11" }, formats: { value: "percent" } },
      { label: "Mortgage Amount", cells: { value: "Returns!D12" }, formats: { value: "money" } },
      { label: "CMHC Premium", cells: { value: "Returns!D13" }, formats: { value: "percent" } },
      {
        label: "Mortgage Amount incl. CMHC Premium",
        cells: { value: "Returns!D14" },
        formats: { value: "money" },
        emphasis: true,
      },
    ],
  },
  {
    title: "Monthly Cash Flow",
    rows: [
      { label: "Income", cells: { value: "Returns!D18" }, formats: { value: "money" } },
      { label: "Expenses", cells: {}, emphasis: true },
      { label: "Vacancy", cells: { pct: "Returns!C21", value: "Returns!D21" }, formats: { pct: "percent", value: "money" } },
      { label: "Mortgage", cells: { value: "Returns!D22" }, formats: { value: "money" } },
      { label: "Property Tax", cells: { value: "Returns!D23" }, formats: { value: "money" } },
      { label: "Insurance", cells: { value: "Returns!D24" }, formats: { value: "money" } },
      { label: "Utilities", cells: { value: "Returns!D25" }, formats: { value: "money" } },
      { label: "Repairs & Maintenance", cells: { pct: "Returns!C26", value: "Returns!D26" }, formats: { pct: "percent", value: "money" } },
      { label: "Property Manager", cells: { pct: "Returns!C27", value: "Returns!D27" }, formats: { pct: "percent", value: "money" } },
      { label: "Asset Manager", cells: { pct: "Returns!C28", value: "Returns!D28" }, formats: { pct: "percent", value: "money" } },
      { label: "Bookkeeping/Accounting/Legal", cells: { value: "Returns!D29" }, formats: { value: "money" } },
      { label: "Pest Control/Snow Removal", cells: { value: "Returns!D30" }, formats: { value: "text" } },
      { label: "Advertising/Bank Fees/Tenant Gifts", cells: { value: "Returns!D31" }, formats: { value: "money" } },
      { label: "Total Expenses", cells: { value: "Returns!D32" }, formats: { value: "money" }, emphasis: true },
      { label: "Annual Cash Flow", cells: { value: "Returns!D33" }, formats: { value: "money" }, emphasis: true },
      { label: "Monthly Cash Flow", cells: { value: "Returns!D34" }, formats: { value: "money" }, emphasis: true },
    ],
  },
  {
    title: "Total Investment",
    rows: [
      { label: "Down Payment", cells: { value: "Returns!D37" }, formats: { value: "money" } },
      { label: "Appraisal", cells: { value: "Returns!D39" }, formats: { value: "money" } },
      { label: "Legal Fees (ours)", cells: { value: "Returns!D41" }, formats: { value: "money" } },
      { label: "Legal Fees (lender)", cells: { value: "Returns!D42" }, formats: { value: "money" } },
      { label: "Mortgage Broker Fees", cells: { value: "Returns!D43" }, formats: { value: "money" } },
      { label: "CMHC Fees ($150/door)", cells: { value: "Returns!D44" }, formats: { value: "money" } },
      { label: "Incorporation Fees / USA", cells: { value: "Returns!D46" }, formats: { value: "money" } },
      { label: "Title Insurance", cells: { value: "Returns!D47" }, formats: { value: "money" } },
      { label: "Insurance Review", cells: { value: "Returns!D48" }, formats: { value: "money" } },
      { label: "Property Tax Adjustment", cells: { value: "Returns!D49" }, formats: { value: "money" } },
      { label: "Late Closing Interest", cells: { value: "Returns!D50" }, formats: { value: "money" } },
      { label: "Lender Fee", cells: { value: "Returns!D51" }, formats: { value: "money" } },
      { label: "Property Inspection", cells: { value: "Returns!D52" }, formats: { value: "money" } },
      { label: "Reserve Fund / Capex", cells: { value: "Returns!D53" }, formats: { value: "money" } },
      { label: "GST Rebate", cells: { pct: "Returns!C55", value: "Returns!D55" }, formats: { pct: "percent", value: "money" } },
      { label: "Acquisition Fee (yours)", cells: { pct: "Returns!C56", value: "Returns!D56" }, formats: { pct: "percent", value: "money" } },
      { label: "Finance Fee (yours)", cells: { pct: "Returns!C57", value: "Returns!D57" }, formats: { pct: "percent", value: "money" } },
      { label: "Total Investment", cells: { value: "Returns!D58" }, formats: { value: "money" }, emphasis: true },
    ],
  },
  {
    title: "Annual Return",
    rows: [
      { label: "From Cash Flow (Cash-on-Cash)", cells: { value: "Returns!H49" }, formats: { value: "percent" } },
      { label: "From Mortgage Principal Reduction", cells: { value: "Returns!H50" }, formats: { value: "percent" } },
      { label: "From Appreciation", cells: { value: "Returns!H51" }, formats: { value: "percent" } },
      { label: "Investor Annual ROI", cells: { value: "Returns!H58" }, formats: { value: "percent" }, emphasis: true },
    ],
  },
  {
    title: "Ownership & Value Outlook",
    rows: [
      { label: "Percent Ownership", cells: { value: "Returns!D61" }, formats: { value: "percent" } },
      { label: "Estimated Appreciation / Inflation", cells: { value: "Returns!H61" }, formats: { value: "percent" } },
      { label: "Value in 5 Years", cells: { value: "Returns!H62" }, formats: { value: "money" } },
      { label: "Value in 10 Years", cells: { value: "Returns!H63" }, formats: { value: "money" } },
    ],
  },
  {
    title: "Projection Breakdown",
    rows: [
      { label: "Capital", cells: { value: "Returns!D72" }, formats: { value: "money" } },
      { label: "Principal Reduction", cells: { value: "Returns!H76" }, formats: { value: "money" } },
      { label: "Principal Reduction ROI", cells: { value: "Returns!H77" }, formats: { value: "percent" } },
      { label: "Property Appreciation", cells: { value: "Returns!H80" }, formats: { value: "money" } },
      { label: "Appreciation ROI", cells: { value: "Returns!H81" }, formats: { value: "percent" } },
      { label: "Cash Flow", cells: { value: "Returns!H84" }, formats: { value: "money" } },
      { label: "Cash Flow ROI", cells: { value: "Returns!H85" }, formats: { value: "percent" } },
      { label: "Profit (PR + PA + CF)", cells: { value: "Returns!H89" }, formats: { value: "money" } },
      { label: "ROI (PR + PA + CF)", cells: { value: "Returns!H90" }, formats: { value: "percent" }, emphasis: true },
    ],
  },
];
const REI_RATIO_DEFS = [
  { id: "cap_rate", label: "Cap Rate", format: "percent" },
  { id: "monthly_cashflow_building", label: "Monthly Cashflow Building", format: "money" },
  { id: "monthly_per_door_cashflow", label: "Monthly Per Door Cashflow", format: "money" },
  { id: "cash_on_cash_return", label: "Cash on Cash Return", format: "percent" },
  { id: "debt_service_coverage_ratio", label: "Debt Service Coverage Ratio (DSCR)", format: "number" },
  { id: "gross_rent_multiplier", label: "Gross Rent Multiplier (GRM)", format: "number" },
  { id: "operating_expense_ratio", label: "Operating Expense Ratio (OER)", format: "percent" },
  { id: "one_percent_rule", label: "1% Rule", format: "percent", threshold: 0.01 },
  {
    id: "purchase_price_break_even_range",
    label: "Purchase Price Cash Flow To Break Even Range",
    format: "range_money",
  },
];
const SENSITIVITY_RESULT_FIELDS = [
  { id: "value_cap", label: "Value Based on Cap Rate", format: "money" },
  { id: "ltc", label: "Loan-to-Cost (LTC)", format: "percent" },
  { id: "monthly_cashflow", label: "Cash Flow (Monthly)", format: "money" },
  { id: "down_payment", label: "Downpayment", format: "money" },
  { id: "cash_to_close", label: "Cash to Close", format: "money" },
  { id: "actual_dscr", label: "Actual DSCR", format: "number" },
  { id: "total_property_roi", label: "Total Property ROI", format: "percent" },
  { id: "investor_roi", label: "Investor ROI", format: "percent" },
];

let workbookModel = null;
const inputElements = new Map();
const formulaElements = new Map();
const adminFormulaInputs = new Map();
const requiredInputKeys = new Set();

const statusEl = document.getElementById("status");
const tabsEl = document.getElementById("sheet-tabs");
const panelsEl = document.getElementById("sheet-panels");
const globalMetricsEl = document.getElementById("global-metrics");
const calculateBtn = document.getElementById("calculate-btn");
const resetBtn = document.getElementById("reset-btn");

const defaultInputValues = new Map();
let globalLastRunValueEl = null;
let adminSummaryValueEl = null;
let lastValidationMissingCount = null;
let hasSuccessfulCalculation = false;
let latestFormulaValues = {};
const derivedMetricElements = new Map();
const inputDisplayElements = new Map();
let pdfReportSummaryEl = null;

let rentRollState = null;
let rentRollDefaultState = null;
let sensitivityState = {
  rentChangePct: 0,
  interestRateChangeBps: 0,
  vacancyPct: null,              // null = baseline; decimal e.g. 0.05 = 5%
  purchasePriceOverride: null,   // null = baseline; absolute dollars
  gstRebate: "baseline",         // "baseline" | "yes" | "no"
  scenarios: [],                 // saved scenario objects, max 10
};
let sensitivityView = null;

function isPresentValue(value) {
  return String(value ?? "").trim() !== "";
}

function setInputMissingState(inputEl, missing) {
  if (!inputEl) return;
  inputEl.classList.toggle("is-missing", missing);
  const parent = inputEl.closest(
    ".field-row, .rowform-field, .valuation-assumption-row, .valuation-cell, .returns-cell, .rentroll-config-field, td"
  );
  if (parent) parent.classList.toggle("required-missing", missing);
}

function wireInputElement(input, key, required = false) {
  input.dataset.key = key;
  inputElements.set(key, input);
  if (required) {
    input.classList.add("is-required");
    input.dataset.required = "1";
  }
  input.addEventListener("input", () => {
    updateInputDisplayElements(key, input.value);
    if (VALUATION_RENTROLL_AUTOFILL_KEYS.has(key)) {
      syncValuationRentRollAutofillManualOverrideFlag(key, input);
    }
    if (key === "Valuation!E11" || key === "Valuation!E12" || key === "Valuation!D5") {
      updateValuationParkingOtherDerivedDisplays();
    }
    updateValidationState();
  });
  input.addEventListener("blur", () => updateValidationState());
}

function registerInputDisplayElement(key, element, format = "text") {
  if (!inputDisplayElements.has(key)) inputDisplayElements.set(key, []);
  element.dataset.format = format;
  inputDisplayElements.get(key).push(element);
}

function formatInputValueForDisplay(key, rawValue, displayType = "text") {
  if (!isPresentValue(rawValue)) return "";
  let valueForDisplay = rawValue;
  if (displayType === "percent" && PERCENT_WHOLE_INPUT_KEYS.has(key)) {
    valueForDisplay = normalizeWorkbookInputNumber(key, rawValue, Number.NaN);
  } else if (displayType === "money" || displayType === "number" || displayType === "compact") {
    valueForDisplay = toNumber(rawValue, Number.NaN);
  }
  return formatDisplayByType(valueForDisplay, displayType);
}

function updateInputDisplayElements(key, rawValue = undefined) {
  const elements = inputDisplayElements.get(key) || [];
  if (!elements.length) return;
  const sourceRawValue =
    rawValue !== undefined ? rawValue : inputElements.get(key)?.value ?? getInputDefaultValue(key, "");
  for (const el of elements) {
    const displayType = el.dataset.format || "text";
    el.textContent = formatInputValueForDisplay(key, sourceRawValue, displayType);
  }
}

function registerDerivedMetricElement(metricId, element, format = "text") {
  if (!derivedMetricElements.has(metricId)) derivedMetricElements.set(metricId, []);
  element.dataset.format = format;
  derivedMetricElements.get(metricId).push(element);
}

function initializeRequiredInputKeys(model) {
  requiredInputKeys.clear();
  const hiddenSheets = new Set(Array.isArray(model.hidden_sheets) ? model.hidden_sheets : []);
  for (const entry of model.input_cells || []) {
    if (!entry || !entry.key) continue;
    if (hiddenSheets.has(entry.sheet)) continue;
    if (entry.sheet === "Rent Roll") continue;
    if (entry.sheet === "Sensitivity Analysis") continue;
    if (entry.key === "Valuation!D4" || entry.key === "Valuation!D5") continue;
    requiredInputKeys.add(entry.key);
  }
}

function buildInitialFormulaSnapshot(model) {
  const snapshot = {};
  for (const sheet of model.sheets || []) {
    for (const cell of sheet.cells || []) {
      if (!cell || !cell.has_formula || !cell.key) continue;
      snapshot[cell.key] = cell.value;
    }
  }
  return snapshot;
}

function deepClone(value) {
  return JSON.parse(JSON.stringify(value));
}

function colToLetters(index) {
  let i = index;
  let out = "";
  while (i > 0) {
    const rem = (i - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    i = Math.floor((i - 1) / 26);
  }
  return out;
}

function cellAddress(col, row) {
  return `${colToLetters(col)}${row}`;
}

function setStatus(text) {
  statusEl.textContent = text;
}

function formatValue(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") {
    if (Number.isInteger(value)) return String(value);
    return value.toLocaleString(undefined, { maximumFractionDigits: 6 });
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function formatMoney(value) {
  const amount = Number(value) || 0;
  return amount.toLocaleString(undefined, {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  });
}

function formatCompactValue(value) {
  if (typeof value !== "number") return formatValue(value);
  const abs = Math.abs(value);
  if (abs >= 1_000_000_000) return `${(value / 1_000_000_000).toFixed(2)}B`;
  if (abs >= 1_000_000) return `${(value / 1_000_000).toFixed(2)}M`;
  if (abs >= 1_000) return `${(value / 1_000).toFixed(1)}K`;
  return formatValue(value);
}

function formatPercentValue(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return formatValue(value);
  return `${(numeric * 100).toFixed(2)}%`;
}

function formatNumberValue(value, decimals = 2) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return formatValue(value);
  return numeric.toLocaleString(undefined, {
    minimumFractionDigits: 0,
    maximumFractionDigits: decimals,
  });
}

function toNumber(value, fallback = 0) {
  if (typeof value === "number") return Number.isFinite(value) ? value : fallback;
  if (typeof value === "string") {
    const text = value.trim().replace(/,/g, "");
    if (!text) return fallback;
    const parsed = Number(text);
    return Number.isFinite(parsed) ? parsed : fallback;
  }
  return fallback;
}

function normalizeWorkbookInputNumber(key, value, fallback = Number.NaN) {
  const numeric = toNumber(value, Number.NaN);
  if (!Number.isFinite(numeric)) return fallback;
  if (PERCENT_WHOLE_INPUT_KEYS.has(key)) return numeric / 100;
  return numeric;
}

function normalizeInputForSubmit(key, rawValue) {
  const text = String(rawValue ?? "").trim();
  if (!text) return "";
  if (PERCENT_WHOLE_INPUT_KEYS.has(key)) {
    const numeric = toNumber(text, Number.NaN);
    if (Number.isFinite(numeric)) return numeric / 100;
  }
  return rawValue;
}

function toPositiveInt(value, fallback = 1) {
  const numeric = Math.round(toNumber(value, fallback));
  if (!Number.isFinite(numeric) || numeric < 1) return fallback;
  return numeric;
}

function toNonNegativeInt(value, fallback = 0) {
  const numeric = Math.round(toNumber(value, fallback));
  if (!Number.isFinite(numeric) || numeric < 0) return fallback;
  return numeric;
}

function createMetricCard(label, value, valueClass = "") {
  const card = document.createElement("article");
  card.className = "metric-card";

  const labelEl = document.createElement("p");
  labelEl.className = "metric-label";
  labelEl.textContent = label;
  card.appendChild(labelEl);

  const valueEl = document.createElement("p");
  valueEl.className = `metric-value ${valueClass}`.trim();
  valueEl.textContent = value;
  card.appendChild(valueEl);

  return { card, valueEl };
}

function registerFormulaElement(key, element) {
  if (!formulaElements.has(key)) formulaElements.set(key, []);
  formulaElements.get(key).push(element);
}

function getInputDefaultValue(key, fallback = "") {
  if (defaultInputValues.has(key)) return defaultInputValues.get(key);
  return fallback;
}

function parsePropertySplit(value) {
  const text = String(value || "").trim();
  if (!text) return { property_name: "", property_address: "" };
  if (text.includes("|")) {
    const [name, ...rest] = text.split("|");
    return {
      property_name: name.trim(),
      property_address: rest.join("|").trim(),
    };
  }
  if (text.includes(" - ")) {
    const [name, ...rest] = text.split(" - ");
    return {
      property_name: name.trim(),
      property_address: rest.join(" - ").trim(),
    };
  }
  return {
    property_name: "",
    property_address: text,
  };
}

function createEmptyRentUnit(index) {
  return {
    row_index: index + 1,
    tenant_name: "",
    unit: "",
    regular_rent: "",
    utilities: "",
    parking: "",
    pet_fee: "",
    projected_rent: "",
    total_rent: 0,
  };
}

function normalizeRentRollState(state) {
  const normalized = {
    property_name: String(state.property_name || "").trim(),
    property_address: String(state.property_address || "").trim(),
    unit_count: toNonNegativeInt(state.unit_count, 0),
    units: [],
    totals: {
      utilities: 0,
      parking: 0,
      pet_fee: 0,
      total_rent: 0,
      projected_rent: 0,
    },
  };

  const sourceUnits = Array.isArray(state.units) ? state.units : [];
  for (let i = 0; i < normalized.unit_count; i += 1) {
    const source = sourceUnits[i] || createEmptyRentUnit(i);
    const row = {
      row_index: i + 1,
      tenant_name: String(source.tenant_name || "").trim(),
      unit: String(source.unit || "").trim(),
      regular_rent: String(source.regular_rent ?? "").trim(),
      utilities: String(source.utilities ?? "").trim(),
      parking: String(source.parking ?? "").trim(),
      pet_fee: String(source.pet_fee ?? "").trim(),
      projected_rent: String(source.projected_rent ?? "").trim(),
      total_rent: 0,
    };
    const regular = toNumber(row.regular_rent, 0);
    const utilities = toNumber(row.utilities, 0);
    const parking = toNumber(row.parking, 0);
    const petFee = toNumber(row.pet_fee, 0);
    const projected = toNumber(row.projected_rent, 0);
    row.total_rent = regular + utilities + parking + petFee;

    normalized.totals.utilities += utilities;
    normalized.totals.parking += parking;
    normalized.totals.pet_fee += petFee;
    normalized.totals.total_rent += row.total_rent;
    normalized.totals.projected_rent += projected;

    normalized.units.push(row);
  }

  return normalized;
}

function buildInitialRentRollState() {
  const propertyCombined = getInputDefaultValue("Valuation!D4", "");
  const property = parsePropertySplit(propertyCombined);
  const rawUnitCount = String(getInputDefaultValue("Valuation!D5", "") || "").trim();
  const unitCount = rawUnitCount ? toPositiveInt(rawUnitCount, 1) : 0;

  const readRentRollNumericInput = (key) => {
    const raw = getInputDefaultValue(key, "");
    if (!isPresentValue(raw)) return "";
    return toNumber(raw, 0);
  };

  const rows = [];
  for (let idx = 0; idx < unitCount; idx += 1) {
    const excelRow = 6 + idx;
    const row = createEmptyRentUnit(idx);
    if (excelRow <= 13) {
      row.tenant_name = String(getInputDefaultValue(`Rent Roll!A${excelRow}`, row.tenant_name) || "").trim();
      row.unit = String(getInputDefaultValue(`Rent Roll!B${excelRow}`, row.unit) || "").trim();
      row.regular_rent = readRentRollNumericInput(`Rent Roll!C${excelRow}`);
      row.utilities = readRentRollNumericInput(`Rent Roll!D${excelRow}`);
      row.parking = readRentRollNumericInput(`Rent Roll!E${excelRow}`);
      row.pet_fee = readRentRollNumericInput(`Rent Roll!F${excelRow}`);
      row.projected_rent = readRentRollNumericInput(`Rent Roll!H${excelRow}`);
    }
    rows.push(row);
  }

  return normalizeRentRollState({
    property_name: property.property_name,
    property_address: property.property_address,
    unit_count: unitCount,
    units: rows,
  });
}

function rentRollPayload() {
  if (!rentRollState) return null;
  return {
    property_name: rentRollState.property_name,
    property_address: rentRollState.property_address,
    unit_count: rentRollState.unit_count,
    units: rentRollState.units.map((row) => ({
      tenant_name: row.tenant_name,
      unit: row.unit,
      regular_rent: row.regular_rent,
      utilities: row.utilities,
      parking: row.parking,
      pet_fee: row.pet_fee,
      projected_rent: row.projected_rent,
    })),
  };
}

function hasRentRollFieldProvided(fieldName) {
  if (!rentRollState || !Array.isArray(rentRollState.units)) return false;
  return rentRollState.units.some((row) => isPresentValue(row?.[fieldName]));
}

function setFormulaDisplayLocal(key, value, format = "text") {
  latestFormulaValues[key] = value;
  const elements = formulaElements.get(key) || [];
  for (const el of elements) {
    let formatted;
    if (format === "percent" || el.dataset.format === "percent") {
      formatted = formatPercentValue(value);
    } else if (format === "money" || el.dataset.format === "money") {
      formatted = formatMoney(value);
    } else if (format === "number" || el.dataset.format === "number") {
      formatted = formatNumberValue(value);
    } else if (format === "compact" || el.dataset.format === "compact") {
      formatted = formatCompactValue(value);
    } else {
      formatted = formatValue(value);
    }
    el.textContent = formatted;
    el.classList.toggle("error", typeof formatted === "string" && formatted.startsWith("#"));
  }
}

function updateValuationParkingOtherDerivedDisplays() {
  const parkingAnnual = normalizeWorkbookInputNumber(
    "Valuation!E11",
    inputElements.get("Valuation!E11")?.value,
    0
  );
  const otherAnnual = normalizeWorkbookInputNumber(
    "Valuation!E12",
    inputElements.get("Valuation!E12")?.value,
    0
  );

  const units =
    toPositiveInt(rentRollState?.unit_count, 0) ||
    toPositiveInt(getLiveInputNumber("Valuation!D5", 0), 0) ||
    0;

  const parkingPerUnitYear = units > 0 ? parkingAnnual / units : 0;
  const parkingPerMonth = parkingAnnual / 12;
  const otherPerUnitYear = units > 0 ? otherAnnual / units : 0;
  const otherPerMonth = otherAnnual / 12;

  setFormulaDisplayLocal("Valuation!F11", parkingPerUnitYear, "money");
  setFormulaDisplayLocal("Valuation!G11", parkingPerMonth, "money");
  setFormulaDisplayLocal("Valuation!F12", otherPerUnitYear, "money");
  setFormulaDisplayLocal("Valuation!G12", otherPerMonth, "money");
}

function getValuationRentRollAutofillAnnualValue(key) {
  const cfg = VALUATION_RENTROLL_AUTOFILL_KEYS.get(key);
  if (!cfg) return null;
  if (!hasRentRollFieldProvided(cfg.rentRollField)) return null;
  const annualValue = toNumber(rentRollState?.totals?.[cfg.rentRollField], 0) * 12;
  return Number(annualValue.toFixed(2));
}

function syncValuationRentRollAutofillManualOverrideFlag(key, input) {
  if (!input) return;
  const autofillAnnualValue = getValuationRentRollAutofillAnnualValue(key);
  if (autofillAnnualValue === null) {
    input.dataset.manualRentRollOverride = "0";
    return;
  }

  const currentAnnualValue = normalizeWorkbookInputNumber(key, input.value, Number.NaN);
  const isManualOverride =
    Number.isFinite(currentAnnualValue) && Math.abs(currentAnnualValue - autofillAnnualValue) > 0.005;
  input.dataset.manualRentRollOverride = isManualOverride ? "1" : "0";
}

function syncValuationRentRollAutofillInputs() {
  for (const [key, cfg] of VALUATION_RENTROLL_AUTOFILL_KEYS.entries()) {
    const input = inputElements.get(key);
    if (!input) continue;

    const hasSourceValue = hasRentRollFieldProvided(cfg.rentRollField);
    const annualValue = getValuationRentRollAutofillAnnualValue(key);
    const parent = input.closest(".valuation-cell, .valuation-assumption-row, .field-row, td");

    if (hasSourceValue) {
      const isManualOverride = input.dataset.manualRentRollOverride === "1";
      if (!isManualOverride && annualValue !== null) {
        input.value = String(annualValue);
        input.dataset.manualRentRollOverride = "0";
      }
      input.disabled = false;
      input.dataset.lockedByRentRoll = "1";
      input.title = isManualOverride
        ? `Default available from Rent Roll ${cfg.label}. Manual override active.`
        : `Auto-filled from Rent Roll ${cfg.label}.`;
      parent?.classList.add("derived-from-rentroll");
      continue;
    }

    const wasLocked = input.dataset.lockedByRentRoll === "1";
    input.disabled = false;
    input.dataset.lockedByRentRoll = "0";
    input.dataset.manualRentRollOverride = "0";
    if (wasLocked) input.value = "";
    input.title = "";
    parent?.classList.remove("derived-from-rentroll");
  }

  updateValuationParkingOtherDerivedDisplays();
}

function isConditionallyRequiredInput(key) {
  if (key === "Valuation!E10") return false;
  if (key === "Valuation!E11") return !hasRentRollFieldProvided("parking");
  if (key === "Valuation!E12") return !hasRentRollFieldProvided("pet_fee");
  return true;
}

function renderGlobalMetrics(model) {
  globalMetricsEl.innerHTML = "";

  const cards = [
    createMetricCard("Mode", isAdminMode ? "Admin" : "Analyst"),
    createMetricCard("Last Run", "Not run"),
  ];

  adminSummaryValueEl = cards[0].valueEl;
  globalLastRunValueEl = cards[1].valueEl;
  for (const card of cards) globalMetricsEl.appendChild(card.card);
}

function formatMaybeMoney(value) {
  const numeric = toNumber(value, Number.NaN);
  if (!Number.isFinite(numeric)) return "--";
  return formatMoney(numeric);
}

function formatMaybePercent(value) {
  const numeric = toNumber(value, Number.NaN);
  if (!Number.isFinite(numeric)) return "--";
  return formatPercentValue(numeric);
}

function formatMaybeNumber(value) {
  const numeric = toNumber(value, Number.NaN);
  if (!Number.isFinite(numeric)) return "--";
  return formatNumberValue(numeric);
}

function getInputDecimalValue(key) {
  const raw = inputElements.get(key)?.value ?? getInputDefaultValue(key, "");
  return normalizeWorkbookInputNumber(key, raw, Number.NaN);
}

function buildPdfInvestorSummaryItems() {
  const egi = latestFormulaValues["Valuation!E15"];
  const oe = latestFormulaValues["Valuation!E26"];
  const oer = latestFormulaValues["Valuation!H26"];
  const noi = latestFormulaValues["Valuation!E28"];
  const valueByCap = latestFormulaValues["Valuation!E30"];
  const propertyCap = latestFormulaValues["Valuation!E47"];
  const dscr = latestFormulaValues["Valuation!E46"];
  const ltc = latestFormulaValues["Valuation!E44"];
  const annualCashflow = latestFormulaValues["Returns!D33"];
  const monthlyCashflow = latestFormulaValues["Returns!D34"];
  const investorRoi = latestFormulaValues["Returns!H58"];
  const areaCapRate = getInputDecimalValue("Valuation!E29");

  const notes = [
    `Effective Gross Income (EGI): ${formatMaybeMoney(egi)} per year.`,
    `Operating Expenses (OE): ${formatMaybeMoney(oe)} (${formatMaybePercent(oer)} of EGI).`,
    `Net Operating Income (NOI): ${formatMaybeMoney(noi)} per year.`,
    `Value based on area cap rate: ${formatMaybeMoney(valueByCap)} (area cap rate ${formatMaybePercent(areaCapRate)}).`,
    `Property cap rate: ${formatMaybePercent(propertyCap)}. Debt metrics: DSCR ${formatMaybeNumber(dscr)}, LTC ${formatMaybePercent(ltc)}.`,
    `Cash flow: ${formatMaybeMoney(monthlyCashflow)} monthly / ${formatMaybeMoney(annualCashflow)} annually. Investor annual ROI: ${formatMaybePercent(investorRoi)}.`,
  ];

  const dscrNum = toNumber(dscr, Number.NaN);
  if (Number.isFinite(dscrNum)) {
    notes.push(
      dscrNum < 1.2
        ? "Risk flag: DSCR is below 1.20, indicating tighter debt coverage."
        : "Strength: DSCR is at or above 1.20, indicating healthier debt coverage."
    );
  }

  const monthlyNum = toNumber(monthlyCashflow, Number.NaN);
  if (Number.isFinite(monthlyNum)) {
    notes.push(
      monthlyNum < 0
        ? "Risk flag: Monthly cash flow is negative in the current scenario."
        : "Strength: Monthly cash flow is positive in the current scenario."
    );
  }

  return notes;
}

function updatePdfReportSummary() {
  if (!pdfReportSummaryEl) return;
  if (!hasSuccessfulCalculation) {
    pdfReportSummaryEl.innerHTML =
      '<p class="panel-subtitle">Run Analysis to generate an investor summary and enable PDF export.</p>';
    return;
  }
  const items = buildPdfInvestorSummaryItems();
  pdfReportSummaryEl.innerHTML = `<ul class="pdf-summary-list">${items.map((t) => `<li>${t}</li>`).join("")}</ul>`;
}

function escapeHtml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getReportTabOrder() {
  return Array.from(tabsEl.querySelectorAll(".tab-btn"))
    .map((btn) => btn.dataset.sheet || "")
    .filter((sheet) => sheet && sheet !== PDF_REPORT_SHEET && !sheet.startsWith("__admin_"));
}

function buildPrintablePanelHtml(sheetName) {
  const panel = panelsEl.querySelector(`.sheet-panel[data-sheet="${sheetName}"]`);
  if (!panel) return "";
  const clone = panel.cloneNode(true);

  clone.querySelectorAll(".hidden, .panel-search, .btn").forEach((el) => el.remove());
  clone.querySelectorAll("details").forEach((el) => (el.open = true));

  clone.querySelectorAll("input, textarea, select").forEach((input) => {
    const span = document.createElement("span");
    span.className = "pdf-input-value";
    const raw = input.value ?? "";
    if (input.dataset.percentInput === "1") {
      const num = toNumber(raw, Number.NaN);
      span.textContent = Number.isFinite(num) ? `${formatNumberValue(num, 2)}%` : "--";
    } else {
      span.textContent = isPresentValue(raw) ? String(raw) : "--";
    }
    input.replaceWith(span);
  });

  return clone.innerHTML;
}

function openPdfReportWindow() {
  if (!hasSuccessfulCalculation) {
    setStatus("Run Analysis before generating the PDF report.");
    return;
  }

  const reportTabs = getReportTabOrder();
  const generatedAt = new Date().toLocaleString();
  const summaryItems = buildPdfInvestorSummaryItems();
  const summaryHtml = `<ul>${summaryItems.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;

  const pagesHtml = reportTabs
    .map((sheetName) => {
      const content = buildPrintablePanelHtml(sheetName);
      return `
        <section class="report-page">
          <header class="report-page-head">
            <h2>${escapeHtml(sheetName)}</h2>
          </header>
          <div class="report-page-body">${content}</div>
        </section>
      `;
    })
    .join("");

  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Underwriting Report</title>
  <style>
    body { font-family: "Avenir Next", "Segoe UI", sans-serif; margin: 0; color: #10253e; }
    .report-page { page-break-after: always; padding: 24px 28px; }
    .report-page:last-child { page-break-after: auto; }
    .report-summary h1 { margin: 0 0 6px; font-size: 24px; }
    .report-summary p { margin: 0 0 14px; color: #3c5472; font-size: 13px; }
    .report-summary ul { margin: 0; padding-left: 18px; display: grid; gap: 8px; }
    .report-page-head h2 { margin: 0 0 10px; font-size: 20px; border-bottom: 2px solid #d7e3f1; padding-bottom: 6px; }
    .report-page .sheet-panel { display: block !important; box-shadow: none; border: none; padding: 0; background: #fff; }
    .report-page .panel-header { margin-bottom: 8px; }
    .report-page .panel-subtitle { color: #4f6580; }
    .report-page table { width: 100%; border-collapse: collapse; }
    .report-page td, .report-page th { border: 1px solid #d9e4f0; padding: 6px 8px; vertical-align: top; }
    .report-page .field-output, .report-page .returns-output, .report-page .valuation-output { font-weight: 600; }
    .pdf-input-value { display: inline-block; min-width: 80px; padding: 2px 4px; border-bottom: 1px solid #cdd9e8; font-weight: 600; }
    .report-page .sheet-tabs, .report-page .toolbar, .report-page .global-metrics { display: none !important; }
    @media print { body { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }
  </style>
</head>
<body>
  <section class="report-page report-summary">
    <h1>Underwriting Summary</h1>
    <p>Generated: ${escapeHtml(generatedAt)}</p>
    ${summaryHtml}
  </section>
  ${pagesHtml}
  <script>window.onload = () => window.print();</script>
</body>
</html>`;

  const w = window.open("", "_blank");
  if (!w) {
    setStatus("Popup blocked. Please allow popups to generate the PDF report.");
    return;
  }
  w.document.open();
  w.document.write(html);
  w.document.close();
}

function renderPdfReportPanel(index) {
  const panel = document.createElement("section");
  panel.className = "sheet-panel pdf-report-panel";
  panel.dataset.sheet = PDF_REPORT_SHEET;
  if (index === 0) panel.classList.add("active");

  const header = document.createElement("div");
  header.className = "panel-header";
  const titleWrap = document.createElement("div");
  titleWrap.className = "panel-title";
  const h2 = document.createElement("h2");
  h2.textContent = "PDF Report";
  const subtitle = document.createElement("p");
  subtitle.className = "panel-subtitle";
  subtitle.textContent =
    "Generate an investor-facing underwriting PDF with one page per tab and an executive summary.";
  titleWrap.appendChild(h2);
  titleWrap.appendChild(subtitle);
  header.appendChild(titleWrap);
  panel.appendChild(header);

  const controls = document.createElement("div");
  controls.className = "pdf-report-controls";
  const btn = document.createElement("button");
  btn.className = "btn btn-primary";
  btn.textContent = "Generate PDF Report";
  btn.addEventListener("click", openPdfReportWindow);
  controls.appendChild(btn);
  panel.appendChild(controls);

  const summary = document.createElement("div");
  summary.className = "pdf-report-summary";
  pdfReportSummaryEl = summary;
  panel.appendChild(summary);
  updatePdfReportSummary();
  return panel;
}

function activateSheet(sheetName) {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.sheet === sheetName);
  });
  document.querySelectorAll(".sheet-panel").forEach((panel) => {
    panel.classList.toggle("active", panel.dataset.sheet === sheetName);
  });
}

function isMeaningfulText(text) {
  if (typeof text !== "string") return false;
  const normalized = text.trim();
  if (!normalized) return false;
  const lower = normalized.toLowerCase();
  if (lower.includes("inputs are shaded cells")) return false;
  if (lower.includes("do not touch this screen")) return false;
  return true;
}

function normalizeLabel(value) {
  if (typeof value !== "string") return "";
  return value.replace(/\s+/g, " ").trim();
}

function isLabelCandidate(cell) {
  return (
    cell &&
    !cell.has_formula &&
    !cell.is_input &&
    cell.value_type === "text" &&
    isMeaningfulText(cell.value)
  );
}

function findBestLabel(sheet, cell, cellMap) {
  for (let c = cell.col - 1; c >= sheet.min_col; c -= 1) {
    const left = cellMap.get(cellAddress(c, cell.row));
    if (isLabelCandidate(left)) return normalizeLabel(left.value);
  }

  for (let dr = 1; dr <= 3; dr += 1) {
    const row = cell.row - dr;
    if (row < sheet.min_row) break;
    const aboveSameCol = cellMap.get(cellAddress(cell.col, row));
    if (isLabelCandidate(aboveSameCol)) return normalizeLabel(aboveSameCol.value);
    for (let c = cell.col - 1; c >= Math.max(sheet.min_col, cell.col - 4); c -= 1) {
      const aboveLeft = cellMap.get(cellAddress(c, row));
      if (isLabelCandidate(aboveLeft)) return normalizeLabel(aboveLeft.value);
    }
  }

  return `${sheet.name} ${cellAddress(cell.col, cell.row)}`;
}

function detectHeadingRows(sheet, rowTextMap) {
  const headings = [];
  for (const [rowStr, cells] of rowTextMap.entries()) {
    const row = Number(rowStr);
    if (!cells.length) continue;
    const first = cells.slice().sort((a, b) => a.col - b.col)[0];
    if (!first) continue;
    const text = normalizeLabel(first.value);
    if (!text) continue;
    const uppercaseLike = /^[A-Z0-9 ()/%.-]+$/.test(text);
    const emphasized = first.fill_id !== 0 || text.endsWith(":") || uppercaseLike;
    if (!emphasized) continue;
    headings.push({ row, label: text.replace(/:$/, "") });
  }
  headings.sort((a, b) => a.row - b.row);
  return headings;
}

function findGroupName(cell, headingRows) {
  let fallback = "General";
  for (const heading of headingRows) {
    if (heading.row > cell.row) break;
    if (cell.row - heading.row <= 20) fallback = heading.label;
  }
  return fallback;
}

function buildSheetViewModel(sheet) {
  const cellMap = new Map();
  const rowTextMap = new Map();

  for (const cell of sheet.cells) {
    cellMap.set(cell.address, cell);
    if (isLabelCandidate(cell)) {
      const key = String(cell.row);
      if (!rowTextMap.has(key)) rowTextMap.set(key, []);
      rowTextMap.get(key).push(cell);
    }
  }

  const headingRows = detectHeadingRows(sheet, rowTextMap);
  const inputs = [];
  const outputs = [];

  for (const cell of sheet.cells) {
    if (!cell.is_input && !cell.has_formula) continue;
    if (sheet.name === "Valuation" && cell.key === "Valuation!D4") continue;
    if (sheet.name === "Valuation" && cell.key === "Valuation!D5") continue;

    const label = findBestLabel(sheet, cell, cellMap);
    const group = findGroupName(cell, headingRows);
    const field = {
      key: cell.key,
      address: cell.address,
      label,
      group,
      row: cell.row,
      col: cell.col,
      value: cell.value,
      value_type: cell.value_type,
      fill_id: cell.fill_id,
      is_input: cell.is_input,
      has_formula: cell.has_formula,
    };
    if (cell.is_input) inputs.push(field);
    if (cell.has_formula) outputs.push(field);
  }

  inputs.sort((a, b) => a.row - b.row || a.col - b.col);
  outputs.sort((a, b) => a.row - b.row || a.col - b.col);

  const metrics = outputs
    .filter((f) => typeof f.value === "number" && !f.label.startsWith(sheet.name))
    .slice(0, 4);

  return {
    name: sheet.name,
    inputs,
    outputs,
    metrics,
  };
}

function groupFields(fields) {
  const grouped = new Map();
  for (const field of fields) {
    if (!grouped.has(field.group)) grouped.set(field.group, []);
    grouped.get(field.group).push(field);
  }

  const output = [];
  for (const [name, items] of grouped.entries()) {
    items.sort((a, b) => a.row - b.row || a.col - b.col);
    if (items.length > 120) {
      let chunkIndex = 0;
      for (let i = 0; i < items.length; i += 120) {
        chunkIndex += 1;
        output.push({
          name: `${name} (${chunkIndex})`,
          items: items.slice(i, i + 120),
        });
      }
    } else {
      output.push({ name, items });
    }
  }
  output.sort((a, b) => {
    const aRow = a.items[0]?.row ?? 0;
    const bRow = b.items[0]?.row ?? 0;
    return aRow - bRow;
  });
  return output;
}

function createFieldRow(field) {
  const row = document.createElement("div");
  row.className = "field-row";
  row.dataset.search = `${field.label} ${field.address} ${field.key}`.toLowerCase();
  const isRequired = field.is_input && requiredInputKeys.has(field.key);
  if (isRequired) row.classList.add("required-field");

  const labelWrap = document.createElement("div");
  labelWrap.className = "field-label";
  const labelMain = document.createElement("div");
  labelMain.className = "field-label-main";
  if (isRequired) labelMain.classList.add("required-label");
  labelMain.textContent = field.label;
  const labelMeta = document.createElement("div");
  labelMeta.className = "field-label-meta";
  labelMeta.textContent = field.address;
  labelWrap.appendChild(labelMain);
  labelWrap.appendChild(labelMeta);
  row.appendChild(labelWrap);

  if (field.is_input) {
    const input = document.createElement("input");
    input.className = "field-input";
    input.type = field.value_type === "number" ? "number" : "text";
    input.step = "any";
    if (PERCENT_WHOLE_INPUT_KEYS.has(field.key)) {
      input.value = formatPercentInput(getInputDefaultValue(field.key, 0));
      input.placeholder = "Enter percent";
      input.dataset.percentInput = "1";
    } else {
      input.value = formatValue(getInputDefaultValue(field.key, ""));
    }
    row.appendChild(input);
    wireInputElement(input, field.key, isRequired);
  } else {
    const output = document.createElement("div");
    output.className = "field-output";
    const text = formatValue(field.value);
    output.textContent = text;
    if (typeof text === "string" && text.startsWith("#")) output.classList.add("error");
    registerFormulaElement(field.key, output);
    row.appendChild(output);
  }

  return row;
}

function createFieldGroup(groupName, fields, openByDefault = false) {
  const group = document.createElement("details");
  group.className = "field-group";
  if (openByDefault) group.open = true;

  const summary = document.createElement("summary");
  const labelEl = document.createElement("span");
  labelEl.textContent = groupName;
  const countEl = document.createElement("span");
  countEl.className = "group-count";
  countEl.textContent = `${fields.length} fields`;
  summary.appendChild(labelEl);
  summary.appendChild(countEl);
  group.appendChild(summary);

  const list = document.createElement("div");
  list.className = "field-list";
  for (const field of fields) list.appendChild(createFieldRow(field));
  group.appendChild(list);

  return group;
}

function renderValuationKeycards(sheetView, panel) {
  const outputMap = new Map(sheetView.outputs.map((field) => [field.key, field.value]));
  const inputMap = new Map(sheetView.inputs.map((field) => [field.key, field]));
  const resultsWrap = document.createElement("section");
  resultsWrap.className = "valuation-keycards";

  const head = document.createElement("div");
  head.className = "valuation-keycards-head";

  const title = document.createElement("h3");
  title.className = "valuation-keycards-title";
  title.textContent = "Key Results";
  head.appendChild(title);

  const subtitle = document.createElement("p");
  subtitle.className = "valuation-keycards-subtitle";
  subtitle.textContent = "Live snapshot of core underwriting metrics";
  head.appendChild(subtitle);

  resultsWrap.appendChild(head);

  const list = document.createElement("div");
  list.className = "valuation-keycards-grid";
  for (const [idx, item] of VALUATION_RESULT_FIELDS.entries()) {
    const card = document.createElement("article");
    card.className = "valuation-keycard";
    card.dataset.tone = String((idx % 4) + 1);

    const label = document.createElement("div");
    label.className = "valuation-keycard-label";
    label.textContent = item.label;
    card.appendChild(label);

    const value = document.createElement("div");
    value.className = "valuation-keycard-value";
    const mainFormat = item.format || "money";
    value.dataset.format = mainFormat;
    if (outputMap.has(item.key)) {
      value.textContent = formatCalculatedOrBlank(outputMap.get(item.key), mainFormat);
      registerFormulaElement(item.key, value);
    } else if (inputMap.has(item.key)) {
      value.textContent = formatInputValueForDisplay(item.key, getInputDefaultValue(item.key, ""), mainFormat);
      registerInputDisplayElement(item.key, value, mainFormat);
    } else {
      value.textContent = "";
    }
    card.appendChild(value);

    if (item.secondaryKey) {
      const secondary = document.createElement("div");
      secondary.className = "valuation-keycard-secondary";

      const secondaryLabel = document.createElement("span");
      secondaryLabel.className = "valuation-keycard-secondary-label";
      secondaryLabel.textContent = item.secondaryLabel || "";
      secondary.appendChild(secondaryLabel);

      const secondaryValue = document.createElement("span");
      secondaryValue.className = "valuation-keycard-secondary-value";
      if (item.secondaryFormat) secondaryValue.dataset.format = item.secondaryFormat;
      if (outputMap.has(item.secondaryKey)) {
        secondaryValue.textContent = formatCalculatedOrBlank(
          outputMap.get(item.secondaryKey),
          item.secondaryFormat || "text"
        );
        registerFormulaElement(item.secondaryKey, secondaryValue);
      } else if (inputMap.has(item.secondaryKey)) {
        secondaryValue.textContent = formatInputValueForDisplay(
          item.secondaryKey,
          getInputDefaultValue(item.secondaryKey, ""),
          item.secondaryFormat || "text"
        );
        registerInputDisplayElement(item.secondaryKey, secondaryValue, item.secondaryFormat || "text");
      } else {
        secondaryValue.textContent = "";
      }
      secondary.appendChild(secondaryValue);

      card.appendChild(secondary);
    }

    list.appendChild(card);
  }
  resultsWrap.appendChild(list);
  panel.appendChild(resultsWrap);
}

function renderValuationCellContent(cellKey, columnFormat, inputMap, outputMap, displayFormatOverride = null) {
  const wrap = document.createElement("div");
  wrap.className = "valuation-cell";

  if (!cellKey) {
    wrap.classList.add("empty");
    wrap.textContent = "";
    return wrap;
  }

  const inputField = inputMap.get(cellKey);
  if (inputField) {
    const isRequired = requiredInputKeys.has(inputField.key) && isConditionallyRequiredInput(inputField.key);
    if (isRequired) wrap.classList.add("required-field");
    const input = document.createElement("input");
    input.className = "field-input valuation-input";
    input.type = inputField.value_type === "number" ? "number" : "text";
    input.step = "any";
    if (PERCENT_WHOLE_INPUT_KEYS.has(inputField.key)) {
      input.value = formatPercentInput(getInputDefaultValue(inputField.key, 0));
      input.placeholder = "Enter percent";
      input.dataset.percentInput = "1";
    } else {
      input.value = formatValue(getInputDefaultValue(inputField.key, ""));
    }
    const hintText = VALUATION_INPUT_HINTS.get(inputField.key);
    if (hintText) input.title = hintText;
    wrap.appendChild(input);
    wireInputElement(input, inputField.key, isRequired);
    return wrap;
  }

  const outputField = outputMap.get(cellKey);
  const output = document.createElement("div");
  output.className = "valuation-output";
  const displayFormat = displayFormatOverride || columnFormat || "text";
  if (displayFormat === "percent") output.dataset.format = "percent";
  if (displayFormat === "money") output.dataset.format = "money";
  if (displayFormat === "number") output.dataset.format = "number";
  output.textContent = formatCalculatedOrBlank(outputField?.value, displayFormat);
  if (outputField) registerFormulaElement(outputField.key, output);
  wrap.appendChild(output);
  return wrap;
}

function formatDisplayByType(value, displayType) {
  if (displayType === "percent") return formatPercentValue(value);
  if (displayType === "money") return formatMoney(value);
  if (displayType === "number") return formatNumberValue(value);
  if (displayType === "compact") return formatCompactValue(value);
  if (displayType === "range_money") {
    if (!value || typeof value !== "object") return "--";
    const low = toNumber(value.low, NaN);
    const high = toNumber(value.high, NaN);
    if (!Number.isFinite(low) || !Number.isFinite(high)) return "--";
    return `${formatMoney(low)} - ${formatMoney(high)}`;
  }
  return formatValue(value);
}

function formatCalculatedOrBlank(value, displayType = "text") {
  if (!hasSuccessfulCalculation) return "";
  return formatDisplayByType(value, displayType);
}

function renderValuationSectionTable(section, inputMap, outputMap, usedInputKeys, usedOutputKeys) {
  const block = document.createElement("section");
  block.className = "valuation-section";
  block.dataset.section = section.title.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "");

  const title = document.createElement("h4");
  title.className = "valuation-section-title";
  title.textContent = section.title;
  block.appendChild(title);

  const tableWrap = document.createElement("div");
  tableWrap.className = "valuation-table-wrap";
  const table = document.createElement("table");
  table.className = "valuation-table";

  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  const firstTh = document.createElement("th");
  firstTh.textContent = "";
  hr.appendChild(firstTh);
  for (const col of VALUATION_TABLE_COLUMNS) {
    const th = document.createElement("th");
    th.textContent = col.label;
    hr.appendChild(th);
  }
  thead.appendChild(hr);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const rowCfg of section.rows) {
    const tr = document.createElement("tr");
    if (rowCfg.emphasis) tr.classList.add("emphasis");

    const labelTd = document.createElement("td");
    labelTd.className = "valuation-label-cell";

    const labelMain = document.createElement("div");
    labelMain.className = "valuation-label-main";
    labelMain.textContent = rowCfg.label;
    labelTd.appendChild(labelMain);

    const rowInputHints = [];
    for (const col of VALUATION_TABLE_COLUMNS) {
      const key = rowCfg.cells?.[col.id] || null;
      if (!key || !inputMap.has(key)) continue;
      const hint = VALUATION_INPUT_HINTS.get(key);
      if (hint && !rowInputHints.includes(hint)) rowInputHints.push(hint);
    }

    if (
      rowCfg.description ||
      (rowCfg.descriptionSupplementals && rowCfg.descriptionSupplementals.length) ||
      rowInputHints.length
    ) {
      const note = document.createElement("div");
      note.className = "valuation-label-note";

      if (rowCfg.description) {
        const noteText = document.createElement("span");
        noteText.textContent = rowCfg.description;
        note.appendChild(noteText);
      }

      for (const supplemental of rowCfg.descriptionSupplementals || []) {
        const chunk = document.createElement("span");
        chunk.className = "valuation-label-note-chunk";

        if (supplemental.label) {
          const chunkLabel = document.createElement("span");
          chunkLabel.className = "valuation-label-note-prefix";
          chunkLabel.textContent = `${supplemental.label}:`;
          chunk.appendChild(chunkLabel);
        }

        if (supplemental.key) {
          const supplementalValue = document.createElement("span");
          supplementalValue.className = "valuation-note-value";
          if (supplemental.format) supplementalValue.dataset.format = supplemental.format;
          const supplementalOutput = outputMap.get(supplemental.key);
          supplementalValue.textContent = formatDisplayByType(supplementalOutput?.value, supplemental.format);
          chunk.appendChild(supplementalValue);
          if (supplementalOutput) registerFormulaElement(supplementalOutput.key, supplementalValue);
          if (supplemental.key) usedOutputKeys.add(supplemental.key);
        }

        if (supplemental.suffix) {
          const suffix = document.createElement("span");
          suffix.className = "valuation-label-note-suffix";
          suffix.textContent = supplemental.suffix;
          chunk.appendChild(suffix);
        }

        note.appendChild(chunk);
      }

      if (rowInputHints.length) {
        const hint = document.createElement("span");
        hint.className = "valuation-label-note-chunk";
        hint.textContent = `Input: ${rowInputHints.join(" ")}`;
        note.appendChild(hint);
      }

      labelTd.appendChild(note);
    }
    tr.appendChild(labelTd);

    for (const col of VALUATION_TABLE_COLUMNS) {
      const td = document.createElement("td");
      const key = rowCfg.cells[col.id] || null;
      const displayFormat = rowCfg.formats?.[col.id] || col.format;
      if (key && inputMap.has(key)) {
        usedInputKeys.add(key);
        if (requiredInputKeys.has(key) && isConditionallyRequiredInput(key)) tr.classList.add("has-required");
      }
      if (key && outputMap.has(key)) usedOutputKeys.add(key);
      td.appendChild(renderValuationCellContent(key, col.format, inputMap, outputMap, displayFormat));
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  tableWrap.appendChild(table);
  block.appendChild(tableWrap);

  return block;
}

function renderValuationSecondaryResults(outputMap, usedOutputKeys) {
  const section = document.createElement("section");
  section.className = "valuation-section valuation-secondary";
  section.dataset.section = "secondary-results";

  const title = document.createElement("h4");
  title.className = "valuation-section-title";
  title.textContent = "Secondary Results";
  section.appendChild(title);

  const tableWrap = document.createElement("div");
  tableWrap.className = "valuation-table-wrap";

  const table = document.createElement("table");
  table.className = "valuation-secondary-table";

  const tbody = document.createElement("tbody");
  for (const rowCfg of VALUATION_SECONDARY_ROWS) {
    const tr = document.createElement("tr");

    const labelTd = document.createElement("td");
    labelTd.className = "valuation-secondary-label";
    labelTd.textContent = rowCfg.label;
    tr.appendChild(labelTd);

    const valueTd = document.createElement("td");
    valueTd.className = "valuation-secondary-value";
    const valueEl = document.createElement("span");
    valueEl.className = "valuation-output";
    if (rowCfg.format) valueEl.dataset.format = rowCfg.format;
    const outputField = outputMap.get(rowCfg.key);
    valueEl.textContent = formatCalculatedOrBlank(outputField?.value, rowCfg.format || "text");
    if (outputField) registerFormulaElement(outputField.key, valueEl);
    usedOutputKeys.add(rowCfg.key);
    valueTd.appendChild(valueEl);
    tr.appendChild(valueTd);

    const supplementalTd = document.createElement("td");
    supplementalTd.className = "valuation-secondary-supplemental";
    if (rowCfg.supplementalKey) {
      const supplementalValue = document.createElement("span");
      supplementalValue.className = "valuation-secondary-supplemental-value";
      if (rowCfg.supplementalFormat) supplementalValue.dataset.format = rowCfg.supplementalFormat;
      const supplementalField = outputMap.get(rowCfg.supplementalKey);
      supplementalValue.textContent = formatCalculatedOrBlank(
        supplementalField?.value,
        rowCfg.supplementalFormat || "text"
      );
      if (supplementalField) registerFormulaElement(supplementalField.key, supplementalValue);
      usedOutputKeys.add(rowCfg.supplementalKey);
      supplementalTd.appendChild(supplementalValue);

      if (rowCfg.supplementalSuffix) {
        const suffix = document.createElement("span");
        suffix.className = "valuation-secondary-supplemental-suffix";
        suffix.textContent = rowCfg.supplementalSuffix;
        supplementalTd.appendChild(suffix);
      }
    }
    tr.appendChild(supplementalTd);

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  tableWrap.appendChild(table);
  section.appendChild(tableWrap);
  return section;
}

function renderValuationAdditionalDetails(remainingOutputs) {
  if (!remainingOutputs.length) return null;

  const section = document.createElement("section");
  section.className = "valuation-section valuation-additional";
  section.dataset.section = "additional-details";

  const title = document.createElement("h4");
  title.className = "valuation-section-title";
  title.textContent = "Additional Details";
  section.appendChild(title);

  const list = document.createElement("div");
  list.className = "valuation-additional-list";
  for (const field of remainingOutputs) {
    const row = document.createElement("div");
    row.className = "valuation-additional-row";

    const label = document.createElement("div");
    label.className = "valuation-additional-label";
    label.textContent = `${field.label} (${field.address})`;
    row.appendChild(label);

    const value = document.createElement("div");
    value.className = "valuation-additional-value";
    value.textContent = hasSuccessfulCalculation ? formatValue(field.value) : "";
    registerFormulaElement(field.key, value);
    row.appendChild(value);

    list.appendChild(row);
  }
  section.appendChild(list);
  return section;
}

function renderValuationAssumptions(inputFields, usedInputKeys) {
  const remaining = inputFields
    .filter((field) => !usedInputKeys.has(field.key))
    .sort((a, b) => a.row - b.row || a.col - b.col);
  if (!remaining.length) return null;

  const wrap = document.createElement("aside");
  wrap.className = "valuation-assumptions";
  wrap.dataset.section = "valuation-assumptions";

  const title = document.createElement("h4");
  title.className = "valuation-assumptions-title";
  title.textContent = "Key Valuation Assumptions";
  wrap.appendChild(title);

  const subtitle = document.createElement("p");
  subtitle.className = "valuation-assumptions-subtitle";
  subtitle.textContent = "These inputs drive valuation, debt sizing, returns, and sensitivity outputs.";
  wrap.appendChild(subtitle);

  const list = document.createElement("div");
  list.className = "valuation-assumptions-list";
  for (const field of remaining) {
    const row = document.createElement("div");
    row.className = "valuation-assumption-row";

    const label = document.createElement("label");
    label.className = "valuation-assumption-label";
    label.textContent = `${field.label} (${field.address})`;
    row.appendChild(label);

    const noteText = VALUATION_ASSUMPTION_NOTES.get(field.key);
    const inputHint = VALUATION_INPUT_HINTS.get(field.key);
    if (noteText || inputHint) {
      if (noteText) row.classList.add("is-highlight");
      const note = document.createElement("div");
      note.className = "valuation-assumption-note";
      const parts = [];
      if (noteText) parts.push(noteText);
      if (inputHint) parts.push(`Input: ${inputHint}`);
      note.textContent = parts.join(" ");
      row.appendChild(note);
    }

    const input = document.createElement("input");
    input.className = "field-input";
    input.type = field.value_type === "number" ? "number" : "text";
    input.step = "any";
    const isRequired = requiredInputKeys.has(field.key) && isConditionallyRequiredInput(field.key);
    if (isRequired) {
      row.classList.add("required-field");
      label.classList.add("required-label");
    }
    if (PERCENT_WHOLE_INPUT_KEYS.has(field.key)) {
      input.value = formatPercentInput(getInputDefaultValue(field.key, 0));
      input.placeholder = "Enter percent";
      input.dataset.percentInput = "1";
    } else {
      input.value = formatValue(getInputDefaultValue(field.key, ""));
    }
    if (inputHint) input.title = inputHint;
    wireInputElement(input, field.key, isRequired);
    row.appendChild(input);

    list.appendChild(row);
  }
  wrap.appendChild(list);
  return wrap;
}

function renderValuationPanel(sheetView, index) {
  const panel = document.createElement("section");
  panel.className = "sheet-panel rowform-panel valuation-panel";
  panel.dataset.sheet = sheetView.name;
  if (index === 0) panel.classList.add("active");

  const header = document.createElement("div");
  header.className = "panel-header";
  const titleWrap = document.createElement("div");
  titleWrap.className = "panel-title";
  const h2 = document.createElement("h2");
  h2.textContent = sheetView.name;
  const subtitle = document.createElement("p");
  subtitle.className = "panel-subtitle";
  subtitle.textContent = `${sheetView.inputs.length} inputs · ${sheetView.outputs.length} calculated outputs`;
  titleWrap.appendChild(h2);
  titleWrap.appendChild(subtitle);
  header.appendChild(titleWrap);

  const search = document.createElement("input");
  search.className = "panel-search";
  search.type = "search";
  search.placeholder = "Search valuation rows or assumptions...";
  header.appendChild(search);
  panel.appendChild(header);

  renderValuationKeycards(sheetView, panel);

  const inputMap = new Map(sheetView.inputs.map((f) => [f.key, f]));
  const outputMap = new Map(sheetView.outputs.map((f) => [f.key, f]));
  const usedInputKeys = new Set();
  const usedOutputKeys = new Set(
    VALUATION_RESULT_FIELDS.flatMap((item) =>
      item.secondaryKey ? [item.key, item.secondaryKey] : [item.key]
    )
  );

  for (const section of VALUATION_SECTIONS) {
    for (const row of section.rows) {
      for (const key of Object.values(row.cells || {})) {
        if (key && inputMap.has(key)) usedInputKeys.add(key);
      }
    }
  }

  const layout = document.createElement("div");
  layout.className = "valuation-layout";

  const main = document.createElement("div");
  main.className = "valuation-main";
  for (const section of VALUATION_SECTIONS) {
    main.appendChild(
      renderValuationSectionTable(section, inputMap, outputMap, usedInputKeys, usedOutputKeys)
    );
  }
  main.appendChild(renderValuationSecondaryResults(outputMap, usedOutputKeys));

  const remainingOutputs = sheetView.outputs
    .filter((field) => !usedOutputKeys.has(field.key))
    .sort((a, b) => a.row - b.row || a.col - b.col);
  const additional = renderValuationAdditionalDetails(remainingOutputs);
  if (additional) main.appendChild(additional);

  layout.appendChild(main);
  const assumptions = renderValuationAssumptions(sheetView.inputs, usedInputKeys);
  if (assumptions) layout.appendChild(assumptions);

  panel.appendChild(layout);

  search.addEventListener("input", () => {
    const query = search.value.trim().toLowerCase();
    panel.querySelectorAll(".valuation-section tbody tr").forEach((row) => {
      if (!query) {
        row.classList.remove("hidden");
        return;
      }
      const text = row.textContent.toLowerCase();
      row.classList.toggle("hidden", !text.includes(query));
    });
    panel.querySelectorAll(".valuation-assumption-row").forEach((row) => {
      if (!query) {
        row.classList.remove("hidden");
        return;
      }
      const text = row.textContent.toLowerCase();
      row.classList.toggle("hidden", !text.includes(query));
    });
    panel.querySelectorAll(".valuation-additional-row").forEach((row) => {
      if (!query) {
        row.classList.remove("hidden");
        return;
      }
      const text = row.textContent.toLowerCase();
      row.classList.toggle("hidden", !text.includes(query));
    });
  });

  return panel;
}

function getLiveInputNumber(key, fallback = Number.NaN) {
  const input = inputElements.get(key);
  if (!input) return fallback;
  if (!isPresentValue(input.value)) return fallback;
  return normalizeWorkbookInputNumber(key, input.value, fallback);
}

function getLiveFormulaNumber(key, fallback = Number.NaN) {
  if (!Object.prototype.hasOwnProperty.call(latestFormulaValues, key)) return fallback;
  return toNumber(latestFormulaValues[key], fallback);
}

function computePurchasePriceBreakEvenRange() {
  const purchasePrice = getLiveInputNumber("Valuation!D6", Number.NaN);
  const monthlyIncome = getLiveFormulaNumber("Returns!D18", Number.NaN);
  const totalExpenses = getLiveFormulaNumber("Returns!D32", Number.NaN);
  const mortgageMonthly = getLiveFormulaNumber("Returns!D22", Number.NaN);
  const ltv = getLiveFormulaNumber("Returns!D11", Number.NaN);
  const loanAmount = getLiveFormulaNumber("Returns!D12", Number.NaN);

  if (
    !Number.isFinite(purchasePrice) ||
    !Number.isFinite(monthlyIncome) ||
    !Number.isFinite(totalExpenses) ||
    !Number.isFinite(mortgageMonthly) ||
    !Number.isFinite(ltv) ||
    !Number.isFinite(loanAmount) ||
    loanAmount <= 0 ||
    ltv <= 0
  ) {
    return null;
  }

  const paymentFactor = mortgageMonthly / loanAmount;
  const denominator = paymentFactor * ltv;
  if (!Number.isFinite(denominator) || denominator <= 0) return null;

  const nonDebtExpenses = totalExpenses - mortgageMonthly;
  const breakEvenPrice = (monthlyIncome - nonDebtExpenses) / denominator;
  if (!Number.isFinite(breakEvenPrice)) return null;

  const toleranceMonthlyCashflow = 100;
  const delta = toleranceMonthlyCashflow / denominator;
  const low = Math.max(0, breakEvenPrice - delta);
  const high = Math.max(0, breakEvenPrice + delta);
  return { low: Math.min(low, high), high: Math.max(low, high) };
}

function computeReiRatios() {
  const out = {};
  if (!hasSuccessfulCalculation) {
    for (const def of REI_RATIO_DEFS) out[def.id] = null;
    return out;
  }

  const purchasePrice = getLiveInputNumber("Valuation!D6", Number.NaN);
  const monthlyCashflow = getLiveFormulaNumber("Returns!D34", Number.NaN);
  const annualCashflow = getLiveFormulaNumber("Returns!D33", Number.NaN);
  const annualPgi = getLiveFormulaNumber("Valuation!E13", Number.NaN);
  const monthlyIncome = getLiveFormulaNumber("Returns!D18", Number.NaN);
  const units = rentRollState?.unit_count || toPositiveInt(getLiveInputNumber("Valuation!D5", 0), 0);

  out.cap_rate = getLiveFormulaNumber("Valuation!E47", Number.NaN);
  out.monthly_cashflow_building = monthlyCashflow;
  out.monthly_per_door_cashflow =
    Number.isFinite(monthlyCashflow) && Number.isFinite(units) && units > 0
      ? monthlyCashflow / units
      : Number.NaN;
  out.cash_on_cash_return = getLiveFormulaNumber("Returns!H49", Number.NaN);
  out.debt_service_coverage_ratio = getLiveFormulaNumber("Valuation!E46", Number.NaN);
  out.gross_rent_multiplier =
    Number.isFinite(purchasePrice) && Number.isFinite(annualPgi) && annualPgi > 0
      ? purchasePrice / annualPgi
      : Number.NaN;
  out.operating_expense_ratio = getLiveFormulaNumber("Valuation!H26", Number.NaN);
  out.one_percent_rule =
    Number.isFinite(monthlyIncome) && Number.isFinite(purchasePrice) && purchasePrice > 0
      ? monthlyIncome / purchasePrice
      : Number.NaN;
  out.purchase_price_break_even_range = computePurchasePriceBreakEvenRange();
  out.annual_cashflow_reference = annualCashflow;

  return out;
}

function updateDerivedMetricElements() {
  const ratios = computeReiRatios();
  for (const def of REI_RATIO_DEFS) {
    const value = ratios[def.id];
    const elements = derivedMetricElements.get(def.id) || [];
    for (const el of elements) {
      const format = el.dataset.format || def.format || "text";
      const isEmpty =
        value === null ||
        value === undefined ||
        (typeof value === "number" && !Number.isFinite(value));
      el.textContent = isEmpty ? "--" : formatDisplayByType(value, format);

      if (def.id === "one_percent_rule") {
        const numeric = typeof value === "number" ? value : Number.NaN;
        const threshold = Number(def.threshold || 0.01);
        el.classList.toggle("ratio-pass", Number.isFinite(numeric) && numeric >= threshold);
        el.classList.toggle("ratio-fail", Number.isFinite(numeric) && numeric < threshold);
      }
    }
  }
}

function percentInputToDecimal(value, fallback = 0) {
  const numeric = toNumber(value, Number.NaN);
  if (!Number.isFinite(numeric)) return fallback;
  return numeric / 100;
}

function formatPercentInput(decimalValue) {
  const numeric = toNumber(decimalValue, Number.NaN);
  if (!Number.isFinite(numeric)) return "0";
  const asPercent = numeric * 100;
  if (Math.abs(asPercent) < 1e-9) return "0";
  return String(Number(asPercent.toFixed(4)));
}

function formatBpsInput(bpsValue) {
  const numeric = toNumber(bpsValue, Number.NaN);
  if (!Number.isFinite(numeric) || Math.abs(numeric) < 1e-9) return "0";
  return String(Number(numeric.toFixed(2)));
}

function formatSignedByType(value, formatType) {
  const numeric = toNumber(value, Number.NaN);
  if (!Number.isFinite(numeric)) return "--";
  if (numeric === 0) return formatDisplayByType(0, formatType);
  const sign = numeric > 0 ? "+" : "-";
  return `${sign}${formatDisplayByType(Math.abs(numeric), formatType)}`;
}

function monthlyRateFromAnnualCanadian(annualRate) {
  const rate = toNumber(annualRate, Number.NaN);
  if (!Number.isFinite(rate) || rate < 0) return Number.NaN;
  const compoundsPerYear = 2;
  return Math.exp((Math.log(1 + rate / compoundsPerYear) * compoundsPerYear) / 12) - 1;
}

function computeMonthlyMortgagePayment(principal, annualRate, amortYears) {
  const loan = toNumber(principal, Number.NaN);
  const years = toNumber(amortYears, Number.NaN);
  if (!Number.isFinite(loan) || !Number.isFinite(years) || years <= 0 || loan < 0) return Number.NaN;
  const nper = Math.round(years * 12);
  if (nper <= 0) return Number.NaN;
  const monthlyRate = monthlyRateFromAnnualCanadian(annualRate);
  if (!Number.isFinite(monthlyRate)) return Number.NaN;
  if (Math.abs(monthlyRate) < 1e-12) return loan / nper;
  const denominator = 1 - Math.exp(Math.log(1 + monthlyRate) * -nper);
  if (!Number.isFinite(denominator) || Math.abs(denominator) < 1e-12) return Number.NaN;
  return (loan * monthlyRate) / denominator;
}

function computeAnnualDebtService(principal, annualRate, amortYears) {
  const payment = computeMonthlyMortgagePayment(principal, annualRate, amortYears);
  if (!Number.isFinite(payment)) return Number.NaN;
  return payment * 12;
}

function excelPV(rate, nper, pmt, fv = 0, type = 0) {
  const r = toNumber(rate, Number.NaN);
  const n = Math.round(toNumber(nper, Number.NaN));
  const payment = toNumber(pmt, Number.NaN);
  const futureValue = toNumber(fv, 0);
  const paymentType = toNumber(type, 0);
  if (!Number.isFinite(r) || !Number.isFinite(n) || n <= 0 || !Number.isFinite(payment)) return Number.NaN;
  if (Math.abs(r) < 1e-12) return -(futureValue + payment * n);
  const factor = Math.exp(Math.log(1 + r) * n);
  return -(
    futureValue + payment * (1 + r * paymentType) * ((factor - 1) / r)
  ) / factor;
}

function computeYearOnePrincipalReduction(principal, annualRate, amortYears) {
  const loan = toNumber(principal, Number.NaN);
  if (!Number.isFinite(loan) || loan <= 0) return 0;
  const payment = computeMonthlyMortgagePayment(loan, annualRate, amortYears);
  const monthlyRate = monthlyRateFromAnnualCanadian(annualRate);
  if (!Number.isFinite(payment) || !Number.isFinite(monthlyRate)) return Number.NaN;

  let remaining = loan;
  let paid = 0;
  for (let month = 0; month < 12 && remaining > 0; month += 1) {
    const interest = remaining * monthlyRate;
    let principalComponent = payment - interest;
    if (!Number.isFinite(principalComponent)) return Number.NaN;
    if (principalComponent < 0) principalComponent = 0;
    if (principalComponent > remaining) principalComponent = remaining;
    remaining -= principalComponent;
    paid += principalComponent;
  }
  return paid;
}

function readSensitivityAssumptions() {
  const pestInput = inputElements.get("Returns!D30");
  const pestValue = pestInput ? toNumber(pestInput.value, 0) : 0;

  return {
    pgiAnnual: getLiveFormulaNumber("Valuation!E13", Number.NaN),
    vacancyRate: getLiveInputNumber("Valuation!D14", getLiveFormulaNumber("Returns!C21", Number.NaN)),
    propertyTaxMonthly: getLiveFormulaNumber("Returns!D23", Number.NaN),
    insuranceMonthly: getLiveFormulaNumber("Returns!D24", Number.NaN),
    utilitiesMonthly: getLiveInputNumber("Returns!D25", getLiveFormulaNumber("Returns!D25", Number.NaN)),
    repairsPct: getLiveInputNumber("Returns!C26", Number.NaN),
    propertyManagerPct: getLiveInputNumber("Returns!C27", Number.NaN),
    assetManagerPct: getLiveInputNumber("Returns!C28", Number.NaN),
    bookkeepingMonthly: getLiveInputNumber("Returns!D29", Number.NaN),
    pestMonthly: pestValue,
    advertisingMonthly: getLiveInputNumber("Returns!D31", Number.NaN),
    purchasePrice: getLiveInputNumber("Valuation!D6", Number.NaN),
    capRate: getLiveInputNumber("Valuation!E29", Number.NaN),
    marketRate: getLiveInputNumber("Valuation!E32", Number.NaN),
    amortYears: getLiveInputNumber("Valuation!E33", Number.NaN),
    maxLtv: getLiveInputNumber("Valuation!E34", Number.NaN),
    requiredDscr: getLiveInputNumber("Valuation!E35", Number.NaN),
    cmhcPremiumRate: getLiveInputNumber("Returns!D13", Number.NaN),
    ownershipPct: getLiveInputNumber("Returns!D61", Number.NaN),
    appreciationRate: getLiveInputNumber("Returns!H61", Number.NaN),
    baseDownPayment: getLiveFormulaNumber("Returns!D37", Number.NaN),
    baseCashToClose: getLiveFormulaNumber("Returns!D58", Number.NaN),
    baseGstRebateAmount: getLiveFormulaNumber("Returns!D55", 0),
    get baseNonGstClosingCosts() {
      return Number.isFinite(this.baseCashToClose) && Number.isFinite(this.baseDownPayment)
        ? this.baseCashToClose - this.baseDownPayment - this.baseGstRebateAmount
        : Number.NaN;
    },
  };
}

function readSensitivityBaselineOutputs() {
  const investorRoi = getLiveFormulaNumber("Returns!H58", Number.NaN);
  const ownership = getLiveInputNumber("Returns!D61", Number.NaN);
  let totalPropertyRoi = getLiveFormulaNumber("Returns!E90", Number.NaN);
  if (!Number.isFinite(totalPropertyRoi) && Number.isFinite(investorRoi) && Number.isFinite(ownership) && ownership !== 0) {
    totalPropertyRoi = investorRoi / ownership;
  }

  return {
    value_cap: getLiveFormulaNumber("Valuation!E30", Number.NaN),
    ltc: getLiveFormulaNumber("Valuation!E44", Number.NaN),
    monthly_cashflow: getLiveFormulaNumber("Returns!D34", Number.NaN),
    down_payment: getLiveFormulaNumber("Returns!D37", Number.NaN),
    cash_to_close: getLiveFormulaNumber("Returns!D58", Number.NaN),
    actual_dscr: getLiveFormulaNumber("Valuation!E46", Number.NaN),
    total_property_roi: totalPropertyRoi,
    investor_roi: investorRoi,
  };
}

function computeSensitivityScenarioOutputs(
  assumptions,
  rentChangePct,
  interestRateChangeBps,
  vacancyPctOverride = null,
  purchasePriceOverride = null,
  gstRebate = "baseline"
) {
  const rentFactor = Math.max(0, 1 + toNumber(rentChangePct, 0));
  const adjustedRate = Math.max(0, assumptions.marketRate + (toNumber(interestRateChangeBps, 0) / 10000));

  const effectivePurchasePrice = purchasePriceOverride != null ? purchasePriceOverride : assumptions.purchasePrice;
  const effectiveVacancyRate = vacancyPctOverride != null ? vacancyPctOverride : assumptions.vacancyRate;

  const monthlyIncome = (assumptions.pgiAnnual * rentFactor) / 12;
  const vacancyMonthly = monthlyIncome * effectiveVacancyRate;

  const fixedOpsMonthly =
    assumptions.propertyTaxMonthly +
    assumptions.insuranceMonthly +
    assumptions.utilitiesMonthly +
    assumptions.bookkeepingMonthly +
    assumptions.pestMonthly +
    assumptions.advertisingMonthly;
  const variableOpsMonthly =
    (assumptions.repairsPct + assumptions.propertyManagerPct + assumptions.assetManagerPct) * monthlyIncome;
  const operatingMonthly = fixedOpsMonthly + variableOpsMonthly;

  const egiAnnual = (monthlyIncome - vacancyMonthly) * 12;
  const operatingAnnual = operatingMonthly * 12;
  const noiAnnual = egiAnnual - operatingAnnual;

  const valueCap = assumptions.capRate !== 0 ? noiAnnual / assumptions.capRate : Number.NaN;
  const maxLoanLtv = assumptions.maxLtv * valueCap;
  const maxLoanLtc = assumptions.maxLtv * effectivePurchasePrice;
  const availableDebtService = assumptions.requiredDscr !== 0 ? noiAnnual / assumptions.requiredDscr : Number.NaN;
  const maxLoanDscr = excelPV(adjustedRate / 12, assumptions.amortYears * 12, -(availableDebtService / 12), 0, 0);

  const validLoanCandidates = [maxLoanLtv, maxLoanLtc, maxLoanDscr].filter((x) => Number.isFinite(x));
  const lesserLoan = validLoanCandidates.length > 0 ? Math.max(0, Math.min(...validLoanCandidates)) : Number.NaN;
  const ltc = effectivePurchasePrice !== 0 ? lesserLoan / effectivePurchasePrice : Number.NaN;

  const annualDebtService = computeAnnualDebtService(lesserLoan, adjustedRate, assumptions.amortYears);
  const actualDscr = annualDebtService !== 0 ? noiAnnual / annualDebtService : Number.NaN;

  const loanWithPremium = lesserLoan * (1 + assumptions.cmhcPremiumRate);
  const mortgageMonthly = computeMonthlyMortgagePayment(
    loanWithPremium,
    adjustedRate,
    assumptions.amortYears
  );
  const totalExpensesMonthly = vacancyMonthly + mortgageMonthly + operatingMonthly;
  const monthlyCashflow = monthlyIncome - totalExpensesMonthly;
  const annualCashflow = monthlyCashflow * 12;

  const downPayment = effectivePurchasePrice - lesserLoan;
  let scenarioGstRebate;
  if (gstRebate === "yes") {
    scenarioGstRebate = effectivePurchasePrice / 1.05 - effectivePurchasePrice; // negative
  } else if (gstRebate === "no") {
    scenarioGstRebate = 0;
  } else {
    scenarioGstRebate = assumptions.baseGstRebateAmount;
  }
  const cashToClose = downPayment + assumptions.baseNonGstClosingCosts + scenarioGstRebate;

  const principalReductionYear1 = computeYearOnePrincipalReduction(
    loanWithPremium,
    adjustedRate,
    assumptions.amortYears
  );
  const appreciationGainYear1 = effectivePurchasePrice * assumptions.appreciationRate;

  const canComputeRoi = Number.isFinite(cashToClose) && cashToClose > 0;
  const cashflowRoi = canComputeRoi ? annualCashflow / cashToClose : Number.NaN;
  const principalRoi = canComputeRoi ? principalReductionYear1 / cashToClose : Number.NaN;
  const appreciationRoi = canComputeRoi ? appreciationGainYear1 / cashToClose : Number.NaN;
  const totalPropertyRoi =
    Number.isFinite(cashflowRoi) && Number.isFinite(principalRoi) && Number.isFinite(appreciationRoi)
      ? cashflowRoi + principalRoi + appreciationRoi
      : Number.NaN;
  const investorRoi = totalPropertyRoi * assumptions.ownershipPct;

  return {
    value_cap: valueCap,
    ltc,
    monthly_cashflow: monthlyCashflow,
    down_payment: downPayment,
    cash_to_close: cashToClose,
    actual_dscr: actualDscr,
    total_property_roi: totalPropertyRoi,
    investor_roi: investorRoi,
  };
}

function updateSensitivityPanel() {
  if (!sensitivityView) return;

  const baseline = readSensitivityBaselineOutputs();
  if (!hasSuccessfulCalculation) {
    sensitivityView.statusEl.textContent =
      "Run Analysis first to establish the baseline scenario, then use the adjustments below.";
    const currentRate = getLiveInputNumber("Valuation!E32", Number.NaN);
    sensitivityView.baseRateValue.textContent = Number.isFinite(currentRate)
      ? formatPercentValue(currentRate)
      : "--";
    sensitivityView.adjustedRateValue.textContent = "--";
    if (sensitivityView.addScenarioBtn) sensitivityView.addScenarioBtn.disabled = true;
    renderSensitivityTable(null, null);
    return;
  }

  const assumptions = readSensitivityAssumptions();
  const requiredValues = [
    assumptions.pgiAnnual,
    assumptions.vacancyRate,
    assumptions.propertyTaxMonthly,
    assumptions.insuranceMonthly,
    assumptions.utilitiesMonthly,
    assumptions.repairsPct,
    assumptions.propertyManagerPct,
    assumptions.assetManagerPct,
    assumptions.bookkeepingMonthly,
    assumptions.advertisingMonthly,
    assumptions.purchasePrice,
    assumptions.capRate,
    assumptions.marketRate,
    assumptions.amortYears,
    assumptions.maxLtv,
    assumptions.requiredDscr,
    assumptions.cmhcPremiumRate,
    assumptions.ownershipPct,
    assumptions.appreciationRate,
    assumptions.baseDownPayment,
    assumptions.baseCashToClose,
    assumptions.baseNonGstClosingCosts,
  ];
  const hasRequired = requiredValues.every((value) => Number.isFinite(value));
  if (!hasRequired) {
    sensitivityView.statusEl.textContent =
      "Required baseline values are missing. Complete required fields on other tabs and run analysis again.";
    sensitivityView.baseRateValue.textContent = "--";
    sensitivityView.adjustedRateValue.textContent = "--";
    if (sensitivityView.addScenarioBtn) sensitivityView.addScenarioBtn.disabled = true;
    renderSensitivityTable(null, null);
    return;
  }

  sensitivityView.statusEl.textContent =
    "Scenario values are calculated using current analysis outputs plus your adjustments.";
  sensitivityView.baseRateValue.textContent = formatPercentValue(assumptions.marketRate);
  sensitivityView.adjustedRateValue.textContent = formatPercentValue(
    assumptions.marketRate + (sensitivityState.interestRateChangeBps / 10000)
  );

  const atMax = sensitivityState.scenarios.length >= 10;
  if (sensitivityView.addScenarioBtn) {
    sensitivityView.addScenarioBtn.disabled = atMax;
    sensitivityView.addScenarioBtn.title = atMax ? "Maximum 10 scenarios reached" : "";
  }

  renderSensitivityTable(baseline, assumptions);
}

function buildScenarioHeaderLines(sc) {
  const lines = [];
  if (sc.rentChangePct !== 0) {
    const sign = sc.rentChangePct > 0 ? "+" : "";
    lines.push(`Rent ${sign}${(sc.rentChangePct * 100).toFixed(1)}%`);
  }
  if (sc.interestRateChangeBps !== 0) {
    const sign = sc.interestRateChangeBps > 0 ? "+" : "";
    lines.push(`Rate ${sign}${sc.interestRateChangeBps} bps`);
  }
  if (sc.vacancyPct != null) {
    lines.push(`Vac ${formatPercentValue(sc.vacancyPct)}`);
  }
  if (sc.purchasePriceOverride != null) {
    lines.push(`PP ${formatCompactValue(sc.purchasePriceOverride)}`);
  }
  if (sc.gstRebate !== "baseline") {
    lines.push(`GST ${sc.gstRebate === "yes" ? "Yes" : "No"}`);
  }
  if (lines.length === 0) lines.push("Baseline inputs");
  return lines;
}

function renderSensitivityTable(baseline, assumptions) {
  if (!sensitivityView) return;
  const container = sensitivityView.tableContainer;
  container.innerHTML = "";

  if (!baseline || !assumptions) {
    const placeholder = document.createElement("p");
    placeholder.className = "sensitivity-placeholder";
    placeholder.textContent = "Run Analysis to populate scenario comparison.";
    container.appendChild(placeholder);
    return;
  }

  // Compute the live preview (current control values, not yet saved)
  const liveOutputs = computeSensitivityScenarioOutputs(
    assumptions,
    sensitivityState.rentChangePct,
    sensitivityState.interestRateChangeBps,
    sensitivityState.vacancyPct,
    sensitivityState.purchasePriceOverride,
    sensitivityState.gstRebate
  );

  // Build column list: live preview + saved scenarios
  const liveCol = {
    label: "Current",
    headerSublines: buildScenarioHeaderLines(sensitivityState),
    outputs: liveOutputs,
    isLive: true,
    scenarioRef: null,
  };
  const savedCols = sensitivityState.scenarios.map((sc) => ({
    label: sc.label,
    headerSublines: buildScenarioHeaderLines(sc),
    outputs: sc.outputs,
    isLive: false,
    scenarioRef: sc,
  }));
  const allCols = [liveCol, ...savedCols];

  // Build table
  const table = document.createElement("table");
  table.className = "sensitivity-table sensitivity-table--multi";

  // THEAD
  const thead = document.createElement("thead");

  // Row 1: column group headers
  const headerRow1 = document.createElement("tr");
  const thMetric = document.createElement("th");
  thMetric.rowSpan = 2;
  thMetric.className = "sensitivity-label-col";
  thMetric.textContent = "Metric";
  headerRow1.appendChild(thMetric);

  const thBase = document.createElement("th");
  thBase.rowSpan = 2;
  thBase.className = "sensitivity-base-col";
  thBase.textContent = "Baseline";
  headerRow1.appendChild(thBase);

  for (const col of allCols) {
    const th = document.createElement("th");
    th.colSpan = 2;
    th.className = "sensitivity-col-header";
    if (col.isLive) th.classList.add("sensitivity-col-live");

    const nameSpan = document.createElement("span");
    nameSpan.className = "sensitivity-col-name";
    nameSpan.textContent = col.label;
    th.appendChild(nameSpan);

    if (col.headerSublines.length > 0) {
      const sub = document.createElement("span");
      sub.className = "sensitivity-col-sublines";
      sub.textContent = col.headerSublines.join(" · ");
      th.appendChild(document.createElement("br"));
      th.appendChild(sub);
    }

    if (!col.isLive && col.scenarioRef) {
      const removeBtn = document.createElement("button");
      removeBtn.className = "sensitivity-remove-btn";
      removeBtn.textContent = "×";
      removeBtn.title = `Remove ${col.label}`;
      removeBtn.addEventListener("click", () => {
        sensitivityState.scenarios = sensitivityState.scenarios.filter((s) => s !== col.scenarioRef);
        updateSensitivityPanel();
      });
      th.appendChild(removeBtn);
    }

    headerRow1.appendChild(th);
  }
  thead.appendChild(headerRow1);

  // Row 2: sub-headers "Value" | "Δ vs Base" for each col
  const headerRow2 = document.createElement("tr");
  for (let i = 0; i < allCols.length; i++) {
    const thVal = document.createElement("th");
    thVal.className = "sensitivity-subhead";
    thVal.textContent = "Value";
    const thDelta = document.createElement("th");
    thDelta.className = "sensitivity-subhead sensitivity-subhead--delta";
    thDelta.textContent = "Δ vs Base";
    headerRow2.appendChild(thVal);
    headerRow2.appendChild(thDelta);
  }
  thead.appendChild(headerRow2);
  table.appendChild(thead);

  // TBODY
  const tbody = document.createElement("tbody");

  // Helper to build an input display row
  function makeInputRow(label, getVal) {
    const tr = document.createElement("tr");
    tr.className = "sensitivity-input-row";
    const tdLabel = document.createElement("td");
    tdLabel.className = "sensitivity-label";
    tdLabel.textContent = label;
    tr.appendChild(tdLabel);
    const tdBase = document.createElement("td");
    tdBase.className = "sensitivity-base";
    tdBase.textContent = "--";
    tr.appendChild(tdBase);
    for (const col of allCols) {
      const sc = col.isLive ? sensitivityState : col.scenarioRef;
      const tdVal = document.createElement("td");
      tdVal.className = "sensitivity-scenario";
      if (col.isLive) tdVal.classList.add("sensitivity-scenario--live");
      tdVal.textContent = getVal(sc);
      tr.appendChild(tdVal);
      const tdDelta = document.createElement("td");
      tdDelta.className = "sensitivity-delta";
      tdDelta.textContent = "";
      tr.appendChild(tdDelta);
    }
    return tr;
  }

  // Section header helper
  function makeSectionHeader(label) {
    const tr = document.createElement("tr");
    tr.className = "sensitivity-section-header";
    const td = document.createElement("td");
    td.colSpan = 2 + allCols.length * 2;
    td.textContent = label;
    tr.appendChild(td);
    return tr;
  }

  // Inputs section
  tbody.appendChild(makeSectionHeader("Inputs"));

  tbody.appendChild(makeInputRow("Rent Change", (sc) => {
    if (!sc.rentChangePct) return "0%";
    const sign = sc.rentChangePct > 0 ? "+" : "";
    return `${sign}${(sc.rentChangePct * 100).toFixed(1)}%`;
  }));

  tbody.appendChild(makeInputRow("Rate Change (bps)", (sc) => {
    const bps = sc.interestRateChangeBps || 0;
    return bps === 0 ? "0 bps" : `${bps > 0 ? "+" : ""}${bps} bps`;
  }));

  tbody.appendChild(makeInputRow("Vacancy", (sc) => {
    return sc.vacancyPct != null ? formatPercentValue(sc.vacancyPct) : "Baseline";
  }));

  tbody.appendChild(makeInputRow("Purchase Price", (sc) => {
    return sc.purchasePriceOverride != null ? `$${formatMoney(sc.purchasePriceOverride)}` : "Baseline";
  }));

  tbody.appendChild(makeInputRow("GST Rebate", (sc) => {
    if (sc.gstRebate === "yes") return "Yes";
    if (sc.gstRebate === "no") return "No";
    return "Baseline";
  }));

  // Outputs section
  tbody.appendChild(makeSectionHeader("Outputs"));

  for (const field of SENSITIVITY_RESULT_FIELDS) {
    const tr = document.createElement("tr");
    const baseValue = baseline[field.id];

    const tdLabel = document.createElement("td");
    tdLabel.className = "sensitivity-label";
    tdLabel.textContent = field.label;
    tr.appendChild(tdLabel);

    const tdBase = document.createElement("td");
    tdBase.className = "sensitivity-base";
    tdBase.textContent = Number.isFinite(baseValue) ? formatDisplayByType(baseValue, field.format) : "--";
    tr.appendChild(tdBase);

    for (const col of allCols) {
      const scenarioValue = col.outputs[field.id];
      const delta = Number.isFinite(baseValue) && Number.isFinite(scenarioValue)
        ? scenarioValue - baseValue : Number.NaN;

      const tdVal = document.createElement("td");
      tdVal.className = "sensitivity-scenario";
      if (col.isLive) tdVal.classList.add("sensitivity-scenario--live");
      tdVal.textContent = Number.isFinite(scenarioValue)
        ? formatDisplayByType(scenarioValue, field.format) : "--";
      tr.appendChild(tdVal);

      const tdDelta = document.createElement("td");
      tdDelta.className = "sensitivity-delta";
      tdDelta.textContent = formatSignedByType(delta, field.format);
      tdDelta.classList.toggle("delta-positive", Number.isFinite(delta) && delta > 0);
      tdDelta.classList.toggle("delta-negative", Number.isFinite(delta) && delta < 0);
      tr.appendChild(tdDelta);
    }

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  container.appendChild(table);
}

function renderSensitivityPanel(sheetView, index) {
  const panel = document.createElement("section");
  panel.className = "sheet-panel sensitivity-panel";
  panel.dataset.sheet = sheetView.name;
  if (index === 0) panel.classList.add("active");

  const header = document.createElement("div");
  header.className = "panel-header";
  const titleWrap = document.createElement("div");
  titleWrap.className = "panel-title";
  const h2 = document.createElement("h2");
  h2.textContent = sheetView.name;
  const subtitle = document.createElement("p");
  subtitle.className = "panel-subtitle";
  subtitle.textContent =
    "Adjust assumptions to stress-test valuation, cash flow, and ROI. Save multiple scenarios to compare side by side.";
  titleWrap.appendChild(h2);
  titleWrap.appendChild(subtitle);
  header.appendChild(titleWrap);
  panel.appendChild(header);

  const controls = document.createElement("section");
  controls.className = "sensitivity-controls";
  controls.innerHTML = `
    <div class="sensitivity-control-card">
      <label>Rent Change (%)</label>
      <input type="number" step="0.1" class="field-input" id="sensitivity-rent-change" />
      <p class="sensitivity-help">Example: -10 means rents are reduced by 10% from baseline.</p>
    </div>
    <div class="sensitivity-control-card">
      <label>Interest Rate Change (bps)</label>
      <input type="number" step="0.1" class="field-input" id="sensitivity-rate-change" />
      <p class="sensitivity-help">+100 bps raises a 5.25% rate to 6.25%.</p>
    </div>
    <div class="sensitivity-control-card">
      <label>Vacancy (%)</label>
      <input type="number" step="0.1" min="0" max="100" class="field-input" id="sensitivity-vacancy" placeholder="Blank = baseline" />
      <p class="sensitivity-help">Absolute override (e.g. 5 = 5% vacancy). Leave blank to use baseline.</p>
    </div>
    <div class="sensitivity-control-card">
      <label>Purchase Price ($)</label>
      <input type="number" step="1000" min="0" class="field-input" id="sensitivity-purchase-price" placeholder="Blank = baseline" />
      <p class="sensitivity-help">Absolute dollar amount. Leave blank to use baseline purchase price.</p>
    </div>
    <div class="sensitivity-control-card">
      <label>GST Rebate</label>
      <select class="field-input" id="sensitivity-gst-rebate">
        <option value="baseline">Baseline</option>
        <option value="yes">Yes</option>
        <option value="no">No</option>
      </select>
      <p class="sensitivity-help">"Baseline" inherits the current analysis setting.</p>
    </div>
    <div class="sensitivity-control-card sensitivity-rate-summary">
      <p class="sensitivity-summary-label">Base Interest Rate</p>
      <p class="sensitivity-summary-value" id="sensitivity-base-rate">--</p>
      <p class="sensitivity-summary-label">Scenario Interest Rate</p>
      <p class="sensitivity-summary-value" id="sensitivity-adjusted-rate">--</p>
    </div>
  `;
  panel.appendChild(controls);

  const scenarioBar = document.createElement("div");
  scenarioBar.className = "sensitivity-scenario-bar";
  const addBtn = document.createElement("button");
  addBtn.className = "btn btn-primary sensitivity-add-btn";
  addBtn.id = "sensitivity-add-scenario";
  addBtn.textContent = "Add Scenario";
  addBtn.disabled = true;
  const scenarioCount = document.createElement("span");
  scenarioCount.className = "sensitivity-scenario-count";
  scenarioCount.id = "sensitivity-scenario-count";
  scenarioBar.appendChild(addBtn);
  scenarioBar.appendChild(scenarioCount);
  panel.appendChild(scenarioBar);

  const status = document.createElement("div");
  status.className = "sensitivity-status";
  status.textContent =
    "Run Analysis first to establish the baseline scenario, then use the adjustments below.";
  panel.appendChild(status);

  const results = document.createElement("section");
  results.className = "sensitivity-results";
  panel.appendChild(results);

  // Wire inputs
  const rentInput = controls.querySelector("#sensitivity-rent-change");
  const rateInput = controls.querySelector("#sensitivity-rate-change");
  const vacancyInput = controls.querySelector("#sensitivity-vacancy");
  const priceInput = controls.querySelector("#sensitivity-purchase-price");
  const gstSelect = controls.querySelector("#sensitivity-gst-rebate");
  const baseRateValue = controls.querySelector("#sensitivity-base-rate");
  const adjustedRateValue = controls.querySelector("#sensitivity-adjusted-rate");

  // Restore live state
  rentInput.value = formatPercentInput(sensitivityState.rentChangePct);
  rateInput.value = formatBpsInput(sensitivityState.interestRateChangeBps);
  vacancyInput.value = sensitivityState.vacancyPct != null
    ? String(sensitivityState.vacancyPct * 100) : "";
  priceInput.value = sensitivityState.purchasePriceOverride != null
    ? String(sensitivityState.purchasePriceOverride) : "";
  gstSelect.value = sensitivityState.gstRebate;

  rentInput.addEventListener("input", () => {
    sensitivityState.rentChangePct = percentInputToDecimal(rentInput.value, 0);
    updateSensitivityPanel();
  });
  rateInput.addEventListener("input", () => {
    sensitivityState.interestRateChangeBps = toNumber(rateInput.value, 0);
    updateSensitivityPanel();
  });
  vacancyInput.addEventListener("input", () => {
    const raw = vacancyInput.value.trim();
    sensitivityState.vacancyPct = raw === "" ? null : toNumber(raw, 0) / 100;
    updateSensitivityPanel();
  });
  priceInput.addEventListener("input", () => {
    const raw = priceInput.value.trim();
    sensitivityState.purchasePriceOverride = raw === "" ? null : toNumber(raw, 0);
    updateSensitivityPanel();
  });
  gstSelect.addEventListener("change", () => {
    sensitivityState.gstRebate = gstSelect.value;
    updateSensitivityPanel();
  });

  // Add Scenario button
  addBtn.addEventListener("click", () => {
    if (sensitivityState.scenarios.length >= 10 || !hasSuccessfulCalculation) return;
    const assumptions = readSensitivityAssumptions();
    const outputs = computeSensitivityScenarioOutputs(
      assumptions,
      sensitivityState.rentChangePct,
      sensitivityState.interestRateChangeBps,
      sensitivityState.vacancyPct,
      sensitivityState.purchasePriceOverride,
      sensitivityState.gstRebate
    );
    sensitivityState.scenarios.push({
      label: `Scenario ${sensitivityState.scenarios.length + 1}`,
      rentChangePct: sensitivityState.rentChangePct,
      interestRateChangeBps: sensitivityState.interestRateChangeBps,
      vacancyPct: sensitivityState.vacancyPct,
      purchasePriceOverride: sensitivityState.purchasePriceOverride,
      gstRebate: sensitivityState.gstRebate,
      outputs,
    });
    updateSensitivityPanel();
  });

  sensitivityView = {
    statusEl: status,
    baseRateValue,
    adjustedRateValue,
    tableContainer: results,
    addScenarioBtn: addBtn,
  };
  updateSensitivityPanel();
  return panel;
}

function renderReiRatiosSection(titleText, subtitleText = "") {
  const section = document.createElement("section");
  section.className = "returns-section rei-ratios-section";

  const title = document.createElement("h4");
  title.className = "returns-section-title";
  title.textContent = titleText;
  section.appendChild(title);

  if (subtitleText) {
    const subtitle = document.createElement("p");
    subtitle.className = "returns-section-subtitle";
    subtitle.textContent = subtitleText;
    section.appendChild(subtitle);
  }

  const tableWrap = document.createElement("div");
  tableWrap.className = "returns-table-wrap";
  const table = document.createElement("table");
  table.className = "rei-ratios-table";

  const tbody = document.createElement("tbody");
  for (const def of REI_RATIO_DEFS) {
    const tr = document.createElement("tr");

    const labelTd = document.createElement("td");
    labelTd.className = "rei-ratio-label";
    labelTd.textContent = def.label;
    tr.appendChild(labelTd);

    const valueTd = document.createElement("td");
    valueTd.className = "rei-ratio-value-cell";
    const value = document.createElement("span");
    value.className = "rei-ratio-value";
    value.textContent = "--";
    registerDerivedMetricElement(def.id, value, def.format || "text");
    valueTd.appendChild(value);
    tr.appendChild(valueTd);

    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  tableWrap.appendChild(table);
  section.appendChild(tableWrap);
  return section;
}

function renderReturnsKeyResults(sheetView, panel) {
  const inputMap = new Map(sheetView.inputs.map((field) => [field.key, field]));
  const outputMap = new Map(sheetView.outputs.map((field) => [field.key, field]));

  const wrap = document.createElement("section");
  wrap.className = "returns-keycards";

  const title = document.createElement("h3");
  title.className = "returns-keycards-title";
  title.textContent = "Key Results";
  wrap.appendChild(title);

  const grid = document.createElement("div");
  grid.className = "returns-keycards-grid";
  for (const item of RETURNS_KEY_RESULTS) {
    const card = document.createElement("article");
    card.className = "returns-keycard";

    const label = document.createElement("div");
    label.className = "returns-keycard-label";
    label.textContent = item.label;
    card.appendChild(label);

    const value = document.createElement("div");
    value.className = "returns-keycard-value";
    value.dataset.format = item.format || "text";

    if (inputMap.has(item.key)) {
      value.textContent = formatInputValueForDisplay(item.key, getInputDefaultValue(item.key, ""), item.format || "text");
      registerInputDisplayElement(item.key, value, item.format || "text");
    } else if (outputMap.has(item.key)) {
      const out = outputMap.get(item.key);
      value.textContent = formatCalculatedOrBlank(out?.value, item.format || "text");
      registerFormulaElement(item.key, value);
    } else {
      value.textContent = "";
    }

    card.appendChild(value);
    grid.appendChild(card);
  }
  wrap.appendChild(grid);
  panel.appendChild(wrap);
}

function renderReturnsCellContent(cellKey, formatType, inputMap, outputMap, usedInputKeys, usedOutputKeys) {
  const wrap = document.createElement("div");
  wrap.className = "returns-cell";

  if (!cellKey) {
    wrap.classList.add("empty");
    return wrap;
  }

  const inputField = inputMap.get(cellKey);
  if (inputField) {
    const isRequired = requiredInputKeys.has(inputField.key);
    if (isRequired) wrap.classList.add("required-field");

    const input = document.createElement("input");
    input.className = "field-input returns-input";
    input.type = inputField.value_type === "number" ? "number" : "text";
    input.step = "any";
    if (PERCENT_WHOLE_INPUT_KEYS.has(inputField.key)) {
      input.value = formatPercentInput(getInputDefaultValue(inputField.key, 0));
      input.placeholder = "Enter percent";
      input.dataset.percentInput = "1";
    } else {
      input.value = formatValue(getInputDefaultValue(inputField.key, ""));
    }
    wrap.appendChild(input);
    wireInputElement(input, inputField.key, isRequired);
    usedInputKeys.add(inputField.key);
    return wrap;
  }

  const outputField = outputMap.get(cellKey);
  const output = document.createElement("div");
  output.className = "returns-output";
  output.dataset.format = formatType || "text";
  output.textContent = formatCalculatedOrBlank(outputField?.value, formatType || "text");
  if (outputField) {
    registerFormulaElement(outputField.key, output);
    usedOutputKeys.add(outputField.key);
  }
  wrap.appendChild(output);
  return wrap;
}

function renderReturnsSectionTable(section, inputMap, outputMap, usedInputKeys, usedOutputKeys) {
  const block = document.createElement("section");
  block.className = "returns-section";

  const title = document.createElement("h4");
  title.className = "returns-section-title";
  title.textContent = section.title;
  block.appendChild(title);

  const tableWrap = document.createElement("div");
  tableWrap.className = "returns-table-wrap";
  const table = document.createElement("table");
  table.className = "returns-table";

  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  const firstTh = document.createElement("th");
  firstTh.textContent = "";
  hr.appendChild(firstTh);
  for (const col of RETURNS_TABLE_COLUMNS) {
    const th = document.createElement("th");
    th.textContent = col.label;
    hr.appendChild(th);
  }
  thead.appendChild(hr);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const rowCfg of section.rows) {
    const tr = document.createElement("tr");
    if (rowCfg.emphasis) tr.classList.add("emphasis");

    const labelTd = document.createElement("td");
    labelTd.className = "returns-label-cell";
    labelTd.textContent = rowCfg.label;
    tr.appendChild(labelTd);

    for (const col of RETURNS_TABLE_COLUMNS) {
      const td = document.createElement("td");
      const key = rowCfg.cells?.[col.id] || null;
      const formatType = rowCfg.formats?.[col.id] || col.format || "text";
      if (key && inputMap.has(key) && requiredInputKeys.has(key)) tr.classList.add("has-required");
      td.appendChild(
        renderReturnsCellContent(key, formatType, inputMap, outputMap, usedInputKeys, usedOutputKeys)
      );
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  tableWrap.appendChild(table);
  block.appendChild(tableWrap);
  return block;
}

function renderReturnsPanel(sheetView, index) {
  const panel = document.createElement("section");
  panel.className = "sheet-panel returns-panel";
  panel.dataset.sheet = sheetView.name;
  if (index === 0) panel.classList.add("active");

  const header = document.createElement("div");
  header.className = "panel-header";
  const titleWrap = document.createElement("div");
  titleWrap.className = "panel-title";
  const h2 = document.createElement("h2");
  h2.textContent = sheetView.name;
  const subtitle = document.createElement("p");
  subtitle.className = "panel-subtitle";
  subtitle.textContent = `${sheetView.inputs.length} inputs · ${sheetView.outputs.length} calculated outputs`;
  titleWrap.appendChild(h2);
  titleWrap.appendChild(subtitle);
  header.appendChild(titleWrap);

  const search = document.createElement("input");
  search.className = "panel-search";
  search.type = "search";
  search.placeholder = "Search returns rows, ratios, or fields...";
  header.appendChild(search);
  panel.appendChild(header);

  renderReturnsKeyResults(sheetView, panel);

  const inputMap = new Map(sheetView.inputs.map((f) => [f.key, f]));
  const outputMap = new Map(sheetView.outputs.map((f) => [f.key, f]));
  const usedInputKeys = new Set();
  const usedOutputKeys = new Set(RETURNS_KEY_RESULTS.map((item) => item.key));

  const layout = document.createElement("div");
  layout.className = "returns-layout";

  const main = document.createElement("div");
  main.className = "returns-main";
  for (const section of RETURNS_SECTIONS) {
    main.appendChild(
      renderReturnsSectionTable(section, inputMap, outputMap, usedInputKeys, usedOutputKeys)
    );
  }
  main.appendChild(
    renderReiRatiosSection(
      "REI Ratios",
      "Calculated from the current underwriting run."
    )
  );
  layout.appendChild(main);

  const side = document.createElement("aside");
  side.className = "returns-side";

  const remainingInputs = sheetView.inputs
    .filter((field) => !usedInputKeys.has(field.key))
    .sort((a, b) => a.row - b.row || a.col - b.col);
  if (remainingInputs.length) {
    side.appendChild(createFieldGroup("Additional Returns Inputs", remainingInputs, false));
  }

  const remainingOutputs = sheetView.outputs
    .filter((field) => !usedOutputKeys.has(field.key))
    .sort((a, b) => a.row - b.row || a.col - b.col)
    .map((field) => (hasSuccessfulCalculation ? field : { ...field, value: "" }));
  if (remainingOutputs.length) {
    side.appendChild(createFieldGroup("Additional Returns Outputs", remainingOutputs, false));
  }

  if (side.children.length > 0) layout.appendChild(side);
  panel.appendChild(layout);

  search.addEventListener("input", () => {
    const query = search.value.trim().toLowerCase();
    panel.querySelectorAll(".returns-section tbody tr, .rei-ratios-table tbody tr, .field-row").forEach((row) => {
      if (!query) {
        row.classList.remove("hidden");
        return;
      }
      const text = row.textContent.toLowerCase();
      row.classList.toggle("hidden", !text.includes(query));
    });

    panel.querySelectorAll(".field-group").forEach((group) => {
      const anyVisible = group.querySelector(".field-row:not(.hidden)") !== null;
      group.classList.toggle("hidden", !anyVisible);
    });
  });

  return panel;
}

function renderReiRatiosPanel() {
  const panel = document.createElement("section");
  panel.className = "sheet-panel rei-panel";
  panel.dataset.sheet = "REI Ratios";

  const header = document.createElement("div");
  header.className = "panel-header";
  const titleWrap = document.createElement("div");
  titleWrap.className = "panel-title";
  const h2 = document.createElement("h2");
  h2.textContent = "REI Ratios";
  const subtitle = document.createElement("p");
  subtitle.className = "panel-subtitle";
  subtitle.textContent = "Standalone investment ratio view from the current underwriting scenario.";
  titleWrap.appendChild(h2);
  titleWrap.appendChild(subtitle);
  header.appendChild(titleWrap);
  panel.appendChild(header);

  panel.appendChild(
    renderReiRatiosSection(
      "Core Real Estate Investment Ratios",
      "Run Analysis after changing inputs to refresh ratio values."
    )
  );

  return panel;
}

function renderGenericSheetPanel(sheetView, index) {
  const panel = document.createElement("section");
  panel.className = "sheet-panel";
  panel.dataset.sheet = sheetView.name;

  const header = document.createElement("div");
  header.className = "panel-header";
  const titleWrap = document.createElement("div");
  titleWrap.className = "panel-title";
  const h2 = document.createElement("h2");
  h2.textContent = sheetView.name;
  const subtitle = document.createElement("p");
  subtitle.className = "panel-subtitle";
  subtitle.textContent = `${sheetView.inputs.length} inputs · ${sheetView.outputs.length} calculated outputs`;
  titleWrap.appendChild(h2);
  titleWrap.appendChild(subtitle);
  header.appendChild(titleWrap);

  const search = document.createElement("input");
  search.className = "panel-search";
  search.type = "search";
  search.placeholder = "Search fields by label or cell...";
  header.appendChild(search);
  panel.appendChild(header);

  if (sheetView.metrics.length > 0) {
    const metricWrap = document.createElement("div");
    metricWrap.className = "sheet-quick-metrics";
    for (const metricField of sheetView.metrics) {
      const { card, valueEl } = createMetricCard(
        metricField.label,
        formatCompactValue(metricField.value)
      );
      valueEl.classList.add("metric-live");
      registerFormulaElement(metricField.key, valueEl);
      metricWrap.appendChild(card);
    }
    panel.appendChild(metricWrap);
  }

  const layout = document.createElement("div");
  layout.className = "sheet-layout";

  const inputColumn = document.createElement("div");
  inputColumn.className = "panel-column";
  inputColumn.innerHTML = `<h3 class="column-header">Inputs</h3>`;
  const inputGroups = groupFields(sheetView.inputs);
  if (inputGroups.length === 0) {
    const empty = document.createElement("p");
    empty.className = "panel-subtitle";
    empty.textContent = "No direct inputs on this tab.";
    inputColumn.appendChild(empty);
  } else {
    inputGroups.forEach((group, idx) =>
      inputColumn.appendChild(createFieldGroup(group.name, group.items, idx === 0))
    );
  }

  const outputColumn = document.createElement("div");
  outputColumn.className = "panel-column";
  outputColumn.innerHTML = `<h3 class="column-header">Calculated Outputs</h3>`;
  const outputGroups = groupFields(sheetView.outputs);
  outputGroups.forEach((group, idx) =>
    outputColumn.appendChild(createFieldGroup(group.name, group.items, idx === 0))
  );

  layout.appendChild(inputColumn);
  layout.appendChild(outputColumn);
  panel.appendChild(layout);

  if (index === 0) panel.classList.add("active");

  search.addEventListener("input", () => {
    const query = search.value.trim().toLowerCase();
    const rows = panel.querySelectorAll(".field-row");
    rows.forEach((row) => {
      if (!query) {
        row.classList.remove("hidden");
        return;
      }
      const hay = row.dataset.search || "";
      row.classList.toggle("hidden", !hay.includes(query));
    });

    panel.querySelectorAll(".field-group").forEach((group) => {
      const anyVisible = group.querySelector(".field-row:not(.hidden)") !== null;
      group.classList.toggle("hidden", !anyVisible);
    });
  });

  return panel;
}

function buildRowFormRows(sheetView) {
  const byRow = new Map();
  const allFields = [...sheetView.inputs, ...sheetView.outputs];
  for (const field of allFields) {
    if (!byRow.has(field.row)) byRow.set(field.row, []);
    byRow.get(field.row).push(field);
  }

  const rows = [];
  for (const [rowNumber, fields] of byRow.entries()) {
    fields.sort((a, b) => a.col - b.col);
    const labels = [];
    for (const field of fields) {
      if (!field.label) continue;
      if (field.label.startsWith(`${sheetView.name} `)) continue;
      if (!labels.includes(field.label)) labels.push(field.label);
    }
    rows.push({
      rowNumber,
      fields,
      title: labels[0] || `Row ${rowNumber}`,
      searchText: `${labels.join(" ")} ${fields.map((f) => `${f.address} ${f.key}`).join(" ")}`.toLowerCase(),
    });
  }
  rows.sort((a, b) => a.rowNumber - b.rowNumber);
  return rows;
}

function createRowFormField(field) {
  const wrap = document.createElement("div");
  wrap.className = `rowform-field ${field.is_input ? "is-input" : "is-output"}`;
  const isRequired = field.is_input && requiredInputKeys.has(field.key);
  if (isRequired) wrap.classList.add("required-field");

  const label = document.createElement("label");
  label.className = "rowform-field-label";
  if (isRequired) label.classList.add("required-label");
  label.textContent = field.label.startsWith(`${field.key.split("!")[0]} `)
    ? field.address
    : `${field.label} (${field.address})`;
  wrap.appendChild(label);

  if (field.is_input) {
    const input = document.createElement("input");
    input.className = "field-input";
    input.type = field.value_type === "number" ? "number" : "text";
    input.step = "any";
    if (PERCENT_WHOLE_INPUT_KEYS.has(field.key)) {
      input.value = formatPercentInput(getInputDefaultValue(field.key, 0));
      input.placeholder = "Enter percent";
      input.dataset.percentInput = "1";
    } else {
      input.value = formatValue(getInputDefaultValue(field.key, ""));
    }
    wrap.appendChild(input);
    wireInputElement(input, field.key, isRequired);
  } else {
    const output = document.createElement("div");
    output.className = "field-output";
    const text = formatValue(field.value);
    output.textContent = text;
    if (typeof text === "string" && text.startsWith("#")) output.classList.add("error");
    wrap.appendChild(output);
    registerFormulaElement(field.key, output);
  }

  return wrap;
}

function renderRowFormSheetPanel(sheetView, index) {
  const panel = document.createElement("section");
  panel.className = "sheet-panel rowform-panel";
  panel.dataset.sheet = sheetView.name;
  if (index === 0) panel.classList.add("active");

  const header = document.createElement("div");
  header.className = "panel-header";
  const titleWrap = document.createElement("div");
  titleWrap.className = "panel-title";
  const h2 = document.createElement("h2");
  h2.textContent = sheetView.name;
  const subtitle = document.createElement("p");
  subtitle.className = "panel-subtitle";
  subtitle.textContent = `${sheetView.inputs.length} inputs · ${sheetView.outputs.length} calculated outputs`;
  titleWrap.appendChild(h2);
  titleWrap.appendChild(subtitle);
  header.appendChild(titleWrap);

  const search = document.createElement("input");
  search.className = "panel-search";
  search.type = "search";
  search.placeholder = "Search rows by label or cell...";
  header.appendChild(search);
  panel.appendChild(header);

  if (sheetView.metrics.length > 0) {
    const metricWrap = document.createElement("div");
    metricWrap.className = "sheet-quick-metrics";
    for (const metricField of sheetView.metrics) {
      const { card, valueEl } = createMetricCard(
        metricField.label,
        formatCompactValue(metricField.value)
      );
      valueEl.classList.add("metric-live");
      registerFormulaElement(metricField.key, valueEl);
      metricWrap.appendChild(card);
    }
    panel.appendChild(metricWrap);
  }

  const rows = buildRowFormRows(sheetView);
  const list = document.createElement("div");
  list.className = "rowform-list";
  for (const row of rows) {
    const card = document.createElement("article");
    card.className = "rowform-row";
    card.dataset.search = row.searchText;

    const head = document.createElement("div");
    head.className = "rowform-row-head";
    const title = document.createElement("h3");
    title.className = "rowform-row-title";
    title.textContent = row.title;
    const meta = document.createElement("p");
    meta.className = "rowform-row-meta";
    meta.textContent = `Row ${row.rowNumber} · ${row.fields.length} fields`;
    head.appendChild(title);
    head.appendChild(meta);
    card.appendChild(head);

    const fieldGrid = document.createElement("div");
    fieldGrid.className = "rowform-fields";
    for (const field of row.fields) {
      fieldGrid.appendChild(createRowFormField(field));
    }
    card.appendChild(fieldGrid);
    list.appendChild(card);
  }

  panel.appendChild(list);

  search.addEventListener("input", () => {
    const query = search.value.trim().toLowerCase();
    list.querySelectorAll(".rowform-row").forEach((rowCard) => {
      if (!query) {
        rowCard.classList.remove("hidden");
        return;
      }
      const hay = rowCard.dataset.search || "";
      rowCard.classList.toggle("hidden", !hay.includes(query));
    });
  });

  return panel;
}

function renderRentRollPanel(index) {
  if (!rentRollState) rentRollState = buildInitialRentRollState();
  const panel = document.createElement("section");
  panel.className = "sheet-panel rentroll-panel";
  panel.dataset.sheet = "Rent Roll";
  if (index === 0) panel.classList.add("active");

  const header = document.createElement("div");
  header.className = "panel-header";
  const titleWrap = document.createElement("div");
  titleWrap.className = "panel-title";
  const h2 = document.createElement("h2");
  h2.textContent = "Rent Roll";
  const subtitle = document.createElement("p");
  subtitle.className = "panel-subtitle";
  subtitle.textContent = `${rentRollState.unit_count} units configured`;
  titleWrap.appendChild(h2);
  titleWrap.appendChild(subtitle);
  header.appendChild(titleWrap);
  panel.appendChild(header);

  const config = document.createElement("div");
  config.className = "rentroll-config";
  config.innerHTML = `
    <div class="rentroll-config-field">
      <label>Property Name</label>
      <input type="text" class="field-input is-required" id="rr-property-name" />
    </div>
    <div class="rentroll-config-field">
      <label>Property Address</label>
      <input type="text" class="field-input is-required" id="rr-property-address" />
    </div>
    <div class="rentroll-config-field rentroll-config-units">
      <label>Number Of Units</label>
      <input type="number" min="1" step="1" class="field-input is-required" id="rr-unit-count" />
    </div>
  `;
  panel.appendChild(config);

  const gridWrap = document.createElement("div");
  gridWrap.className = "rentroll-grid-wrap";
  const table = document.createElement("table");
  table.className = "rentroll-grid";
  table.innerHTML = `
    <thead>
      <tr>
        <th>#</th>
        <th>Tenant Name</th>
        <th>Unit</th>
        <th>Regular Rent</th>
        <th>Utilities</th>
        <th>Parking</th>
        <th>Pet Fee</th>
        <th>Total Rent</th>
        <th>Projected Rent</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;
  gridWrap.appendChild(table);
  panel.appendChild(gridWrap);

  const summary = document.createElement("div");
  summary.className = "rentroll-summary";
  const summaryCards = [
    { label: "Total Utilities", key: "Rent Roll!E15", totalKey: "utilities", formula: true },
    { label: "Total Parking", key: "rr-total-parking", totalKey: "parking", formula: false },
    { label: "Total Pet Fee", key: "rr-total-pet", totalKey: "pet_fee", formula: false },
    { label: "Total Rent", key: "Rent Roll!I15", totalKey: "total_rent", formula: true },
    { label: "Projected Rent Total", key: "Rent Roll!J15", totalKey: "projected_rent", formula: true },
  ];
  const summaryValueEls = {};
  for (const item of summaryCards) {
    const { card, valueEl } = createMetricCard(item.label, "0");
    card.classList.add("metric-card-money");
    summary.appendChild(card);
    summaryValueEls[item.totalKey] = valueEl;
    if (item.formula) registerFormulaElement(item.key, valueEl);
  }
  panel.appendChild(summary);

  const tbody = table.querySelector("tbody");
  const nameInput = config.querySelector("#rr-property-name");
  const addressInput = config.querySelector("#rr-property-address");
  const unitsInput = config.querySelector("#rr-unit-count");

  function recalcRentRollTotals() {
    const totals = {
      utilities: 0,
      parking: 0,
      pet_fee: 0,
      total_rent: 0,
      projected_rent: 0,
    };

    const units = Array.isArray(rentRollState.units) ? rentRollState.units : [];
    for (let i = 0; i < units.length; i += 1) {
      const row = units[i];
      const regularRent = toNumber(row.regular_rent, 0);
      const utilities = toNumber(row.utilities, 0);
      const parking = toNumber(row.parking, 0);
      const petFee = toNumber(row.pet_fee, 0);
      const projectedRent = toNumber(row.projected_rent, 0);
      row.total_rent = regularRent + utilities + parking + petFee;

      totals.utilities += utilities;
      totals.parking += parking;
      totals.pet_fee += petFee;
      totals.total_rent += row.total_rent;
      totals.projected_rent += projectedRent;
    }

    rentRollState.totals = totals;
  }

  function refreshSummary() {
    recalcRentRollTotals();
    syncValuationRentRollAutofillInputs();
    subtitle.textContent = `${rentRollState.unit_count} units configured`;
    summaryValueEls.utilities.textContent = formatMoney(rentRollState.totals.utilities);
    summaryValueEls.parking.textContent = formatMoney(rentRollState.totals.parking);
    summaryValueEls.pet_fee.textContent = formatMoney(rentRollState.totals.pet_fee);
    summaryValueEls.total_rent.textContent = formatMoney(rentRollState.totals.total_rent);
    summaryValueEls.projected_rent.textContent = formatMoney(rentRollState.totals.projected_rent);
  }

  function renderRows() {
    tbody.innerHTML = "";
    rentRollState.units.forEach((row, idx) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td><input type="text" class="field-input rr-cell" data-field="tenant_name" /></td>
        <td><input type="text" class="field-input rr-cell" data-field="unit" /></td>
        <td><input type="number" step="any" class="field-input rr-cell" data-field="regular_rent" /></td>
        <td><input type="number" step="any" class="field-input rr-cell" data-field="utilities" /></td>
        <td><input type="number" step="any" class="field-input rr-cell" data-field="parking" /></td>
        <td><input type="number" step="any" class="field-input rr-cell" data-field="pet_fee" /></td>
        <td><div class="field-output rr-total-cell"></div></td>
        <td><input type="number" step="any" class="field-input rr-cell" data-field="projected_rent" /></td>
      `;

      const totalCell = tr.querySelector(".rr-total-cell");
      totalCell.textContent = formatMoney(row.total_rent);

      tr.querySelectorAll(".rr-cell").forEach((input) => {
        const field = input.dataset.field;
        const isRequired = RENT_ROLL_REQUIRED_ROW_FIELDS.includes(field);
        if (isRequired) {
          input.classList.add("is-required");
          input.closest("td")?.classList.add("required-field");
        }
        input.value = ["tenant_name", "unit"].includes(field)
          ? rentRollState.units[idx][field] || ""
          : String(rentRollState.units[idx][field] ?? "");

        input.addEventListener("input", () => {
          const targetRow = rentRollState.units[idx];
          if (!targetRow) return;
          if (["tenant_name", "unit"].includes(field)) {
            targetRow[field] = String(input.value || "").trim();
          } else {
            targetRow[field] = String(input.value ?? "").trim();
          }
          targetRow.total_rent =
            toNumber(targetRow.regular_rent, 0) +
            toNumber(targetRow.utilities, 0) +
            toNumber(targetRow.parking, 0) +
            toNumber(targetRow.pet_fee, 0);
          totalCell.textContent = formatMoney(targetRow.total_rent);
          refreshSummary();
          updateValidationState();
        });
      });

      tbody.appendChild(tr);
    });
    refreshSummary();
  }

  nameInput.value = rentRollState.property_name;
  addressInput.value = rentRollState.property_address;
  unitsInput.value = rentRollState.unit_count > 0 ? String(rentRollState.unit_count) : "";
  config.querySelectorAll(".rentroll-config-field").forEach((field) =>
    field.classList.add("required-field")
  );
  config.querySelectorAll("label").forEach((label) => label.classList.add("required-label"));

  nameInput.addEventListener("input", () => {
    rentRollState.property_name = String(nameInput.value || "").trim();
    updateValidationState();
  });

  addressInput.addEventListener("input", () => {
    rentRollState.property_address = String(addressInput.value || "").trim();
    updateValidationState();
  });

  function applyUnitCountChange() {
    const raw = String(unitsInput.value || "").trim();
    const next = raw ? toPositiveInt(raw, 1) : 0;
    rentRollState.unit_count = next;
    rentRollState = normalizeRentRollState(rentRollState);
    unitsInput.value = raw ? String(rentRollState.unit_count) : "";
    renderRows();
    updateValidationState();
  }

  unitsInput.addEventListener("change", applyUnitCountChange);
  unitsInput.addEventListener("blur", applyUnitCountChange);
  unitsInput.addEventListener("input", () => updateValidationState());

  renderRows();
  return panel;
}

function renderAdminMortgagePanel(payload) {
  const tabId = "__admin_mortgage";

  const tabBtn = document.createElement("button");
  tabBtn.className = "tab-btn admin-tab";
  tabBtn.dataset.sheet = tabId;
  tabBtn.textContent = "Admin: Mortgage Logic";
  tabBtn.addEventListener("click", () => activateSheet(tabId));
  tabsEl.appendChild(tabBtn);

  const panel = document.createElement("section");
  panel.className = "sheet-panel admin-panel";
  panel.dataset.sheet = tabId;

  const header = document.createElement("div");
  header.className = "panel-header";
  const titleWrap = document.createElement("div");
  titleWrap.className = "panel-title";
  const h2 = document.createElement("h2");
  h2.textContent = "Mortgage Formula Reference (Admin)";
  const subtitle = document.createElement("p");
  subtitle.className = "panel-subtitle";
  subtitle.textContent =
    "Hidden from analyst tabs. Edit formulas here to adjust mortgage and amortization logic used by other tabs.";
  titleWrap.appendChild(h2);
  titleWrap.appendChild(subtitle);
  header.appendChild(titleWrap);

  const search = document.createElement("input");
  search.className = "panel-search";
  search.type = "search";
  search.placeholder = "Search mortgage cell or formula...";
  header.appendChild(search);
  panel.appendChild(header);

  const metrics = document.createElement("div");
  metrics.className = "sheet-quick-metrics";
  const totalCard = createMetricCard("Mortgage Formulas", String(payload.total_records || 0));
  const overrideCard = createMetricCard("Active Overrides", String(payload.overrides_count || 0));
  const pathCard = createMetricCard("Override File", payload.overrides_path || "N/A");
  pathCard.card.classList.add("metric-wide");
  metrics.appendChild(totalCard.card);
  metrics.appendChild(overrideCard.card);
  metrics.appendChild(pathCard.card);
  panel.appendChild(metrics);

  const controls = document.createElement("div");
  controls.className = "admin-controls";
  const saveBtn = document.createElement("button");
  saveBtn.className = "btn btn-primary";
  saveBtn.textContent = "Save Mortgage Overrides";
  const reloadBtn = document.createElement("button");
  reloadBtn.className = "btn";
  reloadBtn.textContent = "Reload From Server";
  controls.appendChild(saveBtn);
  controls.appendChild(reloadBtn);
  panel.appendChild(controls);

  const listWrap = document.createElement("div");
  listWrap.className = "admin-list";
  adminFormulaInputs.clear();

  for (const record of payload.records || []) {
    const row = document.createElement("div");
    row.className = "admin-row";
    if (record.is_overridden) row.classList.add("override");
    row.dataset.search = `${record.key} ${record.address} ${record.current_formula}`.toLowerCase();

    const labelWrap = document.createElement("div");
    labelWrap.className = "field-label";
    const main = document.createElement("div");
    main.className = "field-label-main";
    main.textContent = record.key;
    const meta = document.createElement("div");
    meta.className = "field-label-meta";
    meta.textContent = `Default: ${record.default_formula}`;
    labelWrap.appendChild(main);
    labelWrap.appendChild(meta);
    row.appendChild(labelWrap);

    const input = document.createElement("input");
    input.type = "text";
    input.className = "field-input admin-formula-input";
    input.value = record.current_formula;
    input.dataset.key = record.key;
    input.dataset.defaultFormula = record.default_formula;
    input.dataset.currentFormula = record.current_formula;
    input.addEventListener("input", () => {
      const changed = input.value.trim() !== (input.dataset.currentFormula || "");
      row.classList.toggle("changed", changed);
    });
    row.appendChild(input);
    adminFormulaInputs.set(record.key, input);

    listWrap.appendChild(row);
  }
  panel.appendChild(listWrap);
  panelsEl.appendChild(panel);

  search.addEventListener("input", () => {
    const q = search.value.trim().toLowerCase();
    const rows = listWrap.querySelectorAll(".admin-row");
    rows.forEach((row) => {
      if (!q) {
        row.classList.remove("hidden");
        return;
      }
      row.classList.toggle("hidden", !(row.dataset.search || "").includes(q));
    });
  });

  reloadBtn.addEventListener("click", async () => {
    await loadAdminMortgageData(true);
  });

  saveBtn.addEventListener("click", async () => {
    saveBtn.disabled = true;
    setStatus("Saving mortgage overrides...");
    try {
      const overrides = {};
      for (const [key, input] of adminFormulaInputs.entries()) {
        const formula = input.value.trim();
        const defaultFormula = input.dataset.defaultFormula || "";
        if (formula && formula !== defaultFormula) overrides[key] = formula;
      }

      const response = await fetch(adminMortgageOverrideUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ overrides }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Failed to save overrides");
      setStatus(
        `Mortgage overrides saved (${data.overrides_count} active). Refreshing admin catalog...`
      );
      await loadAdminMortgageData(true);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Error: ${message}`);
    } finally {
      saveBtn.disabled = false;
    }
  });
}

async function loadAdminMortgageData(preserveTab = false) {
  if (!isAdminMode) return;
  const previousActive = preserveTab
    ? document.querySelector(".tab-btn.active")?.dataset.sheet || ""
    : "";

  const response = await fetch(adminMortgageUrl);
  if (!response.ok) throw new Error("Failed to load mortgage admin data");
  const payload = await response.json();

  const oldTab = tabsEl.querySelector('.tab-btn[data-sheet="__admin_mortgage"]');
  const oldPanel = panelsEl.querySelector('.sheet-panel[data-sheet="__admin_mortgage"]');
  if (oldTab) oldTab.remove();
  if (oldPanel) oldPanel.remove();

  renderAdminMortgagePanel(payload);
  if (preserveTab && previousActive === "__admin_mortgage") {
    activateSheet("__admin_mortgage");
  }
}

function renderWorkbook(model) {
  tabsEl.innerHTML = "";
  panelsEl.innerHTML = "";
  inputElements.clear();
  formulaElements.clear();
  derivedMetricElements.clear();
  inputDisplayElements.clear();
  pdfReportSummaryEl = null;
  sensitivityView = null;
  sensitivityState = {
    rentChangePct: 0,
    interestRateChangeBps: 0,
    vacancyPct: null,
    purchasePriceOverride: null,
    gstRebate: "baseline",
    scenarios: [],
  };
  defaultInputValues.clear();
  lastValidationMissingCount = null;
  hasSuccessfulCalculation = false;
  latestFormulaValues = buildInitialFormulaSnapshot(model);
  initializeRequiredInputKeys(model);

  renderGlobalMetrics(model);
  for (const entry of model.input_cells) {
    defaultInputValues.set(entry.key, "");
  }

  rentRollState = buildInitialRentRollState();
  rentRollDefaultState = deepClone(rentRollState);

  const sheetViews = model.sheets.map((sheet) => buildSheetViewModel(sheet));

  let renderIndex = 0;
  for (const sheetView of sheetViews) {
    const tabBtn = document.createElement("button");
    tabBtn.className = "tab-btn";
    tabBtn.dataset.sheet = sheetView.name;
    tabBtn.textContent = sheetView.name;
    tabBtn.addEventListener("click", () => activateSheet(sheetView.name));
    if (renderIndex === 0) tabBtn.classList.add("active");
    tabsEl.appendChild(tabBtn);

    const panel =
      sheetView.name === "Rent Roll"
        ? renderRentRollPanel(renderIndex)
        : sheetView.name === "Valuation"
          ? renderValuationPanel(sheetView, renderIndex)
        : sheetView.name === "Returns"
          ? renderReturnsPanel(sheetView, renderIndex)
        : sheetView.name === "Sensitivity Analysis"
          ? renderSensitivityPanel(sheetView, renderIndex)
        : ROW_FORM_SHEETS.has(sheetView.name)
          ? renderRowFormSheetPanel(sheetView, renderIndex)
          : renderGenericSheetPanel(sheetView, renderIndex);
    panelsEl.appendChild(panel);
    renderIndex += 1;
  }

  const reiTabBtn = document.createElement("button");
  reiTabBtn.className = "tab-btn";
  reiTabBtn.dataset.sheet = "REI Ratios";
  reiTabBtn.textContent = "REI Ratios";
  reiTabBtn.addEventListener("click", () => activateSheet("REI Ratios"));
  tabsEl.appendChild(reiTabBtn);
  panelsEl.appendChild(renderReiRatiosPanel());
  renderIndex += 1;

  const pdfTabBtn = document.createElement("button");
  pdfTabBtn.className = "tab-btn";
  pdfTabBtn.dataset.sheet = PDF_REPORT_SHEET;
  pdfTabBtn.textContent = "PDF Report";
  pdfTabBtn.addEventListener("click", () => activateSheet(PDF_REPORT_SHEET));
  tabsEl.appendChild(pdfTabBtn);
  panelsEl.appendChild(renderPdfReportPanel(renderIndex));

  updateDerivedMetricElements();
  updatePdfReportSummary();
  updateValidationState();
}

function collectInputs() {
  const inputs = {};
  for (const [key, element] of inputElements.entries()) {
    inputs[key] = normalizeInputForSubmit(key, element.value);
  }
  return inputs;
}

function validateWorkbookRequiredInputs() {
  let missing = 0;
  for (const key of requiredInputKeys) {
    const input = inputElements.get(key);
    if (!input) continue;
    const isRequired = isConditionallyRequiredInput(key);
    input.classList.toggle("is-required", isRequired);
    input.dataset.required = isRequired ? "1" : "0";
    const parent = input.closest(
      ".field-row, .rowform-field, .valuation-assumption-row, .valuation-cell, .returns-cell, td"
    );
    if (parent) parent.classList.toggle("required-field", isRequired);
    if (!isRequired) {
      setInputMissingState(input, false);
      continue;
    }
    const isMissing = !isPresentValue(input.value);
    setInputMissingState(input, isMissing);
    if (isMissing) missing += 1;
  }
  return missing;
}

function validateRentRollRequiredInputs() {
  const panel = panelsEl.querySelector('.rentroll-panel[data-sheet="Rent Roll"]');
  if (!panel) return 0;

  let missing = 0;
  for (const selector of RENT_ROLL_CONFIG_REQUIRED_IDS) {
    const input = panel.querySelector(selector);
    if (!input) continue;
    let isMissing = !isPresentValue(input.value);
    if (!isMissing && input.id === "rr-unit-count") {
      const value = toNonNegativeInt(input.value, 0);
      isMissing = value < 1;
    }
    setInputMissingState(input, isMissing);
    if (isMissing) missing += 1;
  }

  const rowInputs = panel.querySelectorAll(".rr-cell");
  rowInputs.forEach((input) => {
    const field = input.dataset.field || "";
    if (!RENT_ROLL_REQUIRED_ROW_FIELDS.includes(field)) return;
    const isMissing = !isPresentValue(input.value);
    setInputMissingState(input, isMissing);
    if (isMissing) missing += 1;
  });

  return missing;
}

function updateValidationState() {
  syncValuationRentRollAutofillInputs();
  const workbookMissing = validateWorkbookRequiredInputs();
  const rentRollMissing = validateRentRollRequiredInputs();
  const totalMissing = workbookMissing + rentRollMissing;
  calculateBtn.disabled = totalMissing > 0;

  if (totalMissing > 0) {
    setStatus(`Fill ${totalMissing} required fields to enable Run Analysis.`);
  } else if ((lastValidationMissingCount ?? 0) > 0) {
    setStatus("All required fields are complete. Run Analysis is enabled.");
  }
  lastValidationMissingCount = totalMissing;
}

async function resetInputs() {
  if (!workbookModel) return;
  const active = document.querySelector(".tab-btn.active")?.dataset.sheet || "";
  renderWorkbook(workbookModel);
  if (isAdminMode) await loadAdminMortgageData(false);
  if (active) {
    const exists = document.querySelector(`.tab-btn[data-sheet="${active}"]`);
    if (exists) activateSheet(active);
  }
  if (adminSummaryValueEl) adminSummaryValueEl.textContent = isAdminMode ? "Admin" : "Analyst";
  if (globalLastRunValueEl) globalLastRunValueEl.textContent = "Not run";
  setStatus("Inputs reset.");
  updateValidationState();
}

function renderFormulaUpdates(formulaValues) {
  for (const [key, value] of Object.entries(formulaValues || {})) {
    latestFormulaValues[key] = value;
    const elements = formulaElements.get(key) || [];
    for (const el of elements) {
      let formatted;
      if (el.dataset.format === "percent") {
        formatted = formatPercentValue(value);
      } else if (el.dataset.format === "money") {
        formatted = formatMoney(value);
      } else if (el.dataset.format === "number") {
        formatted = formatNumberValue(value);
      } else if (el.dataset.format === "compact") {
        formatted = formatCompactValue(value);
      } else if (el.classList.contains("metric-value") || el.classList.contains("metric-live")) {
        formatted = formatCompactValue(value);
      } else {
        formatted = formatValue(value);
      }
      el.textContent = formatted;
      el.classList.toggle("error", typeof formatted === "string" && formatted.startsWith("#"));
    }
  }
  updateSensitivityPanel();
  updateDerivedMetricElements();
  updatePdfReportSummary();
}

async function calculate() {
  if (calculateBtn.disabled) return;
  calculateBtn.disabled = true;
  setStatus("Running analysis...");
  try {
    rentRollState = normalizeRentRollState(rentRollState || rentRollDefaultState || buildInitialRentRollState());
    const response = await fetch(calculateUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        inputs: collectInputs(),
        rent_roll: rentRollPayload(),
      }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Calculation failed");

    hasSuccessfulCalculation = true;
    renderFormulaUpdates(data.formula_values);

    const runText = `Updated ${new Date().toLocaleTimeString()}`;
    setStatus(runText);
    if (globalLastRunValueEl) globalLastRunValueEl.textContent = runText;
    activateSheet(PDF_REPORT_SHEET);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`Error: ${message}`);
  } finally {
    updateValidationState();
  }
}

async function bootstrap() {
  setStatus("Loading underwriting model...");
  const response = await fetch(modelUrl);
  if (!response.ok) {
    setStatus("Failed to load model metadata.");
    return;
  }
  workbookModel = await response.json();
  renderWorkbook(workbookModel);

  if (isAdminMode) {
    await loadAdminMortgageData(false);
    setStatus("Loaded analyst tabs + admin mortgage logic panel.");
  } else {
    const hidden = Array.isArray(workbookModel.hidden_sheets)
      ? workbookModel.hidden_sheets.join(", ")
      : "Mortgage";
    setStatus(`Loaded ${workbookModel.sheets.length} tabs. Hidden internal sheets: ${hidden}.`);
  }
  updateValidationState();
}

calculateBtn.addEventListener("click", calculate);
resetBtn.addEventListener("click", () => {
  resetInputs().catch((error) => {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`Error: ${message}`);
  });
});

bootstrap().catch((error) => {
  const message = error instanceof Error ? error.message : String(error);
  setStatus(`Error: ${message}`);
});
