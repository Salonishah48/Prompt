# Florida Sales Tax App (DR-15 Auto-Filer)

A Python application that calculates Florida state and discretionary sales
tax and produces a filled **DR-15 Sales and Use Tax Return**, based on the
rules in:

- **DR-15N (R. 10/25)** — the official FDOR instructions
- **Florida Sales Tax Compliance Guide** — practitioner-level walkthrough

You drop Excel tax exports into a folder, run one command, and get back:

1. A **PDF summary** of the filled DR-15 return (professional, ready to share)
2. An **Excel workbook** with the same data plus the full transaction audit trail
3. A **JSON payload** for downstream use

## Scope of this build

- **Input format:** Excel (`.xlsx`) or CSV/TSV files
- **Supported schemas:**
  - **Shopify tax export** — the full 38-column format with `Gross sales on line items`, `Taxable amount`, `Tax amount`, destination address columns, etc. This is what Shopify's "Sales by tax" report produces.
  - **Simple schema** — `date`, `amount`, `exempt`, `delivery_county`, `is_single_tpp_item` (for non-Shopify users)
- **Output:** PDF + Excel + JSON
- **DR-15 lines enabled:** Line A (regular sales/services) only
- **Shopify reconciliation:** Automatically compares Shopify's collected tax against what Florida actually requires and flags any under- or over-collection on the PDF.
- Other DR-15 lines (B, D, E) are wired in and can be turned on later
  by editing `rules/business_config.json`

---

## Project structure

```
florida_sales_tax_app/
├── rules/
│   ├── tax_rules.json          ← All FL tax rates, surtax by county, etc.
│   └── business_config.json    ← YOUR business info (edit ONCE)
├── sales_data/                 ← Drop your Shopify export(s) here
│   └── shopify_july_2025.xlsx  ← Sample showing expected column layout
├── output/                     ← PDF + Excel + JSON land here after a run
├── tax_engine.py               ← Core DR-15 calculation engine
├── data_loader.py              ← Reads your sales files, filters by period
├── output_formatter.py         ← Renders PDF + Excel
├── run.py                      ← Main entry point
└── README.md                   ← You are here
```

---

## One-time setup

### 1. Install dependencies

```bash
pip install openpyxl reportlab
```

### 2. Fill in your business info

Edit **`rules/business_config.json`**:

```json
{
  "business_info": {
    "business_name": "Your Business Name LLC",
    "certificate_number": "78-8012345678-9",
    "fein": "12-3456789",
    "city": "Tampa",
    "county": "Hillsborough",
    ...
  },
  "filing_preferences": {
    "files_electronically": true,
    "pays_electronically": true
  }
}
```

The `county` value drives your default discretionary sales surtax rate.
Use the county name exactly as it appears in `rules/tax_rules.json`
(e.g., `Miami_Dade`, `Palm_Beach`, `St_Johns` — underscores for spaces).

---

## Workflow

### 1. Drop your sales data into `sales_data/`

Supported formats: **Excel `.xlsx`**, **CSV**, or **TSV**.

The file can cover any period — a day, a week, a month, a quarter, a year.
Drop as many files as you like; the app reads them all and filters by
reporting period. You don't need to update daily — just drop a fresh export
whenever you're ready to file (typically once per month or quarter).

The app auto-detects the schema. Two layouts are supported:

#### Option A — Shopify tax export (recommended)

Use Shopify's built-in tax report in Shopify Admin → Analytics → Reports →
Finances → "Sales by tax". Export to Excel and drop it in `sales_data/`.
The app reads the full 38-column format as-is:

```
Is filed by channel | Channel | Order number | Line item ID | Sale date |
Gross sales on line items | Discounts | Returns | Net sales on line items |
Shipping | Exempt amount | Non-taxable amount | Non-taxed amount |
Taxable amount | Tax rate | Tax amount | Tax jurisdiction type |
Tax jurisdiction | Tax county | Tax jurisdiction code |
Destination country | Destination state | Destination city |
Destination address | Destination zip | Billing ... | Origin ... |
Product category code | Tax exemptions | Shopify reference ID
```

Delivery county is inferred from `Tax county` when populated, otherwise from
`Destination city` via a built-in Florida city-to-county lookup (extend in
`data_loader.py` as needed).

See `sales_data/shopify_july_2025.xlsx` for a working example.

#### Option B — Simple schema (non-Shopify users)

Use this for manual entry or if you're not on Shopify. First row = header:

| Column                 | Required | Notes                                                             |
| ---------------------- | -------- | ----------------------------------------------------------------- |
| `date`                 | yes      | `YYYY-MM-DD`, `MM/DD/YYYY`, or native Excel date cell             |
| `amount`               | yes      | Dollar amount. `$` signs and commas OK.                           |
| `exempt`               | no       | `true` / `false`. Default `false`.                                |
| `delivery_county`      | no       | County the item was delivered to. Default = business county.      |
| `is_single_tpp_item`   | no       | `true` / `false`. Default `true`. Controls the $5,000 surtax cap. |

`is_single_tpp_item`:

- `true` — this is a single item of **tangible personal property**. Surtax
  applies only to the first $5,000.
- `false` — this is a **service**, **rental**, or **admission** (the $5,000 cap
  does NOT apply — surtax applies to the full amount).

### 2. Run the app

The simplest case — **auto-detect the period from your data**:

```bash
python run.py
```

The app scans every file in `sales_data/`, finds the earliest and latest
transaction, and reports on exactly that period. No flags needed.

**For a specific month only** (ignores rows outside the month):

```bash
python run.py --period 2025-07
```

**For a quarter or any custom date range:**

```bash
python run.py --from 2025-07-01 --to 2025-09-30
```

**Mark as late** (applies 10% or $50 min penalty):

```bash
python run.py --period 2025-07 --late
```

**With lawful deductions** (e.g. refunded tax on returned goods):

```bash
python run.py --period 2025-07 --lawful-deductions 125.00
```

**Generate a blank template** (no data):

```bash
python run.py --template
```

Full list of flags:

```
--period YYYY-MM              Single-month reporting period
--from YYYY-MM-DD             Start date of custom range (inclusive)
--to   YYYY-MM-DD             End date   of custom range (inclusive)
--late                        Mark return as late (applies penalty)
--lawful-deductions FLOAT     Line 6 amount
--est-tax-paid FLOAT          Line 8 amount
--est-tax-due FLOAT           Line 9 amount
--scholarship-credits FLOAT   Tax Credit Scholarship Motor Vehicle credits
--high-crime-credits FLOAT    Rural/Urban High Crime Area job credits
--other-credits FLOAT         Other authorized credits
--template                    Produce blank PDF/Excel templates and exit
```

### 3. Review outputs in `output/`

Each run produces three files:

- `DR15_YYYY-MM_filled.pdf` — polished PDF summary; first page mirrors the
  DR-15 layout, second page is the Discretionary Sales Surtax detail and
  footnotes, third page is the full supporting transaction list.
- `DR15_YYYY-MM_filled.xlsx` — Excel workbook with two sheets:
  - `DR-15 Return` — all filled lines and totals, formatted for review
  - `Transactions` — every source transaction with its origin file
- `DR15_YYYY-MM_filled.json` — machine-readable return object for downstream
  systems or archival.

---

## What the engine computes for you

### State tax and surtax

- **6% state sales tax** on taxable Line A amounts
- **Destination-based discretionary sales surtax** — pulled from
  `county_surtax_rates_2026` in `tax_rules.json`
- **$5,000 surtax cap** on single items of tangible personal property

### The back-of-form Discretionary Sales Surtax breakdown

The app computes **Lines 15(a) through 15(d)** automatically — this is where
the DR-15 trips most filers up:

- **15(a)** — portion of each single TPP item over $5,000 (surtax-exempt)
- **15(b)** — amounts delivered into non-surtax counties
- **15(c)** — amounts taxed at a surtax rate different from your home county
- **15(d)** — total discretionary sales surtax due (flows from Col 4)

### Collection allowance, penalties, and interest

- **Collection allowance**: 2.5% of first $1,200 tax due, **capped at $30** —
  applied automatically **only when** `files_electronically` AND
  `pays_electronically` are both `true` AND the return is on time.
- **Late penalty**: 10% of Line 10, minimum $50 (when `--late` is used).
- **Interest**: flagged with a warning; needs the current daily rate from
  [floridarevenue.com/taxes/rates](https://floridarevenue.com/taxes/rates)
  and the number of days late. The app does not guess.

### Built-in protections (per DR-15N)

- **Line 6 ≤ Line 5** — lawful deductions are capped to total tax due;
  excess triggers a warning ("carry to next return").
- **Line 8 ≤ Line 7** — credits are capped to net tax due.
- **Line 10 ≥ 0** — can never be negative.
- **Florida rounding rule** — tax computed to 3 decimal places per
  transaction, then rounded up to the next cent when the third decimal
  is greater than 4.

### Shopify reconciliation (Shopify inputs only)

When the app ingests a Shopify export, it computes **two tax numbers**:

1. **Tax collected by Shopify** — the sum of the `Tax amount` column from
   your export (what was actually charged to customers at checkout).
2. **Tax calculated by app** — what Florida law says is owed, computed
   from `Taxable amount` and the correct destination county rate, with the
   $5,000 TPP surtax cap applied.

The PDF's "Shopify Collection Reconciliation" section compares the two and
highlights:

- **Green (Match)** — collection is correct; file normally.
- **Red (Under-collection)** — Shopify didn't charge enough surtax; you
  still owe FDOR the full amount and absorb the shortfall. Usually means
  Shopify isn't configured for county surtax — fix in Shopify Admin →
  Settings → Taxes and duties → United States → Florida.
- **Blue (Over-collection)** — Shopify charged too much; you must remit the
  full amount collected (you can't keep the excess under Florida law).

Gaps smaller than $1.00 are treated as rounding and not flagged.

### Commercial rent (Line C)

Left in the output form for historical returns, but **always $0 for periods
on or after October 1, 2025** (per TIP 25A01-04, which repealed the
commercial rent tax).

---

## Updating tax rules

FDOR publishes updated county surtax rates (Form DR-15DSS) each year. To
update rates, just edit **`rules/tax_rules.json`** — no code changes:

```json
"county_surtax_rates_2026": {
  "Hillsborough": 0.015,
  "Miami_Dade":   0.01,
  "Orange":       0.005,
  ...
}
```

The same file holds state rates, vending divisors, collection-allowance
rules, penalty/interest rates, and economic-nexus thresholds — all editable
without touching Python.

---

## Updating your data

You don't need to update data daily. The typical workflow is:

1. **Before filing** (monthly or quarterly, depending on your filing
   frequency), log in to Shopify, go to **Analytics → Reports → Finances →
   Sales by tax**, and export the period you need to file.
2. Save the Excel file to `sales_data/` (replace any old files, or add new
   ones — the app reads all files in the folder).
3. Run `python run.py` — the app auto-detects the period from the data.
4. Review the PDF in `output/`, then file at
   [floridarevenue.com/taxes/eServices](https://floridarevenue.com/taxes/eServices).

### Optional automation

If you want the app to run on a schedule instead of on demand:

- **Windows** — Task Scheduler → Create Basic Task → point to `python.exe`
  with argument `run.py`, working directory set to the app folder.
- **Linux/macOS** — a cron job that runs monthly:
  ```cron
  # Run on the 2nd of every month at 9 AM
  0 9 2 * * cd /path/to/florida_sales_tax_app && /usr/bin/python3 run.py
  ```

Most small businesses just run it manually once per filing period.

---

## Important reminders

- Returns are **due the 1st** of the month after the reporting period and
  **late after the 20th**.
- **Electronic payments** must be initiated by **5 p.m. ET the business day
  before the 20th**.
- **Keep records for 3+ years** (DR-15N requirement).
- This app **prepares** your return; you still file it through FDOR's
  e-Services portal at
  [floridarevenue.com/taxes/eServices](https://floridarevenue.com/taxes/eServices).
