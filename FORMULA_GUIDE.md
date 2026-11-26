# Cost Per Use — Google Sheets Guide

## How to Import into Google Sheets
1. Open [Google Sheets](https://sheets.google.com) and create a **Blank spreadsheet** (or open an existing one).
2. Click **File → Import → Upload** and select `Cost_Per_Use_Calculator.xlsx`.
3. Choose **Replace spreadsheet** and click **Import data**.
4. All 4 tabs, formulas, dropdowns, KPI cards, pill badges, and 6 native charts import automatically.

> **Visual Legend:** Blue cells are editable **Inputs**. White and emerald cells are **Auto-Calculated**.

---

## Tab 1: `Quick Calc` (Minimal Single-Item)
A distraction-free, typography-only card for fast single-item checks on desktop or mobile.

### Inputs (`E6:E10`)
- `E6`: **Item** (text)
- `E7`: **Purchase Price** (`$`)
- `E8`: **Usage Frequency** (number)
- `E9`: **Frequency Unit** (`Day` / `Week` / `Month` / `Year`)
- `E10`: **Estimated Life (Years)** (number)

### Outputs
- **Cost per Use (`C13:E13`):** `=IF(OR(E7="",C17="",C17=0),"—",E7/C17)`
- **Total Uses (`C17`):** `=IF(OR(E8="",E9="",E10=""),"",IFERROR(IF(E9="Day",E8*365,IF(E9="Week",E8*52,IF(E9="Month",E8*12,IF(E9="Year",E8,""))))*E10,""))`
- **Cost / Year (`D17`):** `=IF(OR(E7="",E10="",E10=0),"",E7/E10)`
- **Cost / Month (`E17`):** `=IF(D17="","",D17/12)`

---

## Tab 2: `Calculator & Compare`
A clean, gridline-free dashboard with 5 top KPI cards, a side-by-side 3-option comparison table (pre-loaded with Steph Ango's Darn Tough socks example), and 2 live charts on the right:
- **Chart 1:** *Cost per Use Comparison ($)* — color-coded column chart comparing Option A, Option B, Option C, and Rental/Alternative.
- **Chart 2:** *Cumulative Cost Over Uses (Replacement Break-Even)* — dynamic 10-milestone line chart showing the exact crossover point where buying a durable option beats replacing a cheaper one.

### Core Formulas (Column `C` / Option A)
| Row | Metric | Formula |
| :--- | :--- | :--- |
| `22` | **Uses per Year** | `=IF(OR(C16="",C17=""),"",IF(C17="Day",C16*365,IF(C17="Week",C16*52,IF(C17="Month",C16*12,IF(C17="Year",C16,"")))))` |
| `23` | **Total Estimated Uses** | `=IF(OR(C22="",C18=""),"",C22*C18)` |
| `24` | **Total Ongoing Cost ($)** | `=IF(OR(C14="",C18=""),"",C14*C18)` |
| `25` | **Total Net Lifecycle Cost ($)** | `=IF(C13="","",C13+IF(C24="",0,C24)-IF(C15="",0,C15))` |
| `26` | **Cost per Use ($)** | `=IF(OR(C25="",C23="",C23=0),"",C25/C23)` |
| `27` | **Annualized Cost ($/yr)** | `=IF(OR(C25="",C18="",C18=0),"",C25/C18)` |
| `28` | **Monthly Cost ($/mo)** | `=IF(C27="","",C27/12)` |

### Decision & Break-Even Formulas
- **Lowest Cost per Use Winner (`C31:D31`):** `=IF(COUNT(C26:E26)=0,"",INDEX(C11:E11,1,MATCH(MIN(C26:E26),C26:E26,0)))`
- **Break-Even Uses vs. Option A (`D32`, `E32`):** `=IF(OR($C$26="",D26=""),"",IF(D26>=$C$26,"Not cheaper/use",ROUNDUP(D25/$C$26,0)))`
- **Savings per 100 Uses vs. Option A (`D33`, `E33`):** `=IF(OR($C$26="",D26=""),"",($C$26-D26)*100)`
- **Break-Even Uses vs. Renting (`C35:E35`):** `=IF(OR($C$34="",$C$34=0,C25=""),"",IF(C26>=$C$34,"Rent is cheaper",ROUNDUP(C25/$C$34,0)))`

---

## Tab 3: `Master Tracker`
100 pre-formatted rows (`Rows 9–108`) with horizontal-only freeze panes (Rows `1–8`), zebra striping, status/rating/tier color pills, a green-to-red Cost/Use heatmap, and 6 editable sample rows.

### Row Formulas (shown for Row `9`)
- **Uses / Year (`L9`):** `=IF(OR(B9="",H9="",I9=""),"",IF(I9="Day",H9*365,IF(I9="Week",H9*52,IF(I9="Month",H9*12,IF(I9="Year",H9,"")))))`
- **Total Uses (`M9`):** `=IF(OR(L9="",J9=""),"",L9*J9)`
- **Total Net Cost (`N9`):** `=IF(OR(B9="",E9=""),"",E9+IF(F9="",0,F9)*IF(J9="",0,J9)-IF(G9="",0,G9))`
- **Cost / Use (`O9`):** `=IF(OR(N9="",M9="",M9=0),"",N9/M9)`
- **Cost / Year (`P9`):** `=IF(OR(N9="",J9="",J9=0),"",N9/J9)`
- **Cost / Month (`Q9`):** `=IF(P9="","",P9/12)`
- **Value Tier (`R9`):** `=IF(O9="","",IF(O9<=0.25,"Exceptional",IF(O9<=1,"Great Value",IF(O9<=5,"Moderate","High $/Use"))))`

### Logging Subscriptions or Pure Services
1. Set **Purchase Price (`E`)** to `0`.
2. Set **Annual Ongoing (`F`)** to your annual cost (or type `=15*12` for a \$15/month subscription).
3. Set **Est. Resale (`G`)** to `0` and **Est. Life (`J`)** to `1`.
4. Enter **Usage Freq (`H`)** and **Freq Unit (`I`)**.

---

## Tab 4: `Insights & Charts`
Live portfolio rollups synced automatically from `Master Tracker` via `COUNTIF`, `SUMIF`, and `AVERAGEIF`, paired with 4 native charts:
1. **Total Net Cost by Category ($)** — Horizontal bar chart
2. **Average Cost per Use by Category ($)** — Horizontal bar chart
3. **Items by Cost-per-Use Value Tier** — Color-coded column chart
4. **Total Net Cost by Ownership Status ($)** — Color-coded column chart
