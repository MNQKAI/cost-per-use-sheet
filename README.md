# Cost Per Use

A Google-Sheets-ready and Excel-ready Cost Per Use calculator and purchase tracker workbook, along with the Python script that generates it.

## Quick Start (Use the Spreadsheet)

1. Download `Cost_Per_Use_Calculator.xlsx`.
2. Open it directly in Microsoft Excel, or import it into Google Sheets:
   - Open [Google Sheets](https://sheets.google.com) and create a **Blank spreadsheet**.
   - Click **File -> Import -> Upload** and select `Cost_Per_Use_Calculator.xlsx`.
   - Choose **Replace spreadsheet** and click **Import data**.
3. Edit the **blue** input cells; **white** and **emerald** cells auto-calculate.

## What's Included

- **`Quick Calc`**: Minimal single-item Cost Per Use calculator card for fast checks on desktop or mobile.
- **`Calculator & Compare`**: Executive KPI cards, side-by-side 3-option comparison (Budget vs. Mid-Tier vs. Buy-It-For-Life), break-even analysis, a Cost-per-Use column chart, and a Cumulative Replacement Cost crossover line chart.
- **`Master Tracker`**: 100-row purchase and durability tracker with a KPI strip, status/rating/tier pills, zebra striping, a Cost/Use color scale, and horizontal-only freeze panes.
- **`Insights & Charts`**: Live category, ownership status, and value-tier rollups synced with `Master Tracker`, paired with 4 native charts.

## Regenerating or Customizing the Workbook

Requires Python 3.10+ and `openpyxl`:

```bash
pip install -r requirements.txt
python build_sheet.py
```

## Documentation

See [FORMULA_GUIDE.md](FORMULA_GUIDE.md) for the full formula reference, tab-by-tab breakdown, and subscription logging instructions.

## Inspiration

The Darn Tough socks cost-per-use framing is inspired by Steph Ango's [*Buy wisely*](https://stephango.com/buy-wisely) essay.

## License

MIT License — see [LICENSE](LICENSE).
