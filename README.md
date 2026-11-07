# Amazon Search Term Report Analyzer

Turn the Sponsored Products **Search term report** into a prioritized action list: what to negate, what to harvest, and where the budget is leaking. One command, one Excel file out.

Built by [Mohammed Zohaib Ali] {www.linkedin.com/in/mohammed-zohaib-ali} from running Amazon PPC accounts.

## The problem

A search term report has thousands of rows. Most are 1-2 click terms that tell you nothing. A few are bleeding real money. A few are quietly winning and aren't targeted as exact match yet. Finding them by hand in a filtered spreadsheet takes hours, and most people end up using a blunt rule like "10 clicks, no orders = negative."

That rule is too aggressive when your conversion rate is around 2%: at that rate, you expect roughly 50 clicks per order, so 10 clicks without a sale is normal.

## How this tool decides

It works out your **break-even spend per order**: `average order value x target ACOS`. A search term that has spent that much with zero orders has had a fair chance and gets flagged.

| Tier | Rule |
|---|---|
| 1 - Negate now | Zero orders, spend at or above break-even spend |
| 2 - Likely negate | Zero orders, spend at or above 50% of break-even |
| 3 - Watch | Zero orders, 15+ clicks or 25% of break-even |
| Harvest | 2+ orders at or under target ACOS (1 order = "Test") |

All thresholds are adjustable.

## What you get

| Sheet | What it shows |
|---|---|
| Summary | Totals, settings used, size of the opportunity (live formulas) |
| Negatives | Campaign + search term pairs to negate, by priority, with "converts elsewhere" flag |
| Word Waste | Words and phrases that keep burning spend (negative phrase ideas) |
| Harvest | Converting terms to add as exact match, with a suggested bid |
| ASIN Targets | ASINs to add as product targets or negate (your own ASINs are separated out) |
| Campaigns | Spend, ACOS and wasted spend by campaign |
| Targeting Types | Auto vs Phrase vs Broad vs Exact vs Product targeting, plus auto sub-groups |

## How to use it (no coding background needed)

1. Install Python 3.9 or newer from python.org.
2. Download this repository (green **Code** button, then **Download ZIP**) and unzip it.
3. Open a terminal in the folder and run:
   ```
   pip install -r requirements.txt
   ```
4. In Amazon Ads, download your **Sponsored Products > Search term** report (Excel or CSV) for a period of 30 days or more. Put the file in the same folder.
5. Run:
   ```
   python search_term_analyzer.py your_report.xlsx
   ```
6. Open the Excel file it creates (`your_report_analysis.xlsx`).

## Options

```
python search_term_analyzer.py report.xlsx --target-acos 0.35 --aov 45 --out results.xlsx
```

| Option | Default | Meaning |
|---|---|---|
| `--target-acos` | 0.30 | Your target ACOS as a fraction |
| `--aov` | sales / orders | Average order value, if you want to override it |
| `--watch-clicks` | 15 | Clicks with no order before a term goes on the watch list |
| `--harvest-orders` | 2 | Orders needed for the "Harvest" tier |
| `--own-asins` | auto | Comma-separated ASINs you advertise (also detected from campaign names) |
| `--out` | `<report>_analysis.xlsx` | Output file |

## Limits you should know about

- **Attribution lag.** Amazon credits sales within 7 days of a click, so the last week of any report is incomplete. Recent zero-order terms may still convert.
- **Account-wide average order value.** If your products vary a lot in price or margin, set `--aov` or analyze one product at a time.
- **Sales include halo sales** from other SKUs, which makes ACOS look slightly better than it is.
- **Negatives are suggestions.** Check the "Converts Elsewhere" column first. A term can lose money in one campaign and win in another.
- **Rules of thumb, not statistics.** A tool can tell you where to look; it can't know your margins, seasonality, or strategy.

## Privacy

The tool runs entirely on your computer. Never commit real client reports or outputs to a public repository.

## License

MIT
