# Kraft Heinz · Microsoft 365 Copilot Executive Decision Sprint

A self-contained, browser-only hands-on lab (Level 100–200) that walks executives from
**signal → preparation → analysis → decision → communication → follow-through**
using Microsoft 365 Copilot.

**Live site:** https://brmoinea-microsoft.github.io/kraftheinz-copilot-executive-lab/

## What's in it

Eight activities, all navigable from the left rail or with the ← / → arrow keys:

| # | Activity | Focus |
|---|----------|-------|
| 1 | Welcome | Objectives and flow |
| 2 | Prompting pattern | Goal + Context + Source + Expectations |
| 3 | Catch up | Copilot Chat weekly signal scan |
| 4 | Prepare | Meeting pre-read grounded with Work IQ |
| 5 | Analyze | Copilot in Excel **or** the Analyst agent |
| 6 | Create | Decision brief in Copilot in Word |
| 7 | Follow through | Decision record + Outlook draft |
| 8 | Lead responsibly | Verification and accountability |

Every prompt has a one-click **Copy** button. Activity 5 ends with a common
"Decision Brief Handoff" that feeds Activity 6, so both analysis routes converge.

## Repository layout

```
index.html                          Entire site — no build step, no dependencies
assets/
  Demo_Weekly_Customer_Category_Performance.xlsx   Excel route practice data
  Demo_Consumer_Care_Feedback.csv                  Analyst route practice data
  executive_use_case_prompts.csv                   Executive prompt inventory
tools/
  generate_demo_data.py             Regenerates the two demo datasets
.nojekyll                           Serve files as-is on GitHub Pages
```

## Practice data

All practice data is **fictional and synthetic**, generated deterministically by
`tools/generate_demo_data.py`. It describes a fictitious packaged-food manufacturer,
**Lakeshore Foods Co.**, selling to fictitious retail customers (Contoso Markets,
Fabrikam Club, Northwind Grocers, Adventure Works Foods, Tailspin Convenience).

It is **not** Kraft Heinz data, financials, consumer records, or performance results.

The workbook contains four tabs:

- **Customer Performance** — 150 rows across 5 regions × 5 retail customers × 6 weeks
  (net sales, YoY, volume, price/mix, on-shelf availability, case fill rate, promo
  compliance, distribution points, trade spend, forecast accuracy, escalations).
- **Category Performance** — 138 rows across 5 category platforms and 23 brand groups
  (net sales, YoY, gross margin, volume, days of supply, sell-through, promo flag,
  eCommerce share, market share).
- **Regional Summary** — narrative opportunity/action per region, with the risk and
  evidence strings **computed from the generated rows** so an executive who verifies the
  numbers finds them consistent.
- **Notes and Assumptions** — field definitions and data-quality caveats.

Deliberate outliers are seeded (South Central underperforms; Boxed Macaroni, Canned Pasta
and Drink Mix Canister are weak) so the analysis exercises have something real to find.

To regenerate:

```powershell
python tools/generate_demo_data.py     # requires openpyxl
```

## Running locally

No build. Open `index.html`, or serve the folder so the `download` links behave the same
way they do in production:

```powershell
python -m http.server 8080
# http://127.0.0.1:8080
```

## Deploying

GitHub Pages is served from the `main` branch root. Push to `main` and Pages redeploys.

## Branding note

This lab uses a Kraft Heinz-inspired colour palette (red `#e4002b`, gold `#ffb81c`) and a
text wordmark only — no Kraft Heinz logo files are included. Swap in official brand assets
before customer delivery if required. Kraft Heinz names and marks are the property of
The Kraft Heinz Company and are used here solely to identify the audience for this
enablement session.

## Credits

Structure and activity flow adapted from the Microsoft 365 Copilot VIP hands-on lab by
[@sgreenlee11](https://github.com/sgreenlee11/m365-copilot-vip-hands-on), rebuilt and
re-scoped for a Kraft Heinz executive audience.
