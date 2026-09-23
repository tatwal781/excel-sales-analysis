# E-commerce Sales Analysis (Excel)

An end-to-end sales analysis built in **Microsoft Excel**, turning 520 raw order
records into a one-page summary dashboard of revenue drivers — by category,
region, channel, and month.

The focus is a realistic analyst workflow in Excel: aggregating raw transactions
with formulas, surfacing the key figures, and presenting them as charts a
business could act on.

## Summary dashboard

![Summary dashboard](images/summary-01.jpg)

## Key findings

- **Total revenue £251,715 across 520 orders**, with an average order value of ~£484.
- **Electronics dominates at 63% of revenue** (£157,854), driven by laptops and
  phones. Furniture is a distant second at 25%.
- **London is the top region** (£60,559), well ahead of North West (£38,484) and
  the Midlands (£34,982).
- **Mobile App is the leading sales channel** (£116,704), ahead of Website
  (£104,374), with Marketplace a minor third (£30,637).
- **Revenue is weighted to the second half of the year.** H2 brought in £155,209
  against £96,506 in H1, and the strongest months were July and October–November.

## What this demonstrates

| Skill | Where |
|-------|-------|
| Conditional aggregation | `SUMIFS` / `COUNTIFS` across category, region, channel, and month |
| Lookups | `INDEX`/`MATCH` to find the top region |
| KPIs | `SUM`, `COUNTA`, average order value, % of total |
| Data modelling | Structured `Orders` table with a derived `Month` field for trend analysis |
| Visualisation | Column chart (revenue by category) and line chart (monthly trend) |
| Communication | A single summary sheet a stakeholder can read at a glance |

## Files

```
ecommerce_analysis.xlsx   Workbook: Summary sheet (formulas + charts) and Data sheet (520 orders)
images/summary-01.jpg     Screenshot of the summary dashboard
```

> GitHub can't preview Excel files in the browser, so the screenshot above shows
> the analysis. Download `ecommerce_analysis.xlsx` to see the live formulas and charts.

## Method

The raw data is 520 e-commerce orders from 2025. Each order has a region,
channel, customer segment, product, quantity, and price. The Summary sheet
builds every figure with formulas, with no hardcoded totals, so it recalculates
when the data changes:

- Revenue by category, channel, and region via `SUMIFS` (tables sorted by revenue)
- Monthly revenue and order counts via `SUMIFS` / `COUNTIFS` on the `Month` column
- Top region via `INDEX(range, MATCH(MAX(...), ..., 0))`
- KPIs via `SUM` / `COUNTA`

*Data is synthetic, generated for this project.*

## Tools

Microsoft Excel: SUMIFS, COUNTIFS, INDEX/MATCH, SUM/COUNTA, tables, and charts.
