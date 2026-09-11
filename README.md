# Margin & mix bridge

SKU × channel contribution, price-volume-mix, and gross margin walk for a fictional CPG company (**Northline Consumer Products**). Sample Excel file showing how mix shift to Club can dilute GM$ even when volume is up.

**File to open:** [`Northline_Margin_Mix_Bridge.xlsx`](Northline_Margin_Mix_Bridge.xlsx)

`build.py` only regenerates that workbook. The file a hiring manager should open is the `.xlsx`.

## What you will see

Prior → current gross margin bridge with clear steps: **Prior GM → Volume → Price → Mix → Cost → Current GM**, plus contribution by SKU and by channel (Grocery, Mass, Club, E-comm). Yellow / blue cells are inputs; black font is formulas. Amounts in $000s where labeled.

## How to use (≈8 minutes)

1. Open `01_Assumptions`. Edit prior/current volume and net price by SKU × channel (yellow).
2. Scan `02_Volume_Price` for revenue and contribution builds.
3. Read `03_PVM_Bridge` — revenue bridge and GM$ bridge should reconcile to Δ totals.
4. Check mix % and CM $/unit on `04_Contribution` (Club share rising is the headwind).
5. Close on `05_Dashboard`: GM walk chart, mix callouts, RAG on margin %.

## Screen-share test

Raise Club volume further on `01_Assumptions` and leave prices fixed. Volume step on the GM bridge should improve; Mix should worsen. Then cut Grocery net price — Price step turns more negative; Cost stays ~0 (std COGS/unit unchanged in the sample).

## Tabs

| Tab | Role |
| --- | --- |
| `00_Cover` | Purpose, legend, fictional disclaimer |
| `01_Assumptions` | Close period, channels, SKUs, COGS/unit, prior/current volume & price |
| `02_Volume_Price` | SKU×channel volume, price, revenue, contribution |
| `03_PVM_Bridge` | Price / volume / mix (and cost) bridges on revenue and GM$ |
| `04_Contribution` | Contribution by SKU and by channel; mix % |
| `05_Dashboard` | One-pager: GM walk, mix callouts, RAG on GM% |
| `06_Data_Dictionary` | Field definitions |

Excel formulas only. No VBA, no live ERP or POS feed, no employer data. This is the margin / mix / PVM slice — not a full three-statement model or driver-based P&L.

Fictional company and sample data for portfolio demonstration only.

[Profile](https://github.com/saisiri-bandaru) · [Portfolio](https://saisiri-bandaru.github.io) · [LinkedIn](https://www.linkedin.com/in/bandarusaisiri) · [bandarusaisiri1207@gmail.com](mailto:bandarusaisiri1207@gmail.com)

Sai Siri Bandaru — Financial Analyst | FP&A
