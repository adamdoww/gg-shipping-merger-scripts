# Build Brief: Merge Shopify Shipping Charges into Orders Export

## Objective
Take two Shopify CSV exports (`orders` and `charges`), pull each order's actual shipping cost from the charges file, and insert **two** new columns into the orders file directly to the right of the existing Shipping column:
1. **Company Shipping Cost** — what the carrier charged us to ship the order.
2. **Shipping Difference** — Shipping (customer paid) minus Company Shipping Cost, i.e. the money made (positive) or lost (negative) on shipping that order.

Then add a styled totals row at the bottom and save the result as a polished, professional `.xlsx` file.

The purpose: put **revenue shipping** (what the customer paid), **company shipping cost** (what the carrier charged us), and the **difference** side by side, so per-order and overall shipping margin is visible at a glance.

---

## Inputs
- `orders_export.csv` — one row per line item. Order-level fields appear only on the first row of each order (see Critical Rule #1).
- `charges_export.csv` — one row per individual charge. Multiple shipping charges can belong to a single order.

(Use the real filenames the client provides; the structures below are what to map against.)

## Output
- A single `.xlsx` file (NOT csv — CSV cannot hold styling or a formatted totals row).
- Suggested name: `orders_with_shipping_cost_<date>.xlsx`.

---

## Column map

### Charges file — columns to read
| Column | Header | Use |
|---|---|---|
| E | Charge category | Filter: only use rows where this equals `shipping_fee` |
| G | Amount | The dollar value to sum |
| L | Order | Join key (order number, e.g. `#GG24474`) |

### Orders file — relevant columns (as delivered, 79 columns, A–CA)
| Column | Header |
|---|---|
| A | Name (this is the order number) |
| I | Subtotal |
| **J** | **Shipping** ← existing |
| **K** | **Taxes** |
| L | Total |
| … | … through CA |

### After the two new columns are inserted
Insert **two** new columns starting at position **K**. Everything from the old K (Taxes) through CA shifts two columns to the right. The file goes from 79 to **81** columns (A–CC).

| Column | Header (after insert) |
|---|---|
| J | Shipping (customer paid) |
| **K** | **Company Shipping Cost** ← NEW |
| **L** | **Shipping Difference** ← NEW (= J − K) |
| M | Taxes (was K) |
| N | Total (was L) |
| … | … old CA now lands at CC |

---

## Merge logic
1. From the charges file, take only rows where **Charge category (E) = `shipping_fee`**.
2. Group those rows by **Order (L)** and **sum the Amount (G)** within each group. A single order can have several shipping charges that must be added together (in the sample data, one order had four separate charges totaling \$91.16).
3. Match each order's summed shipping cost to the orders file by joining **charges Order (L) → orders Name (A)**. The order-number format is identical in both files (leading `#`, e.g. `#GG24474`), so it's a direct string match with no reformatting needed.
4. Write the summed value into the new **Company Shipping Cost** column (K).
5. Compute **Shipping Difference** (L) as **Shipping (J) − Company Shipping Cost (K)**. Positive = made money on shipping; negative = lost money.

---

## Critical rules (these are the things that will silently break the output if missed)

**1. Write the new values to the FIRST row of each order only.**
Multi-line-item orders span multiple rows. The order number in column A repeats on every row, but the order-level fields (Subtotal, Shipping, Taxes, Total) are populated **only on the first row**; the follow-on line-item rows leave them blank. Both Company Shipping Cost and Shipping Difference must be written **only to that same first row** (the one where Shipping is populated). Do NOT write them to every row that matches the order number, or they will appear on blank line-item rows and get double-counted in the totals.

**2. Orders with no matching shipping charge → leave BOTH new columns blank.**
Some orders (typically the most recent, not yet billed/shipped) have no shipping charge in the charges file. For these, leave **both** Company Shipping Cost (K) **and** Shipping Difference (L) blank. Do not fill in `0.00`. A zero cost would make the Difference equal the full customer-paid shipping and read as 100% profit, which is misleading for an order that simply hasn't been billed yet. Blank cells correctly signal "no cost data yet" and are skipped by the totals.

**3. Filter on `shipping_fee` explicitly.**
Do not assume non-shipping charges can be ignored just because they currently lack an order number. Other categories (e.g. `application_fee`) must be excluded by the category filter, not by chance.

---

## Totals row
Add one row at the bottom, below the last order, that totals three columns:
- **Column J (Shipping)** — total customer-paid shipping (money in).
- **Column K (Company Shipping Cost)** — total company shipping cost (money out).
- **Column L (Shipping Difference)** — total money made or lost on shipping.

Sum each column straight down, treating blank cells as zero. Because the new values only land on each order's first row (and only on billed orders), each order is counted at most once with no deduping needed.

**Note on what ties out:** unbilled orders are blank in K and L but still carry a Shipping value in J. So the **Shipping Difference total (sum of column L) reflects margin only on orders that have actually been billed**, and it will not exactly equal (total J − total K), since total J includes customer shipping on not-yet-billed orders. This is intentional and consistent with Rule #2 — it avoids counting shipping on unbilled orders as profit. *(If the client would rather the row tie out exactly, restrict all three totals to orders that have a Company Shipping Cost; optionally add a separate reference cell for total customer shipping across all orders.)*

Label the totals row clearly (e.g. `TOTALS` in column A or column I), leave the non-totaled cells empty, and currency-format the three total cells.

---

## Styling (goal: clean, professional, readable)
- **Header row:** bold white text on a dark fill (navy or dark slate). Freeze the header row so it stays visible when scrolling.
- **Highlight the two new columns:** give `Company Shipping Cost` and `Shipping Difference` a subtle light fill across the whole column so it's obvious they're the added data.
- **Difference column color cues:** conditionally format `Shipping Difference` (and its total) so positive values are green and negative values are red. Blank cells stay unformatted. This makes "made vs lost on shipping" readable at a glance. Keep the shades restrained.
- **Currency formatting:** format all money columns as `$#,##0.00` — at minimum Subtotal, Shipping, Company Shipping Cost, Shipping Difference, Taxes, and Total.
- **Zebra striping:** light alternating row shading for readability.
- **Column widths:** auto-fit to contents so nothing is cut off (cap very wide text columns at a reasonable max).
- **Totals row:** bold, with a top border separating it from the data, and a light fill. Currency-format the totaled cells to match the columns above.
- **Autofilter:** enable filter dropdowns on the header row.
- Keep it restrained and business-like — no heavy colors, just enough structure to make it scannable.

---

## Validation checklist (build these as sanity checks)
- [ ] Output has exactly **81** columns; old column CA now sits at **CC**.
- [ ] Company Shipping Cost is at column K and Shipping Difference at column L, between Shipping (J) and Taxes (M).
- [ ] Shipping Difference equals Shipping (J) − Company Shipping Cost (K) on every billed order's first row.
- [ ] For any multi-line-item order, the new values appear on the first row only, blank on the others.
- [ ] An order with multiple shipping charges shows their **sum**, not just one.
- [ ] Orders absent from the charges file have **both** K and L blank (no `0.00`, no value).
- [ ] Totals row: total J and total K match independent sums of those columns (each order counted once).
- [ ] Totals row: total Shipping Difference equals the sum of column L (billed orders only).
- [ ] No `application_fee` (or other non-shipping) amounts leaked into the new columns.
- [ ] Spot-check 2–3 known orders by hand against the source files.
