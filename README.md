# Universe Invoice PWA v0.14.16

## RO1220 customer-specific Invoice footer

Only when the current document is an **Invoice** and Customer Code is **RO1220**, Excel output from the current imported Invoice Master Template adds the following blocks after `Total Amount`:

1. `Payment Term :`
   - PAID BY BANK TT 20%
   - TERMS 90 / 120 / 150 / 180 DAYS, each 20%
   - due dates follow invoice date + 3 / 4 / 5 / 6 calendar months
   - the first four installments are rounded to whole currency units; the final installment balances exactly to the Invoice Total
2. `Remark :`
   - TOTAL GOLD WEIGHT
   - TOTAL STONES WEIGHT (CARATS), shown as grams and carats
   - TOTAL GROSS WEIGHT
3. Natural-stone declaration
4. `STONE DESCRIPTION :`
   - uses the currently imported Stone List only
   - only stone codes actually used in the Invoice
   - code order A–Z
   - two-column index layout: C:D on the left, E:H on the right
   - column-major order (read down the left column, then down the right)
   - one blank row after the final Stone Description line

All inserted RO1220 footer content inherits the current Invoice Master Template Remark style; no font name/size/bold is hard-coded in the Template export path.

For RO1220 Invoice export, both the current Stone List and Invoice Master Template are required. Other customers keep the normal Invoice footer and do not receive automatic Stone Description / RO1220 footer blocks.

## Weight calculation

- Gold weight: leading gram value from DESC1 × Quantity.
- Stone carats: every `...ct` value in DESC2–DESC6 × Quantity.
- Stone grams: total carats × 0.2.
- Gross grams: gold grams + stone grams.

## Existing v0.14.15 behaviour retained

- LOTNO / ARTNO item styles inherit the Invoice Master Template without a forced bold override.
- Universe Records remain embedded in `jmsdata.xlsx` Sheet 2.
- New-exhibition record isolation and cache-busting / network-first update behaviour remain unchanged.
