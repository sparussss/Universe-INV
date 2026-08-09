# Universe Invoice PWA v0.14.17

## RO1220 customer-specific Invoice footer

Only when the current document is an **Invoice** and Customer Code is **RO1220**, Excel output from the currently imported Invoice Master Template adds the customer-specific footer after `Total Amount`.

1. `Payment Term :` shares the same row as `PAID BY BANK TT 20% ...`.
   - The first four 20% installments are rounded **up** to the next whole currency unit.
   - The final installment is the exact balance so all five installments equal the Invoice Total.
   - Example for USD 27,322: 5,465 / 5,465 / 5,465 / 5,465 / 5,462.
   - 90 / 120 / 150 / 180-day dates remain Invoice Date + 3 / 4 / 5 / 6 calendar months.
2. `Remarks :` shares the same row as `TOTAL GOLD WEIGHT`.
   - Weight rows use separate Excel blocks like the old customer file: label in C:E, numeric value in F, unit / carat text in G:I.
   - TOTAL GOLD WEIGHT
   - TOTAL STONES WEIGHT (CARATS)
   - TOTAL GROSS WEIGHT
   - Any manually entered PWA Remark, if present, is kept as a separate row after the three weight rows.
3. Natural-stone declaration follows the weight section.
4. `Stone Decsription :` shares the same row as the first Stone Description entry.
   - uses the currently imported Stone List only
   - only stone codes actually used in the Invoice
   - code order A–Z
   - two-column index layout: C:D on the left, E:H on the right
   - column-major order (read down the left column, then down the right)
   - one blank row after the final Stone Description line

All inserted RO1220 footer content inherits the current Invoice Master Template Remark style; no font name/size/bold is hard-coded in the Template export path.

For RO1220 Invoice export, both the current Stone List and Invoice Master Template are required. Other customers keep the normal Invoice footer and do not receive these automatic customer-specific blocks.

## Weight calculation

- Gold weight: leading gram value from DESC1 × Quantity.
- Stone carats: every `...ct` value in DESC2–DESC6 × Quantity.
- Stone grams: total carats × 0.2.
- Gross grams: gold grams + stone grams.

## Existing behaviour retained

- LOTNO / ARTNO item styles inherit the Invoice Master Template without a forced bold override.
- Universe Records remain embedded in `jmsdata.xlsx` Sheet 2.
- New-exhibition record isolation and cache-busting / network-first update behaviour remain unchanged.
