# Universe Invoice PWA v0.14.21

## Excel export add-on chooser

Invoice Excel export now opens an **optional add-on** dialog before creating the workbook. Every checkbox resets to **unchecked** on every export, for every customer.

Available add-ons:

- Payment Term
- Remark weight details
  - Gold Weight
  - Semi-Precious
  - Diamond
  - All Stone
- Natural-stone declaration
- Stone Description

The previous customer-code restriction is removed. `Total Amount` uses the normal shared Invoice Master Template logic and is not part of the optional add-ons.

When weight details are selected:

- Gold Weight = DESC1 gold weight × Qty
- Semi-Precious = non-diamond DESC2–DESC6 carat weight × Qty
- Diamond = diamond DESC2–DESC6 carat weight × Qty
- All Stone = all DESC2–DESC6 carat weight × Qty
- carats are also shown in grams using ct × 0.2

Diamond / Semi-Precious classification and Stone Description use the **currently imported Stone List**. Stone Description keeps A–Z ordering and two-column output. Optional footer formatting inherits the current Invoice Master Template Remark styles.

If any add-on is selected, the current Invoice Master Template is required. Stone Description and Diamond / Semi-Precious weight options also require the current Stone List.

All other v0.14.20 functionality remains unchanged.
