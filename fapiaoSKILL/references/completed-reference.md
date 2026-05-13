# Completed Invoice Reference

Use `assets/reference-completed-kaixin-invoice.xlsx` as a sample completed invoice reference. It is not current-task data; it shows the target quality bar for a finished carrier invoice.

## Observed KaiXin Reference

- Workbook sheets: `发票`, `FBA地址库编码表`, `服务名称`.
- Main sheet: `发票`.
- Header fields are filled above the detail table, including address-code lookup formulas, customer order number, battery/magnet flags, customs method, and box count.
- Detail header row: 27.
- Detail rows: 28-48 in the sample.
- Product pictures are real embedded picture shapes anchored in column `P`.
- The ASIN column is filled at the right side of the detail table.
- Sales links are filled from product detail data.
- The template instruction says formulas are not allowed in detail fields; finished numeric values should therefore be fixed values for this template family.
- Row heights, column widths, comments, and image sizing matter. Copy or imitate these from the completed reference when repairing a KaiXin export.

## How To Use A Completed Reference

1. Inspect the reference workbook before repairing the current export.
2. Compare the reference detail headers to the current exported workbook. If they differ, map fields explicitly before writing.
3. Use the reference for:
   - formatting
   - row heights and column widths
   - visible/hidden fields
   - default values
   - whether formulas are acceptable
   - image size and placement
4. Do not use the reference for current shipment quantities, box numbers, ASINs, or product values unless the user explicitly says the current shipment is the same.

## Same-Shipment Comparison Mode

When the user provides a known-good completed invoice after a draft or repaired output, first compare in read-only mode. Report differences by category:

- workbook sheets and UsedRange
- header values and number formats
- detail row count, FBA编号, Reference ID grouping, box count, quantity, unit price, total value
- product fields such as brand, material, purpose, links, and images
- formula errors and missing required fields
- row heights, column widths, borders, fonts, wrapping, image anchors, and blank formatted tail area

If the user asks for strict matching, use the completed reference for presentation and workbook structure only. Do not treat it as the source for product facts, images, links, or prices unless the user explicitly confirms that the completed reference is the current data authority for the same shipment.

## Image Standard

For the reference sample, product pictures are small, centered, and fit inside the product image cell. The repaired invoice should contain embedded images, not text placeholders or external paths, when the carrier template/reference requires pictures.
