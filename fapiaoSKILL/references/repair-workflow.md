# Exported Invoice Repair Workflow

The base file is the Saihu exported invoice. Repair it on a copy. When the user provides a completed invoice, use that completed invoice as the finished-format reference and the online `发票产品详情` workbook as the product-data and product-image source.

## Final File Naming

Use this filename shape for repaired carrier invoices:

```text
{出口方式}{物流服务商}发票-{FBA货件号1[-FBA货件号2...]}-{件数}件{YY-M-D}.xlsx
```

Example: `普船海卡宝通达发票-FBA19C55YJQF-FBA19C5BK0SG-47件26-4-24.xlsx`.

Derive `出口方式` from the user's wording or stock-plan `物流商及渠道` by removing the carrier keyword and keeping the service/channel part, such as `普船海卡`. Derive `物流服务商` from the matched carrier template, such as `宝通达` or `大森林`. Use FBA shipment IDs, not per-carton `FBA编号` values. Use the invoice header box count for `件数`. Use the user's shipment/export date in `YY-M-D` filename form, so `26/4/24` becomes `26-4-24` because `/` is not valid in Windows filenames. If any naming component is missing, ask for it before final delivery.

## Source Completeness Gate

Stop before workbook generation when the current task does not include either:

- a current Saihu exported invoice workbook plus the current `发票产品详情` source, or
- an explicit create-from-template request with all required carrier fields, item fields, and product facts.

Only treat a source as current when the user explicitly provides or confirms it for this task. Do not search for or reuse local files by generic names such as `发票产品详情.xlsx`, prior `.analysis` outputs, generated corrections JSON, or unrelated completed invoices. If a plausible file is found locally but was not named by the user, ask for confirmation before using it.

Do not treat stock-plan rows or pasted shipment rows as a complete invoice source. They can confirm intended products, carton counts, country, and dates, but they do not supply the official FBA box numbers, Reference ID, warehouse/address fields, per-box allocation, HS code, declared product facts, product links, or images needed for a finished carrier invoice.

If a completed invoice workbook is provided only for comparison or reference, use it to identify differences and missing inputs. Do not backfill current shipment values from it unless the user explicitly says it is the current source.

If the user provides a half-finished invoice plus a shipment/product list, the invoice may supply shipment identifiers and address fields, but missing product facts still require the current `发票产品详情` source. In that case, report the missing product-detail source instead of producing a finished workbook with untraceable values.

## Inspection

1. Identify the main invoice sheet. It usually contains a detail header row with labels such as `FBA编号`, `Reference ID`, `海关编码`, `中文品名`, `英文品名`, `单价`, or `总价值`.
2. Identify the detail start row as the row after the header row.
3. Record the existing line count. Keep rows unless the user confirms deletion.
4. Scan formulas with cached values for `#VALUE!`, `#REF!`, `#DIV/0!`, `#NAME?`, and `#N/A`.
5. Compare the detail fields against the online `发票产品详情` source.
6. If a completed invoice reference is provided, inspect its detail header row, row heights, column widths, image placement, formulas vs fixed values, default fields, and comments. The finished workbook should look like the completed reference, not like the raw Saihu export.

## Common Saihu Export Problems

- The exported product/detail area can be broadly wrong, not just incomplete. Do not trust product columns just because they are nonblank.
- Unit price is blank, so total-value formulas evaluate to `#VALUE!`.
- Numeric fields often arrive as text and show the WPS/Excel warning "数字是文本类型". The manual fix is "批量转换为数字"; the repair script must perform the same conversion automatically for numeric detail columns.
- Repeated rows in one box do not fill down shared fields such as `Reference ID`, `箱数`, `是否已注册`, `规格型号`, and yes/no flags.
- Material and purpose are too coarse from Saihu and need the online table's bilingual fields.
- HS code, declared name, weight, or dimensions need replacement from the online product table.
- Product image cells are blank even when the carrier template marks them required.
- The raw export may retain a carrier upload template shape that is not the final company-facing invoice shape. In that case, use the completed invoice reference to decide final formatting and visible fields.
- Long Saihu exports can include extra repeated detail rows beyond the selected shipment's actual packing/declaration quantity.
- The last valid rows can lose borders, row height, wrapping, or other detail-row formatting after a long export or after extra rows are deleted.

## Fixing Rules

- Write corrections only to the invoice copy.
- Prefer cell updates through header names instead of fixed column letters. Carrier templates move columns.
- Preserve shipment/box identifiers from the export, but re-source product-detail columns from `发票产品详情.xlsx`.
- If there is no export, require an explicit current source for shipment/box identifiers before using a carrier template. Never synthesize FBA编号, Reference ID, recipient address, warehouse code, or box-number ranges from carton counts or example files.
- Before writing row corrections, confirm every non-export product value can be traced to a user-provided current product-detail source. If the trace is missing, leave the field unresolved and ask for the source.
- Always check and, when available, overwrite these product/detail fields from the online table: `中文品名`, `英文品名`, `总数量(PCS)`, `商品单重(KG)`, `单价`, `总价值`, `品牌`, `是否已注册`, `规格型号`, `材质中英文`, `用途中英文`, `产品图片`, `产品网络链接`, `ASIN`, `海关编码`, `净重`, `毛重`, and dimensions.
- Fill blank shared fields from the nearest valid previous line only when it is the same physical box or the same shipment-level field.
- If the source page/export provides an actual packing/declaration quantity, sum `总数量(PCS)` from the first detail row downward. The row where the cumulative sum equals that source quantity is the valid final detail row. Delete later exported rows.
- After deleting extra rows, fix any tail-row formatting gaps by copying formats and row height from the nearest normal detail row, without copying values or pictures.
- Keep formulas where the carrier template expects formulas. For example, `总价值` should stay a formula when the template calculates it from `单价 * 总数量`.
- Before recalculating formulas, convert number-like text to real numeric values in math/input columns such as `箱数`, `单箱毛重(KGS)`, `总数量(PCS)`, `商品单重(KG)`, `单价`, and `总价值`. Do not convert identifier/code columns such as FBA number, Reference ID, ASIN, SKU, or HS code unless the template owner explicitly requires those cells as numbers.
- If the template says formulas are not allowed, write the calculated number instead of a formula.
- For image-required columns, do not claim completion unless the image is embedded or the missing image is explicitly reported.
- Copy actual product pictures from `发票产品详情.xlsx`, normally from the `图片` column. Pictures must be floating shapes placed over the picture cell, not Excel in-cell pictures. Keep every picture inside the target cell borders; shrink oversized pictures to fit and do not enlarge smaller pictures.
- If the completed reference uses fixed values rather than formulas for totals, write fixed values. If it keeps formulas, preserve formulas and recalculate.

## Example: 宝通达 Export

The sample `报关资料20260425-20260425154028913.xlsx` is a 宝通达-style workbook:

- Main sheet: `发票信息`
- Header row: 11
- Detail rows observed: 12-33
- Important columns:
  - `A` FBA编号
  - `B` Reference ID
  - `C` 材积CM(长*宽*高)
  - `D` 箱数
  - `E` 单箱毛重(KGS)
  - `F` 海关编码HSCODE
  - `G` 中文品名
  - `H` 英文品名
  - `I` 总数量(PCS)
  - `K` 单价
  - `L` 总价值
  - `M` 品牌
  - `N` 是否已注册
  - `O` 规格型号
  - `P` 材质中英文
  - `Q` 用途中英文
  - `U:Y` battery/magnet/liquid/paste/powder flags

Typical corrections for this template:

- Fill `单价` from online table column `单价（美国）/USD` or the carrier-approved price field.
- Set `总价值` to the existing row formula `=K{row}*I{row}` after price is present.
- Fill blank `Reference ID` from the shipment-level reference when rows belong to the same shipment.
- Fill blank `箱数` as `1` for each physical box/detail line unless the source indicates a merged carton.
- Fill default flags: `是否已注册=无`, `规格型号=无`, `是否带电=否`, `是否带磁=否`, `是否是液体=否`, `是否是膏状=否`, `是否粉末=否`.
- Replace product fields from the online table when they differ: `中文品名`, `英文品名`, `海关编码HSCODE`, `材质中英文`, `用途中英文`, dimensions, and gross weight.

## Completed Invoice Reference Handling

When a completed invoice reference is provided:

- Treat it as a specimen, not as a source of current shipment quantities.
- Copy formatting only when it is the same template family or the user explicitly wants that finished layout.
- Compare header labels and detail columns. If the completed reference omits or reorders columns, document the intended mapping before applying corrections.
- In comparison-only mode, do not change files. Report value, formula, sheet-structure, image-placement, and formatting differences separately.
- In strict match mode, apply the completed reference's workbook structure and presentation choices when they are not shipment facts. For 宝通达 this can include phone default `0`, preserving or removing extra sheets to match the reference, retaining a blank formatted tail area instead of physically deleting it, and copying row heights, column widths, borders, fonts, wrapping, and number formats across the used invoice block.
- Do not use the completed reference to fill unit price, product links, product images, HS code, material, purpose, or brand unless the user explicitly says the reference is the current source for the same shipment.
- Match image behavior:
  - If the reference has images in `产品图片*` or `产品图片`, copy floating pictures into the repaired workbook.
  - Keep pictures centered inside the cell and sized to fit without stretching or crossing borders.
  - Shrink oversized pictures only; do not enlarge smaller source pictures.
  - Copy one image per detail row unless the template/reference shows a repeated image pattern.
- For KaiXin-style completed invoices, observed finished rows use:
  - detail header row 27
  - detail rows from row 28
  - product images anchored in column `P`
  - ASIN in the rightmost detail column
  - sales links filled from the online table
  - formula-free numeric fields where the template instructions say formulas are not allowed

## Review Output

Return a compact audit with:

- invoice source path
- repaired output path
- final output filename components: 出口方式, 物流服务商, FBA货件号, 件数, 日期
- sheet and row range
- rows changed
- fields changed
- quantity-based end row and deleted overrun rows
- tail-row formatting fixes
- text-number conversions performed
- formula errors before and after
- missing required fields, especially product images
