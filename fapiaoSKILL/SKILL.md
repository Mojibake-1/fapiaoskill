---
name: amz-invoice-template
description: Repair AMZ/FBA logistics invoice and packing-list workbooks exported from Saihu by using the online 发票产品详情 table as the correction source. Use when the user asks to 做发票, 修正发票, 修正赛狐导出的报关资料/箱单发票, or fill/verify carrier invoice templates for 万逊通、中南通、华洋达、大森林、宁致、宝通达、凯鑫、德鑫淼、斑马、海光、阿拉丁.
---

# AMZ Invoice Repair

Use this skill for the company workflow where operations first exports a carrier invoice workbook from Saihu, then Codex repairs that workbook from the online `发票产品详情` table. The exported workbook is the base file; preserve its carrier template, sheet layout, formulas, merged cells, dropdowns, and styles.

## Hard Input Gate

Do not create, repair, export, or "draft" an invoice workbook until the current task has enough source data.

The normal required source set is:

- the current unrepaired Saihu export workbook, usually `报关资料YYYYMMDD-*.xlsx`
- the current online `发票产品详情` source, preferably an `.xlsx` export so product images can be copied

`Current source` means a path, attachment, export, screenshot, or copied table explicitly provided by the user for this task. Do not silently search `Downloads`, `.analysis`, prior outputs, old screenshots, or generic files such as `发票产品详情.xlsx`. If a plausible local candidate exists but the user did not name it, ask the user to confirm that exact file before using it.

A pasted shipment/product list by itself is not enough, even if it includes carrier, country, cartons, PCS/CTN, total quantity, and dates. For a carrier invoice it normally lacks FBA box numbers, Reference ID, warehouse/address fields, per-box allocation, HS code, declared names, unit price, dimensions, weight, material, purpose, product links, and product images.

Only use a bundled carrier template to create an invoice from scratch when the user explicitly says there is no Saihu export and provides all required top-level and item fields from `references/template-map.md` plus product facts from `发票产品详情`. If any required field is missing, stop and ask for the missing source data instead of filling placeholders or inventing values.

Completed invoice references are presentation or comparison references. Do not use a completed reference as current shipment data unless the user explicitly says it is the same shipment and should be treated as the source.

## Required Workflow

1. Get the current unrepaired Saihu export workbook. This is usually named like `报关资料YYYYMMDD-*.xlsx`.
2. Get the current online `发票产品详情` source or an export/screenshot of it. Prefer an `.xlsx` export because product images must be copied from the workbook.
3. Get a completed invoice reference when the user provides one. Treat it as authoritative for finished formatting, row height, column width, image sizing, default values, and whether formulas should remain or be replaced by values.
4. Identify the carrier template from the workbook structure and/or stock-plan `物流商及渠道`. Read `references/template-map.md` only when the template is unclear.
5. Inspect the invoice detail header row, current formula errors, and the exported actual packing/declaration quantity. Read `references/repair-workflow.md` for the expected fixes.
6. Build a correction JSON using the schema in `references/repair-json.md`.
7. Verify desktop Excel COM:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/check_excel_com.ps1
```

8. Apply corrections to a copy of the exported workbook:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/repair_exported_invoice.ps1 `
  -InvoicePath "C:\path\to\报关资料20260425-20260425154028913.xlsx" `
  -CorrectionsPath "C:\path\to\invoice-corrections.json" `
  -ReferenceWorkbookPath "C:\path\to\完成版发票.xlsx" `
  -OutputPath "C:\path\to\报关资料20260425-已修正.xlsx" `
  -FailOnFormulaErrors
```

9. Review the script JSON report. If required fields are missing, product images are not copied, or formula errors remain, fix the correction source before finalizing.
10. Confirm the report's `textNumbersConverted` count. This replaces the manual WPS/Excel action "批量转换为数字" for numeric detail columns before formulas are recalculated.
11. Record source provenance in the audit summary: which user-provided file supplied the Saihu export, which user-provided file/table supplied `发票产品详情`, and which file was only a completed-format reference.
12. Name the final repaired workbook with the carrier invoice naming rule below.

## Non-Negotiable Rules

- Never modify the original Saihu export, original carrier templates, or SOP workbook. Always write a new repaired workbook.
- Use Excel COM for final repair and save. Do not save official `.xls`/`.xlsx` carrier invoices with `openpyxl`, `pandas`, `xlsxwriter`, or generic spreadsheet exporters.
- Use the current user-provided online product-detail table as the authority for product facts: unit price, declared value, net/gross weight, dimensions, material CN/EN, purpose CN/EN, HS code, product image, product code/name, sales link, SKU, and ASIN.
- Do not source product facts or images from an unmentioned local workbook, a previous run's corrections JSON, or a completed reference unless the user explicitly identifies that file as the current `发票产品详情` source.
- Use the completed invoice reference as the authority for presentation: header layout, hidden/visible columns, row heights, column widths, borders, fonts, image placement, and carrier-specific default values.
- Keep shipment facts from the Saihu export unless the user or online table explicitly corrects them: FBA box number, Reference ID, FBA warehouse, recipient address, box count, and destination.
- Treat the Saihu-exported product/detail area as unreliable. Product columns such as 中文品名, 英文品名, 总数量, 商品单重, 单价, 总价值, 品牌, 是否已注册, 规格型号, 材质中英文, 用途中英文, 产品图片, 产品网络链接, ASIN, and special flags must be checked against `发票产品详情.xlsx` and overwritten when the online table has a value.
- Product pictures must be floating pictures positioned inside the picture cell, not Excel's in-cell/embedded-cell picture mode. Pictures must stay inside the cell borders; oversized pictures are shrunk to fit, and smaller pictures are not enlarged.
- Do not invent compliance-sensitive fields. If HS code, material, purpose, weight, dimensions, unit price, or product image is missing from both sources, ask for it.
- Do not infer shipment identifiers or address fields from filenames, product names, cartons, or prior examples. Missing FBA编号, Reference ID, FBA仓库代码, recipient address, city, state/province, postal code, or country are blockers unless the Saihu export or user-provided current source supplies them.
- Preserve every invoice detail line unless the user confirms rows should be deleted or merged.
- For repeated lines in the same box, fill missing shared fields such as `Reference ID`, `箱数`, brand/default flags, and total formulas consistently.
- When the export shows an actual packing/declaration quantity, use it to determine the valid detail range by cumulatively summing `总数量(PCS)`. Delete rows after the row where the cumulative sum equals that quantity; do not keep overrun rows just because Saihu exported them.
- If the tail valid rows lose borders/row formatting after a long export or row deletion, copy formats from the nearest normal detail row to those tail rows.
- Convert number-like text in numeric detail columns before recalculating formulas. This is the scripted equivalent of clicking WPS/Excel "批量转换为数字"; do it for quantity, carton count, weight, unit price, and total-value columns, but do not convert identifiers such as FBA number, Reference ID, ASIN, SKU, or HS code unless the carrier explicitly requires numeric HS cells.
- Treat unresolved formula errors as blockers unless the user explicitly asks for a draft.

## Output Naming

Final repaired invoice workbook names must use:

```text
{出口方式}{物流服务商}发票-{FBA货件号1[-FBA货件号2...]}-{件数}件{YY-M-D}.xlsx
```

Example:

```text
普船海卡宝通达发票-FBA19C55YJQF-FBA19C5BK0SG-47件26-4-24.xlsx
```

Rules:

- `出口方式` is the service/channel text such as `普船海卡`, `海卡`, `空派`, `卡派`, `快船`, or other terms from the user's message, stock-plan `物流商及渠道`, or the Saihu export. Remove the carrier name from this component when it repeats the carrier.
- `物流服务商发票` is the normalized carrier name plus `发票`, such as `大森林发票`, `宝通达发票`, `万逊通发票`, or `凯鑫发票`.
- `FBA货件号` values come from the current shipment/export, not from `FBA编号` carton suffixes. Preserve source order and join multiple shipment IDs with `-`.
- `件数` is the physical carton/package count used in the invoice header, e.g. `47件`.
- `日期` should use the user's shipment/export date in `YY-M-D` filename form, e.g. user notation `26/4/24` becomes filename text `26-4-24`. Windows filenames cannot contain `/`, so replace slashes with hyphens. If only an 8-digit export date is available, convert `YYYYMMDD` to `YY-M-D`.
- If `出口方式`, carrier, FBA shipment ID, box count, or date cannot be determined from current sources, ask for the missing naming field before finalizing. Do not fall back to generic names such as `已修正`, `output`, or date-only names.

## Source Files

- Repair workflow and completed-reference handling: `references/repair-workflow.md`
- Completed invoice reference notes: `references/completed-reference.md`
- Correction JSON format: `references/repair-json.md`
- Online product table schema: `references/input-schema.md`
- SOP summary: `references/sop-summary.md`
- Template selection and row map: `references/template-map.md`
- Carrier templates: `assets/templates/`
- Completed KaiXin reference sample: `assets/reference-completed-kaixin-invoice.xlsx`
- Original SOP workbook copy: `assets/amz-invoice-sop.xlsx`

## Daily Invocation

If the user says only "帮我做发票" or "修一下这个发票", ask only for missing current-task inputs:

- the unrepaired Saihu export workbook
- the online `发票产品详情` export, link, screenshot, or copied table
- any logistics/channel confirmation only if the exported workbook does not identify the carrier clearly

Do not ask for original template paths unless the exported workbook is corrupt or the bundled reference template is stale.

If the user provides only a product/shipment table, first classify it. Rows shaped like `仓库/产品名/SKU/负责人/国家或渠道/PCS per CTN/箱数/总数量/日期` are stock-plan or shipment-planning data, not a complete invoice source. Reply with the missing inputs instead of generating a workbook. For example, say that 宝通达 invoice generation still needs the Saihu export or, if creating from template, the FBA box-number sequence, Reference ID per shipment, warehouse/address fields, and `发票产品详情` product facts.

If the user provides a half-finished invoice plus a product/shipment list but no current `发票产品详情` source, inspect the workbook and report what can be fixed from the invoice itself. Do not fill missing unit price, brand, product links, product images, HS/material/purpose corrections, or other product facts from unrelated local files. Ask for the current product-detail source before producing a finished repaired workbook.

If the user later provides a known-good completed invoice for comparison, treat it as a reference unless they explicitly say it is the current data source. Use it to report differences and to improve formatting rules; do not retroactively treat it as proof that earlier unprovided product data was authorized.

## Final Checks

Before returning the repaired workbook:

- Confirm the original export was copied and not overwritten.
- Confirm item-line count before/after repair.
- Confirm `总数量(PCS)` across visible detail rows equals the exported actual packing/declaration quantity when that source value is provided.
- Confirm all required corrected fields are filled for each visible detail row.
- Confirm formulas have recalculated and no `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, or `#N/A` remains.
- Confirm text-number warnings are cleared for numeric detail columns, or that the script report has a nonzero/expected `textNumbersConverted` value.
- Confirm `总价值` equals `单价 * 总数量` where the template uses those fields.
- Confirm any image requirement is either satisfied, explicitly marked for manual insertion, or disclosed as missing.
- Provide a compact audit summary: source workbook, selected sheet, detail rows repaired, product-detail source with provenance, completed reference if used, unresolved fields, and formula errors.
