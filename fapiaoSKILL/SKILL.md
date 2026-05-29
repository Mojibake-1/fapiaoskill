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
- the current online `发票产品详情` source. By default this is the pinned Tencent Docs sheet `https://docs.qq.com/sheet/DY0hGbmd2Q1ZTVFVD?tab=BB08J2`; fetch it yourself through the Tencent Docs MCP instead of asking the user to provide it. Prefer an `.xlsx` export when product images must be copied.
- when the task requires choosing which shipment/invoice lines to process from the planning sheet, a clear user-specified work scope from the default `AMZ备货计划及出货安排表` source below. If the user does not specify the part to do, or the scope is vague, ask before selecting rows.

For the product-detail source, the pinned Tencent Docs sheet counts as the current source unless the user explicitly provides a different current `发票产品详情` link, workbook, screenshot, or copied table for this task. For all other files, `Current source` means a path, attachment, export, screenshot, or copied table explicitly provided by the user for this task. Do not silently search `Downloads`, `.analysis`, prior outputs, old screenshots, or generic files such as `发票产品详情.xlsx`. If a plausible local candidate exists but the user did not name it, ask the user to confirm that exact file before using it.

A pasted shipment/product list by itself is not enough, even if it includes carrier, country, cartons, PCS/CTN, total quantity, and dates. For a carrier invoice it normally lacks FBA box numbers, Reference ID, warehouse/address fields, per-box allocation, HS code, declared names, unit price, dimensions, weight, material, purpose, product links, and product images.

Only use a bundled carrier template to create an invoice from scratch when the user explicitly says there is no Saihu export and provides all required top-level and item fields from `references/template-map.md` plus product facts from `发票产品详情`. If any required field is missing, stop and ask for the missing source data instead of filling placeholders or inventing values.

Completed invoice references are presentation or comparison references. Do not use a completed reference as current shipment data unless the user explicitly says it is the same shipment and should be treated as the source.

## Tencent Docs MCP Dependency

This skill includes its own Tencent Docs MCP runbook at `references/tencent-docs-mcp.md`. Use it when a colleague runs the skill, when MCP authorization is unclear, or when any Tencent Docs command fails.

On the current shared Windows environment, colleagues may reuse the already configured `tencent-docs` MCP authorization under the current Windows user. Do not copy, print, or commit the underlying token or `Authorization` header. If the skill is run under another Windows account or another machine, the owner must authorize that local environment with `mcporter auth tencent-docs` or the official Tencent Docs manual token flow.

Before asking the user to resend online sheets, run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/check_tencent_docs_mcp.ps1
```

Use `--server tencent-docs --tool "..."` for all Tencent Docs calls because dotted tool selectors can be parsed incorrectly by some `mcporter` versions.

## Default Product-Detail Source

When the user asks for an invoice repair and does not provide a separate `发票产品详情` source, use this Tencent Docs source:

- URL: `https://docs.qq.com/sheet/DY0hGbmd2Q1ZTVFVD?aidPos=detail&no_promotion=1&is_blank_or_template=blank&tab=BB08J2&u=4eb605868b504f9a88efe9a066309f91`
- `file_id`: `DY0hGbmd2Q1ZTVFVD`
- `sheet_id` / tab: `BB08J2`
- expected title: `发票产品明细`

Use the Tencent Docs MCP directly. Because its tool names contain dots, prefer the `--server` and `--tool` form:

```powershell
mcporter call --server tencent-docs --tool "manage.query_file_info" file_id=DY0hGbmd2Q1ZTVFVD --output json
mcporter call --server tencent-docs --tool "sheet.get_sheet_info" file_id=DY0hGbmd2Q1ZTVFVD --output json
mcporter call --server tencent-docs --tool "sheet.get_cell_data" file_id=DY0hGbmd2Q1ZTVFVD sheet_id=BB08J2 start_row=0 start_col=0 return_csv=true --output json
```

For product images, export the online sheet to a local `.xlsx` first, then pass that workbook as `imageSource.workbookPath` in the repair JSON. Use the helper script when possible:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/fetch_product_detail_from_tencent_docs.ps1
```

The helper script writes a sidecar metadata file next to the local workbook. On later runs it compares Tencent Docs `last_modify_time`, `file_id`, `sheet_id`, and sheet row/column counts with that metadata, and also verifies the local workbook SHA-256 hash has not changed; when all checks match, it returns `skippedDownload=true` and reuses the local workbook instead of downloading the same online document again. Use `-Force` only when you intentionally need a fresh export despite no detected difference.

If Tencent Docs MCP authorization is missing, the export fails, the target sheet ID changes, or the online export lacks required images, then ask the user for either Tencent Docs access/auth repair or a current exported `发票产品详情.xlsx`. Do not fall back to an old local product-detail workbook.

## Default Work-Scope Source

When the user says the invoice details to make are in the planning sheet, use this Tencent Docs source to locate the requested shipment rows:

- URL: `https://docs.qq.com/sheet/DRE1ZTlhoZVZBVkdL?u=2433b3ce0ab54f1e8b74acb4b0e2c643&tab=000001`
- URL file token: `DRE1ZTlhoZVZBVkdL`
- `sheet_id` / tab: `000001`
- expected title: `AMZ备货计划及出货安排表`
- expected sheet name: `备货详情`
- observed header row: row 3, with data starting at row 4

Use this source only to identify the user-requested part and shipment planning fields such as product name, SKU, operator, site, PCS/CTN, carton count, total quantity, dates, destination warehouse, FBA shipment ID, carrier/channel, declaration method, total cartons, inbound number, notes, packing status, and shipping status. It does not replace the Saihu export for official invoice workbook structure, FBA box-number sequences, Reference IDs, address fields, or per-box allocation, and it does not replace `发票产品详情` for product facts or images.

Use Tencent Docs MCP directly:

```powershell
mcporter call --server tencent-docs --tool "manage.query_file_info" file_id=DRE1ZTlhoZVZBVkdL --output json
mcporter call --server tencent-docs --tool "sheet.get_sheet_info" file_id=DRE1ZTlhoZVZBVkdL --output json
mcporter call --server tencent-docs --tool "sheet.get_cell_data" file_id=DRE1ZTlhoZVZBVkdL sheet_id=000001 start_row=0 start_col=0 return_csv=true --output json
```

If a local `.xlsx` copy is useful, use the helper wrapper:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/fetch_invoice_work_scope_from_tencent_docs.ps1
```

The helper uses the same metadata and SHA-256 cache rule as the product-detail export helper: if the online document and local workbook have no detected difference, it returns `skippedDownload=true` and does not download again.

The user must specify the part to do each time, for example by row number/range, FBA shipment ID, carrier/channel plus exact product rows, product/SKU list, or another unambiguous selection. If the user only says "做发票", "做这里的", "最新的", "今天的", "这部分", "上面这些", or provides any scope that can map to multiple rows/blocks, ask which rows or range to use. Do not silently process the whole sheet, all visible rows, the latest rows, or the first matching carrier.

## Required Workflow

1. Get the current unrepaired Saihu export workbook. This is usually named like `报关资料YYYYMMDD-*.xlsx`.
2. If the task refers to the planning sheet or does not otherwise make the shipment scope obvious, get the user-specified rows/range from the default work-scope source. If the requested part is missing or vague, ask before doing any workbook generation.
3. Get the current online `发票产品详情` source. Default to the pinned Tencent Docs sheet above and fetch it through MCP; only ask the user for a product-detail file when MCP/export access is blocked or the user says to use a different current source. Prefer an `.xlsx` export because product images must be copied from the workbook.
4. Get a completed invoice reference when the user provides one. Treat it as authoritative for finished formatting, row height, column width, image sizing, default values, and whether formulas should remain or be replaced by values.
5. Identify the carrier template from the workbook structure and/or stock-plan `物流商及渠道`. Read `references/template-map.md` only when the template is unclear.
6. Inspect the invoice detail header row, current formula errors, and the exported actual packing/declaration quantity. Read `references/repair-workflow.md` for the expected fixes.
7. Before filling corrections, calculate the final invoice detail row count. Treat the first 10 valid detail rows in the current export as the normal style sample unless a completed reference proves otherwise. If more detail rows are needed, copy row height and cell formats from the nearest normal sample row to the later valid rows before filling values and product pictures. Do not hardcode row 35; that boundary depends on the exported template.
8. Build a correction JSON using the schema in `references/repair-json.md`.
9. Verify desktop Excel COM:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/check_excel_com.ps1
```

10. Apply corrections to a copy of the exported workbook:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/repair_exported_invoice.ps1 `
  -InvoicePath "C:\path\to\报关资料20260425-20260425154028913.xlsx" `
  -CorrectionsPath "C:\path\to\invoice-corrections.json" `
  -ReferenceWorkbookPath "C:\path\to\完成版发票.xlsx" `
  -OutputPath "C:\path\to\报关资料20260425-已修正.xlsx" `
  -FailOnFormulaErrors
```

11. Review the script JSON report. If required fields are missing, product images are not copied, or formula errors remain, fix the correction source before finalizing.
12. Confirm the report's `textNumbersConverted` count. This replaces the manual WPS/Excel action "批量转换为数字" for numeric detail columns before formulas are recalculated.
13. Record source provenance in the audit summary: which user-provided file supplied the Saihu export, which user-specified work-scope rows/range were used from `AMZ备货计划及出货安排表`, which Tencent Docs `file_id`/`sheet_id` or user-provided file/table supplied `发票产品详情`, the export time/path when applicable, and which file was only a completed-format reference.
14. Name the final repaired workbook with the carrier invoice naming rule below.

## Non-Negotiable Rules

- Never modify the original Saihu export, original carrier templates, or SOP workbook. Always write a new repaired workbook.
- Never infer the requested invoice scope from the planning sheet. If the user has not specified exact rows/range/IDs, ask before selecting rows.
- Use Excel COM for final repair and save. Do not save official `.xls`/`.xlsx` carrier invoices with `openpyxl`, `pandas`, `xlsxwriter`, or generic spreadsheet exporters.
- Use the current online product-detail table as the authority for product facts: unit price, declared value, net/gross weight, dimensions, material CN/EN, purpose CN/EN, HS code, product image, product code/name, sales link, SKU, and ASIN. By default, fetch this table from Tencent Docs `file_id=DY0hGbmd2Q1ZTVFVD`, `sheet_id=BB08J2`.
- Do not source product facts or images from an unmentioned local workbook, a previous run's corrections JSON, or a completed reference unless the user explicitly identifies that file as the current `发票产品详情` source. The pinned Tencent Docs sheet is the only standing default.
- Use the completed invoice reference as the authority for presentation: header layout, hidden/visible columns, row heights, column widths, borders, fonts, image placement, and carrier-specific default values.
- Keep shipment facts from the Saihu export unless the user or online table explicitly corrects them: FBA box number, Reference ID, FBA warehouse, recipient address, box count, and destination.
- For 海光 (`发票模板`) invoices, fill top-level `服务*` with the channel/service text without the carrier name. For example, when `物流商及渠道` is `海光普船海卡`, write `普船海卡`, not `海光普船海卡` and not a generic channel such as `美西OA卡派含税`. For this 海光 workflow, set `报关方式*` to `报关退税`. Other carrier templates may name and interpret service/declaration fields differently, so inspect the template header and source wording before applying carrier-specific defaults.
- Follow each carrier template's country field wording. If a field asks for `国家代码`, `二字代码`, or similar ISO two-letter country code, write the two-letter code such as `US`, not the Chinese country name `美国`. If another carrier template explicitly asks for country name or address text, keep the required wording for that template instead of applying this rule blindly.
- Treat the Saihu-exported product/detail area as unreliable. Product columns such as 中文品名, 英文品名, 总数量, 商品单重, 单价, 总价值, 品牌, 是否已注册, 规格型号, 材质中英文, 用途中英文, 产品图片, 产品网络链接, ASIN, and special flags must be checked against `发票产品详情.xlsx` and overwritten when the online table has a value.
- Whenever an invoice template has a detail-row product-brand header such as `产品品牌`, `产品品牌*`, `品牌`, or `Brands 品牌`, set every visible detail-row brand cell from the current shipment's `发货站点`: if `发货站点` contains `uc` case-insensitively, use `Ucoolbe`; otherwise use `MXZONE`. Apply it as an overwrite across the final detail range.
- Whenever an invoice template has a detail-row brand-type header such as `品牌类型` or `品牌类型*`, set every visible detail-row brand-type cell to `境外品牌` unless the user explicitly provides a different approved value. Apply it as an overwrite across the final detail range, not as a blank-only default, because Saihu exports can contain mixed `无`, `无品牌`, or `境外品牌(贴牌生产)` values.
- Before saving or returning the final invoice, format customs/HS-code columns such as `海关编码`, `海关编码HSCODE`, `产品海关编码`, and `产品海关编码*` as Excel text format `@`, and write nonblank HS-code values as strings. Do not include these columns in `numericHeaders`; keeping them as text prevents scientific notation and preserves code formatting when copied into other systems.
- Product pictures must be floating pictures positioned inside the picture cell, not Excel's in-cell/embedded-cell picture mode. Prefer exporting the source picture to a temporary bitmap and inserting it with `Shapes.AddPicture` at calculated coordinates; use clipboard paste only as a fallback. Size each picture from the final target cell or merge-area bounds, keep about 3 pt visible padding on every side, default to roughly 90% max cell width and 90% max cell height, preserve the source aspect ratio, and allow normal-sized pictures to be enlarged to that fitted box. After all pictures are copied, reopen or rescan the workbook and verify by geometry, not only by `TopLeftCell`/`BottomRightCell`: no zero-size pictures, exactly one picture per required row, center point in the target picture cell, and all edges inside the target cell bounds.
- Do not invent compliance-sensitive fields. If HS code, material, purpose, weight, dimensions, unit price, or product image is missing from both sources, ask for it.
- Do not infer shipment identifiers or address fields from filenames, product names, cartons, or prior examples. Missing FBA编号, Reference ID, FBA仓库代码, recipient address, city, state/province, postal code, or country are blockers unless the Saihu export or user-provided current source supplies them.
- Preserve every invoice detail line unless the user confirms rows should be deleted or merged.
- For repeated lines in the same box, fill missing shared fields such as `Reference ID`, `箱数`, brand/default flags, and total formulas consistently.
- When the export shows an actual packing/declaration quantity, use it to determine the valid detail range by cumulatively summing `总数量(PCS)`. Delete rows after the row where the cumulative sum equals that quantity; do not keep overrun rows just because Saihu exported them.
- After the valid detail range is known and before filling product corrections, compare the final detail row count with the first 10 valid detail rows from the current export. Those first 10 rows are the normal style sample by default. If later valid rows exist, copy row height and cell formats from the nearest normal sample row to those later rows before writing values or product pictures. This must happen before image insertion so pictures are sized and centered against the final row height.
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
- Tencent Docs MCP authorization and command runbook: `references/tencent-docs-mcp.md`
- Online product table schema: `references/input-schema.md`
- Work-scope planning sheet schema: `references/work-scope-schema.md`
- SOP summary: `references/sop-summary.md`
- Template selection and row map: `references/template-map.md`
- Carrier templates: `assets/templates/`
- Completed KaiXin reference sample: `assets/reference-completed-kaixin-invoice.xlsx`
- Original SOP workbook copy: `assets/amz-invoice-sop.xlsx`

## Daily Invocation

If the user says only "帮我做发票" or "修一下这个发票", ask only for missing current-task inputs:

- the unrepaired Saihu export workbook
- the exact rows/range or unambiguous shipment selection from the default `AMZ备货计划及出货安排表` work-scope sheet, when the task depends on that sheet
- any logistics/channel confirmation only if the exported workbook does not identify the carrier clearly

Do not ask for the `发票产品详情` document by default; fetch the pinned Tencent Docs sheet through MCP. Do not ask for original template paths unless the exported workbook is corrupt or the bundled reference template is stale.

If the user provides only a product/shipment table, first classify it. Rows shaped like `仓库/产品名/SKU/负责人/国家或渠道/PCS per CTN/箱数/总数量/日期` are stock-plan or shipment-planning data, not a complete invoice source. Reply with the missing inputs instead of generating a workbook. For example, say that 宝通达 invoice generation still needs the Saihu export or, if creating from template, the FBA box-number sequence, Reference ID per shipment, warehouse/address fields, and `发票产品详情` product facts.

If the user provides a half-finished invoice plus a product/shipment list but no separate `发票产品详情` source, fetch the pinned Tencent Docs sheet through MCP and use it as the product-detail source. If MCP/export is unavailable, inspect the workbook and report what can be fixed from the invoice itself. Do not fill missing unit price, brand, product links, product images, HS/material/purpose corrections, or other product facts from unrelated local files. Ask for Tencent Docs access repair or a current product-detail export before producing a finished repaired workbook.

If the user later provides a known-good completed invoice for comparison, treat it as a reference unless they explicitly say it is the current data source. Use it to report differences and to improve formatting rules; do not retroactively treat it as proof that earlier unprovided product data was authorized.

## Final Checks

Before returning the repaired workbook:

- Confirm the original export was copied and not overwritten.
- Confirm item-line count before/after repair.
- Confirm `总数量(PCS)` across visible detail rows equals the exported actual packing/declaration quantity when that source value is provided.
- For 海光 invoices, confirm `服务*` and `报关方式*` match the carrier-specific final values, e.g. `服务*=普船海卡` and `报关方式*=报关退税`.
- Confirm template country-code fields that mention `国家代码` or `二字代码` use ISO two-letter codes such as `US`, while carrier templates that request country names retain the carrier-required wording.
- Confirm any detail rows beyond the first 10 valid detail rows have the same row height and visible cell formatting as the normal sample rows before accepting the workbook.
- Confirm all required corrected fields are filled for each visible detail row.
- Confirm formulas have recalculated and no `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, or `#N/A` remains.
- Confirm text-number warnings are cleared for numeric detail columns, or that the script report has a nonzero/expected `textNumbersConverted` value.
- Confirm customs/HS-code cells are stored with Excel text format `@`, not general or numeric/scientific formatting.
- Confirm `总价值` equals `单价 * 总数量` where the template uses those fields.
- Confirm any image requirement is either satisfied, explicitly marked for manual insertion, or disclosed as missing. For embedded images, verify one nonzero-size image per visible detail row, verify every image center and edge stays inside its target cell border using actual `Left/Top/Width/Height` geometry, and visually preview the product-image column when the user has reported image overflow or sizing issues.
- Provide a compact audit summary: source workbook, selected sheet, user-requested work-scope rows/range when used, detail rows repaired, product-detail source with provenance, completed reference if used, unresolved fields, and formula errors.
