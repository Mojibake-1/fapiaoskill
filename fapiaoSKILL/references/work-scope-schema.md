# Invoice Work-Scope Schema

The default work-scope source tells Codex which shipment/planning rows the user wants to make invoices for. It is not the product-detail authority and it is not a replacement for the Saihu export.

## Default Tencent Docs Source

| Field | Value |
| --- | --- |
| URL | `https://docs.qq.com/sheet/DRE1ZTlhoZVZBVkdL?tab=000001` |
| URL file token | `DRE1ZTlhoZVZBVkdL` |
| `sheet_id` / tab | `000001` |
| expected title | `AMZ备货计划及出货安排表` |
| expected sheet name | `备货详情` |
| observed header row | row 3 |
| observed first data row | row 4 |

Use Tencent Docs MCP commands in this form:

```powershell
mcporter call --server tencent-docs --tool "manage.query_file_info" file_id=DRE1ZTlhoZVZBVkdL --output json
mcporter call --server tencent-docs --tool "sheet.get_sheet_info" file_id=DRE1ZTlhoZVZBVkdL --output json
mcporter call --server tencent-docs --tool "sheet.get_cell_data" file_id=DRE1ZTlhoZVZBVkdL sheet_id=000001 start_row=0 start_col=0 return_csv=true --output json
```

If a local export is useful:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/fetch_invoice_work_scope_from_tencent_docs.ps1
```

The helper stores a metadata sidecar and skips repeat download when the online document metadata and local workbook hash still match.

## Scope Selection Rule

The user must specify which part to do every time this sheet is used. Accept clear selectors such as:

- exact row numbers or row ranges from `备货详情`
- FBA shipment IDs from `FBA货件号`
- a carrier/channel plus exact product rows or SKU list
- another unambiguous selection that maps to one contiguous block or an explicit row set

If the user does not specify the part, or the selector can map to multiple rows/blocks, ask which rows or range to use. Do not process the whole sheet, all visible rows, latest rows, today's rows, first matching carrier, or all rows with the same shipping status by default.

## Observed Columns

The observed header row includes:

| Header | Meaning |
| --- | --- |
| 箱号参考 | box-reference marker |
| 参考 | reference group |
| 代码 | product code |
| 产品名称 | product name |
| SKU | SKU |
| 运营人员 | operator |
| 发货站点 | shipping site/country lane |
| PCS/CTN | units per carton |
| 件数 | carton count for the row |
| 总数 | total units |
| 计划发货时间 | planned shipping date |
| 实际发货时间 | actual shipping date |
| 预计到货时间 | estimated arrival date |
| 目的仓 | destination fulfillment center |
| FBA货件号 | FBA shipment ID |
| 物流商及渠道 | carrier and channel |
| 报关方式 | declaration mode |
| 总件数 | shipment total cartons |
| 入仓号 | inbound number |
| 司机信息 | driver information |
| 其它信息 | other information |
| 备注 | notes |
| 采购到货日期 | purchase arrival date |
| 打包情况 | packing status |
| 出货状态 | shipping status |

## Field Authority

Use this sheet to identify the requested work scope and to cross-check shipment planning facts: product rows, SKU, operator, site, carton counts, total quantity, shipping dates, destination warehouse, FBA shipment ID, carrier/channel, declaration method, inbound number, notes, packing status, and shipping status.

Do not use this sheet to invent or replace:

- official Saihu export workbook structure
- FBA box-number sequences
- Reference ID per box
- recipient address, city, state, postal code, or country fields
- per-box allocation details not present in the export
- product facts such as unit price, HS code, material, purpose, weight, dimensions, or images

Those still come from the Saihu export and `发票产品详情`.
