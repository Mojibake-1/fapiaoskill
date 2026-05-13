# AMZ Invoice SOP Summary

This reference summarizes `assets/amz-invoice-sop.xlsx`.

## Saihu Export Flow

1. In Saihu, go to `FBA` -> `FBA货件`.
2. Search by `货件编号`.
3. Select the shipment row.
4. Use `导出` -> `导出报关资料（按模板）`.
5. In the export dialog, choose the carrier-specific invoice/packing-list template.
6. Use the AMZ stock-plan workbook `AMZ备货计划及出货安排表` to match:
   - `FBA货件号`
   - `物流商及渠道`

## Export Dialog Settings Observed in SOP

- Template field used for invoices: `装箱单`.
- Merge dimension: `所选单据合并导出一票报关资料`.
- Product merge row: enabled, with `SKU` checked.
- Report pages: `1`.
- Export file: `分开导出并放在一个压缩包中`.

## Working Rule

The logistics company/channel in the AMZ stock plan determines which invoice template to select. If the logistics text is not an exact carrier name, match by carrier keyword first and preserve the channel/service text in the audit notes.

Example from SOP screenshot:

- AMZ stock-plan `FBA货件号`: `FBA15LF9YXFV`
- `物流商及渠道`: `万逊通海卡自税`
- Saihu export template: `英国万逊通发票`

## Data Sources

Prefer the freshest current-task export over old generated workbooks. If the user provides a screenshot, extract only visible fields and ask for missing required fields before final generation.
