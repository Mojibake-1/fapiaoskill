# Carrier Template Map

Use this reference to identify the carrier template and understand which fields the original template consumes. The repair workflow usually starts from a Saihu export, so this map is mainly for carrier recognition and for checking whether an exported workbook still matches the expected template family.

## Create-From-Template Gate

Carrier templates are not enough to make a finished invoice. Use them without a Saihu export only when the user explicitly requests template-based creation and provides every required top-level token, every required item token, and the product facts normally sourced from `发票产品详情`.

If the user provides only a shipment/product list with product names, SKUs, PCS/CTN, cartons, total quantity, country, and dates, stop and ask for the Saihu export or the missing template fields. Do not generate placeholder invoices from that list.

For 宝通达 specifically, a pasted 47件 product list still lacks the finished-invoice sources unless it also includes:

- FBA编号 for every physical carton, or the exact shipment IDs and box-number expansion rule approved by operations
- Reference ID for each shipment group
- recipient country, postal code, address, city, province/state, and FBA warehouse code
- per-line dimensions, gross weight, HS code, declared Chinese/English names, brand, material, purpose, unit price, product link, and product image source
- confirmation that row order and carton allocation match the intended Saihu export

When a half-finished 宝通达 invoice is provided, it can satisfy the shipment/address side of this gate if it contains the header fields and FBA box rows. It still does not satisfy the product-detail side when unit price, brand, links, images, HS/material/purpose corrections, or product names need to be overwritten. Require the current `发票产品详情` source before finishing those fields.

When a known-good 宝通达 completed invoice is provided as a strict visual reference for the same shipment, align presentation to it without treating it as product-data authority:

- phone cell `H8` should follow the reference value and number format, commonly `0` rather than `00000000000`
- if the reference has only `发票信息` and `国家`, report or remove extra sheets such as `关联单据` when strict matching is requested
- if the reference preserves a blank formatted tail area after the last valid row, do not assume physical row deletion is the desired final layout
- preserve reference line breaks in material/purpose cells when strict matching is requested

## Selection Rules

1. Match the AMZ stock-plan `物流商及渠道` by carrier keyword.
2. Preserve the full channel/service text in the audit summary.
3. If multiple templates could match, ask the user to confirm.
4. `凯鑫` requires country/channel disambiguation:
   - UK / 英国 -> `kaixin-uk-invoice-template.xlsx`
   - DE / 德国 / 欧洲德国渠道 -> `kaixin-de-invoice-template.xlsx`

## Templates

| Carrier keyword | Asset | Main sheet | Item rows | Required top-level tokens | Required item tokens |
| --- | --- | --- | --- | --- | --- |
| 万逊通 | `assets/templates/wanxuntong-invoice-template.xls` | `发票模板` | 19-38 | `fbaFulfillmentCenter`, `boxCount` | `fbaBoxNo`, `referenceId`, `declareNameEn`, `declareNameCh`, `boxPcs`, `boxWeight`, `boxSizeLength`, `boxSizeWidth`, `boxSizeHeight`, `hsCode`, `material`, `purpose`, `salesLink` |
| 中南通 | `assets/templates/zhongnantong-invoice-template.xls` | `出货信息表` | 22-40 | `fbaShipmentId`, `fbaFulfillmentCenter`, `referenceId` | `fbaBoxNo`, `declareNameEn`, `declareNameCh`, `hsCode`, `material`, `purpose`, `boxPcs`, `boxWeight`, `boxSizeLength`, `boxSizeWidth`, `boxSizeHeight`, `brand`, `salesLink` |
| 华洋达 | `assets/templates/huayangda-invoice-template.xlsx` | `Sheet1` | 18-37 | `fbaFulfillmentCenter`, `recipientCountry`, `boxCount`, `referenceId`, `fbaShipmentId`, `recipientAllAddress` | `declareNameCh`, `declareNameEn`, `material`, `purpose`, `hsCode`, `brand`, `boxPcs`, `boxSizeLength`, `boxSizeWidth`, `boxSizeHeight`, `boxWeight`, `salesLink` |
| 大森林 | `assets/templates/dasenlin-invoice-template.xlsx` | `装箱单及发票（清关用）` | 7-19 | `fbaShipmentId`, `referenceId`, `recipientAllAddress` | `declareNameEn`, `declareNameCh`, `hsCode`, `material`, `purpose`, `boxPcs`, `boxWeight`, `totalVolume` |
| 宁致 | `assets/templates/ningzhi-invoice-template.xlsx` | `发票填写模板` | 25-43 | `fbaFulfillmentCenter`, `recipientAddress1`, `recipientCity`, `recipientStateProvince`, `recipientPostalCode`, `recipientCountry`, `referenceId`, `boxCount` | `fbaBoxNo`, `boxWeight`, `sku`, `declareNameEn`, `declareNameCh`, `boxPcs`, `boxSizeLength`, `boxSizeWidth`, `boxSizeHeight`, `material`, `purpose`, `hsCode`, `brand`, `salesLink` |
| 宝通达 | `assets/templates/baotongda-invoice-template.xlsx` | `发票信息` | 12-21 | `referenceId`, `boxCount`, `recipientPostalCode`, `recipientAddress1`, `recipientCity`, `recipientStateProvince`, `fbaFulfillmentCenter` | `fbaBoxNo`, `boxSizeLength`, `boxSizeWidth`, `boxSizeHeight`, `boxWeight`, `hsCode`, `declareNameCh`, `declareNameEn`, `boxPcs`, `brand`, `material`, `purpose` |
| 德国凯鑫 | `assets/templates/kaixin-de-invoice-template.xlsx` | `发票` | 28-39 | `fbaFulfillmentCenter`, `fbaShipmentId`, `boxCount`, `referenceId` | `fbaBoxNo`, `declareNameEn`, `declareNameCh`, `boxPcs`, `boxWeight`, `boxSizeLength`, `boxSizeWidth`, `boxSizeHeight`, `hsCode`, `brand`, `material`, `purpose`, `salesLink`, `asin` |
| 英国凯鑫 | `assets/templates/kaixin-uk-invoice-template.xlsx` | `发票` | 28-39 | `fbaFulfillmentCenter`, `fbaShipmentId`, `boxCount`, `referenceId` | `fbaBoxNo`, `declareNameEn`, `declareNameCh`, `boxPcs`, `boxWeight`, `boxSizeLength`, `boxSizeWidth`, `boxSizeHeight`, `hsCode`, `brand`, `material`, `purpose`, `salesLink`, `asin` |
| 德鑫淼 | `assets/templates/dexinmiao-invoice-template.xlsx` | `发票模板` | 20-34 | `fbaShipmentId`, `referenceId`, `boxCount`, `fbaFulfillmentCenter`, `recipientAddress1`, `recipientCity`, `recipientPostalCode`, `recipientStateProvince` | `fbaBoxNo`, `declareNameCh`, `declareNameEn`, `boxPcs`, `boxWeight`, `hsCode`, `brand`, `brandType`, `material`, `purpose`, `salesLink` |
| 斑马 | `assets/templates/banma-invoice-template.xlsx` | `美.加专线发票` | 22-41 | `fbaShipmentId`, `fbaFulfillmentCenter`, `recipientAddress1`, `recipientCity`, `recipientStateProvince`, `recipientCountry`, `recipientPostalCode`, `referenceId` | `fbaBoxNo`, `declareNameCh`, `declareNameEn`, `material`, `purpose`, `hsCode`, `boxPcs`, `boxWeight`, `boxSizeLength`, `boxSizeWidth`, `boxSizeHeight` |
| 海光 | `assets/templates/haiguang-invoice-template.xls` | `发票模板` | 16-35 | `fbaFulfillmentCenter`, `recipientAddress1`, `recipientCity`, `recipientStateProvince`, `recipientPostalCode`, `recipientCountry`, `boxCount`, `referenceId` | `fbaBoxNo`, `boxWeight`, `boxSizeLength`, `boxSizeWidth`, `boxSizeHeight`, `declareNameCh`, `declareNameEn`, `boxPcs`, `material`, `purpose`, `hsCode`, `brand`, `brandType`, `salesLink` |
| 阿拉丁 | `assets/templates/aladdin-invoice-template.xlsx` | `Sheet1` | 6-18 | `fbaShipmentId`, `referenceId`, `fbaFulfillmentCenter`, `country` | `declareNameEn`, `declareNameCh`, `material`, `purpose`, `brand`, `boxPcs`, `boxWeight`, `hsCode`, `boxSizeLength`, `boxSizeWidth`, `boxSizeHeight`, `salesLink`, `asin` |

## Placeholder Notes

- Top-level placeholders use `{field}`.
- Item placeholders use `{.field}`.
- The fill script treats `{.referenceId}` as item-level first, then falls back to top-level `referenceId`.
- If `boxCount` is omitted, the fill script uses `items.length`.
- If `totalVolume` is omitted, the fill script derives it from dimensions when possible.
- If `recipientAllAddress` is omitted, the fill script joins recipient address components.

## Common Ambiguities

- `物流商及渠道` may include service details such as 海卡, 卡派, 自税, 包税, or country-specific terms. Use those terms for audit notes, not as a reason to switch carrier unless they disambiguate 凯鑫.
- Some templates include blank unit-price columns but formula-based total columns. If the workbook shows formula errors after filling, request the missing unit price or confirm the file is only a draft.
- Some templates have built-in address-code sheets. Do not rewrite those dictionary sheets unless the user provides a new official template.
