# Online Product Detail Schema

The second source for invoice repair is the online `发票产品详情` table. By default, pull the latest table from Tencent Docs through MCP instead of asking the user to provide a file. Prefer an exported `.xlsx` version when product images must be copied. Screenshots can only support visible rows and visible columns.

## Default Tencent Docs Source

Use this source unless the user explicitly provides a different current product-detail source:

| Field | Value |
| --- | --- |
| URL | `https://docs.qq.com/sheet/DY0hGbmd2Q1ZTVFVD?tab=BB08J2` |
| `file_id` | `DY0hGbmd2Q1ZTVFVD` |
| `sheet_id` / tab | `BB08J2` |
| expected title | `发票产品明细` |

Use Tencent Docs MCP commands in this form because the tool names contain dots:

```powershell
mcporter call --server tencent-docs --tool "manage.query_file_info" file_id=DY0hGbmd2Q1ZTVFVD --output json
mcporter call --server tencent-docs --tool "sheet.get_sheet_info" file_id=DY0hGbmd2Q1ZTVFVD --output json
mcporter call --server tencent-docs --tool "sheet.get_cell_data" file_id=DY0hGbmd2Q1ZTVFVD sheet_id=BB08J2 start_row=0 start_col=0 return_csv=true --output json
```

For image-copy workflows, export the sheet and download the `.xlsx`:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/fetch_product_detail_from_tencent_docs.ps1
```

The script stores a metadata sidecar beside the downloaded workbook. If the online `last_modify_time`, `file_id`, `sheet_id`, and sheet row/column counts still match the local metadata and the local workbook SHA-256 still matches the metadata, the script skips the export/download and returns `skippedDownload=true`. Use `-Force` to refresh deliberately.

The downloaded workbook path should be used as `imageSource.workbookPath`. If the export fails or lacks embedded images, ask for Tencent Docs access repair or a current exported workbook; do not use an old local copy.

## Observed Table Shape

The online sheet uses two header rows. The important visible columns are:

| Column | Header | Meaning |
| --- | --- | --- |
| A | 老版品名 / 中文品名 | old Chinese product name |
| B | 老版品名 / 英文品名 | old English product name |
| C | NEW品名 / 中文 | corrected Chinese product name |
| D | NEW品名 / 英文 | corrected English product name |
| E | 数量 | quantity or pack count text as maintained by operations |
| F | 装箱数 / 单箱 | units per box |
| G | 售价 / 单价 | sales price |
| H | 单价（美国） / USD | invoice unit price for USD templates |
| I | 加币（申报） | CAD declared value or Canada declaration value |
| J | 净重 / KGS | net weight |
| K | 毛重 / KGS | gross weight |
| L | 长 | box length cm |
| M | 宽 | box width cm |
| N | 高 | box height cm |
| O | 备用（材质分开版） / 材质(中文) | material Chinese |
| P | 备用（材质分开版） / 材质(英文) | material English |
| Q | 必填 / 材质(中英文) | material bilingual value |
| R | 必填 / 用途(中英文) | purpose bilingual value |
| S | 商品编码 必写 / HS CODE | HS code |
| T | 必填 / 图片 | product image |
| U | 产品名称 | product display name |
| V | ASIN | ASIN |

## Normalized Fields

When preparing repair data, normalize product rows into these keys:

```json
{
  "oldNameCh": "滚刷+滤网",
  "oldNameEn": "Brush roll+Filter",
  "declareNameCh": "吸尘器配件",
  "declareNameEn": "Vacuum accessories",
  "quantityText": "2+2",
  "boxPcs": 49,
  "salePrice": 5,
  "unitPriceUsd": 5,
  "declaredCad": 19.4,
  "netWeight": 19.4,
  "grossWeight": 20.4,
  "boxSizeLength": 51,
  "boxSizeWidth": 42,
  "boxSizeHeight": 42,
  "materialCh": "超细纤维/海帕滤纸",
  "materialEn": "Microfiber/Hepa",
  "material": "超细纤维/Microfiber Hepa/海帕滤纸",
  "purpose": "拖地/Mopping Replacement for Vacuum/吸尘器过滤配件",
  "hsCode": "8508709000",
  "image": "image present or file path",
  "productName": "3882滚刷2+2",
  "asin": "B0DHKYL6T5"
}
```

## Matching Guidance

Prefer direct keys in this order:

1. ASIN, if the exported invoice or user-provided correction includes it.
2. Product code/name from the online table if present in the exported invoice.
3. Exact old Chinese/English name plus dimensions/quantity.
4. Row order only when the user confirms both sources are sorted the same way.

If multiple online rows can match one invoice row, show the candidates and ask the user to choose. Do not silently pick a row when HS code, material, dimensions, or unit price differ.

## Image Handling

In the observed Tencent Docs export, product images are embedded Excel picture shapes anchored to the `图片` column, usually column `T`. The workbook may not load cleanly with openpyxl because of online/WPS style records, so use Excel COM or direct package parsing for image work.

For each matched online row, keep the source worksheet row number. The repair script can copy the picture shape by using:

- `imageSource.sourceImageColumn`: `图片`
- row correction `sourceRow`: the online product-detail row
- row correction `copyImage`: `true`

Do not replace picture requirements with text such as `image present`. The final invoice must contain a real floating picture placed inside the target picture cell when the completed reference/template shows one. Do not use Excel's in-cell picture mode. Do not let pictures exceed cell borders; shrink oversized pictures to fit and do not enlarge smaller pictures.

## Field Authority

Use online table values to correct:

- Chinese/English product names
- unit price and declared value
- net/gross weight
- dimensions
- material CN/EN and combined material text
- purpose text
- HS code
- product image reference
- ASIN/product display name when the template supports it

Use the Saihu export values to preserve:

- FBA box number
- Reference ID unless blank or obviously not filled down
- recipient address and warehouse code
- base carrier template and workbook layout
