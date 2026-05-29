# Repair JSON Format

Use this JSON file with `scripts/repair_exported_invoice.ps1`.

```json
{
  "sourceProvenance": {
    "invoiceWorkbookUserProvided": true,
    "productDetailSource": "tencent-docs",
    "productDetailFileId": "DY0hGbmd2Q1ZTVFVD",
    "productDetailSheetId": "BB08J2",
    "productDetailWorkbookPath": "C:\\path\\to\\product-detail-latest.xlsx",
    "productDetailMetadataPath": "C:\\path\\to\\product-detail-latest.metadata.json",
    "completedReferenceUserProvided": true,
    "completedReferenceRole": "format-reference-only"
  },
  "sheetName": "发票信息",
  "headerRow": 11,
  "detailStartRow": 12,
  "detailEndRow": 33,
  "truncateByQuantity": {
    "enabled": true,
    "quantityHeader": "总数量(PCS)",
    "expectedTotal": 453,
    "deleteRowsAfterMatch": true
  },
  "requiredHeaders": [
    "FBA编号",
    "Reference ID",
    "箱数",
    "海关编码HSCODE",
    "中文品名",
    "英文品名",
    "总数量(PCS)",
    "单价",
    "总价值",
    "材质中英文",
    "用途中英文"
  ],
  "defaults": {
    "Reference ID": "5QGEYPUS",
    "箱数": 1,
    "是否已注册": "无",
    "规格型号": "无",
    "是否带电": "否",
    "是否带磁": "否",
    "是否是液体": "否",
    "是否是膏状": "否",
    "是否粉末": "否"
  },
  "forceValues": {
    "产品品牌*": "MXZONE",
    "品牌类型*": "境外品牌"
  },
  "formulaRules": [
    {
      "target": "总价值",
      "formulaR1C1": "=RC[-1]*RC[-3]",
      "onlyWhenBlank": false
    }
  ],
  "numericHeaders": [
    "箱数",
    "单箱毛重(KGS)",
    "总数量(PCS)",
    "商品单重(KG)",
    "单价",
    "总价值"
  ],
  "textHeaders": [
    "海关编码HSCODE",
    "产品海关编码*"
  ],
  "autoFormatRows": {
    "enabled": true,
    "sampleDetailRows": 10
  },
  "reference": {
    "workbookPath": "C:\\path\\to\\完成版发票.xlsx",
    "sheetName": "发票"
  },
  "imageSource": {
    "workbookPath": "C:\\path\\to\\发票产品详情.xlsx",
    "sheetName": "FBA专线出货资料模块",
    "sourceImageColumn": "图片",
    "targetImageColumn": "产品图片"
  },
  "rows": [
    {
      "row": 12,
      "sourceRow": 4,
      "copyImage": true,
      "sourceKey": "online row 2 / ASIN B0C3HB14SC",
      "values": {
        "中文品名": "吸尘器配件",
        "英文品名": "Vacuum accessories",
        "海关编码HSCODE": "8508700090",
        "单价": 1.7,
        "单箱毛重(KGS)": 14.3,
        "材质中英文": "Hepa/海帕滤纸",
        "用途中英文": "for Vacuum/吸尘器过滤配件"
      }
    }
  ]
}
```

## Top-Level Fields

| Field | Required | Meaning |
| --- | --- | --- |
| `sourceProvenance` | required for production repairs | Auditable source gate. Use it to record whether the invoice workbook, product-detail source, and completed reference were explicitly provided or confirmed by the user for the current task |
| `sheetName` | optional | Invoice sheet name. If omitted, the first worksheet is used |
| `headerRow` | optional | Row containing detail headers. If omitted, the script tries to find it |
| `detailStartRow` | optional | First detail row. Defaults to `headerRow + 1` |
| `detailEndRow` | optional | Last detail row. If omitted, the script stops at the last nonblank row in the first detail column |
| `truncateByQuantity` | optional | Finds the valid final detail row by summing a quantity column until `expectedTotal`; can delete later overrun rows |
| `requiredHeaders` | optional | Headers that must be nonblank after repair |
| `defaults` | optional | Header-value pairs to fill on blank detail cells |
| `forceValues` | optional | Header-value pairs to overwrite across every final detail row after row corrections. Use this for required constants such as `产品品牌*=MXZONE` / `产品品牌*=Ucoolbe` and `品牌类型=境外品牌` / `品牌类型*=境外品牌` |
| `formulaRules` | optional | Header-targeted formulas to write across detail rows |
| `numericHeaders` | optional | Headers where number-like text should be converted to real numbers before formulas recalculate. If omitted, the script uses a conservative default list for quantity, weight, unit price, and total fields |
| `textHeaders` | optional | Headers that must be stored as Excel text format `@` before final save. If omitted, the script uses default customs/HS-code headers such as `海关编码HSCODE`, `海关编码`, `产品海关编码`, and `产品海关编码*` |
| `autoFormatRows` | optional | Automatically copies row height and formats from the first normal detail rows to later valid detail rows before values and pictures are filled. Defaults to enabled with `sampleDetailRows: 10` when `formatRows` is not supplied |
| `formatRows` | optional | Explicit row-format repair rules. Copies formats and row height from `sourceRow` to `targetStartRow` through `targetEndRow` without copying values. Overrides `autoFormatRows` when supplied |
| `reference` | optional | Completed invoice workbook used as formatting reference |
| `imageSource` | optional | Online product-detail workbook used for copying product pictures |
| `rows` | optional | Row-specific corrections |

## Source Provenance Gate

Before running `repair_exported_invoice.ps1` for a production repair, include `sourceProvenance` in the corrections JSON and check it manually:

```json
{
  "sourceProvenance": {
    "invoiceWorkbookUserProvided": true,
    "productDetailSource": "tencent-docs",
    "productDetailFileId": "DY0hGbmd2Q1ZTVFVD",
    "productDetailSheetId": "BB08J2",
    "productDetailWorkbookPath": "C:\\path\\to\\product-detail-latest.xlsx",
    "productDetailMetadataPath": "C:\\path\\to\\product-detail-latest.metadata.json",
    "completedReferenceUserProvided": false,
    "completedReferenceRole": "none"
  }
}
```

If row corrections contain product facts that are not already reliable in the Saihu export, `productDetailSource` must be either `tencent-docs` for the pinned online sheet or `user-provided` for an explicitly supplied current replacement. A generic local file found in `Downloads`, a stale `.analysis` file, a previous corrections JSON, or a completed invoice reference does not satisfy this gate. For the Tencent Docs default, prefer the helper script output and metadata; if it returns `skippedDownload=true`, the existing local workbook is acceptable because its metadata matches the current online document. If the gate is not satisfied, stop and ask for Tencent Docs access repair or a current `发票产品详情` replacement instead of generating a finished workbook.

## Row Corrections

Each row correction can target a physical row number:

```json
{ "row": 12, "values": { "单价": 1.7 } }
```

To copy a product picture from the online product-detail workbook, include `sourceRow` and `copyImage`:

```json
{
  "row": 12,
  "sourceRow": 19,
  "copyImage": true,
  "values": { "单价": 4 }
}
```

Or the first row matching existing cell values:

```json
{
  "match": { "FBA编号": "FBA19BCVR3C7U000001" },
  "values": { "单价": 1.7 }
}
```

If the same `FBA编号` appears multiple times, prefer explicit `row` numbers unless the match object includes additional fields such as `总数量(PCS)` or `中文品名`.

## Force Values

Use `forceValues` for fields that must be uniform across the final detail range. Unlike `defaults`, these values overwrite nonblank exported or row-correction values.

For any invoice template that has a product-brand detail header such as `产品品牌`, `产品品牌*`, `品牌`, or `Brands 品牌`, derive the forced brand from the current shipment's `发货站点`: contains `uc` case-insensitively -> `Ucoolbe`; otherwise -> `MXZONE`.

For any invoice template that has a `品牌类型` or `品牌类型*` detail header, include the matching header name with `境外品牌`:

```json
{
  "forceValues": {
    "产品品牌*": "MXZONE",
    "品牌类型*": "境外品牌"
  }
}
```

## Quantity Truncation

Use `truncateByQuantity` when the Saihu page/export shows the shipment's actual packing/declaration quantity and the workbook contains extra repeated rows.

```json
{
  "truncateByQuantity": {
    "enabled": true,
    "quantityHeader": "总数量(PCS)",
    "expectedTotal": 453,
    "deleteRowsAfterMatch": true
  }
}
```

The script sums `quantityHeader` from `detailStartRow` downward. When the running total equals `expectedTotal`, that row becomes the final valid detail row. If `deleteRowsAfterMatch` is true, later rows through the original `detailEndRow` are deleted.

## Row Format Repairs

The script applies row formatting before writing defaults, row corrections, formulas, or product pictures so images are inserted after the row heights are already correct.

Preferred default: omit `formatRows` and let `autoFormatRows` use the first 10 valid detail rows as the normal style sample. This avoids hardcoding a template-specific cutoff row:

```json
{
  "autoFormatRows": {
    "enabled": true,
    "sampleDetailRows": 10
  }
}
```

Use explicit `formatRows` only when inspection shows a better source row for that template:

```json
{
  "formatRows": [
    { "sourceRow": 25, "targetStartRow": 26, "targetEndRow": 42 }
  ]
}
```

This copies formats and row height only. It does not copy values or product pictures.

## Completed Reference Formatting

The optional `reference` object can also be supplied through the command line with `-ReferenceWorkbookPath`. The script copies formatting, column widths, and row heights from the reference sheet to matching cells in the repaired workbook. Use it only when the reference is the same carrier/template family or the user explicitly wants that completed layout.

## Image Source

The optional `imageSource` object opens the online product-detail workbook with Excel COM and copies picture shapes from the source row to the target invoice row.

| Field | Meaning |
| --- | --- |
| `workbookPath` | Path to `发票产品详情.xlsx` |
| `sheetName` | Usually `FBA专线出货资料模块` |
| `sourceImageColumn` | Source image column header or column number, usually `图片` / column `T` |
| `targetImageColumn` | Target image column header or column number, such as `产品图片*`, `产品图片`, or `产品图片上传` |

The script copies the first visible picture anchored to `sourceRow` and `sourceImageColumn`, then places it as a floating picture centered inside the target image cell. It removes any pre-existing shape in the target image cell, does not use Excel's in-cell picture mode, and prefers `CopyPicture -> temporary PNG -> Shapes.AddPicture` over raw clipboard paste so WPS/Excel COM does not collapse pictures to `0 x 0`. After copying all images, it runs a second geometry pass and refits any image whose center or edge falls outside the target cell.

Optional sizing override:

```json
{
  "imageSizing": {
    "margin": 3,
    "maxWidthRatio": 0.9,
    "maxHeightRatio": 0.9
  }
}
```

Defaults are `margin: 3`, `maxWidthRatio: 0.9`, and `maxHeightRatio: 0.9`. The script preserves source aspect ratio and sizes to the largest fitted rectangle inside the target cell or merge-area bounds.

## Value Types

- Strings and numbers are written with Excel `Value2`.
- Values beginning with `=` are written as formulas.
- Empty string `""` clears a cell.
- Use header names as they appear in the invoice detail header row.

## Text Numbers

The script automatically performs the WPS/Excel equivalent of "批量转换为数字" for `numericHeaders` before applying formula rules. Use this for math/input columns such as `箱数`, `单箱毛重(KGS)`, `总数量(PCS)`, `商品单重(KG)`, `单价`, and `总价值`.

Do not include identifier/code columns such as FBA number, Reference ID, ASIN, SKU, or HS code unless the carrier/template owner explicitly requires those cells to be numeric. Keeping identifiers as text avoids losing leading zeroes or changing code formatting.

## Text Code Columns

The script formats customs/HS-code columns as Excel text format `@` before final save, then rewrites nonblank code values as strings. This prevents codes from appearing as scientific notation after copy/paste or import elsewhere.

If the template uses a nonstandard customs-code header, add it to `textHeaders`:

```json
{
  "textHeaders": [
    "海关编码HSCODE",
    "产品海关编码*"
  ]
}
```

## Formula Rules

`formulaRules[].formulaR1C1` is written with Excel R1C1 notation. This avoids hardcoding column letters across carrier templates.

Example for 宝通达 `总价值 = 单价 * 总数量`:

```json
{
  "target": "总价值",
  "formulaR1C1": "=RC[-1]*RC[-3]"
}
```

If `onlyWhenBlank` is `true`, the script writes the formula only to blank cells.
