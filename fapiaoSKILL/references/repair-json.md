# Repair JSON Format

Use this JSON file with `scripts/repair_exported_invoice.ps1`.

```json
{
  "sourceProvenance": {
    "invoiceWorkbookUserProvided": true,
    "productDetailUserProvided": true,
    "productDetailWorkbookPath": "C:\\path\\to\\发票产品详情.xlsx",
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
  "formatRows": [
    {
      "sourceRow": 21,
      "targetStartRow": 22,
      "targetEndRow": 24
    }
  ],
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
| `formulaRules` | optional | Header-targeted formulas to write across detail rows |
| `numericHeaders` | optional | Headers where number-like text should be converted to real numbers before formulas recalculate. If omitted, the script uses a conservative default list for quantity, weight, unit price, and total fields |
| `formatRows` | optional | Row-format repair rules. Copies formats and row height from `sourceRow` to `targetStartRow` through `targetEndRow` without copying values |
| `reference` | optional | Completed invoice workbook used as formatting reference |
| `imageSource` | optional | Online product-detail workbook used for copying product pictures |
| `rows` | optional | Row-specific corrections |

## Source Provenance Gate

Before running `repair_exported_invoice.ps1` for a production repair, include `sourceProvenance` in the corrections JSON and check it manually:

```json
{
  "sourceProvenance": {
    "invoiceWorkbookUserProvided": true,
    "productDetailUserProvided": true,
    "productDetailWorkbookPath": "C:\\Users\\admin\\Downloads\\发票产品详情.xlsx",
    "completedReferenceUserProvided": false,
    "completedReferenceRole": "none"
  }
}
```

If row corrections contain product facts that are not already reliable in the Saihu export, `productDetailUserProvided` must be `true`. A generic local file found in `Downloads`, a prior `.analysis` file, or a completed invoice reference does not satisfy this gate unless the user explicitly names or confirms it for the current task. If the gate is not satisfied, stop and ask for the current `发票产品详情` source instead of generating a finished workbook.

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

Use `formatRows` when tail rows lose borders, row height, wrapping, or number formatting after a long export or row deletion.

```json
{
  "formatRows": [
    { "sourceRow": 21, "targetStartRow": 22, "targetEndRow": 24 }
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

The script copies the first visible picture anchored to `sourceRow` and `sourceImageColumn`, then places it as a floating picture centered inside the target image cell. It does not use Excel's in-cell picture mode. It shrinks oversized pictures to stay within the cell borders and does not enlarge smaller pictures.

## Value Types

- Strings and numbers are written with Excel `Value2`.
- Values beginning with `=` are written as formulas.
- Empty string `""` clears a cell.
- Use header names as they appear in the invoice detail header row.

## Text Numbers

The script automatically performs the WPS/Excel equivalent of "批量转换为数字" for `numericHeaders` before applying formula rules. Use this for math/input columns such as `箱数`, `单箱毛重(KGS)`, `总数量(PCS)`, `商品单重(KG)`, `单价`, and `总价值`.

Do not include identifier/code columns such as FBA number, Reference ID, ASIN, SKU, or HS code unless the carrier/template owner explicitly requires those cells to be numeric. Keeping identifiers as text avoids losing leading zeroes or changing code formatting.

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
