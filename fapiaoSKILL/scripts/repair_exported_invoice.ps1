param(
    [Parameter(Mandatory=$true)]
    [string]$InvoicePath,

    [Parameter(Mandatory=$true)]
    [string]$CorrectionsPath,

    [Parameter(Mandatory=$true)]
    [string]$OutputPath,

    [string]$ReferenceWorkbookPath,
    [string]$ReferenceSheetName,

    [switch]$FailOnFormulaErrors,
    [switch]$Strict
)

$ErrorActionPreference = "Stop"

function Get-JsonProp($Object, [string]$Name, $Default = $null) {
    if ($null -eq $Object) { return $Default }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop) { return $Default }
    return $prop.Value
}

function Normalize-Text($Value) {
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function Normalize-Key($Value) {
    if ($null -eq $Value) { return "" }
    return (([string]$Value).Trim() -replace "\s+", "").ToLowerInvariant()
}

function Is-BlankCell($Cell) {
    $value = $Cell.Value2
    if ($null -eq $value) { return $true }
    return ([string]$value).Trim().Length -eq 0
}

function Set-ExcelCell($Cell, $Value) {
    if ($null -eq $Value) {
        $Cell.ClearContents() | Out-Null
        return
    }

    if ($Value -is [string] -and $Value.StartsWith("=")) {
        $Cell.Formula = $Value
        return
    }

    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or
        $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        $Cell.Value2 = [double]$Value
    } else {
        $Cell.Value2 = [string]$Value
    }
}

function Convert-NumericTextValue($Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -isnot [string]) { return $null }

    $text = ([string]$Value).Trim()
    if ($text.Length -eq 0) { return $null }
    if ($text -match "[`r`n]") { return $null }

    $normalized = $text -replace "[,\s]", ""
    if ($normalized -notmatch "^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$") { return $null }

    $number = 0.0
    $styles = [System.Globalization.NumberStyles]::Float
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    if ([double]::TryParse($normalized, $styles, $culture, [ref]$number)) {
        return $number
    }
    return $null
}

function Get-DefaultNumericHeaders() {
    return @(
        "箱数",
        "件数",
        "单箱毛重(KGS)",
        "毛重",
        "净重",
        "总数量(PCS)",
        "数量",
        "商品单重(KG)",
        "单价",
        "申报单价",
        "总价值",
        "总价",
        "金额"
    )
}

function Get-DefaultTextHeaders() {
    return @(
        "海关编码HSCODE",
        "海关编码",
        "海关编码*",
        "产品海关编码",
        "产品海关编码*",
        "HSCODE",
        "HS CODE",
        "HS Code",
        "HS编码",
        "HS编码*"
    )
}

function Convert-CodeValueToText($Value) {
    if ($null -eq $Value) { return "" }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or
        $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        $number = [double]$Value
        if ([Math]::Abs($number - [Math]::Round($number)) -lt 0.0000001) {
            return $number.ToString("0", [System.Globalization.CultureInfo]::InvariantCulture)
        }
        return $number.ToString("0.################", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return ([string]$Value).Trim()
}

function Format-TextColumns($Worksheet, [int]$StartRow, [int]$EndRow, $HeaderMap, $NormalizedHeaderMap, $Headers) {
    $formatted = 0
    $seenColumns = @{}

    foreach ($header in (Get-Array $Headers)) {
        $col = Resolve-HeaderColumn $HeaderMap $NormalizedHeaderMap $header
        if ($null -eq $col) { continue }
        if ($seenColumns.ContainsKey([string]$col)) { continue }
        $seenColumns[[string]$col] = $true

        $range = $Worksheet.Range($Worksheet.Cells.Item($StartRow, $col), $Worksheet.Cells.Item($EndRow, $col))
        $range.NumberFormat = "@"

        for ($row = $StartRow; $row -le $EndRow; $row++) {
            $cell = $Worksheet.Cells.Item($row, $col)
            $cell.NumberFormat = "@"
            if ([bool]$cell.HasFormula) { continue }
            if (Is-BlankCell $cell) { continue }

            $cell.Value2 = [string](Convert-CodeValueToText $cell.Value2)
            $formatted++
        }
    }

    return $formatted
}

function Convert-TextNumbersInColumns($Worksheet, [int]$StartRow, [int]$EndRow, $HeaderMap, $NormalizedHeaderMap, $Headers, $Warnings) {
    $converted = 0
    $seenColumns = @{}

    foreach ($header in (Get-Array $Headers)) {
        $col = Resolve-HeaderColumn $HeaderMap $NormalizedHeaderMap $header
        if ($null -eq $col) { continue }
        if ($seenColumns.ContainsKey([string]$col)) { continue }
        $seenColumns[[string]$col] = $true

        for ($row = $StartRow; $row -le $EndRow; $row++) {
            $cell = $Worksheet.Cells.Item($row, $col)
            if ([bool]$cell.HasFormula) { continue }
            $number = Convert-NumericTextValue $cell.Value2
            if ($null -eq $number) { continue }

            $cell.Value2 = [double]$number
            $converted++
        }
    }

    return $converted
}

function Get-CellNumber($Cell) {
    $value = $Cell.Value2
    if ($value -is [byte] -or $value -is [int16] -or $value -is [int32] -or $value -is [int64] -or
        $value -is [single] -or $value -is [double] -or $value -is [decimal]) {
        return [double]$value
    }

    $number = Convert-NumericTextValue $value
    if ($null -ne $number) { return [double]$number }
    return $null
}

function Get-Array($Value) {
    if ($null -eq $Value) { return @() }
    if ($Value -is [System.Array]) { return @($Value) }
    return @($Value)
}

function Resolve-HeaderColumn($HeaderMap, $NormalizedHeaderMap, [string]$HeaderName) {
    if ($HeaderMap.ContainsKey($HeaderName)) { return $HeaderMap[$HeaderName] }
    $key = Normalize-Key $HeaderName
    if ($NormalizedHeaderMap.ContainsKey($key)) { return $NormalizedHeaderMap[$key] }
    return $null
}

function Resolve-ColumnSpec($Worksheet, $HeaderMap, $NormalizedHeaderMap, $Spec, [int[]]$HeaderRows = @(1, 2)) {
    if ($null -eq $Spec) { return $null }
    if ($Spec -is [byte] -or $Spec -is [int16] -or $Spec -is [int32] -or $Spec -is [int64]) {
        return [int]$Spec
    }
    $text = Normalize-Text $Spec
    if ($text -match "^\d+$") { return [int]$text }

    if ($null -ne $HeaderMap -and $null -ne $NormalizedHeaderMap) {
        $col = Resolve-HeaderColumn $HeaderMap $NormalizedHeaderMap $text
        if ($null -ne $col) { return $col }
    }

    $used = $Worksheet.UsedRange
    $maxCol = [int]$used.Columns.Count
    $targetKey = Normalize-Key $text
    foreach ($row in $HeaderRows) {
        for ($col = 1; $col -le $maxCol; $col++) {
            if ((Normalize-Key $Worksheet.Cells.Item($row, $col).Text) -eq $targetKey) {
                return $col
            }
        }
    }
    return $null
}

function Build-HeaderMaps($Worksheet, [int]$HeaderRow) {
    $used = $Worksheet.UsedRange
    $maxCol = [int]$used.Columns.Count
    $headerMap = @{}
    $normalizedMap = @{}

    for ($col = 1; $col -le $maxCol; $col++) {
        $text = Normalize-Text $Worksheet.Cells.Item($HeaderRow, $col).Text
        if ($text.Length -eq 0) { continue }
        if (-not $headerMap.ContainsKey($text)) {
            $headerMap[$text] = $col
        }
        $key = Normalize-Key $text
        if ($key.Length -gt 0 -and -not $normalizedMap.ContainsKey($key)) {
            $normalizedMap[$key] = $col
        }
    }

    return @($headerMap, $normalizedMap)
}

function Find-HeaderRow($Worksheet, $RequiredHeaders) {
    $used = $Worksheet.UsedRange
    $maxRow = [Math]::Min([int]$used.Rows.Count, 100)
    $maxCol = [Math]::Min([int]$used.Columns.Count, 100)
    $required = Get-Array $RequiredHeaders

    if ($required.Count -gt 0) {
        $requiredKeys = @($required | ForEach-Object { Normalize-Key $_ })
        $bestRow = 0
        $bestScore = -1
        for ($row = 1; $row -le $maxRow; $row++) {
            $seen = @{}
            for ($col = 1; $col -le $maxCol; $col++) {
                $key = Normalize-Key $Worksheet.Cells.Item($row, $col).Text
                if ($key.Length -gt 0) { $seen[$key] = $true }
            }
            $score = 0
            foreach ($key in $requiredKeys) {
                if ($seen.ContainsKey($key)) { $score++ }
            }
            if ($score -gt $bestScore) {
                $bestScore = $score
                $bestRow = $row
            }
        }
        if ($bestScore -gt 0) { return $bestRow }
    }

    $fallbackRow = 1
    $fallbackCount = -1
    for ($row = 1; $row -le $maxRow; $row++) {
        $count = 0
        for ($col = 1; $col -le $maxCol; $col++) {
            if ((Normalize-Text $Worksheet.Cells.Item($row, $col).Text).Length -gt 0) {
                $count++
            }
        }
        if ($count -gt $fallbackCount) {
            $fallbackCount = $count
            $fallbackRow = $row
        }
    }
    return $fallbackRow
}

function Find-DetailEndRow($Worksheet, [int]$StartRow, [int]$FirstColumn) {
    $used = $Worksheet.UsedRange
    $maxRow = [int]$used.Rows.Count
    $last = $StartRow - 1
    for ($row = $StartRow; $row -le $maxRow; $row++) {
        if (-not (Is-BlankCell $Worksheet.Cells.Item($row, $FirstColumn))) {
            $last = $row
        }
    }
    return $last
}

function Resolve-DetailEndRowByQuantity($Worksheet, [int]$StartRow, [int]$EndRow, $HeaderMap, $NormalizedHeaderMap, $Rule, $Warnings) {
    $quantityHeader = Get-JsonProp $Rule "quantityHeader" "总数量(PCS)"
    $quantityColumn = Resolve-HeaderColumn $HeaderMap $NormalizedHeaderMap $quantityHeader
    if ($null -eq $quantityColumn) {
        $Warnings.Add("quantity truncate header not found: $quantityHeader") | Out-Null
        return $null
    }

    $expected = Get-JsonProp $Rule "expectedTotal"
    if ($null -eq $expected) {
        $Warnings.Add("quantity truncate skipped: expectedTotal missing") | Out-Null
        return $null
    }
    $expectedNumber = [double]$expected
    $tolerance = [double](Get-JsonProp $Rule "tolerance" 0.0001)
    $runningTotal = 0.0

    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $number = Get-CellNumber $Worksheet.Cells.Item($row, $quantityColumn)
        if ($null -eq $number) { continue }

        $runningTotal += [double]$number
        if ([Math]::Abs($runningTotal - $expectedNumber) -le $tolerance) {
            return [ordered]@{
                row = $row
                total = $runningTotal
                expected = $expectedNumber
            }
        }
        if ($runningTotal -gt ($expectedNumber + $tolerance)) {
            $Warnings.Add("quantity truncate exceeded expected total at row ${row}: $runningTotal > $expectedNumber") | Out-Null
            return [ordered]@{
                row = $null
                total = $runningTotal
                expected = $expectedNumber
            }
        }
    }

    $Warnings.Add("quantity truncate did not reach expected total $expectedNumber; accumulated $runningTotal through row $EndRow") | Out-Null
    return [ordered]@{
        row = $null
        total = $runningTotal
        expected = $expectedNumber
    }
}

function Delete-WorksheetRows($Worksheet, [int]$StartRow, [int]$EndRow) {
    if ($EndRow -lt $StartRow) { return 0 }
    for ($row = $EndRow; $row -ge $StartRow; $row--) {
        $Worksheet.Rows.Item($row).Delete() | Out-Null
    }
    return ($EndRow - $StartRow + 1)
}

function Find-MatchingRow($Worksheet, [int]$StartRow, [int]$EndRow, $HeaderMap, $NormalizedHeaderMap, $MatchObject, $Warnings) {
    $matches = @()
    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $ok = $true
        foreach ($prop in $MatchObject.PSObject.Properties) {
            $col = Resolve-HeaderColumn $HeaderMap $NormalizedHeaderMap $prop.Name
            if ($null -eq $col) {
                $Warnings.Add("match header not found: $($prop.Name)") | Out-Null
                $ok = $false
                break
            }
            $actual = Normalize-Text $Worksheet.Cells.Item($row, $col).Text
            $expected = Normalize-Text $prop.Value
            if ($actual -ne $expected) {
                $ok = $false
                break
            }
        }
        if ($ok) { $matches += $row }
    }

    if ($matches.Count -gt 1) {
        $Warnings.Add("multiple rows matched; using first row $($matches[0])") | Out-Null
    }
    if ($matches.Count -eq 0) { return $null }
    return $matches[0]
}

function Scan-FormulaErrors($Worksheet) {
    $patterns = @("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A")
    $used = $Worksheet.UsedRange
    $maxRow = [int]$used.Rows.Count
    $maxCol = [int]$used.Columns.Count
    $errors = @()

    for ($row = 1; $row -le $maxRow; $row++) {
        for ($col = 1; $col -le $maxCol; $col++) {
            $cell = $Worksheet.Cells.Item($row, $col)
            $text = Normalize-Text $cell.Text
            if ($patterns -contains $text) {
                $errors += [ordered]@{
                    address = $cell.Address($false, $false)
                    text = $text
                    formula = Normalize-Text $cell.Formula
                }
            }
        }
    }
    return $errors
}

function Scan-MissingRequired($Worksheet, [int]$StartRow, [int]$EndRow, $HeaderMap, $NormalizedHeaderMap, $RequiredHeaders) {
    $missing = @()
    foreach ($header in (Get-Array $RequiredHeaders)) {
        $col = Resolve-HeaderColumn $HeaderMap $NormalizedHeaderMap $header
        if ($null -eq $col) {
            $missing += [ordered]@{ header = $header; row = $null; address = $null; issue = "header not found" }
            continue
        }
        for ($row = $StartRow; $row -le $EndRow; $row++) {
            if (Is-BlankCell $Worksheet.Cells.Item($row, $col)) {
                $missing += [ordered]@{
                    header = $header
                    row = $row
                    address = $Worksheet.Cells.Item($row, $col).Address($false, $false)
                    issue = "blank"
                }
            }
        }
    }
    return $missing
}

function Apply-ReferenceFormats($SourceWorksheet, $TargetWorksheet) {
    $warnings = @()
    $sourceUsed = $SourceWorksheet.UsedRange
    $address = $sourceUsed.Address($false, $false)
    try {
        $targetRange = $TargetWorksheet.Range($address)
        $sourceUsed.Copy() | Out-Null
        $targetRange.PasteSpecial(-4122) | Out-Null # xlPasteFormats
    } catch {
        $warnings += "reference format paste skipped: $($_.Exception.Message)"
    }

    $sourceRows = [int]$sourceUsed.Rows.Count
    $sourceCols = [int]$sourceUsed.Columns.Count
    $targetRows = [int]$TargetWorksheet.UsedRange.Rows.Count
    $targetCols = [int]$TargetWorksheet.UsedRange.Columns.Count
    try {
        for ($col = 1; $col -le [Math]::Min($sourceCols, $targetCols); $col++) {
            $TargetWorksheet.Columns.Item($col).ColumnWidth = $SourceWorksheet.Columns.Item($col).ColumnWidth
        }
        for ($row = 1; $row -le [Math]::Min($sourceRows, $targetRows); $row++) {
            $TargetWorksheet.Rows.Item($row).RowHeight = $SourceWorksheet.Rows.Item($row).RowHeight
        }
    } catch {
        $warnings += "reference row/column sizing partially skipped: $($_.Exception.Message)"
    }
    return $warnings
}

function Apply-RowFormatRules($Worksheet, $Rules, $Warnings) {
    $applied = 0
    foreach ($rule in (Get-Array $Rules)) {
        $sourceRow = Get-JsonProp $rule "sourceRow"
        $targetStartRow = Get-JsonProp $rule "targetStartRow"
        $targetEndRow = Get-JsonProp $rule "targetEndRow" $targetStartRow
        if ($null -eq $sourceRow -or $null -eq $targetStartRow) {
            $Warnings.Add("format row rule skipped: sourceRow or targetStartRow missing") | Out-Null
            continue
        }

        $sourceRow = [int]$sourceRow
        $targetStartRow = [int]$targetStartRow
        $targetEndRow = [int]$targetEndRow
        for ($row = $targetStartRow; $row -le $targetEndRow; $row++) {
            $Worksheet.Rows.Item($sourceRow).Copy() | Out-Null
            $Worksheet.Rows.Item($row).PasteSpecial(-4122) | Out-Null # xlPasteFormats
            $Worksheet.Rows.Item($row).RowHeight = $Worksheet.Rows.Item($sourceRow).RowHeight
            $applied++
        }
        $Worksheet.Parent.Application.CutCopyMode = $false
    }
    return $applied
}

function Apply-ForceValues($Worksheet, [int]$StartRow, [int]$EndRow, $HeaderMap, $NormalizedHeaderMap, $Values, $Warnings) {
    if ($null -eq $Values) { return 0 }

    $applied = 0
    foreach ($prop in $Values.PSObject.Properties) {
        $col = Resolve-HeaderColumn $HeaderMap $NormalizedHeaderMap $prop.Name
        if ($null -eq $col) {
            $Warnings.Add("force value header not found: $($prop.Name)") | Out-Null
            continue
        }

        for ($row = $StartRow; $row -le $EndRow; $row++) {
            Set-ExcelCell $Worksheet.Cells.Item($row, $col) $prop.Value
            $applied++
        }
    }

    return $applied
}

function Apply-AutoDetailFormatRows($Worksheet, [int]$DetailStartRow, [int]$DetailEndRow, $Rule, $Warnings) {
    $enabled = $true
    if ($null -ne $Rule) {
        $enabled = [bool](Get-JsonProp $Rule "enabled" $true)
    }
    if (-not $enabled) { return 0 }

    $sampleDetailRows = 10
    if ($null -ne $Rule) {
        $sampleDetailRows = [int](Get-JsonProp $Rule "sampleDetailRows" $sampleDetailRows)
    }
    if ($sampleDetailRows -lt 1) {
        $Warnings.Add("autoFormatRows skipped: sampleDetailRows must be at least 1") | Out-Null
        return 0
    }

    $detailCount = $DetailEndRow - $DetailStartRow + 1
    if ($detailCount -le $sampleDetailRows) { return 0 }

    $sourceRow = $DetailStartRow + $sampleDetailRows - 1
    $targetStartRow = $sourceRow + 1
    if ($null -ne $Rule) {
        $sourceRow = [int](Get-JsonProp $Rule "sourceRow" $sourceRow)
        $targetStartRow = [int](Get-JsonProp $Rule "targetStartRow" $targetStartRow)
    }
    if ($sourceRow -lt $DetailStartRow -or $sourceRow -gt $DetailEndRow) {
        $Warnings.Add("autoFormatRows skipped: sourceRow $sourceRow outside detail range $DetailStartRow-$DetailEndRow") | Out-Null
        return 0
    }
    if ($targetStartRow -lt ($sourceRow + 1)) { $targetStartRow = $sourceRow + 1 }
    if ($targetStartRow -gt $DetailEndRow) { return 0 }

    $ruleObject = [pscustomobject]@{
        sourceRow = $sourceRow
        targetStartRow = $targetStartRow
        targetEndRow = $DetailEndRow
    }
    return Apply-RowFormatRules $Worksheet @($ruleObject) $Warnings
}

function Find-SourcePicture($SourceWorksheet, [int]$SourceRow, [int]$SourceImageColumn) {
    for ($i = 1; $i -le $SourceWorksheet.Shapes.Count; $i++) {
        $shape = $SourceWorksheet.Shapes.Item($i)
        if ([int]$shape.Type -ne 13) { continue } # msoPicture
        $cell = $shape.TopLeftCell
        if ([int]$cell.Row -eq $SourceRow -and [int]$cell.Column -eq $SourceImageColumn -and
            [double]$shape.Width -gt 1 -and [double]$shape.Height -gt 1) {
            return $shape
        }
    }
    return $null
}

function Remove-ShapesFromCell($Worksheet, [int]$Row, [int]$Column) {
    $removed = 0
    $targetCell = $Worksheet.Cells.Item($Row, $Column)
    $cellLeft = [double]$targetCell.Left
    $cellTop = [double]$targetCell.Top
    $cellRight = $cellLeft + [double]$targetCell.Width
    $cellBottom = $cellTop + [double]$targetCell.Height
    for ($i = $Worksheet.Shapes.Count; $i -ge 1; $i--) {
        $shape = $Worksheet.Shapes.Item($i)
        $topLeftMatch = $false
        try {
            $topLeftMatch = ([int]$shape.TopLeftCell.Row -eq $Row -and [int]$shape.TopLeftCell.Column -eq $Column)
        } catch {
            $topLeftMatch = $false
        }
        $centerX = [double]$shape.Left + ([double]$shape.Width / 2)
        $centerY = [double]$shape.Top + ([double]$shape.Height / 2)
        $centerMatch = ($centerX -ge $cellLeft -and $centerX -le $cellRight -and $centerY -ge $cellTop -and $centerY -le $cellBottom)
        if ($topLeftMatch -or $centerMatch) {
            $shape.Delete()
            $removed++
        }
    }
    return $removed
}

function Get-ImageSizing($Corrections) {
    $rule = Get-JsonProp $Corrections "imageSizing"
    [pscustomobject]@{
        margin = [double](Get-JsonProp $rule "margin" 3)
        maxWidthRatio = [double](Get-JsonProp $rule "maxWidthRatio" 0.90)
        maxHeightRatio = [double](Get-JsonProp $rule "maxHeightRatio" 0.90)
    }
}

function Get-TargetPictureArea($TargetWorksheet, [int]$TargetRow, [int]$TargetColumn) {
    $targetCell = $TargetWorksheet.Cells.Item($TargetRow, $TargetColumn)
    $area = $targetCell
    try {
        if ([bool]$targetCell.MergeCells) { $area = $targetCell.MergeArea }
    } catch {
        $area = $targetCell
    }
    return [pscustomobject]@{
        left = [double]$area.Left
        top = [double]$area.Top
        width = [double]$area.Width
        height = [double]$area.Height
    }
}

function Get-FitPictureRect($TargetWorksheet, [int]$TargetRow, [int]$TargetColumn, [double]$SourceWidth, [double]$SourceHeight, $Sizing) {
    $area = Get-TargetPictureArea $TargetWorksheet $TargetRow $TargetColumn
    $margin = [Math]::Max(0, [double]$Sizing.margin)
    $maxWidthRatio = [Math]::Max(0.1, [double]$Sizing.maxWidthRatio)
    $maxHeightRatio = [Math]::Max(0.1, [double]$Sizing.maxHeightRatio)
    $maxWidth = [Math]::Max(1, [Math]::Min($area.width - ($margin * 2), $area.width * $maxWidthRatio))
    $maxHeight = [Math]::Max(1, [Math]::Min($area.height - ($margin * 2), $area.height * $maxHeightRatio))
    $aspect = [Math]::Max(0.01, [double]$SourceWidth / [Math]::Max(1, [double]$SourceHeight))

    $width = $maxWidth
    $height = $width / $aspect
    if ($height -gt $maxHeight) {
        $height = $maxHeight
        $width = $height * $aspect
    }

    return [pscustomobject]@{
        left = $area.left + (($area.width - $width) / 2)
        top = $area.top + (($area.height - $height) / 2)
        width = $width
        height = $height
        area = $area
        margin = $margin
        maxWidth = $maxWidth
        maxHeight = $maxHeight
    }
}

function Verify-PictureInsideCell($Shape, $TargetWorksheet, [int]$TargetRow, [int]$TargetColumn, $Sizing) {
    if ($null -eq $Shape) { return $false }
    $area = Get-TargetPictureArea $TargetWorksheet $TargetRow $TargetColumn
    $margin = [Math]::Max(0, [double]$Sizing.margin)
    $maxWidth = [Math]::Max(1, [Math]::Min($area.width - ($margin * 2), $area.width * [double]$Sizing.maxWidthRatio))
    $maxHeight = [Math]::Max(1, [Math]::Min($area.height - ($margin * 2), $area.height * [double]$Sizing.maxHeightRatio))
    $left = [double]$Shape.Left
    $top = [double]$Shape.Top
    $width = [double]$Shape.Width
    $height = [double]$Shape.Height
    if ($width -le 1 -or $height -le 1) { return $false }
    $right = $left + $width
    $bottom = $top + $height
    return ($left -ge ($area.left + $margin - 0.75) -and
        $top -ge ($area.top + $margin - 0.75) -and
        $right -le ($area.left + $area.width - $margin + 0.75) -and
        $bottom -le ($area.top + $area.height - $margin + 0.75) -and
        $width -le ($maxWidth + 0.75) -and
        $height -le ($maxHeight + 0.75))
}

function Fit-PictureInsideCell($Shape, $TargetWorksheet, [int]$TargetRow, [int]$TargetColumn, $Sizing, [double]$SourceWidth = 0, [double]$SourceHeight = 0) {
    if ($SourceWidth -le 1) { $SourceWidth = [Math]::Max(1, [double]$Shape.Width) }
    if ($SourceHeight -le 1) { $SourceHeight = [Math]::Max(1, [double]$Shape.Height) }
    $rect = Get-FitPictureRect $TargetWorksheet $TargetRow $TargetColumn $SourceWidth $SourceHeight $Sizing
    try { $Shape.LockAspectRatio = 0 } catch {}
    $Shape.Left = [single]$rect.left
    $Shape.Top = [single]$rect.top
    $Shape.Width = [single]$rect.width
    $Shape.Height = [single]$rect.height
    try { $Shape.LockAspectRatio = -1 } catch {}
    $Shape.Placement = 1 # xlMoveAndSize: floating picture that follows row/column changes.
    return Verify-PictureInsideCell $Shape $TargetWorksheet $TargetRow $TargetColumn $Sizing
}

function Export-ShapeToTempPng($Excel, $Shape, [string]$TempDir) {
    if (-not (Test-Path -LiteralPath $TempDir)) {
        New-Item -ItemType Directory -Force -Path $TempDir | Out-Null
    }
    $png = Join-Path $TempDir ("image_{0}.png" -f ([guid]::NewGuid().ToString("N")))
    $tmpWorkbook = $null
    try {
        $tmpWorkbook = $Excel.Workbooks.Add()
        $tmpWorksheet = $tmpWorkbook.Worksheets.Item(1)
        $chartObject = $tmpWorksheet.ChartObjects().Add(1, 1, [Math]::Max(30, [double]$Shape.Width), [Math]::Max(30, [double]$Shape.Height))
        $Shape.CopyPicture(1, 2) | Out-Null
        Start-Sleep -Milliseconds 150
        $chartObject.Chart.Paste() | Out-Null
        $exported = $chartObject.Chart.Export($png, "PNG", $false)
        if ((-not $exported) -and (-not (Test-Path -LiteralPath $png))) {
            throw "Chart.Export returned false and no file was created"
        }
        if (-not (Test-Path -LiteralPath $png)) {
            throw "Chart.Export did not create $png"
        }
        if ((Get-Item -LiteralPath $png).Length -le 0) {
            throw "Chart.Export created an empty file"
        }
        return $png
    } finally {
        if ($null -ne $tmpWorkbook) {
            $tmpWorkbook.Close($false) | Out-Null
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($tmpWorkbook)
        }
    }
}

function Add-PictureFileToCell([string]$PicturePath, $SourceShape, $TargetWorksheet, [int]$TargetRow, [int]$TargetColumn, $Sizing) {
    $sourceWidth = [Math]::Max(1, [double]$SourceShape.Width)
    $sourceHeight = [Math]::Max(1, [double]$SourceShape.Height)
    $rect = Get-FitPictureRect $TargetWorksheet $TargetRow $TargetColumn $sourceWidth $sourceHeight $Sizing
    $picture = $TargetWorksheet.Shapes.AddPicture($PicturePath, 0, -1, [single]$rect.left, [single]$rect.top, [single]$rect.width, [single]$rect.height)
    try { $picture.LockAspectRatio = -1 } catch {}
    $picture.Placement = 1
    if (-not (Verify-PictureInsideCell $picture $TargetWorksheet $TargetRow $TargetColumn $Sizing)) {
        Fit-PictureInsideCell $picture $TargetWorksheet $TargetRow $TargetColumn $Sizing $sourceWidth $sourceHeight | Out-Null
    }
    return $picture
}

function Copy-PictureToCell($SourceShape, $TargetWorksheet, [int]$TargetRow, [int]$TargetColumn, $Excel, [string]$TempImageDir, $Sizing, $ImageFileCache) {
    $targetCell = $TargetWorksheet.Cells.Item($TargetRow, $TargetColumn)
    Remove-ShapesFromCell $TargetWorksheet $TargetRow $TargetColumn | Out-Null
    $cacheKey = $SourceShape.Name
    try {
        if (-not $ImageFileCache.ContainsKey($cacheKey)) {
            $ImageFileCache[$cacheKey] = Export-ShapeToTempPng $Excel $SourceShape $TempImageDir
        }
        $picture = Add-PictureFileToCell $ImageFileCache[$cacheKey] $SourceShape $TargetWorksheet $TargetRow $TargetColumn $Sizing
        if (Verify-PictureInsideCell $picture $TargetWorksheet $TargetRow $TargetColumn $Sizing) {
            return $picture
        }
        $picture.Delete()
    } catch {
        Remove-ShapesFromCell $TargetWorksheet $TargetRow $TargetColumn | Out-Null
    }

    $SourceShape.Copy() | Out-Null
    Start-Sleep -Milliseconds 100
    $TargetWorksheet.Paste($targetCell) | Out-Null
    $fallbackPicture = $TargetWorksheet.Shapes.Item($TargetWorksheet.Shapes.Count)
    if (-not (Fit-PictureInsideCell $fallbackPicture $TargetWorksheet $TargetRow $TargetColumn $Sizing ([double]$SourceShape.Width) ([double]$SourceShape.Height))) {
        throw "image placement failed at row $TargetRow, column $TargetColumn"
    }
    return $fallbackPicture
}

function Ensure-PictureRecordsInsideCells($PictureRecords, $Warnings, $Sizing) {
    $adjusted = 0
    foreach ($record in (Get-Array $PictureRecords)) {
        if ($null -eq $record -or $null -eq $record.shape) { continue }
        $inside = Verify-PictureInsideCell $record.shape $record.worksheet ([int]$record.row) ([int]$record.column) $Sizing
        if (-not $inside) {
            $inside = Fit-PictureInsideCell $record.shape $record.worksheet ([int]$record.row) ([int]$record.column) $Sizing ([double]$record.sourceWidth) ([double]$record.sourceHeight)
        }
        if ($inside) {
            $adjusted++
        } else {
            $Warnings.Add("image still outside target cell or safe box at row $($record.row), column $($record.column)") | Out-Null
        }
    }
    return $adjusted
}

$invoiceFull = (Resolve-Path -LiteralPath $InvoicePath).Path
$correctionsFull = (Resolve-Path -LiteralPath $CorrectionsPath).Path
$outputFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)

if ($invoiceFull -eq $outputFull) {
    throw "OutputPath must not be the same as InvoicePath."
}

$outputDir = Split-Path -Parent $outputFull
if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}

$corrections = Get-Content -Raw -Encoding UTF8 -LiteralPath $correctionsFull | ConvertFrom-Json
$imageSizing = Get-ImageSizing $corrections
$tempImageDir = Join-Path $outputDir (".image-cache-" + [IO.Path]::GetFileNameWithoutExtension($outputFull))
$imageFileCache = @{}
Copy-Item -LiteralPath $invoiceFull -Destination $outputFull -Force

$excel = $null
$workbook = $null
$referenceWorkbook = $null
$imageWorkbook = $null
$report = [ordered]@{
    source = $invoiceFull
    output = $outputFull
    sheet = $null
    headerRow = $null
    detailStartRow = $null
    detailEndRow = $null
    initialDetailEndRow = $null
    quantityExpected = $null
    quantityAccumulated = $null
    quantityMatchedEndRow = $null
    rowsDeleted = 0
    referenceFormatsApplied = $false
    defaultsApplied = 0
    forceValuesApplied = 0
    rowCorrectionsApplied = 0
    rowFormatsApplied = 0
    formulaCellsWritten = 0
    textNumbersConverted = 0
    textCodeCellsFormatted = 0
    imagesCopied = 0
    imageBoundsAdjusted = 0
    warnings = @()
    missingRequired = @()
    formulaErrorsBefore = @()
    formulaErrorsAfter = @()
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($outputFull, $false, $false)

    $sheetName = Get-JsonProp $corrections "sheetName"
    if ($sheetName) {
        $worksheet = $workbook.Worksheets.Item($sheetName)
    } else {
        $worksheet = $workbook.Worksheets.Item(1)
    }
    $report.sheet = $worksheet.Name

    $requiredHeaders = Get-JsonProp $corrections "requiredHeaders" @()
    $headerRow = Get-JsonProp $corrections "headerRow"
    if ($null -eq $headerRow) {
        $headerRow = Find-HeaderRow $worksheet $requiredHeaders
    }
    $headerRow = [int]$headerRow
    $report.headerRow = $headerRow

    $maps = Build-HeaderMaps $worksheet $headerRow
    $headerMap = $maps[0]
    $normalizedHeaderMap = $maps[1]

    $detailStartRow = Get-JsonProp $corrections "detailStartRow" ($headerRow + 1)
    $detailStartRow = [int]$detailStartRow
    $firstHeaderCol = 1
    if ($headerMap.Count -gt 0) {
        $firstHeaderCol = ($headerMap.Values | Sort-Object | Select-Object -First 1)
    }
    $detailEndRow = Get-JsonProp $corrections "detailEndRow"
    if ($null -eq $detailEndRow) {
        $detailEndRow = Find-DetailEndRow $worksheet $detailStartRow $firstHeaderCol
    }
    $detailEndRow = [int]$detailEndRow
    $report.detailStartRow = $detailStartRow
    $report.initialDetailEndRow = $detailEndRow
    $report.detailEndRow = $detailEndRow

    $truncateRule = Get-JsonProp $corrections "truncateByQuantity"
    if ($null -ne $truncateRule -and [bool](Get-JsonProp $truncateRule "enabled" $true)) {
        $quantityWarnings = New-Object System.Collections.Generic.List[string]
        $quantityResult = Resolve-DetailEndRowByQuantity $worksheet $detailStartRow $detailEndRow $headerMap $normalizedHeaderMap $truncateRule $quantityWarnings
        foreach ($warning in $quantityWarnings) { $report.warnings += $warning }
        if ($null -ne $quantityResult) {
            $report.quantityExpected = $quantityResult.expected
            $report.quantityAccumulated = $quantityResult.total
            if ($null -ne $quantityResult.row) {
                $matchedEndRow = [int]$quantityResult.row
                $report.quantityMatchedEndRow = $matchedEndRow
                $deleteRowsAfterMatch = [bool](Get-JsonProp $truncateRule "deleteRowsAfterMatch" $false)
                if ($deleteRowsAfterMatch -and $matchedEndRow -lt $detailEndRow) {
                    $report.rowsDeleted = Delete-WorksheetRows $worksheet ($matchedEndRow + 1) $detailEndRow
                }
                $detailEndRow = $matchedEndRow
                $report.detailEndRow = $detailEndRow
            }
        }
    }

    $referencePath = $ReferenceWorkbookPath
    $referenceConfig = Get-JsonProp $corrections "reference"
    if (-not $referencePath -and $null -ne $referenceConfig) {
        $referencePath = Get-JsonProp $referenceConfig "workbookPath"
    }
    if ($referencePath) {
        $referenceFull = (Resolve-Path -LiteralPath $referencePath).Path
        $referenceWorkbook = $excel.Workbooks.Open($referenceFull, $false, $true)
        $refSheetName = $ReferenceSheetName
        if (-not $refSheetName -and $null -ne $referenceConfig) {
            $refSheetName = Get-JsonProp $referenceConfig "sheetName"
        }
        if ($refSheetName) {
            $referenceWorksheet = $referenceWorkbook.Worksheets.Item($refSheetName)
        } else {
            $referenceWorksheet = $referenceWorkbook.Worksheets.Item(1)
        }
        $formatWarnings = @(Apply-ReferenceFormats $referenceWorksheet $worksheet)
        foreach ($warning in $formatWarnings) { $report.warnings += $warning }
        $report.referenceFormatsApplied = $true
    }

    $report.formulaErrorsBefore = @(Scan-FormulaErrors $worksheet)

    $formatWarnings = New-Object System.Collections.Generic.List[string]
    $explicitFormatRows = Get-JsonProp $corrections "formatRows"
    if ($null -ne $explicitFormatRows) {
        $report.rowFormatsApplied = Apply-RowFormatRules $worksheet $explicitFormatRows $formatWarnings
    } else {
        $report.rowFormatsApplied = Apply-AutoDetailFormatRows $worksheet $detailStartRow $detailEndRow (Get-JsonProp $corrections "autoFormatRows") $formatWarnings
    }
    foreach ($warning in $formatWarnings) { $report.warnings += $warning }

    $defaults = Get-JsonProp $corrections "defaults"
    if ($null -ne $defaults) {
        foreach ($row in $detailStartRow..$detailEndRow) {
            foreach ($prop in $defaults.PSObject.Properties) {
                $col = Resolve-HeaderColumn $headerMap $normalizedHeaderMap $prop.Name
                if ($null -eq $col) {
                    $report.warnings += "default header not found: $($prop.Name)"
                    continue
                }
                $cell = $worksheet.Cells.Item($row, $col)
                if (Is-BlankCell $cell) {
                    Set-ExcelCell $cell $prop.Value
                    $report.defaultsApplied++
                }
            }
        }
    }

    $imageSource = Get-JsonProp $corrections "imageSource"
    $imageWorksheet = $null
    $sourceImageColumn = $null
    $targetImageColumn = $null
    if ($null -ne $imageSource) {
        $imagePath = Get-JsonProp $imageSource "workbookPath"
        if ($imagePath) {
            $imageFull = (Resolve-Path -LiteralPath $imagePath).Path
            $imageWorkbook = $excel.Workbooks.Open($imageFull, $false, $true)
            $imageSheetName = Get-JsonProp $imageSource "sheetName"
            if ($imageSheetName) {
                $imageWorksheet = $imageWorkbook.Worksheets.Item($imageSheetName)
            } else {
                $imageWorksheet = $imageWorkbook.Worksheets.Item(1)
            }
            $sourceImageColumn = Resolve-ColumnSpec $imageWorksheet $null $null (Get-JsonProp $imageSource "sourceImageColumn" "图片") @(1, 2)
            $targetImageColumn = Resolve-ColumnSpec $worksheet $headerMap $normalizedHeaderMap (Get-JsonProp $imageSource "targetImageColumn" "产品图片*") @($headerRow)
            if ($null -eq $targetImageColumn) {
                $targetImageColumn = Resolve-ColumnSpec $worksheet $headerMap $normalizedHeaderMap "产品图片" @($headerRow)
            }
            if ($null -eq $sourceImageColumn) { $report.warnings += "source image column not found" }
            if ($null -eq $targetImageColumn) { $report.warnings += "target image column not found" }
        }
    }

    $copiedPictureRecords = @()
    foreach ($rowCorrection in (Get-Array (Get-JsonProp $corrections "rows"))) {
        $targetRow = Get-JsonProp $rowCorrection "row"
        if ($null -eq $targetRow) {
            $match = Get-JsonProp $rowCorrection "match"
            if ($null -ne $match) {
                $warningList = New-Object System.Collections.Generic.List[string]
                $targetRow = Find-MatchingRow $worksheet $detailStartRow $detailEndRow $headerMap $normalizedHeaderMap $match $warningList
                foreach ($warning in $warningList) { $report.warnings += $warning }
            }
        }
        if ($null -eq $targetRow) {
            $report.warnings += "row correction skipped: no row or match target"
            continue
        }
        $targetRow = [int]$targetRow
        if ($targetRow -lt $detailStartRow -or $targetRow -gt $detailEndRow) {
            $report.warnings += "row correction skipped: row $targetRow outside detail range $detailStartRow-$detailEndRow"
            continue
        }
        $values = Get-JsonProp $rowCorrection "values"
        if ($null -eq $values) { continue }

        foreach ($prop in $values.PSObject.Properties) {
            $col = Resolve-HeaderColumn $headerMap $normalizedHeaderMap $prop.Name
            if ($null -eq $col) {
                $report.warnings += "correction header not found: $($prop.Name)"
                continue
            }
            Set-ExcelCell $worksheet.Cells.Item($targetRow, $col) $prop.Value
            $report.rowCorrectionsApplied++
        }

        if ($null -ne $imageWorksheet -and $null -ne $sourceImageColumn -and $null -ne $targetImageColumn) {
            $copyImage = [bool](Get-JsonProp $rowCorrection "copyImage" $false)
            $sourceRow = Get-JsonProp $rowCorrection "sourceRow"
            $imageSourceRow = Get-JsonProp $rowCorrection "imageSourceRow" $sourceRow
            if ($copyImage -and $null -ne $sourceRow) {
                $sourceShape = Find-SourcePicture $imageWorksheet ([int]$imageSourceRow) ([int]$sourceImageColumn)
                if ($null -ne $sourceShape) {
                    $pastedPicture = Copy-PictureToCell $sourceShape $worksheet $targetRow ([int]$targetImageColumn) $excel $tempImageDir $imageSizing $imageFileCache
                    $copiedPictureRecords += [pscustomobject]@{
                        shape = $pastedPicture
                        worksheet = $worksheet
                        row = $targetRow
                        column = [int]$targetImageColumn
                        sourceWidth = [double]$sourceShape.Width
                        sourceHeight = [double]$sourceShape.Height
                    }
                    $report.imagesCopied++
                } else {
                    $report.warnings += "source image not found at row $imageSourceRow"
                }
            }
        }
    }

    $imageBoundsWarnings = New-Object System.Collections.Generic.List[string]
    $report.imageBoundsAdjusted = Ensure-PictureRecordsInsideCells $copiedPictureRecords $imageBoundsWarnings $imageSizing
    foreach ($warning in $imageBoundsWarnings) { $report.warnings += $warning }

    $forceValueWarnings = New-Object System.Collections.Generic.List[string]
    $report.forceValuesApplied = Apply-ForceValues $worksheet $detailStartRow $detailEndRow $headerMap $normalizedHeaderMap (Get-JsonProp $corrections "forceValues") $forceValueWarnings
    foreach ($warning in $forceValueWarnings) { $report.warnings += $warning }

    $numericHeaders = Get-JsonProp $corrections "numericHeaders" (Get-DefaultNumericHeaders)
    $report.textNumbersConverted = Convert-TextNumbersInColumns $worksheet $detailStartRow $detailEndRow $headerMap $normalizedHeaderMap $numericHeaders $report.warnings

    foreach ($rule in (Get-Array (Get-JsonProp $corrections "formulaRules"))) {
        $target = Get-JsonProp $rule "target"
        $formulaR1C1 = Get-JsonProp $rule "formulaR1C1"
        if (-not $target -or -not $formulaR1C1) { continue }
        $col = Resolve-HeaderColumn $headerMap $normalizedHeaderMap $target
        if ($null -eq $col) {
            $report.warnings += "formula target header not found: $target"
            continue
        }
        $onlyWhenBlank = [bool](Get-JsonProp $rule "onlyWhenBlank" $false)
        foreach ($row in $detailStartRow..$detailEndRow) {
            $cell = $worksheet.Cells.Item($row, $col)
            if ($onlyWhenBlank -and -not (Is-BlankCell $cell)) { continue }
            $cell.FormulaR1C1 = $formulaR1C1
            $report.formulaCellsWritten++
        }
    }

    $textHeaders = Get-JsonProp $corrections "textHeaders" (Get-DefaultTextHeaders)
    $report.textCodeCellsFormatted = Format-TextColumns $worksheet $detailStartRow $detailEndRow $headerMap $normalizedHeaderMap $textHeaders

    $excel.CalculateFull()
    $report.missingRequired = @(Scan-MissingRequired $worksheet $detailStartRow $detailEndRow $headerMap $normalizedHeaderMap $requiredHeaders)
    $report.formulaErrorsAfter = @(Scan-FormulaErrors $worksheet)

    $workbook.Save()

} finally {
    if ($null -ne $workbook) {
        $workbook.Close($true) | Out-Null
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($null -ne $imageWorkbook) {
        $imageWorkbook.Close($false) | Out-Null
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($imageWorkbook)
    }
    if ($null -ne $referenceWorkbook) {
        $referenceWorkbook.Close($false) | Out-Null
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($referenceWorkbook)
    }
    if ($null -ne $excel) {
        $excel.Quit() | Out-Null
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$report | ConvertTo-Json -Depth 8

if ($Strict -and $report.missingRequired.Count -gt 0) {
    exit 1
}
if ($Strict -and @($report.warnings | Where-Object { $_ -match "image|picture|source image" }).Count -gt 0) {
    exit 1
}
if ($FailOnFormulaErrors -and $report.formulaErrorsAfter.Count -gt 0) {
    exit 1
}
