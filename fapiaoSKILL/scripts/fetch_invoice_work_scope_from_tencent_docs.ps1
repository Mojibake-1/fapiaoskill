param(
    [string]$OutputPath = "",
    [int]$PollSeconds = 5,
    [int]$TimeoutSeconds = 180,
    [switch]$Force
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $downloadDir = Join-Path (Get-Location) ".analysis\tencent-docs"
    $OutputPath = Join-Path $downloadDir "invoice-work-scope-latest.xlsx"
}

$fetchScript = Join-Path $PSScriptRoot "fetch_product_detail_from_tencent_docs.ps1"
$args = @{
    FileId = "DRE1ZTlhoZVZBVkdL"
    SheetId = "000001"
    OutputPath = $OutputPath
    PollSeconds = $PollSeconds
    TimeoutSeconds = $TimeoutSeconds
}

if ($Force) {
    & $fetchScript @args -Force
} else {
    & $fetchScript @args
}
