# Tencent Docs MCP Integration

This skill depends on Tencent Docs MCP for two online sources:

| Purpose | URL file token | `sheet_id` | Expected title | Expected sheet |
| --- | --- | --- | --- | --- |
| Product-detail facts and images | `DY0hGbmd2Q1ZTVFVD` | `BB08J2` | `发票产品明细` | `工作表1` |
| Invoice work-scope selection | `DRE1ZTlhoZVZBVkdL` | `000001` | `AMZ备货计划及出货安排表` | `备货详情` |

## Authorization Model

The local machine already has a `tencent-docs` MCP server configured through `mcporter` under the current Windows user. Colleagues using this same machine and Windows profile may reuse that local authorization directly.

Do not store Tencent Docs tokens, Authorization headers, cookies, or copied `mcporter.json` contents inside this repository, skill zip, chat output, or invoice workbooks. The skill should only document commands and checks.

If a colleague runs this skill under a different Windows user or on another machine, the authorization does not automatically follow them. In that case, the owner should authorize that local environment with:

```powershell
mcporter auth tencent-docs
```

If OAuth is unavailable, follow the official Tencent Docs skill's manual token flow locally. Keep the token in the user's local `mcporter` configuration or environment only.

## Health Check

Run this before invoice work when a colleague is using the skill for the first time, after Windows user changes, or after any Tencent Docs failure:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/check_tencent_docs_mcp.ps1
```

Expected result: JSON with `ok: true`, both documents listed, and no `errors`.

Use `-IncludeCellSample` only for debugging. It reads the first few rows from both sheets and can expose operational table content in terminal output.

## Calling Convention

Use `--server` and `--tool`; do not call dotted tool names as `tencent-docs.sheet.get_sheet_info`, because some `mcporter` versions parse that as server `tencent-docs` and tool `sheet`.

```powershell
mcporter call --server tencent-docs --tool "manage.query_file_info" file_id=DY0hGbmd2Q1ZTVFVD --output json
mcporter call --server tencent-docs --tool "sheet.get_sheet_info" file_id=DY0hGbmd2Q1ZTVFVD --output json
mcporter call --server tencent-docs --tool "sheet.get_cell_data" file_id=DY0hGbmd2Q1ZTVFVD sheet_id=BB08J2 start_row=0 start_col=0 return_csv=true --output json

mcporter call --server tencent-docs --tool "manage.query_file_info" file_id=DRE1ZTlhoZVZBVkdL --output json
mcporter call --server tencent-docs --tool "sheet.get_sheet_info" file_id=DRE1ZTlhoZVZBVkdL --output json
mcporter call --server tencent-docs --tool "sheet.get_cell_data" file_id=DRE1ZTlhoZVZBVkdL sheet_id=000001 start_row=0 start_col=0 return_csv=true --output json
```

## Local Exports And Cache

Use the helper scripts instead of writing ad-hoc export code:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/fetch_product_detail_from_tencent_docs.ps1
powershell -ExecutionPolicy Bypass -File scripts/fetch_invoice_work_scope_from_tencent_docs.ps1
```

Both helpers write `.metadata.json` sidecars next to their downloaded `.xlsx` files under `.analysis/tencent-docs/`. On later runs they compare `file_id`, `sheet_id`, online `last_modify_time`, sheet row/column counts, and local workbook SHA-256. If there is no detected difference, they return `skippedDownload=true` and reuse the local workbook.

Use `-Force` only when a fresh export is intentionally required.

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| `mcporter` not found | Install or expose `mcporter` in PATH for the current Windows user. |
| `Authorization required` / token error | Run `mcporter auth tencent-docs` in the current user profile. |
| `工具: "sheet" 没有注册` or `工具: "manage" 没有注册` | Use `mcporter call --server tencent-docs --tool "sheet.get_sheet_info" ...` rather than a dotted selector. |
| Expected sheet ID missing | Re-check the Tencent Docs file URL/tab and update `SKILL.md` plus the related helper defaults before running invoice generation. |
| Export succeeds but data looks stale | Run the helper with `-Force`, then re-run the health check. |
