# Fapiao Skill

Backup of the Codex AMZ/FBA invoice repair skill used to repair carrier invoice and packing-list workbooks exported from Saihu.

This repository keeps the skill instructions, template assets, references, and helper scripts. Generated invoice workbooks, temporary analysis files, and unrelated skill folders are intentionally excluded.

Skill entrypoint:

- `fapiaoSKILL/SKILL.md`

Internal Codex skill name:

- `amz-invoice-template`

Tencent Docs MCP:

- Run `powershell -ExecutionPolicy Bypass -File fapiaoSKILL/scripts/check_tencent_docs_mcp.ps1` before first use on a colleague workstation/profile.
- The shared Windows user can reuse the locally configured `tencent-docs` authorization.
- Do not commit or share Tencent Docs tokens; see `fapiaoSKILL/references/tencent-docs-mcp.md`.
