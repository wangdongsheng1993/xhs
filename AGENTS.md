# Repository Guidelines

## Project Structure & Module Organization
This repository is a Windows-first Python workspace for Xiaohongshu/抖音 data scripts, now split by business scenario.

- `data_handle/`: 本地 Excel 抓取与处理主流程
  - `xhs_excel_runner.py`: CLI 入口（品牌/电商/KOC）
  - `xhs_extractor.py`, `xhs_ecommerce_extractor.py`, `xhs_koc_extractor.py`, `xhs_steam_kol_extractor.py`
  - `xhs_gui_launcher.py`, `launch_gui.bat`
- `cvs_handle/`: 小红书笔记URL修复（并行重试）
  - `fix_xhs_note_urls.py`: 核心逻辑（查找/点击笔记）
  - `fix_xhs_note_urls_parallel.py`: 并行处理入口
  - `fix_xhs_note_urls_parallel_gui.py`: 主GUI界面
  - `retry_failed.py`: 单进程重试失败数据
  - `retry_failed_parallel.py`: 并行重试失败数据
  - `run_parallel_gui.bat`: 启动入口
  - `run_parallel.bat`: 命令行入口（可选）
- `dianshang_xiaohongshu_sync/`: 电商小红书飞书表同步
  - `sync_kol_execute.py`, `sync_kol_execute_gui.py`, `run_sync_kol.bat`, `dianshang-xiaohongshu.md`
- `pinpai_xiaohongshu_douyin_sync/`: 品牌小红书&抖音飞书表同步
  - `sync_machine_plan.py`, `lark_sync_gui.py`, `lark_sync.bat`, `requirement.md`
- Root-level helpers:
  - `xhs_extractor_feishu.py`, `xhs_ecommerce_extractor_feishu.py`
  - `update_kol_to_execute.py`, `run_update_kol.bat`, `update_kol_gui.hta`

Keep runtime artifacts (screenshots, temp outputs, browser sessions) out of source control.

## Build, Test, and Development Commands
Use PowerShell from repo root.

- `python -m venv .venv`
- `.venv\Scripts\Activate.ps1`
- `pip install pandas openpyxl playwright`
- `python -m playwright install chromium`
- Excel flow example:
  - `python data_handle\xhs_excel_runner.py 品牌 395-421 --excel "<input.xlsx>" --output "<result.xlsx>"`
- 电商小红书飞书同步（GUI）:
  - `dianshang_xiaohongshu_sync\run_sync_kol.bat`
- 品牌小红书&抖音飞书同步（GUI）:
  - `pinpai_xiaohongshu_douyin_sync\lark_sync.bat`
- 快速语法检查:
  - `python -m compileall .`

## Coding Style & Naming Conventions
- Python: 4 spaces, `snake_case` for functions/variables, `UPPER_CASE` for constants.
- Keep Feishu sheet names and Chinese business field names unchanged when used as mapping keys.
- Prefer small helper functions for: text normalization, header mapping, write/update operations.
- GUI scripts should keep startup resilience (venv/pythonw fallback) for Windows double-click usage.

## Testing Guidelines
No formal `tests/` directory yet.

- For Excel scripts: run small row ranges (e.g. `395-399`) and write to a separate output file.
- For Feishu sync scripts: first run against a small, safe data window or test sheet.
- Verify key checks explicitly:
  - Existing non-empty target cells are not overwritten when business rule requires protection.
  - Upsert behavior (update existing + append new) works without duplicate key rows.
  - GUI launchers can execute end-to-end via `.bat`.

## Commit & Pull Request Guidelines
- Keep commit subjects short and specific (Chinese or mixed is fine), e.g.:
  - `电商同步: 保护4月/5月非空单元格不覆盖`
  - `结构调整: 同步脚本按业务场景归档`
- PR should include:
  - affected scenario (`data_handle` / 电商同步 / 品牌抖音同步)
  - validation scope (rows/sheets/tasks)
  - screenshots/log snippets for GUI or Feishu write behavior when relevant

## Security & Configuration Tips
- Never commit customer Excel files, browser profiles, tokens, `.env`, or generated screenshots.
- Do not hardcode personal local paths.
- When using `lark-cli`, prefer environment/config-based auth and avoid exposing access tokens in logs.
