# Repository Guidelines

## Project Structure & Module Organization
This repository is a small Windows-first Python toolkit for scraping Xiaohongshu Pugongying data into Excel workbooks. Keep source files at the repo root:

- `xhs_excel_runner.py`: main CLI entry point; dispatches `brand`, `ecommerce`, and `koc` runs.
- `xhs_extractor.py`: brand-sheet Excel flow.
- `xhs_ecommerce_extractor.py`: ecommerce-sheet Excel flow and image insertion.
- `xhs_koc_extractor.py`: KOC-sheet Excel flow.
- `xhs_*_feishu.py`: Feishu-only variants for smaller online updates.
- `browser_session/`, `ecom_screenshots/`, `koc_screenshots/`: runtime artifacts; do not treat as source.

## Build, Test, and Development Commands
Use PowerShell from the repository root.

- `python -m venv .venv`
  Creates a local virtual environment.
- `.venv\Scripts\Activate.ps1`
  Activates the environment on Windows.
- `pip install pandas openpyxl playwright`
  Installs the libraries used by the scripts.
- `python -m playwright install chromium`
  Installs the browser required by `launch_persistent_context(...)`.
- `python xhs_excel_runner.py 品牌 395-421 --excel "<input.xlsx>" --output "<result.xlsx>"`
  Runs the recommended local Excel workflow for brand sheets.
- `python -m compileall .`
  Fast syntax check before committing when no automated test suite exists.

## Coding Style & Naming Conventions
Follow the existing Python style: 4-space indentation, `snake_case` for functions and variables, `UPPER_CASE` for env-driven constants, and short helper functions for parsing or sheet updates. Keep comments brief and practical; Chinese business terms and sheet names should remain unchanged where they map to workbook headers.

## Testing Guidelines
There is no formal `tests/` directory yet. Validate changes with narrow row ranges such as `326-329` or `395-399` and write results to a separate `*_结果.xlsx` file. For scraper changes, confirm browser login reuse, output cell values, and expected screenshots under `ecom_screenshots/` or `koc_screenshots/`.

## Commit & Pull Request Guidelines
Recent history uses very short Chinese subjects (`提交`, `excel版`). Keep the one-line style, but make it specific: for example, `runner: 修复行号区间解析` or `电商: 调整案例链接回填`. Pull requests should state the affected sheet type, sample rows used for validation, any new env vars, and include screenshots or workbook evidence when browser behavior or image placement changes.

## Security & Configuration Tips
Do not commit customer Excel files, browser profiles, `.env` files, or generated screenshots. Prefer overriding paths with `XHS_EXCEL_PATH`, `XHS_OUTPUT_PATH`, and `XHS_SHEET_NAME` instead of hardcoding local machine changes.
