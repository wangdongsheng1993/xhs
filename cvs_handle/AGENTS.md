# cvs_handle Working Rules

## Scope
- This `AGENTS.md` applies to everything under `cvs_handle/`.
- Follow the repo-root `AGENTS.md` first, then these local rules.

## Current Change Policy
- By default, only work on single-task execution flow.
- Treat `fix_xhs_note_urls.py`, `fix_xhs_note_urls_gui.py`, and `run_fix_xhs_note_urls.bat` as the primary maintained path.
- Do not modify parallel-mode scripts unless the user explicitly asks for it.
- Do not modify parallel retry scripts unless the user explicitly asks for it.

## Out Of Scope By Default
- `fix_xhs_note_urls_parallel.py`
- `fix_xhs_note_urls_parallel_gui.py`
- `run_parallel.bat`
- `run_parallel_gui.bat`
- `retry_failed_parallel.py`

## Review And Fix Strategy
- When the user asks for review or fixes in `cvs_handle`, assume single-task mode first.
- If a behavior differs between single-task and parallel codepaths, prioritize the single-task codepath unless the user explicitly requests parity changes.
- Avoid making “sync” edits to parallel scripts just because similar logic exists there.

## Operational Preference
- Prefer lighter-risk changes that improve single-task throughput and usability for large batches.
- When speed tuning is requested, optimize for “fast mode with light anti-risk-control” in the single-task path first.
