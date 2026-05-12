# Turbo Mode Change Log

Date: 2026-05-12

Purpose:
- Speed up single-task execution in `cvs_handle`.
- Target use case: urgent large-batch runs today.
- Scope intentionally excludes parallel mode and parallel retry mode.

## Files Changed
- `cvs_handle/fix_xhs_note_urls.py`
- `cvs_handle/fix_xhs_note_urls_gui.py`

## Important Conclusion
- Most runtime speed changes in `fix_xhs_note_urls.py` were added specifically for `speed_mode=turbo`.
- Not all changes are turbo-only.
- `fix_xhs_note_urls_gui.py` also changed default GUI values, and those defaults affect startup behavior even before the user manually changes anything.

## Turbo-Only Or Turbo-Gated Changes
- Added `is_turbo_speed_mode(...)`.
- Added turbo-specific effective parameter helpers:
  - `get_effective_slow_mo(...)`
  - `get_effective_detail_wait_ms(...)`
  - `get_effective_max_scrolls(...)`
  - `get_effective_per_item_timeout(...)`
  - `get_effective_batch_pause(...)`
  - `get_effective_save_every(...)`
- Added speed-profile-based wait helpers:
  - `get_note_url_poll_interval_ms()`
  - `get_post_navigation_settle_wait_ms()`
  - `get_post_selector_settle_wait_ms()`
  - `get_scroll_stuck_wait_ms()`
- In `turbo`, changed the anti-detection profile to be much more aggressive:
  - lower between-item delay
  - much lighter micro-rest
  - faster scroll delay
  - lighter warmup
  - weaker mouse pre-hover behavior
  - lighter failure cooldown
  - lighter verification cooldown
- In both single-task CSV and Excel flows:
  - `slow_mo` now becomes `0` when using turbo and the incoming value is still the old default `80`
  - `detail_wait_ms` now becomes `2500` when using turbo and the incoming value is still `>= 8000`
  - `max_scrolls` now becomes `8` when using turbo and the incoming value is still a slower high value (`>= 12`)
  - `per_item_timeout` now becomes `12` seconds when using turbo and the incoming value is still a slower high value (`>= 20`)
  - default batch pause was changed to `100` items + `120` seconds
  - heavy batch pause config is neutralized or capped when using turbo
  - but large batch pauses such as `100` items + `120` seconds are now preserved in turbo instead of being compressed to `3` seconds
  - when multiple `sessions` are provided in single-task mode with `session_mode=rotate`, the script now uses fixed rotation:
    `session1` 100 items -> rest 120s -> `session2` 100 items -> rest 120s -> ...
  - in that fixed multi-session rotation path, the separate global `batch-size/batch-interval` pause is disabled to avoid double resting
  - high-frequency save is disabled when using turbo and the value is still the old default `20`
  - turbo scroll gesture now uses fewer wheel steps per scroll
  - turbo card matching removed the heavy LCS fuzzy match and keeps cheaper exact/contains/prefix-suffix heuristics
  - turbo/single-task search path no longer runs the first-pass heavy debug DOM dump during normal lookup

## Non-Turbo Runtime Changes
- `wait_for_current_note_url(...)` now uses a speed-profile-based polling interval, not a fixed `250ms`.
- Navigation settle wait, selector settle wait, and stuck-scroll wait now depend on `SPEED_PROFILE`, so `fast_batch` also benefits.
- In Excel flow, successful items now call `finish_session_item(item_start_time, index)` instead of `finish_session_item()` so perf/ETA accounting is restored.
- `save_every=20` in `speed_mode=auto` is now promoted to `100` through `get_effective_save_every(...)`.

## GUI Changes
- `fix_xhs_note_urls_gui.py` defaults changed:
  - `login_wait`: `20` -> `12`
  - `max_scrolls`: `28` -> `8`
  - `speed_mode`: `auto` -> `turbo`
- GUI file picker now supports CSV as well as XLSX.
- Selecting an input file now auto-generates a separate output path like `*_结果.csv` or `*_结果.xlsx` instead of defaulting to overwrite the same file immediately.
- Some GUI copy text was adjusted from `Excel` to `Excel/CSV`.

## If You Want To Revert Later
- Revert all of `fix_xhs_note_urls_gui.py` if you want the old GUI defaults and picker behavior back.
- In `fix_xhs_note_urls.py`, the safest rollback is to revert:
  - `is_turbo_speed_mode(...)`
  - all `get_effective_*` helpers
  - all `get_*wait*_ms` helpers
  - the turbo branch inside `apply_speed_profile_for_large_batch(...)`
  - the call-site replacements that use these helpers
  - `FIXED_SESSION_BATCH_SIZE` / `FIXED_SESSION_SWITCH_REST_SEC`
  - the fixed multi-session rotate branch that overrides `batch_size` / `batch_interval`
  - the simplified `find_note_candidate(...)` scoring logic
  - the lighter turbo branch in `human_like_scroll(...)`
  - the removal of the first-pass debug DOM dump in `locate_note_url(...)`
- If you only want to keep the bug fix and drop the speed tuning:
  - keep the Excel `finish_session_item(item_start_time, index)` change
  - revert the rest of the turbo tuning

## Notes
- No parallel-mode files were changed in this speed-up pass.
- No parallel retry files were changed in this speed-up pass.
