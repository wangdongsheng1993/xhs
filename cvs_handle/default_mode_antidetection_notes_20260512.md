# Default Mode Anti-Detection Notes

Date: 2026-05-12

Scope:
- This note describes the anti-detection and pacing logic for the default single-task path in `cvs_handle/fix_xhs_note_urls.py`.
- It focuses on the baseline `default` behavior, not `turbo`.
- It also excludes parallel-mode and parallel-retry design concerns.

## Baseline Position
- `SPEED_PROFILE` starts as `default`.
- If `speed_mode` is `default`, `normal`, or `off`, the script stays in `default`.
- If `speed_mode=auto` and batch size is below `800`, the script also stays effectively in `default`.

## Default Timing And Rhythm
- Base random delay:
  - `min_delay_ms = 1500`
  - `max_delay_ms = 4000`
- Between-item pacing:
  - `between_item_delay_min_ms = 900`
  - `between_item_delay_max_ms = 2600`
- Micro-rest control:
  - after every random `4-7` items
  - rest `8-18` seconds

## Failure Cooldown
- The script tracks consecutive failures.
- When consecutive failures reach `3`, it can trigger an automatic cooldown.
- Default failure cooldown:
  - `45-95` seconds
  - at most once every `180` seconds
- This is managed through:
  - `register_processing_outcome(...)`
  - `maybe_take_failure_cooldown()`

## Verification / Login Popup Handling
- The script actively checks for:
  - captcha / verify modal
  - login modal
- It scans selectors and page text for signals like:
  - `请完成验证`
  - `请通过验证`
  - `滑动验证`
  - login modal presence
- If a popup is detected:
  - it increments a verification-hit counter
  - it may trigger a risk cooldown
  - it waits for the user to complete verification or login
- Default verification cooldown policy:
  - threshold: `2` consecutive hits
  - cooldown: `75-135` seconds
  - gap: `180` seconds
- Suppression window for repeated hit counting:
  - `verification_hit_suppression_sec = 15`

## Mouse / Click Simulation
- Mouse motion is intentionally non-linear and jittered.
- The script stores and reuses mouse position per page.
- Click behavior includes:
  - optional pre-hover with probability `0.45`
  - hover pause `120-420ms`
  - natural mouse move with multiple steps
- Related functions:
  - `move_mouse_naturally(...)`
  - `human_like_click(...)`

## Scroll Simulation
- Randomized scroll is enabled by default:
  - `random_scroll_pixels = True`
- Human-like scroll behavior:
  - each scroll is split into `3-6` wheel steps
  - per-step delay `800-2000ms`
- This makes page descent intentionally slow and less uniform.

## Viewport Randomization
- Viewport variance is enabled by default:
  - base width `1440`
  - base height `1000`
  - width jitter `90`
  - height jitter `70`
- Browser window size is therefore slightly different between sessions.

## Page Warmup
- Before real searching, the page may be “warmed up”.
- Default full warmup:
  - random move points: `2-4`
  - pause between moves: `300-900ms`
  - dwell after warmup: `700-1800ms`
- Light warmup exists for repeated visits:
  - probability `0.75`
  - move points `1-2`
  - pause `120-420ms`
  - dwell `150-450ms`
- Related logic:
  - `choose_warmup_mode(...)`
  - `warm_up_page(...)`

## Notes Tab And Card Interaction
- The script tries multiple selectors to locate the `笔记` tab.
- It prefers visible tabs and screens out suspiciously large / off-screen candidates.
- When clicking note cards:
  - it prefers clicking the visual card area
  - falls back to anchor click by `href`
- This reduces obvious robotic interaction patterns.

## Address Bar Polling
- In default mode, note-url polling interval remains conservative:
  - `250ms`
- This is slower than turbo and contributes to stability.

## Post-Navigation Settling
- In default mode, after opening a homepage:
  - extra settle wait is `600-1200ms`
- After waiting for note anchors:
  - additional settle wait is `500-1000ms`
- If scroll appears stuck:
  - extra wait is `400ms`

## Session Rotation Defaults Present In Code
- Even though you currently do not want to maintain parallel behavior, these default anti-risk values exist in the shared config:
  - `session_batch_min = 10`
  - `session_batch_max = 15`
  - `session_switch_rest_min_sec = 60`
  - `session_switch_rest_max_sec = 180`
- These matter only when that codepath is actually used.

## Summary
- Default mode is intentionally conservative.
- The main anti-detection cost comes from:
  - long between-item delay
  - frequent micro-rest
  - heavy failure cooldown
  - heavy verification cooldown
  - slow scroll cadence
  - full page warmup
  - humanized mouse and hover behavior
  - slower post-action settling

## Main Code References
- Config baseline:
  - `ANTI_DETECTION_CONFIG`
- Risk checks:
  - `check_verification_popup(...)`
- Pace control:
  - `apply_between_item_pacing(...)`
  - `register_processing_outcome(...)`
  - `maybe_take_failure_cooldown()`
- Interaction simulation:
  - `move_mouse_naturally(...)`
  - `human_like_scroll(...)`
  - `human_like_click(...)`
  - `warm_up_page(...)`
