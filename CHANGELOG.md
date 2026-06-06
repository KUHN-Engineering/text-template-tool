# Changelog

## [0.5.0] - 2026-06-06

- **Search algorithm:** completely new multi-dimensional scoring across title, path, keywords, and content; differentiates exact, prefix, and substring matches with configurable weights
- **Configuration:** 17+ parameters configurable via `config.txt`; customisable colors, search weights, startup message, and more
- **User interface:** adaptive terminal width; centred TTT logo; search term highlighting in displayed content
- **Performance:** added template proprocessing for search, progress bar throttled to 5% update intervals
- **Verbose mode:** enabled via config file (default: off); shows config values on the info page, shows search scores and performance metrics
- **Robustness:** corrupt template cache is auto-detected and rebuilt; empty template folder handling; improved path validation
- **Desktop shortcut:** skipped on non-Windows systems; optional PowerShell 7 detection (default: off)
- **Documentation:** new README, sample templates added, help page improved

## [0.4.0] - 2025-10-29

- **New commands:** `m` lists most recently modified templates; `i` shows folder path, template count, and last reload date; `p` opens the folder containing the selected template
- **Startup screen:** TTT logo added
- **Template reload:** progress bar shown during reload; reload on startup removed
- **Bugfixes and performance improvements**

## [0.3.0] - 2025-08-17

- Initial release
