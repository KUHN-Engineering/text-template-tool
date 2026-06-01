# Changelog

## [0.5.0] - 2026-06-xx (DRAFT)
- feat: enhance desktop shortcut with pwsh 7 auto-detection
- error-handling:
  - corrupt json -> force reload / regenerate template
  - empty template folder
- refactor: TUI state implementation
- improved: language of code comment and user messages
- added CHANGELOG.md
- improved README.md
- ? (te be tested) fix: set clipbard freeze issue
- minor fixes and improvements
  - fix: Resolve-Path issues
  - fix: empty file handling
  - set-strictmode
  - global config hash table
  - global variables
  - template import
  - fix: single template reload progress bar issue
  - fix: date time culture invariant implementation
- refactor and improve implementation of config parameters
  - set verbose mode
  - override default config parameters through config file
  - added sample config file with explanations
- feat: configurable start up message
- refactor: search algorithm and scoring
- ui: match highlighting
- added sample templates

## [0.4.0] - 2025-10-29
- Added command `m` to list most recently modified templates
- Added command `i` to show folder path, template count, date of last reload, PowerShell version
- Added command `p` to open folder containing selected template
- Added startup screen with TTT logo
- Removed reload on startup
- Added progress bar during template reload
- Bugfix and performance improvements

## [0.3.0] - 2025-08-17
- Initial release
