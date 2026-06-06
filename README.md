# TTT - Text Template Tool

```
             __________________
            /_  __/_  __/_  __/
             / /   / /   / /
            /_/   /_/   /_/      Text Template Tool by KUHN Engineering
```

A tiny keyboard-driven productivity tool that lets you search, select, and copy text templates to your clipboard in seconds — right from your terminal.

## Motivation

Many everyday tasks at work involve typing the same things over and over: email replies, status updates, standard phrases, form responses. This tool keeps all your text snippets in one place and puts any of them on your clipboard with just a few keystrokes.

**What can you use it for?**
- Ready-to-paste email replies, greetings, and standard phrases
- Lookup tables you reach for often — phone numbers, codes, account names
- Notes or checklists you refer to regularly
- Your best LLM prompts, organised and always one keystroke away
- Any text you find yourself typing repeatedly at work

**Why a PowerShell Text-Based User Interface (TUI) Solution?**

Think of it as the slickest possible add-on to Ctrl+C / Ctrl+V — your entire library of text snippets, instantly on your clipboard with just a few keystrokes. PowerShell is available on every Windows client out of the box — no installation, no dependencies, no admin rights required. A text-based user interface (TUI) starts instantly, stays focused, and works entirely from the keyboard. No window to position, no UI to navigate. You type a few letters, your template appears, you press Enter — done.

## Features

- **Smart search** — finds templates by title, folder path, content, and optional keywords
- **Instant clipboard copy** — select a template and its content is on your clipboard, ready to paste anywhere
- **Keyboard-only workflow** — Search, select, and copy templates using only your keyboard for maximum speed
- **Subfolder organisation** — use folders to group templates by topic; all subfolders are searched automatically
- **Keywords** — tag templates with hidden keywords to make them easier to find
- **Multiple results** — browse and pick from ranked matches when several templates fit your query
- **Open and edit** — open any template directly in your default text editor
- **Desktop shortcut** — created automatically on first run for instant access
- **Works in restricted environments** — no installation, no admin rights, no internet connection required

## Usage

### How to Set Up

#### 1. Place the script
Copy `Start-TextTemplateTool.ps1` to a folder in your personal user folder (e.g. `C:\Users\YourName\text-template-tool\`).

> **Tip:** Pro users can clone the repository directly with `git clone https://github.com/KUHN-Engineering/text-template-tool` — this makes updating to the latest version as simple as `git pull`.

#### 2. Create your template folder
Create a folder where your template files will live (e.g. `C:\Users\YourName\text-templates\`). You can use subfolders to organise by topic — all subfolders are scanned automatically.

#### 3. Create your first template
1. Go to your template folder
2. Right-click → **New > Text Document**
3. Give it a descriptive name — the filename becomes the search title (e.g. `Meeting follow-up.txt`)
4. Open it and type your template text
5. Optionally add keywords on the very first line (see [Keywords](#how-to-use-keywords))

> **Tip:** A few starter templates are included in the `sample-templates` folder. Copy them to your template folder to get started quickly and see how templates and keywords are structured.

> **Tip:** Keep the script and your template folder in the same location — this simplifies upgrades and avoids cache compatibility issues. A shared team folder works well for shared access; a personal documents folder (typically available across all workstations) works for personal use.

#### 4. Run the script for the first time
Right-click `Start-TextTemplateTool.ps1` and select **Run with PowerShell**. On first run, you will be prompted to enter the path to your template folder. A desktop shortcut is then created automatically.

> If Windows blocks the script from running, see [Execution Policy](#execution-policy) in the Advanced section.

### How to Run and Search

| | |
|---|---|
| **Desktop shortcut** | Double-click the shortcut created automatically on first run. |
| **Start menu** | Search for "TTT" and press Enter. |
| **Hotkey** | Right-click the shortcut → **Properties** → click **Shortcut key** → press your combination (e.g. `Ctrl+Alt+T`) → **OK**. |
| **Terminal** | Run `.\Start-TextTemplateTool.ps1` directly from PowerShell or Windows Terminal. |

Once the tool is open, start typing what you are looking for and press enter. The best matching template is selected automatically and its content is copied to your clipboard straight away — ready to paste. If the first result is not what you were looking for, type a number and press enter to pick from the next best matches shown on screen.

### How to Add or Edit Templates

You can manage templates without ever leaving the tool:

- **Edit an existing template** — search for it, then press `o` to open it in your default text editor. Make your changes, save the file, then press `r` to reload.
- **Create a similar template** — press `p` to open the folder containing the selected template in File Explorer. Copy and paste an existing file, rename it, and edit its content. Press `r` when done.
- **Browse your template folder** — press `f` to open your root template folder in File Explorer. Add new files, reorganise subfolders, or delete templates as needed.
- **Reload after any change** — always press `r` after adding, editing, or deleting template files to apply the changes. The tool does not detect file changes automatically.

### How to Use Keywords

Add a `KEYWORDS:` line as the very first line of any template file to boost how easily it is found:

```
KEYWORDS: email, customer, follow-up, request
This is the actual template text that gets copied to the clipboard...
```

The `KEYWORDS:` line is excluded when the template is copied to the clipboard. Keywords are especially useful when the words you naturally search for don't appear in the filename or template content.

### Commands

Everything is controlled with single-key commands:

| Key | Command | Description |
|-----|---------|-------------|
| *type text* | Search | Type anything to search your templates — results update with each search. |
| *type a number* | Select | Type a result number (1–10) to select a different template from the list. |
| `o` | Open | Opens the selected template file with its default Windows application (e.g. Notepad for `.txt` files). Use this to read the full content, make edits, or save a copy as the starting point for a similar new template. |
| `p` | Parent folder | Opens the folder containing the selected template in File Explorer — useful for finding, renaming, or duplicating files. |
| `f` | Folder | Opens your root template folder in File Explorer. |
| `r` | Reload | Rescans your template folder and rebuilds the template list. Run this after adding, editing, or deleting template files. |
| `m` | Modified | Lists the most recently modified templates — handy for picking up where you left off or reviewing recent additions. |
| `i` | Info | Shows the number of templates loaded, the template folder path, and when templates were last reloaded. |
| `h` | Help | Shows available commands. |
| `q` | Quit | Exits the tool. |

## Advanced

### Requirements

- Windows 10 or 11
- PowerShell 5.1 (built into Windows) or PowerShell 7
- Read/write access to the folders where the script and templates are stored

### Execution Policy

If Windows blocks the script with an error about execution policy, open PowerShell and run:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

This only affects your own user account and does not require admin rights. See the [official Microsoft documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies) for full details on execution policies.

> **Note:** In a corporate environment, check with your IT department and your organisation's code of conduct before changing this setting.

### Template Cache and Reload

On first load, all `.txt` files in your template folder are read and saved to a local cache file (`template-cache.json`) next to the script. On subsequent starts, the cache is loaded instead of re-reading every file — this makes startup noticeably faster for larger template collections.

After adding, editing, or deleting template files, press `r` to rescan and rebuild the cache. Optionally, set `ReloadCacheOnStartup: true` in the config file (see below) to always reload on startup.

### Search Behaviour

Each template is scored against your query across four areas: **title**, **keywords**, **folder path**, and **content**. Each area carries a different weight (title is ranked highest by default). Within each area, exact word matches score higher than prefix matches, which score higher than partial substring matches. Results are ranked by total score and the top matches are shown.

### Configuration

No configuration file is needed to get started. `config.txt` is created automatically on first startup in the same folder as the script. At minimum, it must contain the template folder path. To customise behaviour, add more configuration settings to `config.txt`. Settings use a simple `Key: Value` format, one per line. Lines starting with `#` are treated as comments. Only include the settings you want to change — all defaults are defined in the script itself and remain in effect for anything not specified.

Most helpful settings:

| Key | Default | Description |
|-----|---------|-------------|
| `TemplateFolder` | *(prompted on first run)* | Path to your template folder |
| `NumberOfResults` | `10` | Maximum number of search results shown |
| `ReloadCacheOnStartup` | `false` | Rescan template folder on every startup |
| `CheckForPowerShell7` | `false` | If `true`, checks whether PowerShell 7 is installed and creates the desktop shortcut targeting it if found |
| `StartupMessage` | *(none)* | Optional message shown in the header on startup |
| `ColorHighlight` | `Cyan` | Color for selected items |
| `ColorText` | `Gray` | Default terminal foreground color |
| `VerboseMode` | `false` | Show extended diagnostics (search scores, timing stats, settings) |

Available colors for color settings: `Cyan` `Yellow` `Green` `White` `Magenta` `Red` `Blue` `DarkYellow` `DarkCyan` `DarkGreen` `DarkMagenta` `DarkRed` `DarkBlue` `Gray` `DarkGray` `Black`

For further settings or their exact behaviour, refer to the script source.

Example `config.txt`:
```
TemplateFolder: C:\Users\YourName\text-templates
NumberOfResults: 7
ColorHighlight: Green
StartupMessage: Work smarter, not harder!
```

## License

**Text Template Tool** is licensed under the [MIT License](LICENSE). You are free to use, modify, and distribute this software, provided you include the original copyright and license notice.

## Changelog

See [CHANGELOG.md](CHANGELOG.md).