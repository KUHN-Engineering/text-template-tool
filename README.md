# Text Template Tool (TTT)

Compact and efficient PowerShell utility that quickly searches through text templates and copies them directly to the clipboard for immediate use.

## Features
- **Search Templates Effortlessly**: Quickly find text templates in a local folder using intuitive keyword searches.
- **Tag for Faster Access**: Add keywords (e.g., `notes`, `email`, `offer`, `order`) to organize and retrieve templates in seconds.
- **Instant Clipboard Copy**: Select a template, and it’s copied to your clipboard for immediate use—no extra steps.
- **Keyboard-Driven Workflow**: Search, select, and copy templates using only your keyboard for maximum speed.
- **Edit and Reload**: Open templates in your favorite text editor, make changes, and reload them seamlessly.
- **Choose from Similar Results**: Easily pick the right template from multiple matches with a simple interface.

## Requirements
- Windows 10/11
- PowerShell 5.1 or PowerShell 7
- Read/Write access to the template and script folder

## Installation
1. **Create Template Folder**: Make a new folder in your personal user folder for text templates. Use subfolders for better organization and search results (recommended).
2. **Set Up Script Folder**: Place the TTT script (`Start-TextTemplateTool.ps1`) in a separate folder in your personal user folder.
3. **Run the Script**: Right-click the script, select "Run with PowerShell," and follow the prompt to set your template folder. A desktop shortcut will be created automatically.

> **Tip**: Store folders in your personal user folder (e.g., `C:\Users\Username`) to ensure access across workstations in corporate environments.

## Usage
- **Simple**: Double-click the desktop shortcut created during installation.
- **Fast**: Search for "TTT" in the Windows Start menu and press Enter.
- **Clever**: Assign a hotkey to the desktop shortcut for instant access.
- **Pro**: Run the script directly from PowerShell or Terminal with `.\Start-TextTemplateTool.ps1`.

## Text Template File
- **What is a Text File?** In Windows, a text file (`.txt`) is a simple file containing plain text, editable in apps like Notepad.
- **How to Create One**: Right-click in a folder, select **New > Text Document**, name it (e.g., `template.txt`), and add your content.
- **Content as Template**: The file’s content is the template copied to your clipboard when selected in TTT.
- **Add Keywords**: Include a line like `KEYWORDS: customer, mail, request` at the top of the file to improve searchability. This line is ignored during clipboard copying.

## Search Functionality
- **Flexible Search Queries**: Type single characters, full words, multiple words, partial words, or specific expressions to search.
- **How Templates Are Found**: Matches are based on filename, optional keywords (e.g., `KEYWORDS: customer, mail, request`), folder/subfolder names, and template content.
- **Ranked Results**: Search results are sorted by a matching score, prioritizing the most relevant templates.

## License
**Text Template Tool (TTT)** is licensed under the [MIT License](LICENSE). You are free to use, modify, and distribute this software, provided you include the original copyright and license notice. See the [LICENSE](LICENSE) file for details.

## Changelog
### \[0.3.0\] - 2025-08-17
- Initial release of the script.