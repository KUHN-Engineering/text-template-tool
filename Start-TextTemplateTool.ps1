#Requires -Version 5.1
<#
    .SYNOPSIS
    Text Template Tool (TTT) by KUHN Engineering

    .DESCRIPTION
    Compact and efficient PowerShell utility that quickly searches through text templates and copies them directly to the clipboard for immediate use.

    .NOTES
    Version:        0.4.0
    Author:         Christian Kuhn, KUHN Engineering, www.kuhn-engineering.ch
    Date:           2025

    .EXAMPLE
    PS> .\Start-TextTemplateTool.ps1
#>

### < CONTENT >
# - GENERAL
# - CONFIGURATION
# - SUB FUNCTIONS
# - MAIN FUNCTION
# - EXECUTION

### < GENERAL >
$app_name = "TTT - Text Template Tool"
$app_version = "0.4.0"

### < CONFIGURATION >
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$CONFIG_personal_config_filename = "config-personal.txt"
$CONFIG_personal_template_filename = "templates-personal.json"

$CONFIG_factor_title = 10
$CONFIG_factor_relativePath = 10
$CONFIG_factor_keywords = 5
$CONFIG_factor_content = 2
$CONFIG_number_of_results = 10

### < SUB FUNCTIONS >
function Set-ClipboardSafe {
    param([string]$Text)
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.Open()
    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    $null = $ps.AddScript({ param($t) Set-Clipboard -Value $t }).AddArgument($Text)
    $async = $ps.BeginInvoke()
    if (-not $async.AsyncWaitHandle.WaitOne(2000)) {
        $ps.Stop()
        Write-Host "Warning: Could not copy to clipboard." -ForegroundColor Yellow
    }
    $ps.Dispose()
    $rs.Close()
}
function Write-Header {
    process {
        Clear-Host
        Write-Host "################################################################################"
        Write-Host "# TTT - Text Template Tool by KUHN Engineering                          (V$($app_version))"
        Write-Host "################################################################################"
        Write-Host ""
    }
}

function Write-StartupScreen {
    process {
        Write-Host ""
        Write-Host "                             __________________"
        Write-Host "                            /_  __/_  __/_  __/"
        Write-Host "                             / /   / /   / /"
        Write-Host "                            /_/   /_/   /_/"
        Write-Host ""
        Write-Host ""
        Write-Host ""
        Write-Host "Enter command 'h' for help."
    }
}

function Add-DesktopShortcut {
    process {
        $shortcutPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "$($app_name).lnk")
        if (!(Test-Path -Path $shortcutPath)) {

            Write-Host "- Adding desktop shortcut..."

            # auto-detect PowerShell 7
            $pwsh = Get-Command pwsh -ErrorAction SilentlyContinue

            if ($pwsh) {
                $TargetPath     = $pwsh.Source
                $IconLocation   = "$TargetPath,0"
                $Version        = "PowerShell 7+"
            } else {
                $TargetPath     = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
                $IconLocation   = "powershell.exe,0"
                $Version        = "Windows PowerShell"
            }
            
            $scriptPath = $PSCommandPath

            $WShell = New-Object -ComObject WScript.Shell
            $shortcut = $WShell.CreateShortcut($shortcutPath)

            $shortcut.TargetPath = $TargetPath
            $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
            $shortcut.WorkingDirectory = (Split-Path $scriptPath -Parent)
            $shortcut.Description = "Desktop Shortcut for $($app_name) using $($Version)"
            $shortcut.IconLocation = $IconLocation

            $shortcut.Save()
        }
    }
}

function Set-Config {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $FilePath
    )
    process {
        $keyName = "personal-template-folder"
        $fileExists = Test-Path -Path $FilePath

        if ($fileExists) {
            $keyLine = Get-Content -Path $FilePath | Where-Object { $_ -match "^\s*$([regex]::Escape($keyName))\s*:" } | Select-Object -First 1
            if ($keyLine) {
                $existingValue = ($keyLine -split ":", 2)[1].Trim()
                if ($existingValue -and (Test-Path -Path $existingValue -PathType Container)) { return }
                (Get-Content -Path $FilePath -Encoding UTF8 | Where-Object { $_ -notmatch "^\s*$([regex]::Escape($keyName))\s*:" }) | Out-File -FilePath $FilePath -Encoding UTF8
            }
            Write-Host ""
            Write-Host "Key '$keyName' missing in '$($CONFIG_personal_config_filename)'." -ForegroundColor Yellow
        } else {
            Write-Host ""
            Write-Host "Configuration file '$($CONFIG_personal_config_filename)' not found." -ForegroundColor Yellow
        }

        Write-Host "Set your personal template folder (drag and drop, copy/paste or type the path)." -ForegroundColor Yellow
        Write-Host ""

        $resolved = $null
        while (-not $resolved) {
            $path = (Read-Host "> Personal template folder").Trim().Trim('"', "'")
            try {
                $candidate = Resolve-Path -Path $path -ErrorAction Stop
                if (Test-Path -Path $candidate -PathType Container) {
                    $resolved = $candidate
                } else {
                    Write-Host "Path exists but is not a folder. Try again." -ForegroundColor Red
                }
            }
            catch {
                Write-Host "Path not found or access denied. Try again." -ForegroundColor Red
            }
        }

        $line = "$keyName`: $resolved"
        try {
            if ($fileExists) {
                $existing = Get-Content -Path $FilePath -Raw -Encoding UTF8
                "$line`n$existing" | Out-File -FilePath $FilePath -Encoding UTF8 -NoNewline
            } else {
                $line | Out-File -FilePath $FilePath -Encoding UTF8
            }
        }
        catch {
            Write-Host "Unable to write configuration file: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Please add the configuration file manually and restart." -ForegroundColor Red
            Read-Host
            exit
        }
    }
}

function Read-Config {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -Type Leaf -Include "*.txt" })]
        $FilePath
    )
    process {
        $config = @{}
        Get-Content -Path $FilePath | ForEach-Object {
            $line = $_.Trim()
            if ($line -eq "" -or $line.StartsWith("#")) { return }
            $key, $value = $line -split ":", 2 | ForEach-Object { $_.Trim() }
            if ($key -and $value) {
                $config[$key] = $value
            }
        }
        return $config
    }
}

function Search-Template {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Templates,

        [Parameter(Mandatory = $true)]
        [string] $Query
    )
    process {
        # split query into individual terms
        $splitQueries = $Query.Split(" ").Trim()

        # score each template against each query term
        ForEach ($template in $templates) {

            $template.Score = 0
            ForEach ($splitQuery in $splitQueries) {
                $score = 0

                # title
                if ($template.Title -like "*$splitQuery*") {
                    $score += $CONFIG_factor_title
                }

                # relative path
                if ($template.RelativePath -like "*$splitQuery*") {
                    $score += $CONFIG_factor_relativePath
                }

                # keywords
                ForEach ($keyword in $template.Keywords) {
                    if ($keyword -like "*$splitQuery*") {
                        $score += $CONFIG_factor_keywords
                    }
                }

                # content
                if ($template.Content -like "*$splitQuery*") {
                    $score += $CONFIG_factor_content
                }

                $template.Score += $score
            }
        }
        
        $results = $templates `
        | Where-Object { $_.Score -gt 0 } `
        | Sort-Object -Property Score -Descending `
        | Select-Object -First $CONFIG_number_of_results

        return $results
    }
}

function Write-Results {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Templates,
        [Parameter(Mandatory = $false)]
        $Selection = 1,
        [Parameter(Mandatory = $false)]
        $Style = "search"
    )
    process {

        $cnt = 0
        ForEach ($template in $templates) {

            # header
            if ($cnt -eq 0) {
                if ($Style -eq "modified") {
                    Write-Host "Most recently modified templates:"
                    Write-Host ""
                }
            }

            # enumeration
            $cnt += 1
            $cntStr = $cnt.ToString().PadLeft(3, ' ')
            
            # string
            if ($Style -eq "modified") {
                $printStr = " $cntStr | $($template.LastWriteTime) $($template.Title)"
            }
            else {
                # "search"
                $printStr = " $cntStr | $($template.Title) [$($template.Score)]"
            }

            # write + color
            if ($cnt -eq $Selection) {
                Write-Host $printStr -ForegroundColor Yellow
            }
            else {
                Write-Host $printStr
            }
        }
    }
}

function Get-TemplatesFromFolder {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -Type Container })]
        [string] $TemplateFolder
    )
    process {

        $folder = Resolve-Path -Path $TemplateFolder
        $templates = @()

        # get files
        $files = Get-ChildItem -Path $folder -Recurse -File -Include "*.txt"

        # process files
        $cnt = 0
        ForEach ($file in $files) {
            $cnt += 1
            $percentComplete = [math]::Round(($cnt / $files.Count) * 100, 2)
            Write-Progress -Activity "Processing templates" `
                -Status "Template $($cnt) of $($files.Count): $($file.Name)" `
                -PercentComplete $percentComplete

            $templates += Get-TemplateFromFile -FilePath $file -BaseFolder $TemplateFolder
        }
        Write-Progress -Activity "Processing templates" -Completed

        return $templates
    }
}

function Get-TemplateFromFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -Type Leaf -Include "*.txt" })]
        [string] $FilePath,
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -Type Container })]
        [string] $BaseFolder,
        [Parameter(Mandatory = $false)]
        [bool] $Personal = $true
    )
    process {

        # get title
        $title = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)

        # get path
        $relativePath = $FilePath | Split-Path -Parent | ForEach-Object { $_ -replace [regex]::Escape($BaseFolder), "" } | ForEach-Object { $_ -replace "^\\", "" }

        # get lastWriteTime
        $lastWriteTime = (Get-ChildItem -Path $FilePath).LastWriteTime

        # get content
        $rawContent = Get-Content -Path $FilePath -Raw -Encoding UTF8
        
        $keywords = @()
        $content = $rawContent
        
        # extract keywords from first line if present
        if ($rawContent) {
            $lines = $rawContent -split '\r?\n', 2  # Split only on first newline
            $firstLine = $lines[0]
            
            if ($firstLine.StartsWith("KEYWORDS:")) {
                $keywords = $firstLine.Replace("KEYWORDS:", "").Split(",") | ForEach-Object { $_.Trim() }
                $content = if ($lines.Count -gt 1) { $lines[1] } else { "" }
            }
        }

        $template = [PSCustomObject]@{
            Title         = $title
            RelativePath  = $relativePath
            Keywords      = $keywords
            Content       = $content
            Score         = 0
            File          = $FilePath
            LastWriteTime = $lastWriteTime
            Personal      = $Personal
        }

        return $template
    }
}

function Convert-TemplatesToJSON {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateScript({ Test-Path -Path $_ -Type Container })]
        $Folder,

        [Parameter(Mandatory = $true)]
        $JSONFile
    )
    process {
        $templates = Get-TemplatesFromFolder -TemplateFolder $Folder

        if (!$templates) {
            Remove-Item -Path $JSONFile -Force -ErrorAction SilentlyContinue
            Write-Host ""
            Write-Host "Your personal template folder doesn't contain any template text files (*.txt)." -ForegroundColor Yellow
            Write-Host "Add your first template to $($personal_template_folder) and restart $($app_name)." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Press any key to exit." -ForegroundColor Yellow
            Read-Host
            exit
        }

        $templates | ConvertTo-Json | Out-File -FilePath $JSONFile -Encoding UTF8
    }
}

function Import-TemplatesFromJSON {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -Type Leaf })]
        $JSONFile
    )
    process {
        $content = Get-Content -Path $JSONFile -Raw -Encoding UTF8
        try
        {
            $templates = $content | ConvertFrom-Json
        }
        catch
        {
            Remove-Item -Path $JSONFile -Force -ErrorAction SilentlyContinue
            Write-Host "Template cache is corrupt and has been removed. Restart to rebuild it." -ForegroundColor Red
            Write-Host ""
            Write-Host "Press any key to exit." -ForegroundColor Red
            Read-Host
            exit
        }
        return $templates
    }
}

function Import-Templates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -Type Container })]
        $TemplateFolder,

        [Parameter(Mandatory = $true)]
        $TemplateFile,

        [Parameter(Mandatory = $false)]
        [switch]$ForceReload
    )
    process {

        # check template cache
        Write-Host "- Checking templates..."
        $isJSON = Test-Path -Path $TemplateFile -Include "*.json" -Type Leaf

        # reload txt templates to JSON if needed or forced
        if ($ForceReload -or !$isJSON) {
            Write-Host "- Reloading templates from folder..."
            Convert-TemplatesToJSON -Folder $TemplateFolder -JSONFile $TemplateFile
        }

        # import templates from json
        Write-Host "- Importing templates..."
        $templates = Import-TemplatesFromJSON -JSONFile $TemplateFile

        return $templates
    }
}

function Write-Info {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -Type Container })]
        $TemplateFolder,

        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -Type Leaf })]
        $TemplateFile,

        [Parameter(Mandatory = $true)]
        $Templates
    )
    process {

        Write-Host "Personal template folder:           $($TemplateFolder)"
        Write-Host "Personal template count:            $(@($Templates).Count)"

        $lastWriteTime_json_str = (Get-ChildItem -Path $TemplateFile).LastWriteTime.ToString("dd-MM-yyyy HH:mm")
        Write-Host "Last reload of templates:           $($lastWriteTime_json_str)"
        Write-Host "PowerShell version:                 $($PSVersionTable.PSVersion.ToString())"
    }
}


### < MAIN FUNCTION >
function Start-TextTemplateTool {
    process {

        # startup procedure
        Write-Header

        # add desktop shortcut
        Write-Host "- Checking for desktop shortcut..."
        Add-DesktopShortcut

        Write-Host "- Loading configuration..."
        $scriptDir = Split-Path -Parent $PSCommandPath
        $personal_config_file = Join-Path $scriptDir $CONFIG_personal_config_filename

        Set-Config -FilePath $personal_config_file

        $config = Read-Config -FilePath $personal_config_file

        # resolve paths from config
        $personal_template_folder = $config['personal-template-folder']
        if (!(Test-Path -Path $personal_template_folder -PathType Container)) {
            Write-Host "Personal template folder not found: $($personal_template_folder)" -ForegroundColor Red
            Write-Host "Update '$($CONFIG_personal_config_filename)' or delete it to reconfigure." -ForegroundColor Red
            Write-Host ""
            Write-Host "Press any key to exit." -ForegroundColor Red
            Read-Host
            exit
        }
        $personal_template_file = Join-Path $scriptDir $CONFIG_personal_template_filename

        # import templates
        $templates = Import-Templates -TemplateFolder $personal_template_folder -TemplateFile $personal_template_file

        # show startup screen
        Write-Header
        Write-StartupScreen

        # infinite main loop
        $selection = 1
        $style = "search"
        do {
            Write-Host ""
            Write-Host "--------------------------------------------------------------------------------"
            $query = Read-Host "> Search / Select / Command"
            Write-Header

            # empty input -> cycle
            if ($query -eq "") {

                # reset TUI state
                $topResults = $null
                $selection = 1

                Write-StartupScreen
                continue
            }

            # check for commands
            elseif ($query -eq "h") {
                # reset TUI state
                $topResults = $null
                $selection = 1

                Write-Host "Available commands:"
                Write-Host ""
                Write-Host "   h    help       Lists all available commands."
                Write-Host ""
                Write-Host "   i    info       Displays template count and folder info."
                Write-Host ""
                Write-Host "   o    open       Opens the selected template in a text editor."
                Write-Host ""
                Write-Host "   p    parent     Opens the folder containing the selected template."
                Write-Host ""
                Write-Host "   f    folder     Opens your template folder in File Explorer."
                Write-Host ""
                Write-Host "   r    reload     Reloads templates from your template folder."
                Write-Host ""
                Write-Host "   m    modified   Lists the most recently modified templates."
                Write-Host ""
                Write-Host "   q    quit       Exits Text Template Tool."
                
                continue
            }
            elseif ($query -eq "q") {
                Clear-Host
                exit
            }
            elseif ($query -eq "o") {
                if ($topResults) {
                    Start-Process $topResults[$selection - 1].File
                }
                else {
                    Write-Host "No template selected. Search first, then open with 'o'."
                    continue
                }
            }
            elseif ($query -eq "p") {
                if ($topResults) {
                    $parentFolder = Split-Path -Path $topResults[$selection - 1].File -Parent
                    Start-Process $parentFolder
                }
                else {
                    Write-Host "No template selected. Search first, then open the parent folder with 'p'."
                    continue
                }
            }
            elseif ($query -eq "r") {
                # reset TUI state
                $topResults = $null
                $selection = 1

                Write-Header
                $templates = Import-Templates -TemplateFolder $personal_template_folder -TemplateFile $personal_template_file -ForceReload

                Write-Header
                Write-Info -TemplateFolder $personal_template_folder -TemplateFile $personal_template_file -Templates $templates

                continue
            }
            elseif ($query -eq "f") {
                # reset TUI state
                $topResults = $null
                $selection = 1

                Start-Process $personal_template_folder

                Write-StartupScreen
                continue
            }
            elseif ($query -eq "m") {

                $topResults = $templates | Sort-Object LastWriteTime -Descending | Select-Object -First $CONFIG_number_of_results
                $selection = 1
                $style = "modified"

                # no template found
                if (!$topResults) {
                    Write-Host "No matching template found."
                    continue
                }
            }
            elseif ($query -eq "i") {
                # reset TUI state
                $topResults = $null
                $selection = 1

                Write-Info -TemplateFolder $personal_template_folder -TemplateFile $personal_template_file -Templates $templates
                continue
            }

            # check for selection inputs
            elseif ( $topResults -and ($query -match '^(?:[1-9]|[1-9][0-9])$') ) {
                if ([int]$query -gt $topResults.Count) {
                    $selection = 1
                }
                else {
                    $selection = [int]$query
                }
            }

            # search query
            else {
                $topResults = Search-Template -Templates $templates -Query $query
                $selection = 1
                $style = "search"

                # no template found
                if (!$topResults) {
                    Write-Host "No matching template found."
                    continue
                }
            }

            if ($topResults) {
                $template = $topResults[$selection - 1]
                
                if ($template.Content) {
                    Set-ClipboardSafe $template.Content
                    $template.Content | Write-Host
                }else{
                    Write-Host "Empty template file." -ForegroundColor Yellow
                }

                Write-Host ""
                Write-Host "--------------------------------------------------------------------------------"
                Write-Host ""
                Write-Results -Templates $topResults -Selection $selection -Style $style
            }

        } while ( $true )
    }
}

### < EXECUTION >
Start-TextTemplateTool