#Requires -Version 5.1
<#
    .SYNOPSIS
    Text Template Tool (TTT) by KUHN Engineering

    .DESCRIPTION
    Compact and efficient PowerShell utility that quickly searches through text templates and copies them directly to the clipboard for immediate use.

    .NOTES
    Version:        0.5.0
    Author:         Christian Kuhn, KUHN Engineering, www.kuhn-engineering.ch
    Date:           2026

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
Set-StrictMode -Version Latest

$script:AppName = "TTT - Text Template Tool"
$script:AppVersion = "0.5.0"

### < CONFIGURATION >
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$script:Config = @{
    ConfigFilename        = "config.txt"
    TemplateCacheFilename = "template-cache.json"

    # default values can be overridden in config file
    SearchWeightTitle     = 10
    SearchWeightPath      = 10
    SearchWeightKeywords  = 5
    SearchWeightContent   = 2
    NumberOfResults       = 10
    VerboseMode           = $false
    ReloadCacheOnStartup  = $false

    # will be set during runtime in Read-Config
    TemplateFolder        = ""
    TemplateCacheFile     = ""
}

### < SUB FUNCTIONS >
function Set-ClipboardSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,
        
        [int]$TimeoutMs = 2000,
        [int]$MaxRetries = 2
    )

    $attempt = 0
    while ($attempt -le $MaxRetries) {
        $runspace = $null
        $powershell = $null
        
        try {
            $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
            $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
            $runspace.Open()

            $powershell = [PowerShell]::Create()
            $powershell.Runspace = $runspace

            $null = $powershell.AddScript({
                    param($text)
                    # Use lower-level API with retry capability if possible
                    Add-Type -AssemblyName System.Windows.Forms
                    [System.Windows.Forms.Clipboard]::SetText($text)
                }).AddArgument($Text)

            $async = $powershell.BeginInvoke()

            if ($async.AsyncWaitHandle.WaitOne($TimeoutMs)) {
                # Success path
                $null = $powershell.EndInvoke($async)
                return # Exit function
            } 
            else {
                # Timeout
                $powershell.Stop()
                $attempt++
                if ($attempt -le $MaxRetries) {
                    Start-Sleep -Milliseconds 300
                }
            }
        }
        catch {
            Write-Warning "Clipboard operation failed (attempt $attempt): $($_.Exception.Message)"
            $attempt++
            if ($attempt -le $MaxRetries) {
                Start-Sleep -Milliseconds 300
            }
        }
        finally {
            if ($powershell) { 
                $powershell.Dispose() 
            }
            if ($runspace) { 
                $runspace.Close()
                $runspace.Dispose()
            }
        }
    }

    Write-Host "Warning: Could not copy to clipboard after $MaxRetries attempts." -ForegroundColor Yellow
}
function Write-Header {
    process {
        Clear-Host
        Write-Host "################################################################################"
        Write-Host "# TTT - Text Template Tool by KUHN Engineering                          (V$($script:AppVersion))"
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
        $shortcutPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "$($script:AppName).lnk")
        if (!(Test-Path -Path $shortcutPath)) {

            Write-Host "- Adding desktop shortcut..."

            # auto-detect PowerShell 7
            $pwsh = Get-Command pwsh -ErrorAction SilentlyContinue

            if ($pwsh) {
                $TargetPath = $pwsh.Source
                $IconLocation = "$TargetPath,0"
                $Version = "PowerShell 7+"
            }
            else {
                $TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
                $IconLocation = "powershell.exe,0"
                $Version = "Windows PowerShell"
            }
            
            $scriptPath = $PSCommandPath

            $WShell = New-Object -ComObject WScript.Shell
            $shortcut = $WShell.CreateShortcut($shortcutPath)

            $shortcut.TargetPath = $TargetPath
            $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
            $shortcut.WorkingDirectory = (Split-Path $scriptPath -Parent)
            $shortcut.Description = "Desktop Shortcut for $($script:AppName) using $($Version)"
            $shortcut.IconLocation = $IconLocation

            $shortcut.Save()
        }
    }
}

function Read-Config {
    process {
        $scriptDir = Split-Path -Parent $PSCommandPath
        $filePath = Join-Path $scriptDir $script:Config.ConfigFilename
        $keyName = "TemplateFolder"
        $fileExists = Test-Path -Path $filePath
        $needsPrompt = $true

        # check if TemplateFolder key already has a valid value
        if ($fileExists) {
            $keyLine = Get-Content -Path $filePath | Where-Object { $_ -match "^\s*$([regex]::Escape($keyName))\s*:" } | Select-Object -First 1
            if ($keyLine) {
                $existingValue = ($keyLine -split ":", 2)[1].Trim()
                if ($existingValue -and (Test-Path -Path $existingValue -PathType Container)) {
                    $needsPrompt = $false
                }
                else {
                    (Get-Content -Path $filePath -Encoding UTF8 | Where-Object { $_ -notmatch "^\s*$([regex]::Escape($keyName))\s*:" }) | Out-File -FilePath $filePath -Encoding UTF8
                    Write-Host ""
                    Write-Host "Key '$keyName' has an invalid value in '$($script:Config.ConfigFilename)'." -ForegroundColor Yellow
                }
            }
            else {
                Write-Host ""
                Write-Host "Key '$keyName' missing in '$($script:Config.ConfigFilename)'." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host ""
            Write-Host "Configuration file '$($script:Config.ConfigFilename)' not found." -ForegroundColor Yellow
        }

        # prompt for template-folder if needed
        if ($needsPrompt) {
            Write-Host "Set your template folder (drag and drop, copy/paste or type the path)." -ForegroundColor Yellow
            Write-Host ""

            $resolved = $null
            while (-not $resolved) {
                $path = (Read-Host "> Template folder").Trim().Trim('"', "'")
                try {
                    $candidate = Resolve-Path -Path $path -ErrorAction Stop
                    if (Test-Path -Path $candidate -PathType Container) {
                        $resolved = $candidate
                    }
                    else {
                        Write-Host "Path exists but is not a folder. Try again." -ForegroundColor Red
                    }
                }
                catch {
                    Write-Host "Path not found or access denied. Try again." -ForegroundColor Red
                }
            }

            # insert or overwrite key-value pair to config file or create new file if it doesn't exist
            $line = "$keyName`: $resolved"
            try {
                if ($fileExists) {
                    $existing = Get-Content -Path $filePath -Raw -Encoding UTF8
                    "$line`n$existing" | Out-File -FilePath $filePath -Encoding UTF8 -NoNewline
                }
                else {
                    $line | Out-File -FilePath $filePath -Encoding UTF8
                }
            }
            catch {
                Write-Host "Unable to write configuration file: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Please add the configuration file manually and restart." -ForegroundColor Red
                Read-Host
                exit
            }
        }

        # read all key-value pairs from config file
        $config = @{}
        Get-Content -Path $filePath | ForEach-Object {
            $cfgLine = $_.Trim()
            if ($cfgLine -eq "" -or $cfgLine.StartsWith("#")) { return }
            $key, $value = $cfgLine -split ":", 2 | ForEach-Object { $_.Trim() }
            if ($key -and $value) { $config[$key] = $value }
        }

        # apply optional config overrides
        $intValue = 0
        $boolValue = $false
        if ($config.ContainsKey('NumberOfResults') -and [int]::TryParse($config['NumberOfResults'], [ref]$intValue)) { $script:Config.NumberOfResults = $intValue }
        if ($config.ContainsKey('SearchWeightTitle') -and [int]::TryParse($config['SearchWeightTitle'], [ref]$intValue)) { $script:Config.SearchWeightTitle = $intValue }
        if ($config.ContainsKey('SearchWeightPath') -and [int]::TryParse($config['SearchWeightPath'], [ref]$intValue)) { $script:Config.SearchWeightPath = $intValue }
        if ($config.ContainsKey('SearchWeightKeywords') -and [int]::TryParse($config['SearchWeightKeywords'], [ref]$intValue)) { $script:Config.SearchWeightKeywords = $intValue }
        if ($config.ContainsKey('SearchWeightContent') -and [int]::TryParse($config['SearchWeightContent'], [ref]$intValue)) { $script:Config.SearchWeightContent = $intValue }
        if ($config.ContainsKey('VerboseMode') -and [bool]::TryParse($config['VerboseMode'], [ref]$boolValue)) { $script:Config.VerboseMode = $boolValue }
        if ($config.ContainsKey('ReloadCacheOnStartup') -and [bool]::TryParse($config['ReloadCacheOnStartup'], [ref]$boolValue)) { $script:Config.ReloadCacheOnStartup = $boolValue }

        # validate template folder
        $templateFolder = $config[$keyName]
        if (!(Test-Path -Path $templateFolder -PathType Container)) {
            Write-Host "Template folder not found: $templateFolder" -ForegroundColor Red
            Write-Host "Update '$($script:Config.ConfigFilename)' or delete it to reconfigure." -ForegroundColor Red
            Write-Host ""
            Write-Host "Press any key to exit." -ForegroundColor Red
            Read-Host
            exit
        }

        $script:Config.TemplateFolder = $templateFolder
        $script:Config.TemplateCacheFile = Join-Path $scriptDir $script:Config.TemplateCacheFilename

        # set verbose mode if enabled in config
        if ($script:Config.VerboseMode) { $script:VerbosePreference = 'Continue' } else { $script:VerbosePreference = 'SilentlyContinue' }
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
                    $score += $script:Config.SearchWeightTitle
                }

                # relative path
                if ($template.RelativePath -like "*$splitQuery*") {
                    $score += $script:Config.SearchWeightPath
                }

                # keywords
                ForEach ($keyword in $template.Keywords) {
                    if ($keyword -like "*$splitQuery*") {
                        $score += $script:Config.SearchWeightKeywords
                    }
                }

                # content
                if ($template.Content -like "*$splitQuery*") {
                    $score += $script:Config.SearchWeightContent
                }

                $template.Score += $score
            }
        }
        
        $results = $templates `
        | Where-Object { $_.Score -gt 0 } `
        | Sort-Object -Property Score -Descending `
        | Select-Object -First $script:Config.NumberOfResults

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

        # get files
        $files = @(Get-ChildItem -Path $folder -Recurse -File -Include "*.txt")

        # process files
        $cnt = 0
        $templates = ForEach ($file in $files) {
            $cnt += 1
            $percentComplete = [math]::Round(($cnt / $files.Count) * 100, 2)
            Write-Progress -Activity "Processing templates" `
                -Status "Template $($cnt) of $($files.Count): $($file.Name)" `
                -PercentComplete $percentComplete

            Get-TemplateFromFile -FilePath $file -BaseFolder $TemplateFolder
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
        [string] $BaseFolder
    )
    process {

        # get title
        $title = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)

        # get path
        $relativePath = $FilePath | Split-Path -Parent | ForEach-Object { $_ -replace [regex]::Escape($BaseFolder), "" } | ForEach-Object { $_ -replace "^\\", "" }

        # get lastWriteTime
        $lastWriteTime = (Get-ChildItem -Path $FilePath).LastWriteTime.ToString("yyyy-MM-dd HH:mm")

        # get content
        $rawContent = Get-Content -Path $FilePath -Raw -Encoding UTF8
        
        $keywords = @()
        $content = $rawContent
        
        # extract keywords from first line if present
        if ($rawContent) {
            $lines = $rawContent -split '\r?\n', 2  # Split only on first newline
            $firstLine = $lines[0]
            
            if ($firstLine.StartsWith("KEYWORDS:", [System.StringComparison]::OrdinalIgnoreCase)) {
                $keywords = ($firstLine -replace "(?i)^KEYWORDS:", "").Split(",") | ForEach-Object { $_.Trim() }
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
        }

        return $template
    }
}

function Convert-TemplatesToJSON {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
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
            Write-Host "Your template folder doesn't contain any template text files (*.txt)." -ForegroundColor Yellow
            Write-Host "Add your first template to $($Folder) and restart $($script:AppName)." -ForegroundColor Yellow
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
        try {
            $templates = $content | ConvertFrom-Json
        }
        catch {
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
        [Parameter(Mandatory = $false)]
        [switch]$ForceReload
    )
    process {

        # check template cache
        Write-Host "- Checking templates..."
        $isJSON = Test-Path -Path $script:Config.TemplateCacheFile -Type Leaf

        # rebuild cache from .txt files if missing or forced
        if ($ForceReload -or !$isJSON) {
            Write-Host "- Reloading templates from folder..."
            Convert-TemplatesToJSON -Folder $script:Config.TemplateFolder -JSONFile $script:Config.TemplateCacheFile
        }

        # load templates from cache
        Write-Host "- Importing templates..."
        $templates = Import-TemplatesFromJSON -JSONFile $script:Config.TemplateCacheFile

        return $templates
    }
}

function Write-Info {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Templates
    )
    process {

        Write-Host "Template folder:            $($script:Config.TemplateFolder)"
        Write-Host "Template count:             $(@($Templates).Count)"

        $lastWriteTime_json_str = (Get-ChildItem -Path $script:Config.TemplateCacheFile).LastWriteTime.ToString("dd-MM-yyyy HH:mm")
        Write-Host "Last reload of templates:   $($lastWriteTime_json_str)"

        if ($script:Config.VerboseMode) {
            Write-Host ""
            Write-Host "PowerShell version:         $($PSVersionTable.PSVersion.ToString())"  -ForegroundColor Yellow
            Write-Host "Script location:            $(Split-Path -Parent $PSCommandPath)" -ForegroundColor Yellow
            Write-Host "Template cache file:        $($script:Config.TemplateCacheFile)" -ForegroundColor Yello
            Write-Host "Reload cache on startup:    $($script:Config.ReloadCacheOnStartup)" -ForegroundColor Yellow
            Write-Host "Number of results:          $($script:Config.NumberOfResults)" -ForegroundColor Yellow
            Write-Host "Search weights:             title=$($script:Config.SearchWeightTitle)  path=$($script:Config.SearchWeightPath)  keywords=$($script:Config.SearchWeightKeywords)  content=$($script:Config.SearchWeightContent)" -ForegroundColor Yellow
        }
    }
}


### < MAIN FUNCTION >
function Start-TextTemplateTool {
    [CmdletBinding()]
    param()
    process {
        
        # startup procedure
        Write-Header

        # add desktop shortcut
        Write-Host "- Checking for desktop shortcut..."
        Add-DesktopShortcut

        Write-Host "- Loading configuration..."
        Read-Config

        # load templates
        $templates = Import-Templates -ForceReload:$script:Config.ReloadCacheOnStartup

        # show startup screen
        Write-Header
        Write-StartupScreen

        # infinite main loop
        $topResults = $null
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
                $style = "search"

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
                Read-Config
                $templates = Import-Templates -ForceReload

                Write-Header
                Write-Info -Templates $templates

                continue
            }
            elseif ($query -eq "f") {
                # reset TUI state
                $topResults = $null
                $selection = 1

                Start-Process $script:Config.TemplateFolder

                Write-StartupScreen
                continue
            }
            elseif ($query -eq "m") {

                $topResults = $templates | Sort-Object LastWriteTime -Descending | Select-Object -First $script:Config.NumberOfResults
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

                Write-Info -Templates $templates
                continue
            }

            # check for selection inputs
            elseif ( $topResults -and ($query -match '^[1-9]\d*$') ) {
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
                
                if (-not [string]::IsNullOrWhiteSpace($template.Content)) {
                    Set-ClipboardSafe $template.Content
                    Write-Host $template.Content
                }
                else {
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