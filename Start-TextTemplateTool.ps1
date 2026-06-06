#Requires -Version 5.1
<#
    .SYNOPSIS
    TTT- Text Template Tool by KUHN Engineering

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
    ConfigFilename                = "config.txt"
    TemplateCacheFilename         = "template-cache.json"

    # default values can be overridden in config file
    SearchDimensionWeightTitle    = 10
    SearchDimensionWeightPath     = 8
    SearchDimensionWeightKeywords = 6
    SearchDimensionWeightContent  = 2
    SearchMatchWeightExact        = 100
    SearchMatchWeightPrefix       = 50
    SearchMatchWeightSubstring    = 10
    NumberOfResults               = 10
    ColorText                     = "Gray"
    ColorBackground               = "Black"
    ColorHighlight                = "Cyan"
    ColorWarning                  = "Yellow"
    ColorError                    = "Red"
    VerboseMode                   = $false
    ReloadCacheOnStartup          = $false
    StartupMessage                = ""

    # will be set during runtime in Read-Config
    TemplateFolder                = ""
    TemplateCacheFile             = ""

    # TEMP fpor testing different clipboard methods, can be set in config file
    TempSelectClipboardFunction   = 2
}

$script:Stats = @{
    RebuildCacheMs = 0
    PreprocessMs   = 0
}

### < SUB FUNCTIONS >
function Set-ClipboardSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [int]$TimeoutMs = 2000,
        [int]$MaxRetries = 2
    )

    # TEMP for testing different clipboard methods, can be set in config file
    switch ($script:Config.TempSelectClipboardFunction) {

        1 {
            # simple Set-Clipboard cmdlet
            Set-Clipboard -Value $Text
        }

        2 {
            # clip.exe via stdin
            $Text | clip
        }

        default {
            # STA runspace with Windows.Forms (version 2)
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
                            Add-Type -AssemblyName System.Windows.Forms
                            [System.Windows.Forms.Clipboard]::SetText($text)
                        }).AddArgument($Text)

                    $async = $powershell.BeginInvoke()

                    if ($async.AsyncWaitHandle.WaitOne($TimeoutMs)) {
                        $null = $powershell.EndInvoke($async)
                        return
                    }
                    else {
                        $powershell.Stop()
                        $attempt++
                        if ($attempt -le $MaxRetries) { Start-Sleep -Milliseconds 300 }
                    }
                }
                catch {
                    Write-Warning "Clipboard operation failed (attempt $attempt): $($_.Exception.Message)"
                    $attempt++
                    if ($attempt -le $MaxRetries) { Start-Sleep -Milliseconds 300 }
                }
                finally {
                    if ($powershell) { $powershell.Dispose() }
                    if ($runspace) { $runspace.Close(); $runspace.Dispose() }
                }
            }

            Write-Host "Warning: Could not copy to clipboard after $MaxRetries attempts." -ForegroundColor $script:Config.ColorWarning
        }
    }
}
function Write-Header {
    param(
        [switch]$ShowStartupMessage
    )
    process {
        Clear-Host
        Write-Host "################################################################################"
        Write-Host "# TTT - Text Template Tool by KUHN Engineering                          (V$($script:AppVersion))"
        Write-Host "################################################################################"
        if ($ShowStartupMessage -and -not [string]::IsNullOrWhiteSpace($script:Config.StartupMessage)) {
            Write-Host $script:Config.StartupMessage -ForegroundColor $script:Config.ColorHighlight
        }
        else {
            Write-Host ""
        }
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
        Write-Host "Type to search, 'h' for help, 'r' to reload templates."
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
                    Write-Host "Key '$keyName' has an invalid value in '$($script:Config.ConfigFilename)'." -ForegroundColor $script:Config.ColorWarning
                }
            }
            else {
                Write-Host ""
                Write-Host "Key '$keyName' missing in '$($script:Config.ConfigFilename)'." -ForegroundColor $script:Config.ColorWarning
            }
        }
        else {
            Write-Host ""
            Write-Host "Configuration file '$($script:Config.ConfigFilename)' not found." -ForegroundColor $script:Config.ColorWarning
        }

        # prompt for template-folder if needed
        if ($needsPrompt) {
            Write-Host "Set your template folder (drag and drop, copy/paste or type the path)." -ForegroundColor $script:Config.ColorWarning
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
                        Write-Host "Path exists but is not a folder. Try again." -ForegroundColor $script:Config.ColorError
                    }
                }
                catch {
                    Write-Host "Path not found or access denied. Try again." -ForegroundColor $script:Config.ColorError
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
                Write-Host "Unable to write configuration file: $($_.Exception.Message)" -ForegroundColor $script:Config.ColorError
                Write-Host "Please add the configuration file manually and restart." -ForegroundColor $script:Config.ColorError
                Write-Host ""
                Write-Host "Press any key to exit." -ForegroundColor $script:Config.ColorError
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
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
        if ($config.ContainsKey('NumberOfResults') -and [int]::TryParse($config['NumberOfResults'], [ref]$intValue) -and $intValue -gt 0) { $script:Config.NumberOfResults = $intValue }
        if ($config.ContainsKey('SearchDimensionWeightTitle') -and [int]::TryParse($config['SearchDimensionWeightTitle'], [ref]$intValue) -and $intValue -gt 0) { $script:Config.SearchDimensionWeightTitle = $intValue }
        if ($config.ContainsKey('SearchDimensionWeightPath') -and [int]::TryParse($config['SearchDimensionWeightPath'], [ref]$intValue) -and $intValue -gt 0) { $script:Config.SearchDimensionWeightPath = $intValue }
        if ($config.ContainsKey('SearchDimensionWeightKeywords') -and [int]::TryParse($config['SearchDimensionWeightKeywords'], [ref]$intValue) -and $intValue -gt 0) { $script:Config.SearchDimensionWeightKeywords = $intValue }
        if ($config.ContainsKey('SearchDimensionWeightContent') -and [int]::TryParse($config['SearchDimensionWeightContent'], [ref]$intValue) -and $intValue -gt 0) { $script:Config.SearchDimensionWeightContent = $intValue }
        if ($config.ContainsKey('SearchMatchWeightExact') -and [int]::TryParse($config['SearchMatchWeightExact'], [ref]$intValue) -and $intValue -gt 0) { $script:Config.SearchMatchWeightExact = $intValue }
        if ($config.ContainsKey('SearchMatchWeightPrefix') -and [int]::TryParse($config['SearchMatchWeightPrefix'], [ref]$intValue) -and $intValue -gt 0) { $script:Config.SearchMatchWeightPrefix = $intValue }
        if ($config.ContainsKey('SearchMatchWeightSubstring') -and [int]::TryParse($config['SearchMatchWeightSubstring'], [ref]$intValue) -and $intValue -gt 0) { $script:Config.SearchMatchWeightSubstring = $intValue }
        
        if ($config.ContainsKey('ColorText')) { try { $script:Config.ColorText = [string][System.ConsoleColor]$config['ColorText'] } catch {} }
        if ($config.ContainsKey('ColorBackground')) { try { $script:Config.ColorBackground = [string][System.ConsoleColor]$config['ColorBackground'] } catch {} }
        if ($config.ContainsKey('ColorHighlight')) { try { $script:Config.ColorHighlight = [string][System.ConsoleColor]$config['ColorHighlight'] } catch {} }
        if ($config.ContainsKey('ColorWarning')) { try { $script:Config.ColorWarning = [string][System.ConsoleColor]$config['ColorWarning'] } catch {} }
        if ($config.ContainsKey('ColorError')) { try { $script:Config.ColorError = [string][System.ConsoleColor]$config['ColorError'] } catch {} }
        
        if ($config.ContainsKey('VerboseMode') -and [bool]::TryParse($config['VerboseMode'], [ref]$boolValue)) { $script:Config.VerboseMode = $boolValue }
        if ($config.ContainsKey('ReloadCacheOnStartup') -and [bool]::TryParse($config['ReloadCacheOnStartup'], [ref]$boolValue)) { $script:Config.ReloadCacheOnStartup = $boolValue }
        if ($config.ContainsKey('StartupMessage') -and -not [string]::IsNullOrWhiteSpace($config['StartupMessage'])) { $script:Config.StartupMessage = $config['StartupMessage'] }

        # TEMP for testing different clipboard methods, can be set in config file
        if ($config.ContainsKey('TempSelectClipboardFunction') -and [int]::TryParse($config['TempSelectClipboardFunction'], [ref]$intValue) -and $intValue -ge 1 -and $intValue -le 3) { $script:Config.TempSelectClipboardFunction = $intValue }

        # validate template folder
        $templateFolder = $config[$keyName]
        if (!(Test-Path -Path $templateFolder -PathType Container)) {
            Write-Host "Template folder not found: $templateFolder" -ForegroundColor $script:Config.ColorError
            Write-Host "Update '$($script:Config.ConfigFilename)' or delete it to reconfigure." -ForegroundColor $script:Config.ColorError
            Write-Host ""
            Write-Host "Press any key to exit." -ForegroundColor $script:Config.ColorError
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit
        }

        $script:Config.TemplateFolder = $templateFolder
        $script:Config.TemplateCacheFile = Join-Path $scriptDir $script:Config.TemplateCacheFilename

        # set verbose mode if enabled in config
        $global:VerbosePreference = if ($script:Config.VerboseMode) { 'Continue' } else { 'SilentlyContinue' }

        $Host.UI.RawUI.ForegroundColor = $script:Config.ColorText
        $Host.UI.RawUI.BackgroundColor = $script:Config.ColorBackground
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
        # preprocess query: lowercase, normalize, split, unique
        $queryFragments = ConvertTo-SearchWords -Text $Query

        # score each template against each query fragment
        $results = foreach ($template in $templates) {

            $score = 0
            foreach ($fragment in $queryFragments) {

                foreach ($dimension in @(
                        @{ Words = $template.Search.TitleWords; Weight = $script:Config.SearchDimensionWeightTitle },
                        @{ Words = $template.Search.KeywordWords; Weight = $script:Config.SearchDimensionWeightKeywords },
                        @{ Words = $template.Search.PathWords; Weight = $script:Config.SearchDimensionWeightPath },
                        @{ Words = $template.Search.ContentWords; Weight = $script:Config.SearchDimensionWeightContent }
                    )) {
                    $dimensionScore = 0
                    foreach ($word in $dimension.Words) {
                        if ($word -eq $fragment) {
                            $dimensionScore += $script:Config.SearchMatchWeightExact
                        }
                        elseif ($word.StartsWith($fragment)) {
                            $dimensionScore += $script:Config.SearchMatchWeightPrefix
                        }
                        elseif ($word.Contains($fragment)) {
                            $dimensionScore += $script:Config.SearchMatchWeightSubstring
                        }
                    }
                    $score += $dimensionScore * $dimension.Weight
                }
            }

            if ($score -gt 0) {
                [PSCustomObject]@{
                    Template = $template
                    Score    = $score
                }
            }
        }
        
        $results = $results `
        | Sort-Object -Property Score -Descending `
        | Select-Object -First $script:Config.NumberOfResults

        return $results
    }
}

function Write-ContentWithHighlights {
    param(
        [string]$Content,
        [string]$Query
    )
    process {
        if ([string]::IsNullOrWhiteSpace($Query)) {
            Write-Host $Content
            return
        }

        # union of raw terms (preserves diacritics/casing) and preprocessed fragments (diacritics stripped)
        $rawTerms          = @($Query -split '\s+' | Where-Object { $_ -ne '' })
        $preprocessedTerms = @(ConvertTo-SearchWords -Text $Query)
        $terms             = @($rawTerms + $preprocessedTerms | Select-Object -Unique)
        $pattern           = ($terms | ForEach-Object { [regex]::Escape($_) }) -join '|'
        $matchList = [regex]::Matches($Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        if ($matchList.Count -eq 0) {
            Write-Host $Content
            return
        }

        $pos = 0
        foreach ($m in $matchList | Sort-Object Index) {
            if ($m.Index -gt $pos) {
                Write-Host $Content.Substring($pos, $m.Index - $pos) -NoNewline
            }
            Write-Host $Content.Substring($m.Index, $m.Length) -NoNewline -ForegroundColor $script:Config.ColorHighlight
            $pos = $m.Index + $m.Length
        }
        if ($pos -lt $Content.Length) {
            Write-Host $Content.Substring($pos) -NoNewline
        }
        Write-Host ""
    }
}

function Write-Results {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Results,
        [Parameter(Mandatory = $false)]
        $Selection = 1,
        [Parameter(Mandatory = $false)]
        $Style = "search"
    )
    process {

        $cnt = 0
        ForEach ($result in $Results) {

            $template = $result.Template

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
                $printStr = " $cntStr | $($template.Title) [$($result.Score)]"
            }

            # write + color
            if ($cnt -eq $Selection) {
                Write-Host $printStr -ForegroundColor $script:Config.ColorHighlight
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

            Get-TemplateFromFile -FilePath $file -BaseFolder $folder
        }
        Write-Progress -Activity "Processing templates" -Completed

        return $templates
    }
}

function ConvertTo-SearchWords {
    param(
        [string]$Text
    )
    process {
        if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

        $s = $Text.ToLowerInvariant()

        # strip diacritics: decompose to NFD, remove combining marks, recompose
        $s = $s.Normalize([System.Text.NormalizationForm]::FormD)
        $s = [System.Text.RegularExpressions.Regex]::Replace($s, '\p{Mn}', '')
        $s = $s.Normalize([System.Text.NormalizationForm]::FormC)

        # split on any non-alphanumeric run (spaces, _, -, \, /, . etc.)
        $words = [System.Text.RegularExpressions.Regex]::Split($s, '[^a-z0-9]+') |
        Where-Object { $_ -ne '' } |
        Select-Object -Unique

        return @($words)
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
        $rawContent = [System.IO.File]::ReadAllText($FilePath, [System.Text.Encoding]::UTF8)
        
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

        if (@($templates).Count -eq 0) {
            Remove-Item -Path $JSONFile -Force -ErrorAction SilentlyContinue
            Write-Host ""
            Write-Host "Your template folder doesn't contain any template text files (*.txt)." -ForegroundColor $script:Config.ColorWarning
            Write-Host "Add your first template to the following folder and restart." -ForegroundColor $script:Config.ColorWarning
            Write-Host "$($Folder)" -ForegroundColor $script:Config.ColorHighlight
            Write-Host ""
            Write-Host "Press any key to exit." -ForegroundColor $script:Config.ColorWarning
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit
        }

        $templates | ConvertTo-Json -Depth 3 | Out-File -FilePath $JSONFile -Encoding UTF8
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
            Write-Host "Template cache is corrupt and has been removed." -ForegroundColor $script:Config.ColorWarning
            return $null
        }
        return $templates
    }
}

function Add-SearchWords {
    param($Templates)
    process {
        foreach ($template in $Templates) {
            $template | Add-Member -NotePropertyName Search -NotePropertyValue (
                [PSCustomObject]@{
                    TitleWords   = @(ConvertTo-SearchWords -Text $template.Title)
                    KeywordWords = @(@($template.Keywords) | ForEach-Object { ConvertTo-SearchWords -Text $_ } | Select-Object -Unique)
                    ContentWords = @(ConvertTo-SearchWords -Text $template.Content)
                    PathWords    = @(ConvertTo-SearchWords -Text $template.RelativePath)
                }
            )
        }
        return $Templates
    }
}

function Import-Templates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$ForceReload
    )
    process {

        $script:Stats.RebuildCacheMs = 0
        $script:Stats.PreprocessMs = 0

        # check template cache
        Write-Host "- Checking templates..."
        $isJSON = Test-Path -Path $script:Config.TemplateCacheFile -Type Leaf

        # rebuild cache from .txt files if missing or forced
        if ($ForceReload -or !$isJSON) {
            Write-Host "- Reloading templates from folder..."
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            Convert-TemplatesToJSON -Folder $script:Config.TemplateFolder -JSONFile $script:Config.TemplateCacheFile
            $sw.Stop()
            $script:Stats.RebuildCacheMs = $sw.ElapsedMilliseconds
        }

        # load templates from cache
        Write-Host "- Importing templates..."
        $templates = Import-TemplatesFromJSON -JSONFile $script:Config.TemplateCacheFile

        # if cache was corrupt, rebuild and reimport
        if ($null -eq $templates) {
            Write-Host "- Reloading templates from folder..."
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            Convert-TemplatesToJSON -Folder $script:Config.TemplateFolder -JSONFile $script:Config.TemplateCacheFile
            $sw.Stop()
            $script:Stats.RebuildCacheMs = $sw.ElapsedMilliseconds
            Write-Host "- Importing templates..."
            $templates = Import-TemplatesFromJSON -JSONFile $script:Config.TemplateCacheFile
        }

        # preprocess search words in memory
        Write-Host "- Preprocessing search index..."
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $templates = Add-SearchWords -Templates $templates
        $sw.Stop()
        $script:Stats.PreprocessMs = $sw.ElapsedMilliseconds

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
        Write-Host "Template count:             $(@($Templates).Count)"
        
        $lastWriteTimeCache = (Get-ChildItem -Path $script:Config.TemplateCacheFile).LastWriteTime
        $lastWriteTimeCache_str = $lastWriteTimeCache.ToString("dd-MM-yyyy HH:mm")
        Write-Host "Last reload of templates:   " -NoNewline
        if ($lastWriteTimeCache -lt (Get-Date).AddMonths(-1)) {
            Write-Host "$($lastWriteTimeCache_str) Cache outdated. Run 'r' to reload." -ForegroundColor $script:Config.ColorError
        }
        else {
            Write-Host $lastWriteTimeCache_str
        }

        Write-Host "Template folder:            $($script:Config.TemplateFolder)"
        Write-Host "TTT and config location:    $(Split-Path -Parent $PSCommandPath)"

        if ($script:Config.VerboseMode) {
            Write-Host ""
            Write-Host "PowerShell version:         $($PSVersionTable.PSVersion.ToString())"  -ForegroundColor $script:Config.ColorWarning
            Write-Host "Script file:                $($PSCommandPath)" -ForegroundColor $script:Config.ColorWarning
            Write-Host "Config file:                $(Join-Path (Split-Path -Parent $PSCommandPath) $script:Config.ConfigFilename)" -ForegroundColor $script:Config.ColorWarning
            Write-Host "Cache file:                 $($script:Config.TemplateCacheFile)" -ForegroundColor $script:Config.ColorWarning
            Write-Host "Reload cache on startup:    $($script:Config.ReloadCacheOnStartup)" -ForegroundColor $script:Config.ColorWarning
            Write-Host "Startup message:            $($script:Config.StartupMessage)" -ForegroundColor $script:Config.ColorWarning
            Write-Host "Number of results:          $($script:Config.NumberOfResults)" -ForegroundColor $script:Config.ColorWarning
            Write-Host "Search dimension weights:   title=$($script:Config.SearchDimensionWeightTitle)  path=$($script:Config.SearchDimensionWeightPath)  keywords=$($script:Config.SearchDimensionWeightKeywords)  content=$($script:Config.SearchDimensionWeightContent)" -ForegroundColor $script:Config.ColorWarning
            Write-Host "Search match weights:       exact=$($script:Config.SearchMatchWeightExact)  prefix=$($script:Config.SearchMatchWeightPrefix)  substring=$($script:Config.SearchMatchWeightSubstring)" -ForegroundColor $script:Config.ColorWarning
            Write-Host "Colors:                     selection=$($script:Config.ColorHighlight)  warning=$($script:Config.ColorWarning)  error=$($script:Config.ColorError)  text=$($script:Config.ColorText)  background=$($script:Config.ColorBackground)" -ForegroundColor $script:Config.ColorWarning
            Write-Host "Clipboard function (TEMP):  $($script:Config.TempSelectClipboardFunction)" -ForegroundColor $script:Config.ColorWarning
            Write-Host ""
            Write-Host "Rebuild cache (ms):         $($script:Stats.RebuildCacheMs)" -ForegroundColor $script:Config.ColorWarning
            Write-Host "Preprocess templates (ms):  $($script:Stats.PreprocessMs)" -ForegroundColor $script:Config.ColorWarning
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
        Write-Header -ShowStartupMessage
        Write-StartupScreen

        # infinite main loop
        $topResults = $null
        $selection = 1
        $style = "search"
        $currentQuery = ""
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
                $currentQuery = ""

                Write-StartupScreen
                continue
            }

            # check for commands
            elseif ($query -eq "h") {
                # reset TUI state
                $topResults = $null
                $selection = 1
                $currentQuery = ""

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
                    Start-Process $topResults[$selection - 1].Template.File
                }
                else {
                    Write-Host "No template selected. Search first, then open with 'o'." -ForegroundColor $script:Config.ColorWarning
                    continue
                }
            }
            elseif ($query -eq "p") {
                if ($topResults) {
                    $parentFolder = Split-Path -Path $topResults[$selection - 1].Template.File -Parent
                    Start-Process $parentFolder
                }
                else {
                    Write-Host "No template selected. Search first, then open the parent folder with 'p'." -ForegroundColor $script:Config.ColorWarning
                    continue
                }
            }
            elseif ($query -eq "r") {
                # reset TUI state
                $topResults = $null
                $selection = 1
                $currentQuery = ""

                Write-Header
                $templates = Import-Templates -ForceReload

                Write-Header
                Write-Info -Templates $templates

                continue
            }
            elseif ($query -eq "f") {
                # reset TUI state
                $topResults = $null
                $selection = 1
                $currentQuery = ""

                Start-Process $script:Config.TemplateFolder

                Write-StartupScreen
                continue
            }
            elseif ($query -eq "m") {

                $currentQuery = ""
                $topResults = ForEach ($template in $templates) {
                    [PSCustomObject]@{
                        Template      = $template
                        LastWriteTime = $template.LastWriteTime
                    }
                }
                
                $topResults = $topResults | Sort-Object LastWriteTime -Descending | Select-Object -First $script:Config.NumberOfResults
                
                $selection = 1
                $style = "modified"

                # no template found
                if (@($topResults).Count -eq 0) {
                    Write-Host "No matching template found."
                    continue
                }
            }
            elseif ($query -eq "i") {
                # reset TUI state
                $topResults = $null
                $selection = 1
                $currentQuery = ""

                Write-Info -Templates $templates
                continue
            }

            # check for selection inputs
            elseif ( $topResults -and ($query -match '^[1-9]\d*$') -and ([int]$query -le @($topResults).Count) ) {
                $selection = [int]$query
            }

            # search query
            else {
                $currentQuery = $query
                $topResults = Search-Template -Templates $templates -Query $query
                $selection = 1
                $style = "search"

                # no template found
                if (@($topResults).Count -eq 0) {
                    Write-Host "No matching template found."
                    continue
                }
            }

            if ($topResults) {
                $template = $topResults[$selection - 1].Template
                
                if (-not [string]::IsNullOrWhiteSpace($template.Content)) {
                    Set-ClipboardSafe $template.Content
                    Write-ContentWithHighlights -Content $template.Content -Query $currentQuery
                }
                else {
                    Write-Host "Empty template file." -ForegroundColor $script:Config.ColorWarning
                }

                Write-Host ""
                Write-Host "--------------------------------------------------------------------------------"
                Write-Host ""
                Write-Results -Results $topResults -Selection $selection -Style $style
            }

        } while ( $true )
    }
}

### < EXECUTION >
Start-TextTemplateTool