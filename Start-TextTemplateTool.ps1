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
$CONFIG_personal_config_filename = "config-personal.txt"
$CONFIG_personal_template_filename = "templates-personal.json"
$CONFIG_number_of_results = 10
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

### < SUB FUNCTIONS >
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
        $shortcut_path = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "$($app_name).lnk")
        if (!(Test-Path -Path $shortcut_path)) {

            Write-Host "- Adding desktop shortcut..."
            
            $script_path = Resolve-Path -Path $PSCommandPath
            $ps_path = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"

            $WShell = New-Object -ComObject WScript.Shell
            $shortcut = $WShell.CreateShortcut($shortcut_path)

            $shortcut.TargetPath = $ps_path
            $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$script_path`""
            $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($script_path)
            $shortcut.Description = "Desktop Shortcut for $($app_name)"
            $shortcut.IconLocation = "$PowerShellPath, 0"

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
        Write-Host ""
        Write-Host "Configuration file '$($CONFIG_personal_config_filename)' not found." -ForegroundColor Yellow
        Write-Host "Set your desired personal template folder below by drag and drop, copy/paste or typing." -ForegroundColor Yellow
        Write-Host ""
        $path = Read-Host "> Personal template folder"

        try {
            $path = Resolve-Path -Path $path -ErrorAction Stop
            "personal-template-folder: $($path)" | Out-File -FilePath $FilePath -Encoding UTF8
        }
        catch {
            Write-Host "Unable to set configuration. Invalid path or access denied." -ForegroundColor Red
            Write-Host "Please restart or add configuration file manually." -ForegroundColor Red
            Write-Host "Press any key to exit." -ForegroundColor Red
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
            $key, $value = $_ -split ":", 2 | ForEach-Object { $_.Trim() }
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
    begin {
        # search algorithm configuration
        $factor_title = 10
        $factor_relativePath = 10
        $factor_keywords = 5
        $factor_content = 2
    }
    process {
        # preprocess search query
        $splitQueries = $Query.Split(" ").Trim()

        # iterate among templates and calculate query match score
        ForEach ($template in $templates) {

            $template.score = 0
            ForEach ($splitQuery in $splitQueries) {
                $score = 0

                # title
                if ($template.Title -like "*$splitQuery*") {
                    $score += $factor_title
                }

                # relative path
                if ($template.RelativePath -like "*$splitQuery*") {
                    $score += $factor_relativePath
                }

                # keywords
                ForEach ($keyword in $template.Keywords) {
                    if ($keyword -like "*$splitQuery*") {
                        $score += $factor_keywords
                    }
                }

                # content
                ForEach ($content in $template.Content) {
                    if ($content -like "*$splitQuery*") {
                        $score += $factor_content
                    }
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
            Write-Progress  -Activity "Processing templates" `
                -Status "Template $($cnt) of $($files.Count): $($file.Name)" `
                -PercentComplete $percentComplete

            $templates += Get-TemplateFromFile -FilePath $file -BaseFolder $TemplateFolder
        }
        Write-Progress -Activity "Processing Text Files" -Completed

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
        $lines = $rawContent -split '\r?\n'
        $numberOfLines = $lines.Count
        
        # extract keywords
        $firstLine = $lines[0]
        $isKeywords = $firstLine.StartsWith("KEYWORDS:")
        $keywords = @()
        if ($isKeywords) {
            $keywords = $firstLine.Replace("KEYWORDS:", "").Split(",") | ForEach-Object { $_.Trim() }
            $content = $lines[1..($numberOfLines - 1)]
        }
        else {
            $content = $lines
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
        $templates | ConvertTo-Json | Out-File -FilePath $JSONFile -Encoding UTF8
    }
}

function Import-TemplatesFromJSON {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -Type Leaf })]
        $JSON
    )
    process {
        return Get-Content -Path $JSON -Encoding UTF8 | ConvertFrom-Json
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

        # checking situation
        Write-Host "- Checking templates..."
        $isJSON = Test-Path -Path $TemplateFile

        # reload if needed
        if ($ForceReload -or !$isJSON) {
            Write-Host "- Reloading templates from folder..."
            Convert-TemplatesToJSON -Folder $TemplateFolder -JSONFile $TemplateFile
        }

        # import templates from json
        Write-Host "- Importing templates..."
        $templates = Import-TemplatesFromJSON -JSON $TemplateFile

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
        $personal_config_file = Resolve-Path -Path $PSCommandPath | Split-Path -Parent | Join-Path -ChildPath $CONFIG_personal_config_filename

        if (!(Test-Path -Path $personal_config_file)) {
            Set-Config -FilePath $personal_config_file
        }

        $config = Read-Config -FilePath $personal_config_file

        # derive files and folders from config
        $personal_template_folder = Resolve-Path -Path $config['personal-template-folder']
        $personal_template_file = Resolve-Path -Path $PSCommandPath | Split-Path -Parent | Join-Path -ChildPath $CONFIG_personal_template_filename

        # import templates
        $templates = Import-Templates -TemplateFolder $personal_template_folder -TemplateFile $personal_template_file

        # write startup screen
        Write-Header
        Write-StartupScreen

        if (!$templates) {
            Write-Host ""
            Write-Host "Your personal template folder doesn't contain any template text files (*.txt)." -ForegroundColor Yellow
            Write-Host "Add your first template to $($personal_template_folder) and restart $($app_name)." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Press any key to exit." -ForegroundColor Yellow
            Read-Host
            exit
        }

        # infinite main loop
        $selection = 1
        $style = "search"
        $isSelection = $false
        do {
            Write-Host ""
            Write-Host "--------------------------------------------------------------------------------"
            $query = Read-Host "> Search / Select / Command"
            Write-Header

            # empty input -> cycle
            if ($query -eq "") {
                $selection = 1
                $style = "search"
                Write-StartupScreen
                continue
            }

            # check for commands
            elseif ($query -eq "h") {
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
                
                $selection = 0
                continue
            }
            elseif ($query -eq "q") {
                Clear-Host
                exit
            }
            elseif ($query -eq "o") {
                if ($isSelection) {
                    Start-Process $topResults[$Selection - 1].File
                }
                else {
                    Write-Host "Search and select a template before opening with command 'o'."
                    continue
                }
            }
            elseif ($query -eq "p") {
                if ($isSelection) {
                    $parentFolder = Split-Path -Path $topResults[$Selection - 1].File -Parent
                    Start-Process $parentFolder
                }
                else {
                    Write-Host "Search and select a template before opening parent folder with command 'p'."
                    continue
                }
            }
            elseif ($query -eq "r") {
                Write-Header
                $templates = Import-Templates -TemplateFolder $personal_template_folder -TemplateFile $personal_template_file -ForceReload

                Write-Header
                Write-Info -TemplateFolder $personal_template_folder -TemplateFile $personal_template_file -Templates $templates

                $isSelection = $false
                continue
            }
            elseif ($query -eq "f") {
                Start-Process $personal_template_folder
                continue
            }
            elseif ($query -eq "m") {

                $selection = 1
                $style = "modified"
                $topResults = $templates | Sort-Object LastWriteTime -Descending | Select-Object -First $CONFIG_number_of_results

                # no template found
                if (!$topResults) {
                    Write-Host "No matching template found."
                    $isSelection = $false
                    continue
                }
                else {
                    $isSelection = $true
                }
            }
            elseif ($query -eq "i") {
                Write-Info -TemplateFolder $personal_template_folder -TemplateFile $personal_template_file -Templates $templates
                $isSelection = $false
                continue
            }

            # check for selection inputs
            elseif ( $isSelection -and ($query -match '^(?:[1-9]|[1-9][0-9]|[1-9][0-9]{2})$') ) {
                if ([int]$query -gt $topResults.Count) {
                    $selection = 1
                }
                else {
                    $selection = [int]$query
                }
            }

            # search query
            else {
                $selection = 1
                $style = "search"
                $topResults = Search-Template -Templates $templates -Query $query

                # no template found
                if (!$topResults) {
                    Write-Host "No matching template found."
                    $isSelection = $false
                    continue
                }
                else {
                    $isSelection = $true
                }
            }

            if ($isSelection) {
                $template = $topResults[$Selection - 1]
                $template.Content | Set-Clipboard
                $template.Content | Write-Host

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