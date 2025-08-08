<#
    .SYNOPSIS
    Text Template Tool (TTT) by KUHN Engineering

    .DESCRIPTION
    Compact and efficient PowerShell utility that quickly searches through text templates and copies them directly to the clipboard for immediate use.

    .NOTES
    Version:        0.3.0
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
$app_version = "0.3.0"

### < CONFIGURATION >
$CONFIG_personal_config_filename = "config-personal.txt"
$CONFIG_personal_template_filename = "templates-personal.json"
$CONFIG_search_number_of_results = 10
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

### < SUB FUNCTIONS >
function Write-Header {
    process{
        Clear-Host
        Write-Host "################################################################################"
        Write-Host "# Text Template Tool by KUHN Engineering                                (V$($app_version))"
        Write-Host "################################################################################"
        Write-Host ""
    }
}

function Add-DesktopShortcut {
    process {
        $shortcut_path = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "$($app_name).lnk")
        if(!(Test-Path -Path $shortcut_path)){

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
            # TODO: add hotkey to shortcut

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
            "personal-template-folder: $($path)" | Out-File -FilePath $FilePath
        }catch {
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
        [ValidateScript({ Test-Path -Path $_ -Type Leaf })]
        $FilePath
    )
    process {
        if (-not (Test-Path $FilePath)) {
            Write-Error "Config file '$($FilePath)' does not exist."
            return
        }

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
        [string] $Query,

        [Parameter(Mandatory = $true)]
        $Config
    )
    begin{
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
        ForEach($template in $templates){

            $template.score = 0
            ForEach($splitQuery in $splitQueries){
                $score = 0

                # title
                if($template.Title -like "*$splitQuery*"){
                    $score += $factor_title
                }

                # relative path
                if($template.RelativePath -like "*$splitQuery*"){
                    $score += $factor_relativePath
                }

                # keywords
                ForEach($keyword in $template.Keywords){
                    if($keyword -like "*$splitQuery*"){
                        $score += $factor_keywords
                    }
                }

                # content
                ForEach($content in $template.Content){
                    if($content -like "*$splitQuery*"){
                        $score += $factor_content
                    }
                }

                $template.Score += $score
            }
        }
        
        $results = $templates `
            | Where-Object{$_.Score -gt 0} `
            | Sort-Object -Property Score -Descending `
            | Select-Object -First $CONFIG_search_number_of_results

        return $results
    }
}

function Write-SearchResults {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Templates,
        [Parameter(Mandatory = $false)]
        $Selection = 1
    )
    process {

        $cnt =  0
        ForEach($template in $templates){
            $cnt += 1
            # TODO: dynamic padding depending on number_of_results
            $cntStr = $cnt.ToString().PadLeft(3, ' ')
            if($cnt -eq $Selection){
                Write-Host " $cntStr | $($template.Title) [$($template.Score)]" -ForegroundColor Yellow
            }else{
                Write-Host " $cntStr | $($template.Title) [$($template.Score)]"
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
        $files = Get-ChildItem -Path $folder -Recurse -File -Include *.txt

        # process files
        ForEach($file in $files){
            $templates += Get-TemplateFromFile -FilePath $file -BaseFolder $TemplateFolder
        }

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

        # get content
        $rawContent = Get-Content -Path $FilePath -Raw -Encoding UTF8
        $lines = $rawContent -split '\r?\n'
        $numberOfLines = $lines.Count
        
        # extract keywords
        $firstLine = $lines[0]
        $isKeywords = $firstLine.StartsWith("KEYWORDS:")
        $keywords = @()
        if ($isKeywords) {
            $keywords = $firstLine.Replace("KEYWORDS:", "").Split(",") | ForEach-Object { $_.Trim()}
            $content = $lines[1..($numberOfLines-1)]
        }else{
            $content = $lines
        }

        $template = [PSCustomObject]@{
            Title           = $title
            RelativePath    = $relativePath
            Keywords        = $keywords
            Content         = $content
            Score           = 0
            File            = $FilePath
            Personal        = $Personal
        }

        return $template
    }
}

function Convert-Templates2JSON {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateScript({ Test-Path -Path $_ -Type Container })]
        $Folder = $True,

        [Parameter(Mandatory = $true)]
        $JSONFile = $False
    )
    process {
        $templates = Get-TemplatesFromFolder -TemplateFolder $Folder
        $templates | ConvertTo-Json | Out-File -FilePath $JSONFile -Encoding UTF8
    }
}

function Read-TemplateJSON {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -Type Leaf })]
        $FilePath
    )
    process {
        return Get-Content -Path $FilePath -Encoding UTF8 | ConvertFrom-Json
    }
}

### < MAIN FUNCTION >
function Start-TextTemplateTool {
    process {

        # startup procedure
        Write-Header

        Write-Host "- Checking for desktop shortcut..."
        Add-DesktopShortcut

        Write-Host "- Loading configuration..."
        $personal_config_file = Resolve-Path -Path $PSCommandPath | Split-Path -Parent | Join-Path -ChildPath $CONFIG_personal_config_filename

        if(!(Test-Path -Path $personal_config_file)){
            Set-Config -FilePath $personal_config_file
        }

        $config = Read-Config -FilePath $personal_config_file
        $personal_template_folder =  Resolve-Path -Path $config['personal-template-folder']

        Write-Host "- Loading personal templates..."
        $personal_template_file = Resolve-Path -Path $PSCommandPath | Split-Path -Parent | Join-Path -ChildPath $CONFIG_personal_template_filename
        Convert-Templates2JSON   -Folder $personal_template_folder -JSONFile $personal_template_file
        $templates = Read-TemplateJSON -FilePath $personal_template_file

        Write-Header
        Write-Host "Enter command 'h' for help."
        Write-Host ""
        Write-Host "Personal template folder:   $($personal_template_folder)"
        Write-Host "Personal template count:    $(@($templates).Count)"

        if(!$templates){
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
        $isSelection = $false
        do{
            Write-Host ""
            Write-Host "--------------------------------------------------------------------------------"
            $query = Read-Host "> Search / Select / Command"
            Write-Header

            # empty input -> cycle
            if($query -eq ""){
                $selection = 1
                continue
            }

            # check for commands
            elseif($query -eq "h"){
                Write-Host "Available commands:"
                Write-Host ""
                Write-Host "   h    help       Displays all available commands."
                Write-Host "   f    folder     Opens your personal template folder in File Explorer."
                Write-Host "   r    reload     Reloads templates from your personal template folder."
                Write-Host "   o    open       Opens the selected template text file."
                Write-Host "   q    quit       Exits the Text Template Tool."
                continue
            }
            elseif($query -eq "q"){
                Clear-Host
                exit
            }
            elseif($query -eq "o"){
                if($isSelection){
                    Start-Process $topResults[$Selection-1].File
                }else{
                    Write-Host "Search and select a template before opening with command 'o'."
                    continue
                }
            }
            elseif($query -eq "r"){
                Write-Header
                Write-Host "- Reloading personal templates..."
                Convert-Templates2JSON -Folder $personal_template_folder -JSONFile $personal_template_file
                $templates = Read-TemplateJSON -File $personal_template_file

                Write-Host ""
                Write-Host "Personal template folder:   $($personal_template_folder)"
                Write-Host "Personal template count:    $(@($templates).Count)"

                $isSelection = $false
                continue
            }
            elseif($query -eq "f"){
                Start-Process $personal_template_folder
                continue
            }

            # check for selection inputs
            elseif( $isSelection -and ($query -match '^(?:[1-9]|[1-9][0-9]|[1-9][0-9]{2})$') ){
                if ([int]$query -gt $topResults.Count){
                    $selection = 1
                }else{
                    $selection = [int]$query
                }
            }

            # search query
            else{
                $selection = 1
                $topResults = Search-Template -Templates $templates -Query $query -Config $config

                # no template found
                if (!$topResults){
                    Write-Host "No matching template found."
                    $isSelection = $false
                    continue
                }else{
                    $isSelection = $true
                }
            }

            if($isSelection){
                $template = $topResults[$Selection-1]
                $template.Content | Set-Clipboard
                $template.Content | Write-Host

                Write-Host ""
                Write-Host "--------------------------------------------------------------------------------"
                Write-Host ""
                Write-SearchResults -Templates $topResults -Selection $selection
            }

        } while ( $true )
    }
}

### < EXECUTION >
Start-TextTemplateTool