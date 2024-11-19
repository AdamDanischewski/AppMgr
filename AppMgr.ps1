<#PSScriptInfo
 
.VERSION 0.0.1
 
.GUID d0f7effa-e5c1-482c-98bd-f942160d246e
 
.AUTHOR adanisch
 
.COMPANYNAME adanisch
 
.TAGS PowerShell Windows winget win get install installer fix script setup
 
.PROJECTURI https://github.com/AdamDanischewski/AppMgr
 
.RELEASENOTES
[Version 0.0.1] - Initial Release. 
#>

<#
.SYNOPSIS
    Downloads and installs the latest version of winget and its dependencies.
.DESCRIPTION
    Downloads and installs the latest version of winget and its dependencies.
 
This script is designed to be straightforward and easy to use, removing the hassle of manually downloading, installing, and configuring winget. This function should be run with administrative privileges.
.EXAMPLE
    winget-install
.PARAMETER Debug
    Enables debug mode, which shows additional information for debugging.
.PARAMETER Force
    Ensures installation of winget and its dependencies, even if already present.
.PARAMETER ForceClose
    Relaunches the script in conhost.exe and automatically ends active processes associated with winget that could interfere with the installation.
.PARAMETER Wait
    Forces the script to wait several seconds before exiting.
.PARAMETER NoExit
    Forces the script to wait indefinitely before exiting.
.PARAMETER UpdateSelf
    Updates the script to the latest version on PSGallery.
.PARAMETER CheckForUpdate
    Checks if there is an update available for the script.
.PARAMETER Version
    Displays the version of the script.
.PARAMETER Help
    Displays the full help information for the script.
.NOTES
    Version : 5.0.4
    Created by : asheroto
.LINK
    Project Site: https://github.com/asheroto/winget-install
#>
function AppManager {
    [string]$c_InvalidOption = " >> Invalid option ({0}): Valid options are --remove, --reset, --repair, --add or --help." 
    [string]$c_PackageNotFound = "Package ({0}) not found. Are you sure it's installed?"
    [bool]$f_option_error = $false 
    [string]$bad_action = "" 

    ## Sanity check, make sure we have winget  
    $winget_cmd = get-command winget -ErrorAction SilentlyContinue
    if (-not $winget_cmd) {
        Write-host (">> Winget is not installed, would you like to install it via winget-install`n" +
                   "from Powershell gallery?") -ForegroundColor DarkCyan
        $selection = Read-Host "Enter [Y/n]:"
        if ($selection -eq "Y") { 
            $wgi_process = Start-Process powershell -Verb RunAs '-Command', "install-script winget-install -force -confirm" -wait -PassThru
            if ($wgi_process.ExitCode -eq 0) { 
                Write-Host ">> Installed winget-install .." -ForegroundColor Green
            } else { 
                throw ("Couldn't install winget-install exit code was: {0}" -f $wgi_process.ExitCODE) 
            }
            $wgicmd = Get-Command winget-install | Select-Object -ExpandProperty Definition
            $wgi_cmd_process = Start-Process powershell -Verb RunAs "-Command", "$wgicmd -force -wait" -wait -PassThru
            if ($wgi_cmd_process.ExitCode -in @(0,1)) { 
                Write-host ">> Success, winget is now installed!!" -ForegroundColor Green
            } else { 
                throw "Something went wrong installing winget, exit code: $($wgi_cmd_process.ExitCode)"               
            } 
        } else { 
            throw "Couldn't find winget-install - perhaps paths? Not sure ;( .."
        }
    }

    # Parse action, we expect positional args: (first argument)
    $Action = if ($args[0] -match '^-{1,2}(add|remove|reset|repair|help)$') {
        "--$($matches[1])"
    }
    else {
        $f_option_error = $true
        $bad_action = $args[0] ## No match to above regex
        "--help"
    }
    
    $appName = $args[1]
    $msStoreVer = if ($args[2]) { $args[2] -in @($true, "true", 1) } else { $false }
    
    # Write-Host "Value of `$Action is: $Action"  
    $AppName = if ($Action -ne "--help") { 
        if ($Action -in @("--remove", "--reset")) { 
            getAppMapping $AppName $false
        }
        else {             
            getAppMapping $AppName ($msStoreVer -or ($Action -in @("--repair", "--add")))
        }
        Write-Host "[Processing: $AppName with command $Action]" -ForegroundColor Blue
    }

    switch ($Action) {
        "--remove" {
            Write-Host "Removing app: $AppName ..."
            $package = Get-AppxPackage -Name $AppName
            if ($package) {
                Remove-AppxPackage -Package $package -WhatIf:$false -Confirm:$false
                if ($LASTEXITCODE -eq 0) { Write-Host "Successfully removed app: $AppName." -f Yellow }
                else { Write-CustomError ("Package {0} could not be removed." -f $AppName) }
            }
            else { 
                Write-CustomError ($c_PackageNotFound -f $AppName)
            }
        }
        "--reset" {
            Write-Host "Resetting app: $AppName ..."
            $package = Get-AppxPackage -Name $AppName
            if ($package) {
                Reset-AppxPackage -Package $package -WhatIf:$false -Confirm:$false
                if ($LASTEXITCODE -eq 0) { Write-Host "Successfully reset app: $AppName." -f Yellow }
                else { Write-CustomError ("Package {0} could not be reset." -f $AppName) }
            }
            else { 
                Write-CustomError ($c_PackageNotFound -f $AppName)
            }
        }
        "--add" {
            Write-Host "Adding app: $AppName ..."
            $packageID = Get-AppID $AppName $msStoreVer
            if ($null -ne $packageID) { 
                winget install $packageID --accept-source-agreements --accept-package-agreements
            }
            else { 
                Write-CustomError ($c_PackageNotFound -f $AppName)
            }
        }
        "--repair" {
            Write-Host "Repairing app: $packageName ..."
            Write-Host "[Searching for $AppName app id ..]" -ForegroundColor Blue
            $ErrorActionPreference = 'Stop'
            $packageID = (Get-AppID $AppName $msStoreVer)
            if ($null -ne $packageID) { 
                winget repair $packageID --accept-source-agreements --accept-package-agreements
            }
            else { 
                Write-CustomError ($c_PackageNotFound -f $AppName)
            }
        }
        "--help" {
            $heredoc = @"
Usage:
       AppManager (--remove|--reset|--repair|--add) <app_name> [MsStoreVer = (1,`$true,"true")]
       AppManager --help 

Note: Setting the MsStoreVer flag restricts winget to source = msstore.   

Supported short names: 
`t  Short alias       Appx/(Winget(`$msStoreVer=`$true)):
`t  -----------       -----------------------------
`t  paint,mspaint  -  Microsoft.Paint / Paint
`t  calc           -  Microsoft.WindowsCalculator / Windows Calculator
`t  wt,terminal    -  Microsoft.WindowsTerminal / Windows Terminal

Examples:
               Remove Paint app: AppManager --remove paint
  Reset MS Store Calculator app: AppManager --reset calc `$true
  Reset MS Store Calculator app: AppManager --reset calc 1
          Remove Calculator app: AppManager --remove calc
               Repair MS WT app: AppManager --repair terminal 1
               Add MS Paint app: AppManager --add paint 1

"@
            Write-Host $heredoc
            Write-Host ($c_InvalidOption -f $bad_action) -ForegroundColor Yellow

        }
        default {
            # Write-Host ($c_InvalidOption -f $Action) -ForegroundColor Red
            Write-CustomError ($c_InvalidOption -f $Action)
        }
    }
}

function getAppMapping { 
    param(
        [Parameter(Mandatory = $true, Position = 0)][string]$AppName,
        # $WinGetVer determines mapping type:
        # true  -> Returns display name for WinGet search (e.g., "Windows Terminal")
        # false -> Returns AppX package name (e.g., "Microsoft.WindowsTerminal")
        [Parameter(Mandatory = $false, Position = 1)][object]$WinGetVer = $false
    )
    $WinGetVer = $WinGetVer -in @($true, "true", 1)
    $appMap = @{}
    $appMap["paint"] = if ($WinGetVer) { "Paint" } else { "Microsoft.Paint" }
    $appMap["mspaint"] = $appMap["paint"]
    $appMap["calc"] = if ($WinGetVer) { "Windows Calculator" } else { "Microsoft.WindowsCalculator" }
    $appMap["wt"] = if ($WinGetVer) { "Windows Terminal" } else { "Microsoft.WindowsTerminal" }
    $appMap["terminal"] = $appMap["wt"]
    $appMap["windowsterminal"] = $appMap["wt"]
    # Add more mappings as needed...
    if ($appMap.ContainsKey($AppName)) {
        $mapped_val = $appMap[$AppName]
        Write-Host -ForegroundColor yellow ">> Converting short name $AppName to $mapped_val."  
        return $mapped_val 
    }
    else { return $AppName }
}

function Get-AppID {
    param(
        [Parameter(Mandatory = $true, Position = 0)][string]$AppName,
        [Parameter(Mandatory = $false, Position = 1)][object]$msStoreVer
    )

    [string]$c_nomatch = "No package found matching input criteria."
    [string]$c_msstore = "msstore"
    [string]$c_regex_full = '(.*?)\s+(\S+)\s+(\S+)\s+(\S+)$'
    [string]$c_regex_msstore = '(.*?)\s+(\S+)\s+(\S+)$'
    [string]$c_regex = $c_regex_full
    [int]$results_count = 0
    
    if ($msStoreVer -eq $true -or $msStoreVer -eq "true" -or $msStoreVer -eq 1) {
        $msStoreVer = $true
        $c_regex = $c_regex_msstore
    }
    else {
        $msStoreVer = $false
    }

    $AppName = getAppMapping $AppName

    $wingetSrchCmd = "winget search --name `"$AppName`"" 
    if ($msStoreVer) {
        $wingetSrchCmd += " -s `"$c_msstore`""
    }

    # Search for the app and filter out unwanted lines
    try { 
        $searchResults = Invoke-Expression $wingetSrchCmd | Where-Object {
            $_ -and # Remove empty lines
            $_ -notmatch '^\s*[-\\|/]\s*$' -and # Remove spinner characters
            $_ -notmatch '^-+$' -and # Remove separator lines
            $_ -notmatch '^\s*$' -and # Remove whitespace-only lines
            $_ -notmatch 'Windows Package Manager' -and # Remove header
            $_ -notmatch 'Copyright \(c\)' -and # Remove copyright
            $_ -notmatch 'Name\s+Id\s+Version' -and # Remove column headers
            $_ -notmatch 'Name\s+Id\s+Version\s+Source'  # Remove column headers
        } | ForEach-Object {            
            # Convert multiple spaces to a single space and trim
            $line = $_ -replace '\s+', ' ' 
            
            if ($_ -eq $c_nomatch) { 
                Write-Warning "$c_nomatch ($AppName)" 
            }   
            elseif ($line -match $c_regex) {
                $fullName = $Matches[1].Trim()
                $id = $Matches[2]
                $version = $Matches[3]
                $source = if ($msStoreVer) { $c_msstore } else { $Matches[4].Trim() }
                
                [PSCustomObject]@{
                    Name    = $fullName
                    ID      = $id
                    Version = $version
                    Source  = $source
                }
                $results_count++
            }
            else {
                Write-Warning "Unexpected format: $_"
                $null
            }
        } | ConvertTo-Json
    } 
    catch {
        Write-CustomError "Failed to execute winget search: $_"
        return $null
    }
    if ($results_count -eq 0) { 
        return $null
    }
    elseif ($results_count -eq 1) { 
        return ($searchResults | ConvertFrom-Json).ID
    }
    else {
        $apps = $searchResults | convertfrom-json
        [int]$index = $null            
        Write-Host "Multiple matches found. Please select an app (or press Ctrl-C to abort):`n"
        for ($i = 0; $i -lt $results_count; $i++) {
            Write-Host ("{0}.{1} Version: {2}, ID: {3}, Source: {4}" -f 
                    ($i + 1), 
                $apps[$i].Name.Trim('"'), 
                $apps[$i].Version, 
                $apps[$i].ID,
                $apps[$i].Source.Trim('"'))
        }            
        do {
            $selection = Read-Host "`nEnter the number of your selection (1-$results_count)"
            $index = if ([int]::TryParse($selection, [ref]$null)) { [int]$selection - 1 } else { -1 }
        } until ($index -ge 0 -and $index -lt $results_count)
            
        if ($null -ne $index) {
            return $apps[$index].ID
        }
        else {
            Write-Error "No App found matching '$AppName'" -Category ObjectNotFound
        }
    }
}

function Write-CustomError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [System.Management.Automation.ErrorCategory]$Category = 'NotSpecified'
    )
    try {
        Write-Error -Message $Message -Category $Category -ErrorAction Stop
    }
    catch {
        Write-Host "$($_.Exception.Message)" -ForegroundColor Red
        # Write-Host "+ $($_.InvocationInfo.Line)" -ForegroundColor Red
        # Write-Host "    + CategoryInfo          : $($_.CategoryInfo)" -ForegroundColor Red
        # Write-Host "    + FullyQualifiedErrorId : $($_.FullyQualifiedErrorId)" -ForegroundColor Red
    }
}
