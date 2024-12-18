<#PSScriptInfo
.VERSION 1.0.1
.GUID d0f7effa-e5c1-482c-98bd-f942160d246e
.AUTHOR Adam Danischewski
.COMPANYNAME Adam Danischewski
.COPYRIGHT (c) 2024 Adam Danischewski
.TAGS PowerShell Windows
.PROJECTURI https://github.com/AdamDanischewski/AppMgr
.DESCRIPTION Simplifies management (add/remove/repair/reset) of apps using winget. Allows adding, removing, resetting, and repairing apps using shortnames.
.ICONURI
.EXTERNALMODULEDEPENDENCIES
.REQUIREDSCRIPTS
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES 
    [Version 0.0.1] - Initial Release.
    [Version 0.0.2] - Refactored, updated parameter logic for AppMgr.
    [Version 1.0.0] - Removed trailing whitespace and invoke-expression syntax.
    [Version 1.0.1] - Removed trailing whitespace and invoke-expression syntax.
.PRIVATEDATA
#>

function AppMgr {
    <#
.SYNOPSIS
Simplifies management (add/remove/repair/reset) of apps
.DESCRIPTION

This script allows the user to add/remove/reset/repair apps using shortnames, if multiple matches are found they will be listed for the user to select. Requires winget - app will automatically downloads and installs the latest version of winget and its dependencies if necessary.
.EXAMPLE
AppMgr add calc
AppMgr -AppName calc -Action add
.EXAMPLE
AppMgr repair wt 1
AppMgr -AppName terminal -Action repair -msStoreVer $true

Note: "terminal" and "wt" are both short names for Windows Terminal.
.EXAMPLE
AppMgr remove wt 1
AppMgr -AppName wt -Action remove -msStoreVer $true
AppMgr -name wt -a remove -S 1
.EXAMPLE
AppMgr reset notepad 1
AppMgr -AppName wt -Action remove -msStoreVer $true
AppMgr -name wt -a remove -S 1
.PARAMETER Version
Displays the version of the script.
.PARAMETER Help
Displays the full help information for the script.
.NOTES
Version : 1.0.1
Created by : Adam Danischewski
.LINK
Project Site: https://github.com/AdamDanischewski/AppMgr
#>
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingCmdletBinding', '')]
    param(
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'Action')]
        [Alias('a')][string]$Action,
        [Parameter(Mandatory = $false, Position = 1)][Alias('n')][string]$AppName,
        [Parameter(Mandatory = $false, Position = 2)][ValidateSet("true", "1", $true, "false", $false, 0)]
        [Alias('ms', 'store', 's')][object]$msStoreVer = $false,
        [Parameter(Mandatory = $false, Position = 3)][Alias('v')][switch]$Version
    )
    [string]$c_InvalidOption = ">> Invalid option (`"{0}`"): Valid options are: remove, reset, repair, add or help."
    [string]$c_PackageNotFound = "Package ({0}) not found. Are you sure it's installed ?"
    [string]$bad_action = ""
    [string]$c_Version = "1.0.1"

    if ($Version) { $Action = "version" }

    if ($Action -notin @("add", "repair", "remove", "reset", "help", "version")) {
        $bad_action = $Action
        $Action = "help"
    }
    else {
        ## Make sure we have winget
        $winget_cmd = get-command winget -ErrorAction SilentlyContinue
        if (-not $winget_cmd) { Install-AppMgrWinGet }

        $msStoreVer = $msStoreVer -in @($true, "true", 1)
    }

    # Write-Host "Value of `$Action is: $Action"
    $AppName = if ($Action -notin @("help", "version")) {
        if ($Action -in @("remove", "reset")) {
            Get-AppMgrMapping $AppName $false
        }
        else {
            Get-AppMgrMapping $AppName ($msStoreVer -or ($Action -in @("repair", "add")))
        }
        Write-Host "[Processing: $AppName with command $Action]" -ForegroundColor Blue
    }

    switch ($Action) {
        "remove" {
            Write-Host "Removing app: $AppName ..."
            $package = Get-AppxPackage -Name $AppName
            if ($package) {
                Remove-AppxPackage -Package $package -WhatIf:$false -Confirm:$false
                if ($LASTEXITCODE -eq 0) { Write-Host "Successfully removed app: $AppName." -f Yellow }
                else { Write-AppMgrCustomError ("Package {0} could not be removed." -f $AppName) }
            }
            else {
                Write-AppMgrCustomError ($c_PackageNotFound -f $AppName)
            }
        }
        "reset" {
            Write-Host "Resetting app: $AppName ..."
            $package = Get-AppxPackage -Name $AppName
            if ($package) {
                Reset-AppxPackage -Package $package -WhatIf:$false -Confirm:$false
                if ($LASTEXITCODE -eq 0) { Write-Host "Successfully reset app: $AppName." -f Yellow }
                else { Write-AppMgrCustomError ("Package {0} could not be reset." -f $AppName) }
            }
            else {
                Write-AppMgrCustomError ($c_PackageNotFound -f $AppName)
            }
        }
        "add" {
            Write-Host "Adding app: $AppName ..."
            $packageID = Get-AppMgrAppID $AppName $msStoreVer
            if ($null -ne $packageID) {
                winget install $packageID --accept-source-agreements --accept-package-agreements
            }
            else {
                Write-AppMgrCustomError ($c_PackageNotFound -f $AppName)
            }
        }
        "version" {
            Write-Output "AppMgr Version - $c_Version"
        }
        "repair" {
            Write-Host "Repairing app: $packageName ..."
            Write-Host "[Searching for $AppName app id ..]" -ForegroundColor Blue
            $ErrorActionPreference = 'Stop'
            $packageID = (Get-AppMgrAppID $AppName $msStoreVer)
            if ($null -ne $packageID) {
                winget repair $packageID --accept-source-agreements --accept-package-agreements
            }
            else {
                Write-AppMgrCustomError ($c_PackageNotFound -f $AppName)
            }
        }
        "help" {
            $heredoc = @"
Usage:
AppMgr (remove|reset|repair|add) <app_name> [MsStoreVer = (1,`$true,"true")]
AppMgr help

Note: Setting the MsStoreVer flag restricts winget to source = msstore.

Supported short names:
`t  Short alias       Appx/(Winget(`$msStoreVer=`$true)):
`t  -----------       -----------------------------
`t  paint,mspaint  -  Microsoft.Paint / Paint
`t  calc           -  Microsoft.WindowsCalculator / Windows Calculator
`t  wt,terminal    -  Microsoft.WindowsTerminal / Windows Terminal

Examples:
Remove Paint app: AppMgr remove paint
Reset MS Store Calculator app: AppMgr reset calc `$true
Reset MS Store Calculator app: AppMgr reset calc 1
Remove Calculator app: AppMgr -AppName calc -Action remove
Repair MS WT app: AppMgr -Action repair -AppName terminal -msStoreVer 1
Add MS Store Paint app: AppMgr add paint 1
"@
            Write-Host $heredoc
            if (![string]::IsNullOrEmpty($bad_action)) { Write-Host ($c_InvalidOption -f $bad_action) -ForegroundColor Yellow }
        }
        default {
            Write-AppMgrCustomError ($c_InvalidOption -f $Action)
        }
    }
}

function Install-AppMgrWinGet {
    Write-host (">> Winget is not installed, would you like to install it via winget-install`n" +
        "from Powershell gallery?") -ForegroundColor DarkCyan
    $selection = Read-Host "Enter [Y/n]:"
    if ($selection -eq "Y") {
        $wgi_process = Start-Process powershell -Verb RunAs '-Command', "install-script winget-install -force -confirm" -wait -PassThru
        if ($wgi_process.ExitCode -eq 0) {
            Write-Host ">> Installed winget-install .." -ForegroundColor Green
        }
        else {
            throw ("Couldn't install winget-install exit code was: {0}" -f $wgi_process.ExitCODE)
        }
        $wgicmd = Get-Command winget-install | Select-Object -ExpandProperty Definition
        $wgi_cmd_process = Start-Process powershell -Verb RunAs "-Command", "$wgicmd -force -wait" -wait -PassThru
        if ($wgi_cmd_process.ExitCode -in @(0, 1)) {
            Write-host ">> Success, winget is now installed!!" -ForegroundColor Green
        }
        else {
            throw "Something went wrong installing winget, exit code: $($wgi_cmd_process.ExitCode)"
        }
    }
    else {
        throw "Couldn't find winget-install - perhaps paths? Not sure ;( .."
    }
}

function Get-AppMgrMapping {
    param(
        [Parameter(Mandatory = $true, Position = 0)][string]$AppName,
        [Parameter(Mandatory = $false, Position = 1)][object]$msStoreVer = $false
    )
    $msStoreVer = $msStoreVer -in @($true, "true", 1)
    $appMap = @{}
    $appMap["paint"] = if ($msStoreVer) { "Paint" } else { "Microsoft.Paint" }
    $appMap["mspaint"] = $appMap["paint"]
    $appMap["calc"] = if ($msStoreVer) { "Windows Calculator" } else { "Microsoft.WindowsCalculator" }
    $appMap["wt"] = if ($msStoreVer) { "Windows Terminal" } else { "Microsoft.WindowsTerminal" }
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

function Get-AppMgrAppID {
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

    $AppName = Get-AppMgrMapping $AppName

    $wingetArgs = @("search", "--name", "`"$AppName`"")
    if ($msStoreVer) {
        $wingetArgs += "-s", "`"$c_msstore`""
    }

    # Search for the app and filter out unwanted lines
    try {
        $searchResults = & winget @wingetArgs | Where-Object {
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
        Write-AppMgrCustomError "Failed to execute winget search: $_"
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
function Write-AppMgrCustomError {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
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


# Argument completers for AppName (using existing app mappings)
## Add these to your profile if you'd like
Register-ArgumentCompleter -CommandName AppMgr -ParameterName Action -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    # Define the valid actions
    $validActions = @("add", "remove", "repair", "reset", "help", "version")

    $validActions | Where-Object {
        $_ -like "$wordToComplete*"
    } | ForEach-Object {
        [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', $_)
    }
}

Register-ArgumentCompleter -CommandName AppMgr -ParameterName AppName -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    # Define your app mappings (similar to your existing Get-AppMgrMapping function)
    $appMappings = @(
        "paint", "mspaint",
        "calc",
        "wt", "terminal", "windowsterminal"
        # Add more as needed
    )

    $appMappings | Where-Object {
        $_ -like "$wordToComplete*"
    } | ForEach-Object {
        [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', $_)
    }
}
