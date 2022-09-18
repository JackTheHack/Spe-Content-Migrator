Clear-Host

Import-Module -Name SPE -Force

$scriptDirectory = & {
    if ($psISE) {
        Split-Path -Path $psISE.CurrentFile.FullPath        
    } else {
        $PSScriptRoot
    }
}
. "$($scriptDirectory)\Copy-RainbowContent.ps1"

# Specify fields to ignore here
$ignoreFieldIds = "{2C37A308-C849-493E-AA8D-5BD4A443AF6D}|{2C37A308-C849-493E-AA8D-5BD4A443AF6D}";

# Specify items to ignore here
$ignoreItemIds = "{1F52376B-0EA9-470B-B5A4-DCAB0D88CEF3}|{025E2453-C4DF-466E-9535-EBA59FA1732E}";

# Specify root id
$rootId = "{110D559F-DEA5-42EA-9C1C-8A5DF7E70EF9}";

# Specify source and destination Sitecore instances
$sourceSession = New-ScriptSession -user "Admin" -pass 'b' -conn "https://sourcesite.local"
$destinationSession = New-ScriptSession -user "Admin" -pass 'b' -conn "https://destinationsite.local"

$copyProps = @{
    "WhatIf"=$true
    "CopyBehavior"="CompareRevision"
    "Recurse"=$true
    "RemoveNotInSource"=$false
    "ClearAllCaches"=$true
    "LogLevel"="Normal"
    "CheckDependencies"=$false
    "BoringMode"=$false
    "FailOnError"=$false
	"ignoreFieldsIds"= $ignoreFieldIds
    "ignoreItemsIds"= $ignoreItemIds
}

$copyProps["SourceSession"] = $sourceSession
$copyProps["DestinationSession"] = $destinationSession

# Default Home
Copy-RainbowContent @copyProps -RootId $rootId *>&1 | Tee-Object "$($scriptDirectory)\Migration-DefaultHome.log"