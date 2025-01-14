Import-Module -Name SPE -Force

function Copy-RainbowContent {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [pscustomobject]$SourceSession,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [pscustomobject]$DestinationSession,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$RootId,

        [Parameter(Mandatory=$false)]
        [string]$ignoreFieldsIds = "",

        [Parameter(Mandatory=$false)]
        [string]$ignoreItemsIds = "",

        [switch]$Recurse,

        [ValidateSet("SkipExisting", "Overwrite", "CompareRevision")]
        [string]$CopyBehavior,

        [switch]$RemoveNotInSource,

        [switch]$ClearAllCaches,

        [switch]$CheckDependencies,

        [ValidateSet("Normal", "Detailed", "Diagnostic")]
        [string]$LogLevel,

        [switch]$BoringMode,

        [switch]$FailOnError
    )

    function Write-Message {
        param(
            [string]$Message,
            [System.ConsoleColor]$ForegroundColor = [System.ConsoleColor]::White,
            [switch]$Hide
        )

        $timeFormat = "HH:mm:ss:fff"
        if(!$Hide) {
            Write-Host "[$(Get-Date -Format $timeFormat)] $($Message)" -ForegroundColor $ForegroundColor
        }
    }

    function Get-SpecialText {
        param(
            [int]$Character,
            [string]$Fallback
        )
        if($BoringMode) {
            $Fallback
        } else {
            "$([char]::ConvertFromUtf32($Character))"
        }
    }

    if($WhatIfPreference) {
        Write-Message "[WhatIf] Following results will be in the WhatIf scenario" -ForegroundColor Yellow
    }
   
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $recurseChildren = $Recurse.IsPresent
    $skipExisting = $CopyBehavior -eq "SkipExisting"
    $compareRevision = $CopyBehavior -eq "CompareRevision"
    $overwrite = $CopyBehavior -eq "Overwrite"
    $bulkCopy = $true
    $isLogLevelDetailed = $LogLevel -eq "Detailed"
    $isLogLevelDiagnostic = $LogLevel -eq "Diagnostic"

    $dependencyScript = {       
        $result = (Test-Path -Path "$($AppPath)\bin\Unicorn.dll") -and (Test-Path -Path "$($AppPath)\bin\Rainbow.dll")
        if($result) {
            $result = $result -band (@(Get-Command -Name "Import-RainbowItem").Count -gt 0)
        }

        $result
    }

    if($CheckDependencies) {
        Write-Message "[Check] Testing connection with remote servers" -ForegroundColor Green
        Write-Message "- Validating source $($SourceSession.Connection[0].BaseUri)"
        if(-not(Test-RemoteConnection -Session $SourceSession -Quiet)) {
            Write-Message " - Unable to connect to $($SourceSession.Connection[0].BaseUri)"
            return
        }
        Write-Message "- Validating destination $($DestinationSession.Connection[0].BaseUri)"
        if(-not(Test-RemoteConnection -Session $DestinationSession -Quiet)) {
            Write-Message " - Unable to connect to $($DestinationSession.Connection[0].BaseUri)"
            return
        }

        Write-Message "[Check] Verifying prerequisites are installed" -ForegroundColor Green
        $isReady = Invoke-RemoteScript -ScriptBlock $dependencyScript -Session $SourceSession

        if($isReady) {
            $isReady = Invoke-RemoteScript -ScriptBlock $dependencyScript -Session $DestinationSession
        }

        if(!$isReady) {
            Write-Message "- Missing required installation of Rainbow and Unicorn"
            return
        } else {
            Write-Message "- All systems are go!"
        }
    }

    if($bulkCopy) {
        $checkIsMediaScript = {
            $rootId = "{ROOT_ID}"
            $db = Get-Database -Name "master"
            $item = $db.GetItem($rootId)
            if($item) {
                $item.Paths.Path.StartsWith("/sitecore/media library/")
            } else {
                $false
            }
        }
        $checkIsMediaScript = [scriptblock]::Create($checkIsMediaScript.ToString().Replace("{ROOT_ID}", $RootId))
        $bulkCopy = !(Invoke-RemoteScript -ScriptBlock $checkIsMediaScript -Session $SourceSession)
    }

    Write-Message "[Running] Transfer from $($SourceSession.Connection[0].BaseUri) to $($DestinationSession.Connection[0].BaseUri)" -ForegroundColor Green
    Write-Message "[Options] RootId = $($RootId); Recurse = $($Recurse);"        
    Write-Message "[Options] CopyBehavior = $($CopyBehavior); RemoveNotInSource = $($RemoveNotInSource);"

    $compareScript = {
        $rootId = "{ROOT_ID}"
        $recurseChildren = [bool]::Parse("{RECURSE_CHILDREN}")
        Import-Function -Name Invoke-SqlCommand
        $connection = [Sitecore.Configuration.Settings]::GetConnectionString("master")

        $revisionFieldId = "{8CDC337E-A112-42FB-BBB4-4143751E123F}"
        if($recurseChildren) {
            $query = "
                WITH [ContentQuery] AS (SELECT [ID], [Name], [ParentID] FROM [dbo].[Items] WHERE ID='$($rootId)' UNION ALL SELECT  i.[ID], i.[Name], i.[ParentID] FROM [dbo].[Items] i INNER JOIN [ContentQuery] ci ON ci.ID = i.[ParentID])
                SELECT cq.[ID], vf.[Value] AS [Revision], cq.[ParentID], vf.[Language] FROM [ContentQuery] cq INNER JOIN dbo.[VersionedFields] vf ON cq.[ID] = vf.[ItemId] WHERE vf.[FieldId] = '$($revisionFieldId)' AND vf.[Language] != '' AND vf.[Version] = (SELECT MAX(vf2.[Version]) FROM dbo.[VersionedFields] vf2 WHERE vf2.[ItemId] = cq.[Id])
            "
        } else {
            $query = "
                WITH [ContentQuery] AS (SELECT [ID], [Name], [ParentID] FROM [dbo].[Items] WHERE ID='$($rootId)')
                SELECT cq.[ID], vf.[Value] AS [Revision], cq.[ParentID], vf.[Language] FROM [ContentQuery] cq INNER JOIN dbo.[VersionedFields] vf ON cq.[ID] = vf.[ItemId] WHERE vf.[FieldId] = '$($revisionFieldId)' AND vf.[Language] != '' AND vf.[Version] = (SELECT MAX(vf2.[Version]) FROM dbo.[VersionedFields] vf2 WHERE vf2.[ItemId] = cq.[Id])
            "
        }
        $records = Invoke-SqlCommand -Connection $connection -Query $query
        if($records) {
            $itemIds = $records | ForEach-Object { "I:{$($_.ID)}+R:{$($_.Revision)}+P:{$($_.ParentID)}+L:$($_.Language)" }
            $itemIds -join "|"
        }
    }
    $compareScript = [scriptblock]::Create($compareScript.ToString().Replace("{ROOT_ID}", $RootId).Replace("{RECURSE_CHILDREN}", $recurseChildren))
    
    class ShallowItem {
        [string]$ItemId
        [string]$ParentId
    }

    class ShallowItemExtended : ShallowItem {
        [string]$RevisionId
        [string]$Language
        [string]$Key
    }

    Write-Message "- Querying item list from source"
    $s1 = [System.Diagnostics.Stopwatch]::StartNew()
    $sourceTree = [System.Collections.Generic.Dictionary[string,[System.Collections.Generic.List[ShallowItem]]]]([StringComparer]::OrdinalIgnoreCase)
    $sourceTree.Add($RootId, [System.Collections.Generic.List[ShallowItem]]@())
    $sourceRecordsString = Invoke-RemoteScript -Session $SourceSession -ScriptBlock $compareScript -Raw
    if([string]::IsNullOrEmpty($sourceRecordsString)) {
        Write-Message "- No items found in source"
        return
    }
    
    $sourceItemsHash = [System.Collections.Generic.HashSet[string]]([StringComparer]::OrdinalIgnoreCase)
    $sourceItemLanguageRevisionHash = [System.Collections.Generic.HashSet[string]]([StringComparer]::OrdinalIgnoreCase)
    $skipItemsHash = [System.Collections.Generic.HashSet[string]]([StringComparer]::OrdinalIgnoreCase)
    $RootParentId = ""
    foreach($sourceRecord in $sourceRecordsString.Split("|".ToCharArray(), [System.StringSplitOptions]::RemoveEmptyEntries)) {
        $shallowItem = [ShallowItem]@{
            "ItemId"=$sourceRecord.Substring(2,38)
            "ParentId"=$sourceRecord.Substring(84,38)
        }
        $shallowItemExtended = [ShallowItemExtended]@{
            "ItemId"=$sourceRecord.Substring(2,38)
            "RevisionId"=$sourceRecord.Substring(43,38)
            "ParentId"=$sourceRecord.Substring(84,38)
            "Language"=$sourceRecord.Substring(125, $sourceRecord.Length - 125)
            "Key"=$sourceRecord
        }
        if(!$sourceTree.ContainsKey($shallowItem.ItemId)) {
            $sourceTree[$shallowItem.ItemId] = [System.Collections.Generic.List[ShallowItem]]@()
        }

        $skipItemsHash.Add($shallowItem.ItemId) > $null
        if([string]::IsNullOrEmpty($RootParentId) -and $shallowItem.ItemId -eq $RootId) {
            $RootParentId = $shallowItem.ParentId
        }
        $sourceItemLanguageRevisionHash.Add($shallowItemExtended.Key) > $null

        $childCollection = $sourceTree[$shallowItem.ParentId]
        if(!$childCollection) {
            $childCollection = [System.Collections.Generic.List[ShallowItem]]@()
        }
        if(!$sourceItemsHash.Contains($shallowItem.ItemId)) {
            $sourceItemsHash.Add($shallowItem.ItemId) > $null       
            
            $childCollection.Add($shallowItem) > $null
        }
        
        $sourceTree[$shallowItem.ParentId] = $childCollection
    }
    $sourceItemsCount = $sourceItemsHash.Count
    $s1.Stop()
    Write-Message " - Found $($sourceItemsCount) item(s) in $($s1.ElapsedMilliseconds / 1000) seconds"

    $destinationItemsHash = [System.Collections.Generic.HashSet[string]]([StringComparer]::OrdinalIgnoreCase)
    $destinationItemLanguageRevisionHash = [System.Collections.Generic.HashSet[string]]([StringComparer]::OrdinalIgnoreCase)
    if(!$overwrite -or $RemoveNotInSource) {
        Write-Message "- Querying item list from destination"
        $d1 = [System.Diagnostics.Stopwatch]::StartNew()
        $destinationRecordsString = Invoke-RemoteScript -Session $DestinationSession -ScriptBlock $compareScript -Raw
        
        if(![string]::IsNullOrEmpty($destinationRecordsString)) {
            foreach($destinationRecord in $destinationRecordsString.Split("|".ToCharArray(), [System.StringSplitOptions]::RemoveEmptyEntries)) {
                $shallowItem = [ShallowItem]@{
                    "ItemId"=$destinationRecord.Substring(2,38)
                    "ParentId"=$destinationRecord.Substring(84,38)
                }
                $shallowItemExtended = [ShallowItemExtended]@{
                    "ItemId"=$destinationRecord.Substring(2,38)
                    "RevisionId"=$destinationRecord.Substring(43,38)
                    "ParentId"=$destinationRecord.Substring(84,38)
                    "Language"=$destinationRecord.Substring(125, $destinationRecord.Length - 125)
                    "Key"=$destinationRecord
                }

                if(!$destinationItemsHash.Contains($shallowItem.ItemId)) {
                    $destinationItemsHash.Add($shallowItem.ItemId) > $null            
                }
                $destinationItemLanguageRevisionHash.Add($shallowItemExtended.Key) > $null
            }
        }

        if($skipExisting) {
            $skipItemsHash.UnionWith($destinationItemsHash)
        } elseif($compareRevision) {
            $itemsToSkipHash = [System.Collections.Generic.HashSet[string]]([StringComparer]::OrdinalIgnoreCase)
            $itemsToSkipHash.UnionWith($sourceItemLanguageRevisionHash)
            $itemsToSkipHash.ExceptWith($destinationItemLanguageRevisionHash)
            if($itemsToSkipHash.Count -gt 0) {
                foreach($key in $itemsToSkipHash) {
                    $itemId = $key.Split("+")[0].Split(":")[1]
                    $skipItemsHash.Remove($itemId) > $null
                }
            }
        }

        $destinationShallowItemsCount = $destinationItemsHash.Count
        $d1.Stop()
        Write-Message " - Found $($destinationShallowItemsCount) item(s) in $($d1.ElapsedMilliseconds / 1000) seconds"
    }

    if($overwrite) {
        $skipItemsHash.Clear()
    }

    $pullCounter = 0
    $pushCounter = 0
    $errorCounter = 0
    $updateCounter = 0
    
    function New-PowerShellRunspace {
        param(
            [System.Management.Automation.Runspaces.RunspacePool]$Pool,
            [scriptblock]$ScriptBlock,
            [PSCustomObject]$Session,
            [object[]]$Arguments
        )
        
        $runspace = [PowerShell]::Create()
        $runspace.AddScript($ScriptBlock) > $null
        $runspace.AddArgument($Session) > $null
        foreach($argument in $Arguments) {
            $runspace.AddArgument($argument) > $null
        }
        $runspace.RunspacePool = $pool

        $runspace
    }


    $sourceScript = {
        param(
            $Session,
            [string]$RootId,
            [string]$ItemIdListString,
            [string]$ignoreFieldsIds,
            [string]$ignoreItemsIds
        )

        $fieldsToExclude = $fieldIds
        $itemsToExclude = $ignoreItemsIds

        $script = {

            Write-Log "Loading assembly..."

            $itemIdList = $itemIdListString.Split("|".ToCharArray(), [System.StringSplitOptions]::RemoveEmptyEntries)

            try {
                [System.Reflection.Assembly]::LoadFrom("$($AppPath)\bin\Rainbow.ContentCopy.dll") > $null                                       
                $extractor = New-Object -TypeName "Unicorn.PowerShell.BulkItemExtractor" -ArgumentList @($fieldsToExclude, $itemsToExclude)
                $yamlItems = $extractor.LoadItems($rootId, $itemIdList)
                $builder = New-Object System.Text.StringBuilder       
                foreach($yamlItem in $yamlItems) {
                    $builder.Append($yamlItem) > $null
                    $builder.Append('<-item->') > $null
                }
                $builder.ToString()
            } catch {
                Write-Host "Bulk import failed with errors $($Error[0].Exception)" -Log Error
                return $Error[0].Exception
            }
        }

        $scriptString = $script.ToString()
       
        $scriptString = "`$rootId = '$($RootId)';`$itemIdListString = '$($ItemIdListString)';`$fieldsToExclude = '$($fieldsToExclude)';`$itemsToExclude = '$($itemsToExclude)';`n" + $scriptString

        Write-Log $scriptString

        $script = [scriptblock]::Create($scriptString)

        Invoke-RemoteScript -ScriptBlock $script -Session $Session -Raw
    }

    $destinationScript = {
        param(
            $Session,
            [string]$Yaml,
            [string]$fieldIds,
            [string]$ignoreItemsIds
        )


        $rainbowYaml = $Yaml
        $fieldsToExclude = $fieldIds
        $itemsToExclude = $ignoreItemsIds

        $script = {
            $rainbowYamlBytes = [System.Convert]::FromBase64String($rainbowYamlBase64)
            $rainbowYaml = [System.Text.Encoding]::UTF8.GetString($rainbowYamlBytes)
            $rainbowItems = $rainbowYaml -split '<-item->' | Where-Object { ![string]::IsNullOrEmpty($_) } | ConvertFrom-RainbowYaml

            $totalItems = $rainbowItems.Count
            $importedItems = 0
            $errorMessages = New-Object System.Text.StringBuilder
            try {
                [System.Reflection.Assembly]::LoadFrom("$($AppPath)\bin\Rainbow.ContentCopy.dll") > $null                
                $installer = New-Object -TypeName "Unicorn.PowerShell.BulkItemInstaller" -ArgumentList @($fieldsToExclude, $itemsToExclude)
                $importedItems = $installer.LoadItems($rainbowItems)
            } catch {
                Write-Host "Bulk import failed with errors $($Error[0].Exception)" -Log Error
            }

            $errorCount = $totalItems - $importedItems

            "{ TotalItems: $($totalItems), ImportedItems: $($importedItems), ErrorCount: $($errorCount), ErrorMessages: '$($errorMessages.ToString())' }"
        }

        $scriptString = $script.ToString()
        $rainbowYamlBytes = [System.Text.Encoding]::UTF8.GetBytes($rainbowYaml)
        
        $scriptString = "`$rainbowYamlBase64 = '$([System.Convert]::ToBase64String($rainbowYamlBytes))';`$fieldsToExclude = '$($fieldsToExclude)';`$itemsToExclude = '$($itemsToExclude)';" + $scriptString

        Write-Log $scriptString

        $script = [scriptblock]::Create($scriptString)

        Invoke-RemoteScript -ScriptBlock $script -Session $Session -Raw
    }

    $pushedLookup = @{}
    $pullPool = [RunspaceFactory]::CreateRunspacePool()
    $pullPool.Open()
    $pullRunspaces = [System.Collections.Generic.List[PSCustomObject]]@()
    $pushPool = [RunspaceFactory]::CreateRunspacePool()
    $pushPool.Open()
    $pushRunspaces = [System.Collections.Generic.List[PSCustomObject]]@()

    class QueueItem {
        [int]$Level
        [string]$Yaml
    }

    $treeLevels = [System.Collections.Generic.List[System.Collections.Generic.List[ShallowItem]]]@()
    $treeLevelQueue = [System.Collections.Generic.Queue[ShallowItem]]@()
    $treeLevelQueue.Enqueue($sourceTree[$RootParentId][0])
    Write-Message "- Tree Level Counts $(Get-SpecialText -Character 0x1F334)" -Hide:(!$isLogLevelDetailed)
    while($treeLevelQueue.Count -gt 0) {
        if($bulkCopy) {
            $currentLevelItems = [System.Collections.Generic.List[ShallowItem]]@()
            while($treeLevelQueue.Count -gt 0 -and ($currentDequeued = $treeLevelQueue.Dequeue())) {
                $currentLevelItems.Add($currentDequeued) > $null
            }
            $treeLevels.Add($currentLevelItems) > $null
            foreach($currentLevelItem in $currentLevelItems) {
                $currentLevelChildren = $sourceTree[$currentLevelItem.ItemId]
                foreach($currentLevelChild in $currentLevelChildren) {
                    $treeLevelQueue.Enqueue($currentLevelChild)
                }
            }
            Write-Message " - Level $($treeLevels.Count - 1) : $($currentLevelItems.Count)" -Hide:(!$isLogLevelDetailed)
        } else {
            $batchSize = 10
            $initialTreeLevelCount = $treeLevels.Count
            $allcurrentLevelItems = [System.Collections.Generic.List[ShallowItem]]@()
            $currentLevelItems = [System.Collections.Generic.List[ShallowItem]]@()
            while($treeLevelQueue.Count -gt 0 -and ($currentDequeued = $treeLevelQueue.Dequeue())) {
                $currentLevelItems.Add($currentDequeued) > $null
                $allCurrentLevelItems.Add($currentDequeued) > $null
                if($currentLevelItems.Count % $batchSize -eq 0) {
                    $treeLevels.Add($currentLevelItems) > $null
                    $currentLevelItems = [System.Collections.Generic.List[ShallowItem]]@()
                } elseif($treeLevelQueue.Count -eq 0) {
                    $treeLevels.Add($currentLevelItems) > $null
                    $currentLevelItems = [System.Collections.Generic.List[ShallowItem]]@()
                }
            }
            foreach($currentLevelItem in $allcurrentLevelItems) {
                $currentLevelChildren = $sourceTree[$currentLevelItem.ItemId]
                foreach($currentLevelChild in $currentLevelChildren) {
                    $treeLevelQueue.Enqueue($currentLevelChild)
                }
            }
            Write-Message " - Levels $($initialTreeLevelCount) to $($treeLevels.Count - 1) : $($allcurrentLevelItems.Count)" -Hide:(!$isLogLevelDetailed)
        }
    }

    Write-Message "Spinning up jobs to transfer content" -ForegroundColor Green
    
    $skippedCounter = 0
    $currentLevel = 0
    $processedCounter = 0
    
    $stepCount = [Math]::Max(1, [int][Math]::Round(($sourceItemsCount / 10.0)))
    if($sourceItemsCount -lt 20) {
        $stepCount = $sourceItemsCount
    }
    $processedBytes = 0

    $keepProcessing = !$WhatIfPreference
    $loopCounter = 0
    while($keepProcessing) {
        if($currentLevel -lt $treeLevels.Count -and $pullRunspaces.Count -lt 4) {
            if($currentLevel -eq 0) {
                Write-Message "[Status] Getting started"
            }
            $itemIdList = [System.Collections.Generic.List[string]]@()
            $levelItems = $treeLevels[$currentLevel]
            $pushedLookup.Add($currentLevel, [System.Collections.Generic.List[QueueItem]]@()) > $null
            foreach($levelItem in $levelItems) {
                Write-Message "Processing $($levelItem.ItemId)"
                $itemId = $levelItem.ItemId              

                if(($skipExisting -or $compareRevision) -and $skipItemsHash.Contains($itemId)) {
                    $processedCounter++
                    Write-Message "[Skip] $($itemId)" -ForegroundColor Cyan -Hide:(!$isLogLevelDiagnostic)
                    $skippedCounter++
                    if($processedCounter % $stepCount -eq 0) {
                        $percentComplete = [int][Math]::Round((($processedCounter * 100 / $sourceItemsCount))/10.0) * 10
                        if($percentComplete -ne 100) { Write-Message "[Status] $($percentComplete)% complete ($($processedCounter))" }
                    }
                } else {
                    Write-Message "[Pull] $($itemId) added to level $($currentLevel) request" -ForegroundColor Green -Hide:(!$isLogLevelDetailed)
                    $itemIdList.Add($itemId) > $null
                }
            }
            if($itemIdList.Count -gt 0) {           
                Write-Message "[Pull] Level $($currentLevel) with $($itemIdList.Count) item(s)" -ForegroundColor Green -Hide:(!$isLogLevelDetailed)
                $runspaceProps = @{
                    ScriptBlock = $sourceScript
                    Pool = $pullPool
                    Session = $SourceSession
                    Arguments = @($RootId,($itemIdList -join "|"), $ignoreFieldIds, $ignoreItemsIds)
                }
                $runspace = New-PowerShellRunspace @runspaceProps
                $pullRunspaces.Add([PSCustomObject]@{
                    Operation = "Pull"
                    Pipe = $runspace
                    Status = $runspace.BeginInvoke()
                    Level = $currentLevel
                    Time = [datetime]::Now
                }) > $null
            } else {
                if($bulkCopy) {
                    Write-Message "[Skip] Level $($currentLevel)" -ForegroundColor Cyan -Hide:(!$isLogLevelDetailed)
                } 
                if($pushedLookup.Contains($currentLevel) -and $pushedLookup[$currentLevel].Count -eq 0) {
                    $pushedLookup.Remove($currentLevel)
                }
            }
            
            $currentLevel++
        }

        $currentRunspaces = $pushRunspaces.ToArray() + $pullRunspaces.ToArray()
        foreach($currentRunspace in $currentRunspaces) {
            if(!$currentRunspace.Status.IsCompleted) { continue }
            
            [System.Management.Automation.PSDataCollection[PSObject]]$response = $currentRunspace.Pipe.EndInvoke($currentRunspace.Status)[0]            

            if($null -eq $response)
            {
                Write-Message "ERROR - response is null"
                Write-Message $currentRunspace

                $currentRunspace.Pipe.Dispose()

                if($currentRunspace.Operation -eq "Pull") {                    
                    $pullRunspaces.Remove($currentRunspace) > $null
                } elseif ($currentRunspace.Operation -eq "Push") {
                    $pushRunspaces.Remove($currentRunspace) > $null
                }                

                continue;
            }

            if($currentRunspace.Operation -eq "Pull") {
                [System.Threading.Interlocked]::Increment([ref] $pullCounter) > $null
            } elseif ($currentRunspace.Operation -eq "Push") {
                [System.Threading.Interlocked]::Increment([ref] $pushCounter) > $null
            }
            Write-Message "[$($currentRunspace.Operation)] Level $($currentRunspace.Level) completed" -ForegroundColor Gray -Hide:(!$isLogLevelDetailed)
            Write-Message "- Processed in $(([datetime]::Now - $currentRunspace.Time))" -ForegroundColor Gray -Hide:(!$isLogLevelDetailed)

            if($currentRunspace.Operation -eq "Pull") {
                $yaml = $response.Item(0)       
                if(![string]::IsNullOrEmpty($yaml) -and [regex]::IsMatch($yaml,"^---")) {               
                    $processedBytes += [System.Text.Encoding]::UTF8.GetByteCount($yaml)
                    if($pushedLookup.Contains(($currentRunspace.Level - 1))) {
                        Write-Message "[Queue] Level $($currentRunspace.Level)" -ForegroundColor Cyan -Hide:(!$isLogLevelDetailed)
                        $pushedLookup[($currentRunspace.Level - 1)].Add([QueueItem]@{"Level"=$currentRunspace.Level;"Yaml"=$yaml;}) > $null
                    } else {             
                        Write-Message "[Push] Level $($currentRunspace.Level)" -ForegroundColor Gray -Hide:(!$isLogLevelDetailed)                        
                        $runspaceProps = @{
                            ScriptBlock = $destinationScript
                            Pool = $pushPool
                            Session = $DestinationSession
                            Arguments = @($yaml, $ignoreFieldIds, $ignoreItemsIds)
                        }
                        $runspace = New-PowerShellRunspace @runspaceProps  
                        $pushRunspaces.Add([PSCustomObject]@{
                            Operation = "Push"
                            Pipe = $runspace
                            Status = $runspace.BeginInvoke()
                            Level = $currentRunspace.Level
                            Time = [datetime]::Now
                        }) > $null
                    }
                }

                $currentRunspace.Pipe.Dispose()
                $pullRunspaces.Remove($currentRunspace) > $null
                $yaml = $null
            }

            if($currentRunspace.Operation -eq "Push") {
                if(![string]::IsNullOrEmpty($response)) {                    
                    $feedback = $response | ConvertFrom-Json
                    Write-Message "- Imported $($feedback.ImportedItems)/$($feedback.TotalItems)" -ForegroundColor Gray -Hide:(!$isLogLevelDetailed)
                    if($feedback.ImportedItems -gt 0) {
                        1..$feedback.ImportedItems | ForEach-Object { [System.Threading.Interlocked]::Increment([ref] $updateCounter) > $null }
                    }
                    if($feedback.TotalItems -gt 0) {
                        1..$feedback.TotalItems | ForEach-Object {
                            $processedCounter++
                            if($processedCounter % $stepCount -eq 0) {
                                $percentComplete = [int][Math]::Round((($processedCounter * 100 / $sourceItemsCount))/10.0) * 10
                                if($percentComplete -ne 100) { Write-Message "[Status] $($percentComplete)% complete ($($processedCounter))" }
                            }
                        }
                    }
                    if($feedback.ErrorCount -gt 0) {
                        [System.Threading.Interlocked]::Increment([ref] $errorCounter) > $null
                        Write-Message "- Errored $($feedback.ErrorCount)" -ForegroundColor Red -Hide:(!$isLogLevelDetailed)
                        Write-Message $feedback
                        Write-Message " - $($feedback.ErrorMessages)" -ForegroundColor Red -Hide:(!$isLogLevelDetailed)
                    }
                }
                
                $queuedItems = [System.Collections.Generic.List[QueueItem]]@()
                if($pushedLookup.ContainsKey($currentRunspace.Level)) {
                    $queuedItems.AddRange($pushedLookup[$currentRunspace.Level])
                    $pushedLookup.Remove($currentRunspace.Level) > $null
                    if($bulkCopy) {
                        Write-Message "[Pull] Level $($currentRunspace.Level) completed" -ForegroundColor Gray -Hide:(!$isLogLevelDetailed)
                    }
                }
                if($queuedItems.Count -gt 0) {
                    foreach($queuedItem in $queuedItems) {
                        $level = $queuedItem.Level
                        Write-Message "[Dequeue] Level $($level)" -ForegroundColor Cyan -Hide:(!$isLogLevelDetailed)
                        Write-Message "[Push] Level $($level)" -ForegroundColor Green -Hide:(!$isLogLevelDetailed)
                        $runspaceProps = @{
                            ScriptBlock = $destinationScript
                            Pool = $pushPool
                            Session = $DestinationSession
                            Arguments = @($queuedItem.Yaml,$ignoreFieldIds,$ignoreItemsIds)
                        }

                        $runspace = New-PowerShellRunspace @runspaceProps  
                        $pushRunspaces.Add([PSCustomObject]@{
                            Operation = "Push"
                            Pipe = $runspace
                            Status = $runspace.BeginInvoke()
                            Level = $level
                            Time = [datetime]::Now
                        }) > $null
                    }
                }
                
                $currentRunspace.Pipe.Dispose()
                $pushRunspaces.Remove($currentRunspace) > $null
            }

            $response = $null
            $currentRunspace = $null
        }

        Start-Sleep -Milliseconds 50
        $loopCounter++
        if($loopCounter % 1000 -eq 0) {
            [GC]::Collect()
        }
        $keepProcessing = ($currentLevel -lt $treeLevels.Count -or $pullRunspaces.Count -gt 0 -or $pushRunspaces.Count -gt 0)
    }

    if($WhatIfPreference) {
        $processedCounter = $sourceItemsCount
        $updateCounter = $sourceItemsCount - $skipItemsHash.Count
        $skippedCounter = $skipItemsHash.Count

        while($currentLevel -lt $treeLevels.Count) {
            Write-Message "[WhatIf] Level $($currentLevel)" -Hide:(!$isLogLevelDetailed)
            $levelItems = $treeLevels[$currentLevel]
            foreach($levelItem in $levelItems) {
                if(!$skipItemsHash.Contains($levelItem.ItemId)) {
                    Write-Message "[WhatIf] $($levelItem.ItemId) would be updated" -ForegroundColor Yellow -Hide:(!$isLogLevelDetailed)
                }
            }
            $currentLevel++
        }
    }

    Write-Message "[Status] 100% complete ($($processedCounter))"

    $pullPool.Close() 
    $pullPool.Dispose()
    $pullPool = $null
    $pushPool.Close() 
    $pushPool.Dispose()
    $pushPool = $null
    $removedCounter = 0
    if($RemoveNotInSource) {
        $removeItemsHash = [System.Collections.Generic.HashSet[string]]([StringComparer]::OrdinalIgnoreCase)
        $removeItemsHash.UnionWith($destinationItemsHash)
        $removeItemsHash.ExceptWith($sourceItemsHash)

        if($removeItemsHash.Count -gt 0) {
            Write-Message "- Removing items from destination not in source"
            $itemsNotInSource = $removeItemsHash -join "|"
            $removeNotInSourceScript = {
                $sd = New-Object Sitecore.SecurityModel.SecurityDisabler
                $ed = New-Object Sitecore.Data.Events.EventDisabler
                $itemsNotInSource = "{ITEM_IDS}"
                $itemsNotInSourceIds = ($itemsNotInSource).Split("|", [System.StringSplitOptions]::RemoveEmptyEntries)
                $db = Get-Database -Name "master"
                foreach($itemId in $itemsNotInSourceIds) {
                    $db.GetItem($itemId) | Remove-Item -Recurse -ErrorAction 0
                }
                $ed.Dispose() > $null
                $sd.Dispose() > $null
            }
            $removeNotInSourceScript = [scriptblock]::Create($removeNotInSourceScript.ToString().Replace("{ITEM_IDS}", $itemsNotInSource))
            if(!$WhatIfPreference) {
                Invoke-RemoteScript -ScriptBlock $removeNotInSourceScript -Session $DestinationSession -Raw
            }
            Write-Message "- Removed $($removeItemsHash.Count) item(s) from the destination"
            $removedCounter += $removeItemsHash.Count
        }
    }

    if($ClearAllCaches) {
        $clearAllCachesScript = {
            [Sitecore.Caching.CacheManager]::ClearAllCaches()
        }

        if(!$WhatIfPreference) {
            Invoke-RemoteScript -ScriptBlock $clearAllCachesScript -Session $DestinationSession -Raw
        }
        Write-Message "- Clearing all caches in the destination"
    }

    $watch.Stop()
    $totalSeconds = $watch.ElapsedMilliseconds / 1000
    Write-Message "[Done] Completed in a record $($totalSeconds) seconds! $(Get-SpecialText -Character 0x1F525)$(Get-SpecialText -Character 0x1F37B)" -ForegroundColor Green
    if($processedCounter -gt 0) {
        if($updateCounter -gt 0) {
            Write-Message " - Transferred $([Math]::Round($processedBytes / 1MB, 2)) MB of item data"
        }
        Write-Message "Processed $($processedCounter)"
        Write-Message " $(Get-SpecialText -Character 0x2714 -Fallback "-") Updated $($updateCounter)"
        Write-Message " $(Get-SpecialText -Character 0x2714 -Fallback "-") Skipped $($skippedCounter)"
        if($errorCounter -gt 0) {
            Write-Message " $(Get-SpecialText -Character 0x274C -Fallback "-") Errored $($errorCounter)"
        } else {
            Write-Message " $(Get-SpecialText -Character 0x2714 -Fallback "-") Errored $($errorCounter)"
        }
        if($RemoveNotInSource) {
            Write-Message " $(Get-SpecialText -Character 0x2714 -Fallback "-") Removed $($removedCounter)"
        }
        Write-Message " $(Get-SpecialText -Character 0x2714 -Fallback "-") Pulled $($pullCounter)"
        if($pullCounter -ne $pushCounter) {
            Write-Message " $(Get-SpecialText -Character 0x274C -Fallback "-") Pushed $($pushCounter)"
        } else {
            Write-Message " $(Get-SpecialText -Character 0x2714 -Fallback "-") Pushed $($pushCounter)"
        }

        if(($errorCounter -gt 0) -and $FailOnError) {
            throw "Content copy failed. See Sitecore logs for details."
        }
    }
}
