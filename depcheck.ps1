# Check for deprecated actions in Power Automate Cloud Flows or unpacked Canvas Apps
# Note: The solution file structure is proprietary and subject Cloud flows will exist

param (
    [Parameter(Mandatory = $true)]
    [string]$solnfolder,
    [switch]$skipCloudFlows = $false,
    [switch]$skipPowerFx = $false
)

$deprecatedActions = @{}
$usedActions = [System.Collections.Generic.List[ActionUsage]]::new()

class ActionUsage {
    # Prevent invalid values
    [ValidateNotNullOrEmpty()][string]$Connector
    [ValidateNotNullOrEmpty()][string]$Action
    [ValidateNotNullOrEmpty()][string]$Filename
    [ValidateNotNullOrEmpty()][string]$Type
    [string]$ActionUniqueName
    [string]$Description
    [string]$ConnectionRef
    [int]$LineNum
    [bool]$Deprecated

    # Have a constructor to force properties to be set
    ActionUsage($Connector, $Action, $Filename, $Type) {
        $this.Connector = $Connector
        $this.Action = $Action
        $this.Filename = $Filename
        $this.Type = $Type
        $this.ActionUniqueName = BuildKey -connector $Connector -action $Action
        $this.Deprecated = IsActionDeprecated($this.ActionUniqueName)
    }
}

if ($skipCloudFlows -and $skipPowerFx) {
    Write-Error "Nothing to do. Only one of skipCloudFlows or skipPowerFx can be set."
    exit
}

Function IsActionDeprecated([string]$key) {
    
    if ($deprecatedActions.ContainsKey($key)) {
        return $true
    }
    return $false
}

Function Get-DeepProperty([object] $InputObject, [string] $Property) {
    $path = $Property -split '\.'
    $obj = $InputObject
    $path | ForEach-Object { $obj = $obj.$_ }
    $obj
}

Function BuildKey {
    param ([string]$connector, [string]$action)

    $key = "${connector}|${action}"
    return $key
}

Function RemoveLeadingString {
    param ([string]$inputStr, [string]$leading)

    if ($inputStr.StartsWith($leading)) {
        return $inputStr.Replace($leading, "")
    }
    return $inputStr
}

Function LogDeprecatedFlowActions {
    param ([int]$depth, [object]$actions)

    if ($actions.getType().Name -eq "PSCustomObject") {

        # Stackoverflow (ahem) helped with navigating dynamics object
        # https://stackoverflow.com/questions/27195254/dynamically-get-pscustomobject-property-and-values/27195828
        $actionObjects = Get-Member -InputObject $actions -MemberType NoteProperty

        foreach ($actionObject in $actionObjects) {

            $actionBody = $actions | Select-Object -ExpandProperty $actionObject.Name

            $actionName = $actionObject.Name
            $type = $actionBody.type
            $description = $actionBody.description

            # Check if this is using an OpenApiConnection
            if ($type -eq "OpenApiConnection") {
                $connectionRef = $actionBody.inputs.host.connectionName
                $operationId = $actionBody.inputs.host.operationId

                $propertyPath = "properties.connectionReferences." + $connectionRef + ".api.name"
                $connector = Get-DeepProperty -InputObject $flow -Property $propertyPath

                $connector = RemoveLeadingString -inputStr $connector --leading "shared_"

                $connRefLogicalNamePath = "properties.connectionReferences." + $connectionRef + ".connection.connectionReferenceLogicalName"
                $connRefLogicalName = Get-DeepProperty -InputObject $flow -Property $connRefLogicalNamePath

                $type = "${connector}:${operationId} via ${connRefLogicalName}"

                $key = "${connector}|${operationId}"
                if ($deprecatedActions.ContainsKey($key)) { 
                    Write-Host "Deprecated action in flow ${filename}: ${operationId} on connector ${connector} (connectionRef: ${connRefLogicalName}) in step ${actionName}: ${description}"
                }

                # Store action usage
                $usage = [ActionUsage]::new($connector, $action, $filename, "CloudFlow")
                $usage.Description = "${actionName}: ${description}"
                $usage.ConnectionRef = $connRefLogicalName
                $usedActions.add($usage)
            }
        }

        # $message = " " * $depth + "<${type}>: ${description}"
        # Write-Host $message

        # Does this action have an "actions" node below it?
        $actionNodes = Get-Member -InputObject $actionBody -MemberType NoteProperty
        $childActions = $actionNodes | Where-Object -Property Name -eq -Value "actions" 
        if ($null -ne $childActions) {
            # We have child actions - recurse into actions list and drill through these actions
            [int]$newDepth = $depth + 2
            $childActionObject = $actionBody | Select-Object -ExpandProperty $childActions.Name
            # Recurse
            LogDeprecatedFlowActions -depth $newDepth -actions $childActionObject
        }
    }
}

Function ReadDeprecatedActions {

    ## TODO: Give option to pull from the cloud or locally

    ## FIXME: Hardcoded deprecated
    $deprecatedfilename = Join-Path ${PSScriptRoot} "Deprecated.json"

    Get-Content $deprecatedfilename | ConvertFrom-Json | Foreach-Object {
        $connector = $_;
        $uniqueName = $_.UniqueName;

        $_.Actions | ForEach-Object {
            $action = $_
            $operationId = $_.OperationId

            $key = "${uniqueName}|${operationId}"
            $deprecatedActions[$key] = [PSCustomObject]@{
                Connector = $connector
                Action    = $action
            }
        }
    }
    Write-Host "Read" $deprecatedActions.Count "deprecated actions"
}



Function ScanFlows {
    $workflowfolder = Join-Path $solnfolder "Workflows"

    if (!( Test-Path $workflowfolder -PathType Container)) {
        Write-Host "No flows exist in solution"
        return
    }

    Write-Host "Scanning flows in $workflowfolder"
    
    # Perhaps we could check the metadata xml relating to the flow to check if it's a cloud flow.
    # I think cloud flows are <Category>5</Category> (but need to check)
    # We don't need to do that right now, because the all json files are cloud flows.
    Get-ChildItem $workflowfolder -Filter *.json |
    Foreach-Object {
    
        Write-Progress "Checking flow "$_.FullName
    
        $flow = Get-Content $_.FullName | ConvertFrom-Json
    
        $actions = $flow.properties.definition.actions
    
        LogDeprecatedFlowActions -depth 2 -actions $actions
    }
}

Function ScanAllCanvasApps {
    # NOTE: We expect the canvas apps in the solution to have already been exploded with 'pac canvas unpac'
    # An unpacked canvas app can be identified in a solution by the existnace of the 'CanvasManifest.json' file
    Get-ChildItem -Path $solnfolder -Filter CanvasManifest.json -Recurse  | ForEach-Object {
        ScanUnpackedCanvasApp -folder $_.Directory.FullName
    }
}

Function ScanUnpackedCanvasApp {
    Param ([string]$folder)

    Write-Host "Examining Power Fx in $folder"

    # Read "Connections.json" file
    $connectionsFilename = Join-Path $folder "Connections"
    $connectionsFilename = Join-Path $connectionsFilename "Connections.json"
    $connections = Get-Content $connectionsFilename | ConvertFrom-Json
    $connectionObjects = Get-Member -InputObject $connections -MemberType NoteProperty

    # Find all connectors used in Connections.json

    $connectors = @{}

    foreach ($connectionObject in $connectionObjects) {

        $connection = $connections | Select-Object -ExpandProperty $connectionObject.Name

        $connector = $connection.connectionRef.id
        $displayName = $connection.connectionRef.displayName
        $dataSources = $connection.dataSources[0]

        $connector = RemoveLeadingString -inputStr $connector -leading "/providers/microsoft.powerapps/apis/shared_"

        $connectors.Add($dataSources, 1)

        Write-Host "Uses connector $connector ($displayName) $dataSources"
    }

    # Build a regex to search the *.fx.yaml Power Fx files
    $connectors = $connectors.keys | Join-String -Separator "|"
    $regex = "((^|[^\w])(?<connector>(${connectors}))\.(?<action>\w+))+"
    Write-Debug "Searching using regex: $regex"
    [regex]$rx = $regex

    # Search the Power Fx files, and see if they contain any usages of each dataSource
    Get-ChildItem -Path $folder -Filter *.fx.yaml -Recurse -File | ForEach-Object {

        $filepath = $_.FullName
        $linenum = 1
        $c = Get-Content -Path $filepath

        $c | ForEach-Object {
            $results = $rx.Match( $_ )
            
            if ($results.Success) {
                foreach ($actionmatch in $results) {
                    $connector = $actionmatch.Groups["connector"].Value
                    $operationId = $actionmatch.Groups["action"].Value
                    
                    Write-Host "${connector}:${operationId} in ${filepath}:${linenum}" 
                    
                    $usage = [ActionUsage]::new($connector, $operationId, $filepath, "Power Fx")
                    $usage.LineNum = $linenum
                    $usedActions.add($usage)
                }
            }

            $linenum = $linenum + 1
        }
    }
}

## MAIN BODY

ReadDeprecatedActions

if ($skipCloudFlows -ne $true) {
    ScanFlows
}

if ($skipPowerFx -ne $true) {
    ScanAllCanvasApps
}

$numUsedActions = $usedActions.Count

Write-Output "`nSummary of $numUsedActions usages:`n" 

$usedActions | Group-Object -Property ActionUniqueName -NoElement | Sort-Object -Property Count -Descending | ForEach-Object {
    $deprecatedWarning = "";
    if ($deprecatedActions.ContainsKey($_.Name)) {
        $deprecatedWarning = "*** DEPRECATED *** "
    }

    $message = "{0,3} usages of {2}{1}" -f $_.Count, $_.Name, $deprecatedWarning
    Write-Host $message
}

$deprecatedUsages = $usedActions | Where-Object -Property Deprecated -EQ -Value $true

$numDeprecatedUsages = $deprecatedUsages.Count

Write-Output "`nUsages of deprecated actions: $numDeprecatedUsages`n"

$deprecatedUsages | ForEach-Object {
    $message = "{0}|{1} in {2} at {3}:{4}" -f $_.Connector, $_.Action, $_.Type, $_.Filename, $_.LineNum
    Write-Host $message
}
