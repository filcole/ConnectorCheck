<#
  .SYNOPSIS
  Checks connectors and actions used in power platform solutions.
  .DESCRIPTION
  Examines an unpacked solution and reports on connectors and actions used. Deprecated actions are downloaded from https://connectorstatus.com. Deprecated actions used in the solution are output along with their location. Power Automate Cloud Flows and Power Fx (e.g. Canvas Apps) are checked.

  TODO
  Improve to report instances of deprecated triggers

  NOTES/WARNINGS
  1. The power apps solution structure is proprietary and subject to change.
  2. The solution must be exported and unpacked, additionally any .msapp files within the solution must also be unpacked.

  .PARAMETER solnFolder
  Specifies the path to the exploded solution folder.

  .PARAMETER skipCloudFlows
  Indicates if checking of cloud flows should be not be performed.

  .PARAMETER skipPowerFx
  Indicates if checking of Canvas Apps/PowerFx should be not be performed.

  .LINK
  https://philcole.org/post/connector-check

  .EXAMPLE
  PS> ConnectorCheck.ps1 -solnfolder ./SolutionPackage

#>

param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the path to the unpacked solution")]
    [Alias("p", "path")]
    [string]$solnfolder,
    [switch]$skipCloudFlows = $false,
    [switch]$skipPowerFx = $false
)

# We use a list of deprecated actions already generated in the cloud. See https://connectorstatus.com for details.
# If you want to produce your own list, then see https://philcole.org/post/deprecated-actions/ for one method.
Set-Variable -Name DeprecatedActionsUrl -Option Constant -Value "https://connectorstatus.com/Deprecated.json"

# Hash of deprecated actions read in a pre-generated JSON file containing deprecated actions
$deprecatedActions = @{}

# List of actions found in flows and Power Fx (canvas apps) scanned
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

Function ReadDeprecatedActions {


    # Download the latest deprecated actions.
    Write-Host "Downloading deprecated actions from $DeprecatedActionsUrl"

    try {
        $r = Invoke-WebRequest -Uri $DeprecatedActionsUrl
    }
    catch {
        $_.Exception.Message
        exit
    }

    #$deprecatedfilename = Join-Path ${PSScriptRoot} "Deprecated.json"

    $r.Content | ConvertFrom-Json | Foreach-Object {
        $connector = $_;
        $uniqueName = $_.UniqueName;

        $_.Actions | ForEach-Object {
            $action = $_
            $operationId = $_.OperationId

            $key = BuildKey -connector ${uniqueName} -action ${operationId}
            $deprecatedActions[$key] = [PSCustomObject]@{
                Connector = $connector
                Action    = $action
            }
        }
    }
    Write-Host "Read" $deprecatedActions.Count "deprecated actions"
}

Function BuildKey {
    param ([string]$connector, [string]$action)

    return "${connector}|${action}"
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

Function RemoveLeadingString {
    param ([string]$inputStr, [string]$leading)

    if ($inputStr.StartsWith($leading)) {
        return $inputStr.Replace($leading, "")
    }
    return $inputStr
}

Function CheckIsSolutionFolder {
    $solnxml = Join-Path $solnfolder "Other"
    $solnxml = Join-Path $solnxml "Solution.xml"

    if (!( Test-Path $solnxml -PathType Leaf)) {
        Write-Error "Not valid solution folder. $solnxml does not exist"
        exit
    }
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
        
        $flowFilename = $_.FullName

        Write-Host "Scanning flow"$_.Name
    
        $flow = Get-Content $flowFilename | ConvertFrom-Json
    
        $actions = $flow.properties.definition.actions
    
        ScanFlowActions -depth 2 -actions $actions
    }
}

Function ScanFlowActions {
    param ([int]$depth, [object]$actions)

    if ($actions.getType().Name -eq "PSCustomObject") {

        # Stackoverflow (ahem) helped with navigating dynamics object
        # https://stackoverflow.com/questions/27195254/dynamically-get-pscustomobject-property-and-values/27195828
        $actionObjects = Get-Member -InputObject $actions -MemberType NoteProperty

        foreach ($actionObject in $actionObjects) {

            $actionBody = $actions | Select-Object -ExpandProperty $actionObject.Name

            $actionName = $actionObject.Name
            $type = $actionBody.type
            $description = ""
            if ($actionBody.PSobject.Properties.Name -contains "description") {
                $description = $actionBody.description
            }

            # Check if this is using an OpenApiConnection
            if ($type -eq "OpenApiConnection") {
                $connectionRef = $actionBody.inputs.host.connectionName
                $operationId = $actionBody.inputs.host.operationId

                $propertyPath = "properties.connectionReferences." + $connectionRef + ".api.name"
                $connector = Get-DeepProperty -InputObject $flow -Property $propertyPath

                $connector = RemoveLeadingString -inputStr $connector -leading "shared_"

                $connRefLogicalNamePath = "properties.connectionReferences." + $connectionRef + ".connection.connectionReferenceLogicalName"
                $connRefLogicalName = Get-DeepProperty -InputObject $flow -Property $connRefLogicalNamePath

                $type = "${connector}:${operationId} via ${connRefLogicalName}"

                $key = BuildKey -connector $connector -action $operationId
                if ($deprecatedActions.ContainsKey($key)) { 
                    Write-Host "Deprecated action in flow ${flowFilename}: ${operationId} on connector ${connector} (connectionRef: ${connRefLogicalName}) in step ${actionName}: ${description}"
                }

                # Store action usage
                $usage = [ActionUsage]::new($connector, $operationId, $flowFilename, "CloudFlow")
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
            ScanFlowActions -depth $newDepth -actions $childActionObject
        }
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

    Write-Host "Scanning CanvasApp in $folder"

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
        $dataSources = $connection.dataSources[0]

        $connector = RemoveLeadingString -inputStr $connector -leading "/providers/microsoft.powerapps/apis/shared_"

        $connectors.Add($dataSources, 1)

        $displayName = $connection.connectionRef.displayName
        Write-Verbose "Uses connector $connector ($displayName) $dataSources"
    }

    if ($connectors.Keys.Count -eq 0) {
        Write-Verbose "No connectors found in canvas app"
        return
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
                    
                    #Write-Host "${connector}:${operationId} in ${filepath}:${linenum}" 
                    
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

if ($skipCloudFlows -and $skipPowerFx) {
    Write-Error "Nothing to do. Only one of skipCloudFlows or skipPowerFx can be set."
    exit
}

CheckIsSolutionFolder
ReadDeprecatedActions

if ($skipCloudFlows -ne $true) {
    ScanFlows
}

if ($skipPowerFx -ne $true) {
    ScanAllCanvasApps
}

$numUsedActions = $($usedActions | Measure-Object).Count

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

$numDeprecatedUsages = $($deprecatedUsages | Measure-Object).Count

Write-Output "`nUsages of deprecated actions: $numDeprecatedUsages`n"

$deprecatedUsages | ForEach-Object {
    $message = "{0}|{1} in {2} at {3}:{4}" -f $_.Connector, $_.Action, $_.Type, $_.Filename, $_.LineNum
    Write-Host $message
}
