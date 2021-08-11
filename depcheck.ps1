# Check for deprecated actions in Power Automate Cloud Flows or Canvas Apps

param (
    #[Parameter(Mandatory = $true)]
    [string]$solnfolder = 'C:\Dev\AxisK2\SolutionPackage\AxisK2CloudFlows', # FiXME
    [switch]$skipCloudFlows = $false,
    [switch]$skipPowerFx = $false
)

if ($skipCloudFlows -and $skipPowerFx) {
    Write-Error "Nothing to do. Only one of skipCloudFlows or skipPowerFx can be set."
    exit
}

$hash = @{}
$usedAction = @{}

Function Get-DeepProperty([object] $InputObject, [string] $Property) {
    $path = $Property -split '\.'
    $obj = $InputObject
    $path | ForEach-Object { $obj = $obj.$_ }
    $obj
}

Function StoreUsedAction {
    param ([string]$connector, [string]$action)

    $key = "${connector}|${action}"
    
    # FIXME: This is pretty lame, but so is my powershell skills
    if ( $usedAction.ContainsKey(${key}) ) {
        $count = $usedAction[${key}]
        $newCount = $count + 1
        $usedAction[$key] = $newCount
    }
    else {
        $usedAction.add(${key}, 1)
    }
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
                if ($hash.ContainsKey($key)) { 
                    Write-Host "Deprecated action in flow ${filename}: ${operationId} on connector ${connector} (connectionRef: ${connRefLogicalName}) in step ${actionName}: ${description}"
                }

                StoreUsedAction $connector $operationId
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

    $deprecatedActions = Get-Content $deprecatedfilename | ConvertFrom-Json
    $deprecatedActions | Foreach-Object {
        $connector = $_;
        $uniqueName = $_.UniqueName;

        $_.Actions | ForEach-Object {
            $action = $_
            $operationId = $_.OperationId

            $key = "${uniqueName}|${operationId}"
            $hash[$key] = [PSCustomObject]@{
                Connector = $connector
                Action    = $action
            }
        }
    }
    Write-Host "Read" $hash.Count "deprecated actions"
}



Function ScanFlows {
    $workflowfolder = Join-Path $solnfolder "Workflows"
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

Function ScanMsApps {
    # FIXME: Unpac canvas app
    
    # FIXME: Loop for app canvas apps
    # FIXME: Prevent hardcdoed
    ScanUnpackedMsApp -folder "C:\Dev\AxisK2\SolutionPackage\AxisK2MobileApp\CanvasAppsSrc\ebecs_axismobileappwithiamap_3e0f5_DocumentUri_src"
}

Function ScanUnpackedMsApp {
    Param ([string]$folder)

    Write-Debug "Examining PowerFx in $folder"

    $connectionsFilename = Join-Path $folder "Connections"
    $connectionsFilename = Join-Path $connectionsFilename "Connections.json"

    $connections = Get-Content $connectionsFilename | ConvertFrom-Json

    $connectionObjects = Get-Member -InputObject $connections -MemberType NoteProperty

    $connectors = @{}

    foreach ($connectionObject in $connectionObjects) {

        $connection = $connections | Select-Object -ExpandProperty $connectionObject.Name

        $id = $connection.id
        $connector = $connection.connectionRef.id
        $displayName = $connection.connectionRef.displayName
        $dataSources = $connection.dataSources[0]

        $connector = RemoveLeadingString -inputStr $connector -leading "/providers/microsoft.powerapps/apis/shared_"

        $connectors.Add($dataSources, 1)

        Write-Host "Using connector $connector ($displayName) $dataSources"
    }

    # Fixme need to extend this

    $connectorRegex = $connectors.keys | Join-String -Separator "|"

    # Search the PowerFx files, and see if they contain any usages of each dataSource
    ChildItem -Path $folder -Filter *.fx.yaml -Recurse -File | ForEach-Object {

        $c = Get-Content -Path $_.FullName 

        $regex = "(([^\w])($connectorRegex)\.(\w+))+"
        Write-Host "regex=$regex"
        
        $d = $c | Where-Object { $_ -match $regex }

        $d | ForEach-Object {
            $match = $_ -match "(^|[^\w])($connectorRegex)\.(?<action>\w+)\(.*$"
            Write-Host "match"
        }



        #[System.IO.Path]::GetFileNameWithoutExtension($_)
    }


}

## MAIN BODY

ReadDeprecatedActions

if ($skipCloudFlows -ne $true) {
    ScanFlows
}

if ($skipPowerFx -ne $true) {
    ScanMsApps
}

Write-Output "`nSummary:`n"

$usedAction.GetEnumerator() | ForEach-Object {
    $isDeprecated = "";
    if ($hash.ContainsKey($_.key)) {
        $isDeprecated = "*** DEPRECATED *** "
    }
    $message = "{1,3} usages of {2}{0}" -f $_.key, $_.value, $isDeprecated
    Write-Output $message
}