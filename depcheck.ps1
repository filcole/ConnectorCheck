# Check for deprecated actions in Power Automate Cloud Flows or Canvas Apps

param (
    #[Parameter(Mandatory = $true)]
    [string]$solnfolder = 'C:\Dev\AxisK2\SolutionPackage\AxisK2CloudFlows'    # FiXME

)

$hash = @{}
$usedAction = @{}

Function Get-DeepProperty([object] $InputObject, [string] $Property) {
    $path = $Property -split '\.'
    $obj = $InputObject
    $path | % { $obj = $obj.$_ }
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

Function log-DeprecatedActions {
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

                if ($connector.StartsWith("shared_")) {
                    $connector = $connector.Replace("shared_", "")
                }

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
        if ($childActions -ne $null) {
            # We have child actions - recurse into actions list and drill through these actions
            [int]$newDepth = $depth + 2
            $childActionObject = $actionBody | Select-Object -ExpandProperty $childActions.Name
            # Recurse
            log-DeprecatedActions -depth $newDepth -actions $childActionObject
        }
    }
}

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

$workflowfolder = Join-Path $solnfolder "Workflows"
Write-Host "Scanning flows in $workflowfolder"

# Perhaps we could check the metadata xml relating to the flow to check if it's a cloud flow.
# I think cloud flows are <Category>5</Category> (but need to check)
# We don't need to do that right now, because the all json files are cloud flows.
Get-ChildItem $workflowfolder -Filter *.json |
Foreach-Object {
    $filename = $_.FullName

    #Write-Progress "Checking flow "$filename

    $flow = Get-Content $_.FullName | ConvertFrom-Json

    $actions = $flow.properties.definition.actions

    log-DeprecatedActions -depth 2 -actions $actions
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