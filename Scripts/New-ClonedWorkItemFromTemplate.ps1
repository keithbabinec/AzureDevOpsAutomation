<#
    .SYNOPSIS
        A script that uses the Az CLI to clone an Azure DevOps work item.

    .DESCRIPTION
        A script that uses the Az CLI to clone an Azure DevOps work item. 
        
        This script is significantly slower than the clone operation through the Azure DevOps web UI, but offers two distinct advantages over the web:

        1. This script can easily be called in a loop from another process. If you need to perform many clone operations you can 'foreach' through a 
        list of items to clone.
        
        2. This script offers variable expansion support. This is especially helpful for template work items that you may need to create over and over 
        but with a slightly different title, description, or acceptance critera. This allows you to create slightly modified clones, instead of 'exact' clones.

        This script was developed with Azure CLI v2.24.2 and the Azure-Devops v0.18.0 extension. That version of the AzDo extension has a couple bugs 
        in it which have workarounds in the script. Its possible these workarounds may cause breaks in newer SDK versions.
        
        Tutorial instructions: https://keithbabinec.com/2021/06/05/how-to-use-the-azure-cli-to-clone-azure-devops-work-items-with-variable-expansion-support/

        Prerequisites:
        1. The Az CLI is installed.
        2. The Azure-Devops CLI extension is installed.
        3. You have created an Azure DevOps Personal Access Token (PAT) that has full control of work items.
        4. You have stored the PAT in the environment variable (AZURE_DEVOPS_EXT_PAT) so it can accessed by the commands. 

    .PARAMETER TemplateItemId
        Specify the work item ID that should be cloned. Ex: 123456

    .PARAMETER CloneChildWorkItems
        Specify if all the child work items should be cloned with the parent (recursive clones the entire tree).

    .PARAMETER VariableExpansionMap
        An optional parameter to specify variable expansions. How it works:

        1. In the work item to be cloned, add placeholder values using double-curly braces. Example: {{MyVariable}}. Variables should word characters only, 
        with no spaces, special characters, or numbers.

        2. Construct the map and pass it to the script as a parameter. See the example section for usage.

    .EXAMPLE
        # Constructs the variable expansion map, then clones an item.

        $map = new-object -typename 'System.Collections.Generic.Dictionary[System.String,System.String]'
        $map["MyVariable"] = "expanded value!"
        $map["MyOtherVariable"] = "expanded value!"
        
        .\New-ClonedWorkItemFromTemplate.ps1 -TemplateItemId 123456 -CloneChildWorkItems $true -Verbose -VariableSubstitutionMap $map
#>
[CmdletBinding()]
Param
(
    [Parameter(
        Mandatory=$true,
        HelpMessage='Specify the work item ID that should be cloned.')]
    [System.Int32]
    [ValidateScript({$_ -gt 0})]
    $TemplateItemId,

    [Parameter(
        Mandatory=$true,
        HelpMessage='Specify if all the child work items should be cloned with the parent (recursive clones the entire tree).')]
    [System.Boolean]
    $CloneChildWorkItems,

    [Parameter(Mandatory=$false)]
    [System.Collections.Generic.Dictionary[System.String,System.String]]
    $VariableExpansionMap
)

function Expand-InlineVariables([System.String]$Text, [System.Collections.Generic.Dictionary[System.String,System.String]]$Map)
{
    $matchPattern = "{{\w+}}"
    $regex = New-Object -TypeName System.Text.RegularExpressions.Regex -ArgumentList $matchPattern
    
    $matchCollection = $regex.Matches($Text)
    
    if ($matchCollection.Count -gt 0)
    {
        $foundVars = $matchCollection.Captures.Value

        foreach ($foundVar in $foundVars)
        {
            $lookupKey = $foundVar.Substring(2, $foundVar.Length-4)

            if ($Map.ContainsKey($lookupKey))
            {
                $Text = $Text.Replace($foundVar, $Map[$lookupKey])
            }
            else
            {
                Write-Warning -Message "Variable $foundVar was found in the field, but was not provided in the substitution map."
            }
        }
    }

    Write-Output -InputObject $Text
}

function Add-QuoteEscape($Text)
{
    $Text = $Text.Replace("`"", "\`"")
    $Text = $Text.Replace("&quot;", "\`"")
    Write-Output -InputObject $Text
}

# create a table of the extra fields that should be cloned (if present in the work item).

$extraFieldsToClone = @(
    "Microsoft.VSTS.Common.AcceptanceCriteria"
    "Microsoft.VSTS.Common.ValueArea",
    "Microsoft.VSTS.Scheduling.CompletedWork",
    "Microsoft.VSTS.Scheduling.OriginalEstimate",
    "Microsoft.VSTS.Scheduling.RemainingWork",
    "Microsoft.VSTS.Scheduling.StoryPoints"

    # do not include these fields, they are handled automatically:
    # - System.WorkItemType
    # - System.Title
    # - System.Description
    # - System.AreaPath
    # - System.IterationPath
    # - System.TeamProject
    # - Microsoft.VSTS.Common.Priority
)

# start a queue for the work items that need to be cloned.
# add the first item (our initial work item).

$itemsToClone = New-Object -TypeName System.Collections.Generic.Queue[System.String]
$itemsToClone.Enqueue($TemplateItemId)

$visitedItemIds = New-Object -TypeName System.Collections.Generic.List[System.String]

# begin the search and clone operations

while ($itemsToClone.Count -gt 0)
{
    # while the search queue is not empty, grab the next item to work.

    $currentWorkItem = $itemsToClone.Dequeue()

    if ($currentWorkItem.Contains("/"))
    {
        # we are cloning an item that will need a link/ref added back to its cloned parent.
        $currentWorkItemParentId = $currentWorkItem.Split('/')[0]
        $currentWorkItemId = $currentWorkItem.Split('/')[1]
    }
    else
    {
        # we are cloning the top level item (no parents).
        $currentWorkItemId = $currentWorkItem
        $currentWorkItemParentId = $null
    }

    $visitedItemIds.Add($currentWorkItemId)

    Write-Verbose -Message "Fetching work item ID $currentWorkItemId"

    $originalWorkItemJson = az boards work-item show --id $currentWorkItemId --detect false
    if ($LASTEXITCODE -ne 0)
    {
        # last command returned an error to the console, can't proceed.
        return
    }

    # prepare common arguments

    $originalWorkItemObj = $originalWorkItemJson | ConvertFrom-Json
    $workItemType = $originalWorkItemObj.fields."System.WorkItemType"
    $workItemArea = $originalWorkItemObj.fields."System.AreaPath"
    $workItemIteration = $originalWorkItemObj.fields."System.IterationPath"
    $workItemProject = $originalWorkItemObj.fields."System.TeamProject"
    $workItemPriority = $originalWorkItemObj.fields."Microsoft.VSTS.Common.Priority"

    $title = Expand-InlineVariables -Text $originalWorkItemObj.fields."System.Title" -Map $VariableExpansionMap
    $description = Expand-InlineVariables -Text $originalWorkItemObj.fields."System.Description" -Map $VariableExpansionMap

    $newItemArgs = @(
        "--type", $workItemType,
        "--title", $title,
        "--area", $workItemArea,
        "--iteration", $workItemIteration,
        "--project", $workItemProject,
        "--detect", "false"
    )

    if ([System.String]::IsNullOrWhiteSpace($description) -eq $false)
    {
        # there is a bug in the SDK that doesn't handle quote characters correctly for certain fields.
        # explicit escaping allows it to pass through.
        $wrappedDescription = Add-QuoteEscape -Text $description
        
        $newItemArgs += "--description"
        $newItemArgs += $wrappedDescription
    }

    if ([System.String]::IsNullOrWhiteSpace($workItemPriority) -eq $false)
    {
        $newItemArgs += "--fields"
        $newItemArgs += "Microsoft.VSTS.Common.Priority=$workItemPriority"
    }

    # submit request to create new item

    Write-Verbose -Message "Creating new cloned work item."

    $newItemJson = & az boards work-item create $newItemArgs
         
    if ($LASTEXITCODE -ne 0)
    {
        # last command returned an error to the console, can't proceed.
        return
    }

    $newWorkItemId = ($newItemJson | ConvertFrom-Json).id
    Write-Verbose -Message "Cloned work item id: $newWorkItemId"

    # prepare fields arguments

    # there is a bug in the SDK where they don't allow you to specify more than one Field
    # at a time. loop through for now and call update for each field. change later to a single
    # call when the fix is available, because this is terrible performance.

    foreach ($field in $extraFieldsToClone)
    {
        if ([System.String]::IsNullOrWhiteSpace($originalWorkItemObj.fields.$field) -eq $false)
        {
            Write-Verbose -Message "Adding extra field: $field"

            $varExpandedValue = Expand-InlineVariables -Text $originalWorkItemObj.fields.$field -Map $VariableExpansionMap
            $escapedValue = Add-QuoteEscape -Text $varExpandedValue

            $null = az boards work-item update `
                --id $newWorkItemId `
                --fields ("{0}={1}" -f $field, $escapedValue) `
                --detect false
        }
    }

    # are there links to add back to a parent item?
    # then add the relation object.

    if ($currentWorkItemParentId -ne $null)
    {
        Write-Verbose -Message "Adding relation to cloned parent: $currentWorkItemParentId"

        $null = az boards work-item relation add `
            --id $newWorkItemId `
            --relation-type "Parent" `
            --target-id $currentWorkItemParentId `
            --detect false
    }

    # are there child items to clone?
    # then add them to the clone queue.

    $originalWorkItemChildItems = $originalWorkItemObj.relations | Where-Object { $_.attributes.name -eq "Child" }
    foreach ($originalWorkItemChildItem in $originalWorkItemChildItems)
    {
        $originalWorkItemChildItemId = $originalWorkItemChildItem.url.split('/')[-1]

        # queue this work, but append it along with the new work item ID so we know how to parent it correctly.
        # also safety check to make sure we don't revisit/clone the same item ids if parent links
        # are somehow circular on the original items -- I don't think thats even possible, but safety first.

        if ($visitedItemIds.Contains($originalWorkItemChildItemId) -eq $false)
        {
            Write-Verbose -Message "Adding child work item to be cloned: $originalWorkItemChildItemId"
            $itemsToClone.Enqueue("$newWorkItemId/$originalWorkItemChildItemId")
        }
    }
}

Write-Verbose -Message "All clone operations completed."
