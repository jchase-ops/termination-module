<#
GetServiceDeskInfo method for Termination Module
#>

Function Get-ServiceDeskInfo {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ChangeOrder,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $CurrentCount
    )

    $Log.Account     = $env:USERNAME
    $Log.AccountType = "Developer"
    $Log.ChangeOrder = $ChangeOrder
    $Log.DisplayName = $env:USERNAME -Replace "\.", " "
    $Log.Domain      = "All"
    $Log.EmployeeID  = "N/A"
    $Log.Status      = "Start"
    $Log.Step        = $MyInvocation.MyCommand
    $Log.Timestamp   = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount

    $ServiceDeskProxy = New-WebServiceProxy -Uri $ServiceDeskUri
    $SessionID = $ServiceDeskProxy.login($Credential.UserName, $Credential.GetNetworkCredential().Password)
    $ServiceDeskInfo = New-Object System.Collections.Generic.List[System.Object]

    $ChangeOrderHandle = (($ServiceDeskProxy.doSelect($SessionID, "chg", "chg_ref_num LIKE '$ChangeOrder'", 1, "id") -Split "\r?\n" | Select-String -Pattern "Handle") -Replace "<\/{1}.*$", "") -Replace "^.*>", ""
    $ChangeOrderId = $ChangeOrderHandle -Replace "^(chg:)", ""

    $WorkflowTaskHandles = $ServiceDeskProxy.doSelect($SessionID, "wf", "chg = $ChangeOrderId AND group.last_name LIKE 'Operations Center'", 1, "id") -Split "\r?\n" | Select-String "Handle" | ForEach-Object {
        ($_ -Replace "<\/{1}.*$", "") -Replace "^.*>", ""
    }

    $WorkflowTaskAttributes = @(
        "date_created"
        "last_mod_dt"
        "web_url"
    )

    $UnixTime = New-Object -Type DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0

    $WorkflowTaskHandles | ForEach-Object {
        $TaskHandle = $_
        $TaskHash = $null

        $ServiceDeskProxy.getObjectValues($SessionID, $TaskHandle, $WorkflowTaskAttributes) -Split "\r?\n" | Select-String -Pattern "(<AttrName>)" -Context (0, 1) | ForEach-Object {
            $Pair = $_
            $Name = (($Pair -Split "\r?\n" | Select-String -Pattern "AttrName") -Replace "<\/{1}.*$", "") -Replace "^.*>", ""
            $Value = (($Pair -Split "\r?\n" | Select-String -Pattern "AttrValue") -Replace "<\/{1}.*$", "") -Replace "^.*>", ""
            $TaskHash += @{ $Name = $Value }
        }

        $DateTaskCreated = $UnixTime.AddSeconds(${TaskHash.date_created}).ToLocalTime().ToString("%M/dd/yyyy %h:mm tt")
        $DateTaskModified = $UnixTime.AddSeconds(${TaskHash.last_mod_dt}).ToLocalTime().ToString("%M/dd/yyyy %h:mm tt")
        $DateChangeOrderCreated = $UnixTime.AddSeconds($((($ServiceDeskProxy.getObjectValues($SessionID, $ChangeOrderHandle, "open_date") -Split "\r?\n" | Select-String -Pattern "AttrValue") -Replace "<\/{1}.*$", "") -Replace "^.*>", "")).ToLocalTime().ToString("%M/dd/yyyy %h:mm tt")
        $DateChangeOrderModified = $UnixTime.AddSeconds($((($ServiceDeskProxy.getObjectValues($SessionID, $ChangeOrderHandle, "modified_date") -Split "\r?\n" | Select-String -Pattern "AttrValue") -Replace "<\/{1}.*$", "") -Replace "^.*>", "")).ToLocalTime().ToString("%M/dd/yyyy %h:mm tt")
        $TaskUrl = $TaskHash.web_url
        $TaskComment = (($ServiceDeskProxy.getObjectValues($SessionID, $_, "comments") -Split "\r?\n" | Select-String -Pattern "AttrValue") -Replace "<\/{1}.*$", "") -Replace "^.*>", ""

        $TaskDescriptionXML = $ServiceDeskProxy.getObjectValues($SessionID, $TaskHandle, "description")
        $TaskDescriptionLines = ($TaskDescriptionXML -Split "\r?\n" | Select-String -Pattern "AttrValue").LineNumber
        if ($TaskDescriptionLines.Count -gt 1) {
            $TaskDescription = (($TaskDescriptionXML -Split "\r?\n")[$($TaskDescriptionLines[0] - 1)..$($TaskDescriptionLines[1] - 1)] -Replace "(<AttrValue>)", "") -Replace "(</AttrValue>)", ""
        }
        else {
            $TaskDescription = (($TaskDescriptionXML -Split "\r?\n" | Select-String -Pattern "AttrValue") -Replace "<\/{1}.*$", "") -Replace "^.*>", ""
        }

        $Summary = (($ServiceDeskProxy.getObjectValues($SessionID, $ChangeOrderHandle, "summary") -Split "\r?\n" | Select-String -Pattern "AttrValue") -Replace "<\/{1}.*$", "") -Replace "^.*>", ""
        $Category = (($ServiceDeskProxy.getObjectValues($SessionID, $ChangeOrderHandle, "category.sym") -Split "\r?\n" | Select-String -Pattern "AttrValue") -Replace "<\/{1}.*$", "") -Replace "^.*>", ""
        $ChangeOrderUrl = (($ServiceDeskProxy.getObjectValues($SessionID, $ChangeOrderHandle, "web_url") -Split "\r?\n" | Select-String -Pattern "AttrValue") -Replace "<\/{1}.*$", "") -Replace "^.*>", ""

        $ChangeOrderDescriptionXML = $ServiceDeskProxy.getObjectValues($SessionID, $ChangeOrderHandle, "description")
        $ChangeOrderDescriptionLines = ($ChangeOrderDescriptionXML -Split "\r?\n" | Select-String -Pattern "AttrValue").LineNumber
        $ChangeOrderDescription = (($ChangeOrderDescriptionXML -Split "\r?\n")[$($ChangeOrderDescriptionLines[0] + 1)..$($ChangeOrderDescriptionLines[1] - 3)] -Replace "(<AttrValue>)", "") -Replace "(</AttrValue>)", ""

        $Employee = $ChangeOrderDescription[0] -Replace "^.*\s{3}", ""
        $Title = $ChangeOrderDescription[1] -Replace "^.*\s{3}", ""
        $EmployeeID = $ChangeOrderDescription[2] -Replace "^.*\s{3}", ""
        $Manager = ($ChangeOrderDescription[3] -Replace "^(Manager)\s?", "") -Replace "\s*\;$", ""
        if ($null -eq $Manager -or $Manager -eq "") {
            $Manager = $null
        }
        $Department = $ChangeOrderDescription[4] -Replace "^.*\s{3}", ""
        $EmployeeStatus = $ChangeOrderDescription[5] -Replace "^.*\s", ""

        $obj = [PSCustomObject]@{
            ChangeOrderCategory      = $Category
            ChangeOrderHandle        = $ChangeOrderHandle
            ChangeOrderNumber        = $ChangeOrder
            ChangeOrderUrl           = $ChangeOrderUrl
            DateChangeOrderCreated   = $DateChangeOrderCreated
            DateChangeOrderModified  = $DateChangeOrderModified
            DateWorkflowTaskCreated  = $DateTaskCreated
            DateWorkflowTaskModified = $DateTaskModified
            Department               = $Department
            EmployeeID               = $EmployeeID
            EmployeeName             = $Employee
            EmployeeStatus           = $EmployeeStatus
            ManagerName              = $Manager
            Summary                  = $Summary
            Title                    = $Title
            WorkflowTaskComment      = $TaskComment
            WorkflowTaskDescription  = $TaskDescription
            WorkflowTaskHandle       = $TaskHandle
            WorkflowTaskUrl          = $TaskUrl
        }
        $ServiceDeskInfo.Add($obj)
    }
    $ServiceDeskProxy.logout($SessionID)

    $Log.Status    = "Complete"
    $Log.Timestamp = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    return $ServiceDeskInfo
}