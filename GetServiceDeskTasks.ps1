<#
GetServiceDeskTasks for Termination Module
#>

Function Get-ServiceDeskTasks {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential
    )

    $Log.Account     = $env:USERNAME
    $Log.AccountType = "Developer"
    $Log.ChangeOrder = "N/A"
    $Log.DisplayName = $env:USERNAME -Replace "\.", " "
    $Log.Domain      = "All"
    $Log.EmployeeID  = "N/A"
    $Log.Status      = "Start"
    $Log.Step        = $MyInvocation.MyCommand
    $Log.Timestamp   = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    $TaskAttributes = @(
        "chg"
        "comments"
        "date_created"
        "description"
        "last_mod_dt"
        "persistent_id"
        "start_date"
        "web_url"
    )

    $ChangeOrderAttributes = @(
        "category.sym"
        "chg_ref_num"
        "description"
        "modified_date"
        "open_date"
        "persistent_id"
        "summary"
        "web_url"
    )

    $UnixTime = New-Object -TypeName DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0
    $TaskXML = New-Object -TypeName System.Xml.XmlDocument
    $ServiceDeskProxy = New-WebServiceProxy -Uri $ServiceDeskUri
    $SessionID = $ServiceDeskProxy.login($Credential.Username, $Credential.GetNetworkCredential().Password)
    $objectType = "wf"
    $whereClause = "group.last_name = 'Termination Development' AND status = 'PEND' AND (sequence = 10 OR sequence = 11)"
    $TaskList = $ServiceDeskProxy.doQuery($SessionID, $objectType, $whereClause)
    $TaskXML.LoadXml($ServiceDeskProxy.getListValues($SessionID, $TaskList.listHandle, 0, $TaskList.listLength-1, $TaskAttributes))
    $ServiceDeskProxy.freeListHandles($SessionID, $TaskList.listHandle)
    $ServiceDesk = New-Object System.Collections.Generic.List[System.Object]
    $TaskXML.UDSObjectList.UDSObject | ForEach-Object {
        $TaskHandle = $_.Handle
        $TaskValues = $_.Attributes.Attribute
        $ChangeOrderHandle = "chg:" + ($TaskValues | Where-Object AttrName -eq chg).AttrValue 
        $ChangeOrderXML = New-Object -TypeName System.Xml.XmlDocument
        $ChangeOrderXML.LoadXml($ServiceDeskProxy.getObjectValues($SessionID, $ChangeOrderHandle, $ChangeOrderAttributes))
        $Description = ($ChangeOrderXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq description).AttrValue
        $FirstLine = $Description -Split "\r?\n" | Select-Object -First 1
        if ($FirstLine -notlike "*Employment Status Change*") {
            if ($FirstLine -like "*TERMINATION OF REMOTE EMPLOYEE*") {
                $Result = $null
                $Handle = $null
                $RequestXML = New-Object System.Xml.XmlDocument
                $RequestCO = $ChangeOrderHandle -Replace "^.*\:", ""
                $RequestXML.LoadXml($ServiceDeskProxy.doSelect($SessionID, "cr", "change = $RequestCO", -1, @()))
                $GroupXML = New-Object System.Xml.XmlDocument
                $GroupXML.LoadXml($ServiceDeskProxy.doSelect($SessionID, "cnt", "last_name = 'Frontline Support'", -1, @()))
                $RequestStatusXML = New-Object System.Xml.XmlDocument
                $RequestStatusXML.LoadXml($ServiceDeskProxy.doSelect($SessionID, "crs", "code = 'OP'", -1, @()))
                $ChangeCancelXML = New-Object System.Xml.XmlDocument
                $ChangeCancelXML.LoadXml($ServiceDeskProxy.doSelect($SessionID, "chgstat", "sym = 'Cancelled'", -1, @()))
                $ChangeCloseXML = New-Object System.Xml.XmlDocument
                $ChangeCloseXML.LoadXml($ServiceDeskProxy.doSelect($SessionID, "chgstat", "sym = 'Closed'", -1, @()))
                $ActivityDescription = "Cancelling and closing.  CO should not have been created."
                $ActivityType = ($CancelledXML.UDSObjectList.UDSObject.Attributes.Attribute | Where-Object AttrName -eq code).AttrValue
                $ActivityTime = "0"
                $ChangeID = $ChangeOrderHandle -Replace "^.*\:", ""
                $ActivityValueArray = @(
                    "type"
                    $ActivityType
                    "time_spent"
                    $ActivityTime
                    "change_id"
                    $ChangeID
                    "description"
                    $ActivityDescription
                )
                $ActivityXML = New-Object System.Xml.XmlDocument
                $ActivityXML.LoadXml($ServiceDeskProxy.createObject($SessionID, "chgalg", $ActivityValueArray, @(), [ref]$Result, [ref]$Handle))
                $UpdateValueArray = @(
                    "status"
                    ($ClosedXML.UDSObjectList.UDSObject.Attributes.Attribute | Where-Object AttrName -eq persistent_id).AttrValue
                )
                $ResultXML = New-Object System.Xml.XmlDocument
                $ResultXML.LoadXml($ServiceDeskProxy.updateObject($SessionID, $ChangeOrderHandle, $UpdateValueArray, @()))
            }
        }
        $obj = [PSCustomObject]@{
            ChangeOrderCategory = ($ChangeOrderXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq category.sym).AttrValue
            ChangeOrderHandle = ($ChangeOrderXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq persistent_id).AttrValue
            ChangeOrderNumber = ($ChangeOrderXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq chg_ref_num).AttrValue
            ChangeOrderUrl = ($ChangeOrderXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq web_url).AttrValue
            DateChangeOrderCreated = $UnixTime.AddSeconds(($ChangeOrderXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq open_date).AttrValue).ToLocalTime().ToString("%M/dd/yyyy %h:mm tt")
            DateChangeOrderModified = $UnixTime.AddSeconds(($ChangeOrderXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq modified_date).AttrValue).ToLocalTime().ToString("%M/dd/yyyy %h:mm tt")
            DateWorkflowTaskCreated = $UnixTime.AddSeconds(($TaskXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq date_created).AttrValue).ToLocalTime().ToString("%M/dd/yyyy %h:mm tt")
            DateWorkflowTaskModified = $UnixTime.AddSeconds(($TaskXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq last_mod_dt).AttrValue).ToLocalTime().ToString("%M/dd/yyyy %h:mm tt")
            Department = 
            WorkflowTaskComment = ($TaskXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq comments).AttrValue
            WorkflowTaskDescription = ($TaskXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq description).AttrValue
            WorkflowTaskHandle = ($TaskXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq persistent_id).AttrValue
            WorkflowTaskUrl = ($TaskXML.UDSObject.Attributes.Attribute | Where-Object AttrName -eq web_url).AttrValue
        }
    }
    $ServiceDeskProxy.logout($SessionID)

    $Log.Status    = "Complete"
    $Log.Timestamp = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    return $ServiceDesk
}