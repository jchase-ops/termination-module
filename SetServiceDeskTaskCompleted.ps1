<#
SetServiceDeskTaskCompleted for Termination Module
#>

Function Set-ServiceDeskTaskCompleted {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String]
        $WorkflowTaskHandle,

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
    $Log.ChangeOrder = "N/A"
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
    $StatusHandle = (($ServiceDeskProxy.doSelect($SessionID, "tskstat", "sym LIKE 'Complete'", 5, "id") -Split "\r?\n" | Select-String -Pattern "(Handle)") -Replace "<\/{1}.*$", "") -Replace "^.*>", ""
    $attrVals = @("status", $StatusHandle)
    $ServiceDeskProxy.updateObject($SessionID, $WorkflowTaskHandle, $attrVals, "id") | Out-Null
    $ServiceDeskProxy.logout($SessionID)
}