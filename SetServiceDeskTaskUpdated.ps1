<#
SetServiceDeskTaskUpdated for Termination Module
#>

Function Set-ServiceDeskTaskUpdated {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Comment,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [String]
        $WorkflowTaskHandle,

        [Parameter(Mandatory = $true, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 4)]
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
    $attrVals = @("comments", $Comment)
    $ServiceDeskProxy.updateObject($SessionID, $WorkflowTaskHandle, $attrVals, "id") | Out-Null
    $ServiceDeskProxy.logout($SessionID)
}