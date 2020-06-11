<#
SetTerminatedOutOfOffice method for Termination Module
#>

Function Set-TerminatedOutOfOffice {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ChangeOrder,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [PSObject]
        $Mailbox,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $CurrentCount
    )

    $Log.Account     = $Mailbox.Alias
    $Log.AccountType = $Mailbox.AccountType
    $Log.ChangeOrder = $ChangeOrder
    $Log.DisplayName = $Mailbox.DisplayName
    $Log.Domain      = $Mailbox.Domain
    $Log.EmployeeID  = $Mailbox.EmployeeID
    $Log.Status      = "Start"
    $Log.Step        = $MyInvocation.MyCommand
    $Log.Timestamp   = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    $Result.Account     = $Mailbox.Alias
    $Result.AccountType = $Mailbox.AccountType
    $Result.ChangeOrder = $ChangeOrder
    $Result.DisplayName = $Mailbox.DisplayName
    $Result.Domain      = $Mailbox.Domain
    $Result.EmployeeID  = $Mailbox.EmployeeID
    $Result.Status      = "Start"
    $Result.Step        = $MyInvocation.MyCommand
    $Result.Timestamp   = Get-Date -Format FileDateTime

    $Termination.AddResult($Result)

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    
    try {
        Set-MailboxAutoReplyConfiguration -Identity $Mailbox.Name -AutoReplyState Enabled -InternalMessage $Message.$($Mailbox.Domain) -ExternalMessage $Message.$($Mailbox.Domain) -WarningAction SilentlyContinue -ErrorAction Stop

        $Log.Status    = "Success"
        $Log.Timestamp = Get-Date -Format FileDateTime

        $Main.AddLog($Log)

        $Result.Status    = "Success"
        $Result.Timestamp = Get-Date -Format FileDateTime

        $Termination.AddResult($Result)

        return "Enabled"
    }
    catch {
        $Log.Status    = "Failed"
        $Log.Timestamp = Get-Date -Format FileDateTime
                
        $Main.AddLog($Log)

        $Result.Status    = "Failed"
        $Result.Timestamp = Get-Date -Format FileDateTime

        $Termination.AddResult($Result)

        $ErrorLog.Exception             = $_.Exception
        $ErrorLog.ExceptionFullName     = $_.Exception.GetType().FullName
        $ErrorLog.FullyQualifiedErrorId = $_.FullyQualifiedErrorId
        $ErrorLog.Line                  = $_.InvocationInfo.Line
        $ErrorLog.MyCommand             = $_.InvocationInfo.MyCommand
        $ErrorLog.ScriptLineNumber      = $_.InvocationInfo.ScriptLineNumber
        $ErrorLog.ScriptName            = $_.InvocationInfo.ScriptName
        $ErrorLog.Timestamp             = Get-Date -Format FileDateTime

        $Main.AddErrorLog($ErrorLog)

        return $Mailbox.AutoReplyState
    }
}