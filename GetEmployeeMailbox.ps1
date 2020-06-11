<#
GetEmployeeMailbox method for Termination Module
#>

Function Get-EmployeeMailbox {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [PSObject]
        $Account,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ChangeOrder,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $CurrentCount
    )

    $Log.Account     = $Account.SamAccountName
    $Log.AccountType = "Mailbox"
    $Log.ChangeOrder = $ChangeOrder
    $Log.DisplayName = $Account.DisplayName
    $Log.Domain      = $Account.Domain
    $Log.EmployeeID  = $Account.EmployeeID
    $Log.Status      = "Start"
    $Log.Step        = $MyInvocation.MyCommand
    $Log.Timestamp   = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    $Result.Account     = $Account.SamAccountName
    $Result.AccountType = "Mailbox"
    $Result.ChangeOrder = $ChangeOrder
    $Result.DisplayName = $Account.DisplayName
    $Result.Domain      = $Account.Domain
    $Result.EmployeeID  = $Account.EmployeeID
    $Result.Status      = "Start"
    $Result.Step        = $MyInvocation.MyCommand
    $Result.Timestamp   = Get-Date -Format FileDateTime

    $Termination.AddResult($Result)

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount

    if ($Account.Title -ne "Adjunct Faculty" -and $Account.Department -ne "Adjuncts") {
        try {
            $Mailbox = Get-Mailbox -Identity $Account.Name -WarningAction SilentlyContinue -ErrorAction Stop
            $Mailbox = [PSCustomObject]@{
                AccountType                   = "Mailbox"
                Alias                         = $Mailbox.Alias
                AutoReplyState                = (Get-MailboxAutoReplyConfiguration -Identity $Mailbox.Name).AutoReplyState
                Database                      = $Mailbox.Database
                DeliverToMailboxAndForward    = $Mailbox.DeliverToMailboxAndForward
                DisplayName                   = $Mailbox.DisplayName
                DistinguishedName             = $Mailbox.DistinguishedName
                Domain                        = $Account.Domain
                EmployeeID                    = $Account.EmployeeID
                ExchangeGuid                  = $Mailbox.ExchangeGuid.Guid
                ForwardingAddress             = $Mailbox.ForwardingAddress
                Guid                          = $Mailbox.Guid.Guid
                HiddenFromAddressListsEnabled = $Mailbox.HiddenFromAddressListsEnabled
                Id                            = $Mailbox.Id
                Identity                      = $Mailbox.Identity
                MailboxMoveBatchName          = $Mailbox.MailboxMoveBatchName
                MailboxMoveStatus             = $Mailbox.MailboxMoveStatus
                Name                          = $Mailbox.Name
                SamAccountName                = $Mailbox.SamAccountName
                UserPrincipalName             = $Mailbox.UserPrincipalName
            }
            $Log.Status    = "Complete"
            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $Result.Status    = "Complete"
            $Result.Timestamp = Get-Date -Format FileDateTime

            $Termination.AddResult($Result)

            return $Mailbox
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

            return $null
        }
    }
    else {
        $Log.Status    = "Skipped"
        $Log.Timestamp = Get-Date -Format FileDateTime

        $Main.AddLog($Log)

        $Result.Status    = "Skipped"
        $Result.Timestamp = Get-Date -Format FileDateTime

        $Termination.AddResult($Result)

        return $null
    }
}