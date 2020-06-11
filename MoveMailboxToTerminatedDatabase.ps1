<#
MoveMailboxToTerminatedDatabase method for Termination Module
#>

Function Move-MailboxToTerminatedDatabase {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ChangeOrder,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [PSObject[]]
        $Database,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [PSObject]
        $Mailbox,

        [Parameter(Mandatory = $true, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 4)]
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

    $MostEmpty = $Database | Sort-Object Size | Select-Object -First 1
    $DatabaseUpdate.Name = $MostEmpty.Name
    $DatabaseUpdate.Index = $Database.Name.IndexOf($MostEmpty.Name)
    $DatabaseUpdate.Size = $MostEmpty.Size

    $obj = [PSCustomObject]@{
        BatchName = $Mailbox.Alias
        Database = $MostEmpty.Name
        InitialSize = $MostEmpty.Size
        MailboxSize = $null
        FinalSize = $null
    }

    $MoveRequest = Get-MoveRequest -BatchName $Mailbox.Alias
    if ($null -eq $MoveRequest) {
        try {
            $MailboxSize = (((((Get-MailboxStatistics -Identity $Mailbox.Name).TotalItemSize).Value -Replace "^.*\(", "") -Replace "\s.*$", "") -Replace "\,", "").ToDouble($null)
            New-MoveRequest -Identity $Mailbox.Name -TargetDatabase $MostEmpty.Name -BadItemLimit 50 -BatchName $Mailbox.Alias -WarningAction SilentlyContinue -ErrorAction Stop

            $Log.Status    = "Success"
            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $Result.Status    = "Success"
            $Result.Timestamp = Get-Date -Format FileDateTime

            $Termination.AddResult($Result)

            $DatabaseUpdate.Name = $MostEmpty.Name
            $DatabaseUpdate.Size = $MostEmpty.Size + $MailboxSize

            $obj.MailboxSize = $MailboxSize
            $obj.FinalSize = $DatabaseUpdate.Size

            $BoxDatabase.Add($obj)

            return $Mailbox.Alias
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

            $DatabaseUpdate.Name = $MostEmpty.Name
            $DatabaseUpdate.Size = $MostEmpty.Size

            $BoxDatabase.Add($obj)

            return $null
        }
    }
    else {
        $Log.Status = "Skipped"
        $Log.TimeStamp = Get-Date -Format FileDateTime

        $Main.AddLog($Log)

        $Result.Status = "Skipped"
        $Result.TimeStamp = Get-Date -Format FileDateTime

        $Termination.AddResult($Result)

        $DatabaseUpdate.Name = $MostEmpty.Name
        $DatabaseUpdate.Size = $MostEmpty.Size

        $BoxDatabase.Add($obj)

        return $MoveRequest.BatchName
    }
}