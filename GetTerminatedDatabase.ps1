<#
GetTerminatedDatabase method for Termination Module
#>

Function Get-TerminatedDatabase {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $CurrentCount
    )

    $global:BoxDatabase = New-Object System.Collections.Generic.List[System.Object]

    $Log.Account = $env:USERNAME
    $Log.AccountType = "Developer"
    $Log.ChangeOrder = "N/A"
    $Log.DisplayName = $env:USERNAME -Replace "\.", " "
    $Log.Domain = "All"
    $Log.EmployeeID = "N/A"
    $Log.Status = "Start"
    $Log.Step = $MyInvocation.MyCommand
    $Log.Timestamp = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount

    try {
        $Database = New-Object System.Collections.Generic.List[System.Object]

        Get-MailboxDatabase -Status | Where-Object { $_.Name -like "Terminated*" } | Select-Object Name, DatabaseSize | ForEach-Object {
            $obj = [PSCustomObject]@{
                Name = $_.Name
                Size = ((($_.DatabaseSize -Replace "^.*\(", "") -Replace "\s.*$", "") -Replace "\,", "").ToDouble($null)
            }
            $Database.Add($obj)
        }

        $Log.Status = "Complete"
        $Log.Timestamp = Get-Date -Format FileDateTime

        $Main.AddLog($Log)

        return $Database
    }
    catch {
        $Log.Status = "Failed"
        $Log.Timestamp = Get-Date -Format FileDateTime

        $Main.AddLog($Log)

        $ErrorLog.Exception = $_.Exception
        $ErrorLog.ExceptionFullName = $_.Exception.GetType().FullName
        $ErrorLog.FullyQualifiedErrorId = $_.FullyQualifiedErrorId
        $ErrorLog.Line = $_.InvocationInfo.Line
        $ErrorLog.MyCommand = $_.InvocationInfo.MyCommand
        $ErrorLog.ScriptLineNumber = $_.InvocationInfo.ScriptLineNumber
        $ErrorLog.ScriptName = $_.InvocationInfo.ScriptName
        $ErrorLog.Timestamp = Get-Date -Format FileDateTime

        $Main.AddErrorLog($ErrorLog)

        return $null
    }
}
    