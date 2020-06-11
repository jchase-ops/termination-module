<#
DisableEmployeeSIP method for Termination Module
#>

Function Disable-EmployeeSIP {

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

    $Log.Account = $Account.SamAccountName
    $Log.AccountType = $Account.AccountType
    $Log.ChangeOrder = $ChangeOrder
    $Log.DisplayName = $Account.DisplayName
    $Log.Domain = $Account.Domain
    $Log.EmployeeID = $Account.EmployeeID
    $Log.Status = "Start"
    $Log.Step = $MyInvocation.MyCommand
    $Log.Timestamp = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    $Result.Account = $Account.SamAccountName
    $Result.AccountType = $Account.AccountType
    $Result.ChangeOrder = $ChangeOrder
    $Result.DisplayName = $Account.DisplayName
    $Result.Domain = $Account.Domain
    $Result.EmployeeID = $Account.EmployeeID
    $Result.Status = "Start"
    $Result.Step = $MyInvocation.MyCommand
    $Result.Timestamp = Get-Date -Format FileDateTime

    $Termination.AddResult($Result)

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount

    $CsUser = Get-CsUser -Identity $Account.UserPrincipalName -ErrorAction SilentlyContinue
    if ($null -ne $CsUser) {
        try {
            Disable-CsUser -Identity $CsUser.SipAddress -ErrorAction Stop | Out-Null

            $Log.Status = "Success"
            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $Result.Status = "Success"
            $Result.Timestamp = Get-Date -Format FileDateTime

            $Termination.AddResult($Result)

            return $null
        }
        catch {
            $Log.Status = "Failed"
            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $Result.Status = "Failed"
            $Result.Timestamp = Get-Date -Format FileDateTime

            $Termination.AddResult($Result)

            $ErrorLog.Exception = $_.Exception
            $ErrorLog.ExceptionFullName = $_.Exception.GetType().FullName
            $ErrorLog.FullyQualifiedErrorId = $_.FullyQualifiedErrorId
            $ErrorLog.Line = $_.InvocationInfo.Line
            $ErrorLog.MyCommand = $_.InvocationInfo.MyCommand
            $ErrorLog.ScriptLineNumber = $_.InvocationInfo.ScriptLineNumber
            $ErrorLog.ScriptName = $_.InvocationInfo.ScriptName
            $ErrorLog.Timestamp = Get-Date -Format FileDateTime

            $Main.AddErrorLog($ErrorLog)

            return $Account."msRTCSIP-PrimaryUserAddress"
        }
    }
    else {
        $Log.Status = "Skipped"
        $Log.Timestamp = Get-Date -Format FileDateTime

        $Main.AddLog($Log)

        $Result.Status = "Skipped"
        $Result.Timestamp = Get-Date -Format FileDateTime

        $Termination.AddResult($Result)

        return $null
    }
}