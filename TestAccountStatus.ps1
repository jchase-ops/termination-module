<#
TestAccountStatus method for Termination Module
#>

Function Test-AccountStatus {

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
    $Domain = $Account.Domain

    $Log.Account     = $Account.SamAccountName
    $Log.AccountType = $Account.AccountType
    $Log.ChangeOrder = $ChangeOrder
    $Log.DisplayName = $Account.DisplayName
    $Log.Domain      = $Account.Domain
    $Log.EmployeeID  = $Account.EmployeeID
    $Log.Status      = "Start"
    $Log.Step        = $MyInvocation.MyCommand
    $Log.Timestamp   = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount

    Switch ($Domain) {
        "DOMAIN1" {
            $EmployeeStatus = $Account.EmployeeStatus
            $Enabled        = $Account.Enabled

            if ($EmployeeStatus -eq "T" -and $Enabled -eq $false) {
                $Continue      = $true
                $Log.Status    = "Success"
                $Log.Timestamp = Get-Date -Format FileDateTime
            }
            else {
                $Continue      = $false
                $Log.Status    = "Failed"
                $Log.Timestamp = Get-Date -Format FileDateTime
            }
            $Main.AddLog($Log)

            return $Continue, $EmployeeStatus, $Enabled
        }
        "DOMAIN3" {
            $Description    = $Account.Description
            $EmployeeStatus = $Account.EmployeeStatus
            $Enabled        = $Account.Enabled

            if ($Enabled -eq $false -and ($Description -like "*Disabled*" -or $EmployeeStatus -eq "T")) {
                $Continue      = $true
                $Log.Status    = "Success"
                $Log.Timestamp = Get-Date -Format FileDateTime
            }
            else {
                $Continue      = $false
                $Log.Status    = "Failed"
                $Log.Timestamp = Get-Date -Format FileDateTime
            }
            $Main.AddLog($Log)

            return $Continue, $Description, $EmployeeStatus, $Enabled
        }
    }
}