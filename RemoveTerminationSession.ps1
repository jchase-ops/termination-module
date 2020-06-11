<#
RemoveTerminationSession method for Termination Module
#>

Function Remove-TerminationSession {

    [CmdletBinding()]

    Param(
        
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet("DOMAIN1", "O365", "SKYPE", "All")]
        [String]
        $SessionType,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $CurrentCount
    )

    $Log.Account     = $env:USERNAME
    $Log.AccountType = "Developer"
    $Log.ChangeOrder = "N/A"
    $Log.DisplayName = $env:USERNAME -Replace "\.", " "
    $Log.Domain      = $SessionType
    $Log.EmployeeID  = "N/A"
    $Log.Status      = "Start"
    $Log.Step        = $MyInvocation.MyCommand
    $Log.Timestamp   = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount

    Switch ($SessionType) {

        "DOMAIN1" {
            try {
                Get-PSSession | Where-Object { $_.ComputerName -eq $ExchangeHostname } | Remove-PSSession

                $Log.Status    = "Complete"
                $Log.Timestamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)
            }
            catch {
                $Log.Status    = "Failed"
                $Log.Timestamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $ErrorLog.Exception             = $_.Exception
                $ErrorLog.ExceptionFullName     = $_.Exception.GetType().FullName
                $ErrorLog.FullyQualifiedErrorId = $_.FullyQualifiedErrorId
                $ErrorLog.Line                  = $_.InvocationInfo.Line
                $ErrorLog.MyCommand             = $_.InvocationInfo.MyCommand
                $ErrorLog.ScriptLineNumber      = $_.InvocationInfo.ScriptLineNumber
                $ErrorLog.ScriptName            = $_.InvocationInfo.ScriptName
                $ErrorLog.Timestamp             = Get-Date -Format FileDateTime

                $Main.AddErrorLog($ErrorLog)
            }
        }
        "O365" {
            try {
                Get-PSSession | Where-Object { $_.ComputerName -eq "outlook.office365.com" } | Remove-PSSession

                $Log.Status    = "Complete"
                $Log.Timestamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)
            }
            catch {
                $Log.Status    = "Failed"
                $Log.Timestamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $ErrorLog.Exception             = $_.Exception
                $ErrorLog.ExceptionFullName     = $_.Exception.GetType().FullName
                $ErrorLog.FullyQualifiedErrorId = $_.FullyQualifiedErrorId
                $ErrorLog.Line                  = $_.InvocationInfo.Line
                $ErrorLog.MyCommand             = $_.InvocationInfo.MyCommand
                $ErrorLog.ScriptLineNumber      = $_.InvocationInfo.ScriptLineNumber
                $ErrorLog.ScriptName            = $_.InvocationInfo.ScriptName
                $ErrorLog.Timestamp             = Get-Date -Format FileDateTime

                $Main.AddErrorLog($ErrorLog)
            }
        }
        "SKYPE" {
            try {
                Get-PSSession | Where-Object { $_.ComputerName -eq "lyncpool.domain1.com" } | Remove-PSSession

                $Log.Status    = "Complete"
                $Log.Timestamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)
            }
            catch {
                $Log.Status    = "Failed"
                $Log.Timestamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $ErrorLog.Exception             = $_.Exception
                $ErrorLog.ExceptionFullName     = $_.Exception.GetType().FullName
                $ErrorLog.FullyQualifiedErrorId = $_.FullyQualifiedErrorId
                $ErrorLog.Line                  = $_.InvocationInfo.Line
                $ErrorLog.MyCommand             = $_.InvocationInfo.MyCommand
                $ErrorLog.ScriptLineNumber      = $_.InvocationInfo.ScriptLineNumber
                $ErrorLog.ScriptName            = $_.InvocationInfo.ScriptName
                $ErrorLog.Timestamp             = Get-Date -Format FileDateTime

                $Main.AddErrorLog($ErrorLog)
            }
        }
        "All" {
            try {
                Get-PSSession | Remove-PSSession

                $Log.Status    = "Complete"
                $Log.Timestamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)
            }
            catch {
                $Log.Status    = "Failed"
                $Log.Timestamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $ErrorLog.Exception             = $_.Exception
                $ErrorLog.ExceptionFullName     = $_.Exception.GetType().FullName
                $ErrorLog.FullyQualifiedErrorId = $_.FullyQualifiedErrorId
                $ErrorLog.Line                  = $_.InvocationInfo.Line
                $ErrorLog.MyCommand             = $_.InvocationInfo.MyCommand
                $ErrorLog.ScriptLineNumber      = $_.InvocationInfo.ScriptLineNumber
                $ErrorLog.ScriptName            = $_.InvocationInfo.ScriptName
                $ErrorLog.Timestamp             = Get-Date -Format FileDateTime

                $Main.AddErrorLog($ErrorLog)
            }
        }
    }
}