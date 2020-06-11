<#
NewTerminationSession method for Termination Module
#>

Function New-TerminationSession {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateSet("DOMAIN1", "O365", "SKYPE")]
        [String]
        $SessionType,

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
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExchangeHostname/Powershell" -Authentication Kerberos -Credential $Credential
                Import-PSSession -Session $Session -DisableNameChecking -AllowClobber | Out-Null

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
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid" -Authentication Basic -Credential $Credential -AllowRedirection
                Import-PSSession -Session $Session -DisableNameChecking -AllowClobber | Out-Null

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
                $SessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
                $Session = New-PSSession -ConnectionUri "https://lyncpool.domain1.com/OcsPowershell" -SessionOption $SessionOption -Credential $Credential
                Import-PSSession -Session $Session -DisableNameChecking -AllowClobber | Out-Null

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