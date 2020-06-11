<#
GetTerminationCredential method for Termination Module
#>

Function Get-TerminationCredential {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet("DOMAIN1", "DOMAIN2", "DOMAIN3", "O365")]
        [String]
        $Domain,

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
    $Log.Domain      = $Domain
    $Log.EmployeeID  = "N/A"
    $Log.Status      = "Start"
    $Log.Step        = $MyInvocation.MyCommand
    $Log.Timestamp   = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount

    if ($Domain -ne "O365") {
        try {
            $Credential = $Host.UI.PromptForCredential($MyInvocation.MyCommand, "Enter password for $Domain\a.${env:USERNAME}", "$Domain\a.${env:USERNAME}", "")

            $Log.Status    = "Complete"
            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            return $Credential
        }
        catch {
            Write-Host "$($MyInvocation.MyCommand) for $Domain Failed" -ForegroundColor Red

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

            return $null
        }
    }
    else {
        try {
            $Credential = $Host.UI.PromptForCredential($MyInvocation.MyCommand, "Enter password for ${env:USERNAME}@domain3.com", "${env:USERNAME}@domain3.com", "")

            $Log.Status    = "Complete"
            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            return $Credential
        }
        catch {
            Write-Host "$($MyInvocation.MyCommand) for $Domain Failed" -ForegroundColor Red

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

            return $null
        }
    }
}