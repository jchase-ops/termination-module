<#
TestUserAccountAccess method for Termination Module
#>

Function Test-UserAccountAccess {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateSet("DOMAIN1", "DOMAIN2", "DOMAIN3", "O365")]
        [String]
        $Domain,

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
    $Log.Domain      = $Domain
    $Log.EmployeeID  = "N/A"
    $Log.Status      = "Start"
    $Log.Step        = $MyInvocation.MyCommand
    $Log.Timestamp   = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount

    if ($Domain -ne "O365") {
        $Root        = "LDAP://$Domain/$($SearchBase.Admin.Active.$Domain)"
        $Username    = $Credential.UserName -Replace ".*\\{1}", ""
        $Password    = $Credential.GetNetworkCredential().Password
        $DomainEntry = [System.DirectoryServices.DirectoryEntry]::New($Root, $Username, $Password)
        $DomainCheck = (Get-ADUser -Identity $Username -Server $DomainControllers.$Domain.Name).SamAccountName
        if ($DomainEntry -or $DomainCheck) {
            $Log.Status    = "Success"
            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            return $true
        }
        else {
            $Log.Status    = "Failed"
            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            return $false
        }
    }
    else {
        try {
            Connect-AzureAD -Credential $Credential | Out-Null
            $Log.Status    = "Success"
            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            return $true
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
            $ErrorLog.OffsetInLine          = $_.InvocationInfo.OffsetInLine
            $ErrorLog.ScriptLineNumber      = $_.InvocationInfo.ScriptLineNumber
            $ErrorLog.ScriptName            = $_.InvocationInfo.ScriptName
            $ErrorLog.Timestamp             = Get-Date -Format FileDateTime

            $Main.AddErrorLog($ErrorLog)

            return $false
        }
    }
}