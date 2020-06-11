<#
GetEmployeeAccount method for Termination Module
#>

Function Get-EmployeeAccount {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet("Admin", "Manager", "Standard")]
        [String]
        $AccountType,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ChangeOrder,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateSet("DOMAIN1", "DOMAIN2", "DOMAIN3")]
        [String]
        $Domain,

        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = "ID")]
        [ValidateNotNullOrEmpty()]
        [String]
        $EmployeeID,

        [Parameter(Mandatory = $true, Position = 4, ParameterSetName = "Name")]
        [ValidateNotNullOrEmpty()]
        [String]
        $EmployeeName,

        [Parameter(Mandatory = $true, Position = 5)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 6)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $CurrentCount
    )

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    
    Switch ($AccountType) {
        "Admin" {
            Switch ($PSCmdlet.ParameterSetName) {
                "ID" {
                    $Log.Account     = "Searching"
                    $Log.AccountType = $AccountType
                    $Log.ChangeOrder = $ChangeOrder
                    $Log.DisplayName = "Searching"
                    $Log.Domain      = $Domain
                    $Log.EmployeeID  = $EmployeeID
                    $Log.Status      = "Start"
                    $Log.Step        = $MyInvocation.MyCommand
                    $Log.Timestamp   = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    try {
                        $StandardAccount = Get-ADUser -Filter { EmployeeID -eq $EmployeeID } -Server $DomainControllers.$Domain.Name
                        if ($StandardAccount.SamAccountName.Length -gt 18) {
                            $SearchName = "a." + $StandardAccount.SamAccountName.Substring(0, 18)
                        }
                        else {
                            $SearchName = "a.$($StandardAccount.SamAccountName)"
                        }
                        $Account = Get-ADUser -Identity $SearchName -Properties * -Server $DomainControllers.$Domain.Name
                        $Account = [PSCustomObject]@{
                            AccountType       = $AccountType
                            Description       = $Account.Description
                            DisplayName       = $Account.DisplayName
                            DistinguishedName = $Account.DistinguishedName
                            Domain            = $Domain
                            EmployeeID        = $EmployeeID
                            loginShell        = $Account.loginShell
                            msNPAllowDialIn   = $Account.msNPAllowDialIn
                            PasswordLastSet   = $Account.PasswordLastSet
                            SamAccountName    = $Account.SamAccountName
                            unixHomeDirectory = $Account.unixHomeDirectory
                            UserPrincipalName = $Account.UserPrincipalName
                        }

                        $Log.Account     = $Account.SamAccountName
                        $Log.DisplayName = $Account.DisplayName
                        $Log.Status      = "Complete"
                        $Log.Timestamp   = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        return $Account
                    }
                    catch {
                        $Log.Account     = "None"
                        $Log.DisplayName = "None"
                        $Log.Status      = "Failed"
                        $Log.Timestamp   = Get-Date -Format FileDateTime

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
                "Name" {
                    if ($EmployeeName -like "a.*" -and $EmployeeName -notlike "* *") {
                        $Log.Account     = $EmployeeName
                        $Log.AccountType = $AccountType
                        $Log.ChangeOrder = $ChangeOrder
                        $Log.DisplayName = "Searching"
                        $Log.Domain      = $Domain
                        $Log.EmployeeID  = "Admin Account"
                        $Log.Status      = "Start"
                        $Log.Step        = $MyInvocation.MyCommand
                        $Log.Timestamp   = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        try {
                            $Account = Get-ADUser -Identity $EmployeeName -Properties * -Server $DomainControllers.$Domain.Name
                            $Account = [PSCustomObject]@{
                                AccountType       = $AccountType
                                Description       = $Account.Description
                                DisplayName       = $Account.DisplayName
                                DistinguishedName = $Account.DistinguishedName
                                Domain            = $Domain
                                EmployeeID        = (Get-ADUser -Identity (($Account.UserPrincipalName -Replace "\@.*$", "") -Replace "^(a.)", "") -Property EmployeeID).EmployeeID
                                loginShell        = $Account.loginShell
                                msNPAllowDialIn   = $Account.msNPAllowDialIn
                                PasswordLastSet   = $Account.PasswordLastSet
                                SamAccountName    = $Account.SamAccountName
                                unixHomeDirectory = $Account.unixHomeDirectory
                                UserPrincipalName = $Account.UserPrincipalName
                            }
                            $Log.DisplayName = $Account.DisplayName
                            $Log.Status      = "Complete"
                            $Log.Timestamp   = Get-Date -Format FileDateTime

                            $Main.AddLog($Log)

                            return $Account
                        }
                        catch {
                            $Log.DisplayName = "None"
                            $Log.EmployeeID  = "None"
                            $Log.Status      = "Failed"
                            $Log.Timestamp   = Get-Date -Format FileDateTime

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
                        $SearchName = "$EmployeeName (Admin)"

                        $Log.Account     = "Searching"
                        $Log.AccountType = $AccountType
                        $Log.ChangeOrder = $ChangeOrder
                        $Log.DisplayName = $SearchName
                        $Log.Domain      = $Domain
                        $Log.EmployeeID  = "Admin Account"
                        $Log.Status      = "Start"
                        $Log.Step        = $MyInvocation.MyCommand
                        $Log.Timestamp   = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        try {
                            $Account = Get-ADUser -Filter { DisplayName -eq $SearchName } -Properties * -Server $DomainControllers.$Domain.Name
                            if ($Account.Count -gt 1) {
                                $Account = $Account | Select-Object SamAccountName, DisplayName, UserPrincipalName | Out-GridView -Title "Choose Correct Account" -OutputMode Single
                            }
                            else {
                                $Account = Get-ADUser -Identity $Account.SamAccountName -Properties * -Server $DomainControllers.$Domain.Name
                            }
                            $Account = [PSCustomObject]@{
                                AccountType       = $AccountType
                                Description       = $Account.Description
                                DisplayName       = $Account.DisplayName
                                DistinguishedName = $Account.DistinguishedName
                                Domain            = $Domain
                                EmployeeID        = (Get-ADUser -Identity (($Account.UserPrincipalName -Replace "\@.*$", "") -Replace "^(a.)", "") -Property EmployeeID).EmployeeID
                                loginShell        = $Account.loginShell
                                msNPAllowDialIn   = $Account.msNPAllowDialIn
                                PasswordLastSet   = $Account.PasswordLastSet
                                SamAccountName    = $Account.SamAccountName
                                unixHomeDirectory = $Account.unixHomeDirectory
                                UserPrincipalName = $Account.UserPrincipalName
                            }
                            $Log.Account   = $Account.SamAccountName
                            $Log.Status    = "Complete"
                            $Log.Timestamp = Get-Date -Format FileDateTime

                            $Main.AddLog($Log)

                            return $Account
                        }
                        catch {
                            $Log.Account     = "None"
                            $Log.EmployeeID  = "None"
                            $Log.Status      = "Failed"
                            $Log.Timestamp   = Get-Date -Format FileDateTime

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
            }
        }
        "Manager" {
            $Log.Account     = "Searching"
            $Log.AccountType = $AccountType
            $Log.ChangeOrder = $ChangeOrder
            $Log.DisplayName = $EmployeeName
            $Log.Domain      = $Domain
            $Log.EmployeeID  = "Searching"
            $Log.Status      = "Start"
            $Log.Step        = $MyInvocation.MyCommand
            $Log.Timestamp   = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            if ($Domain -eq "DOMAIN3") { $EmployeeName = "$EmployeeName (DOMAIN3)" }

            try {
                $Account = Get-ADUser -Filter { DisplayName -eq $EmployeeName -and Title -ne "Adjunct Faculty" } -Properties * -SearchBase $SearchBase.Default.$Domain -Server $DomainControllers.$Domain.Name
                if ($Account.Count -gt 1) {
                    $Account = $Account | Select-Object SamAccountName, DisplayName, Department, Title, UserPrincipalName | Out-GridView -Title "Choose Correct Account" -OutputMode Single
                }
                else {
                    $Account = Get-ADUser -Identity $Account.SamAccountName -Properties * -Server $DomainControllers.$Domain.Name
                }
                $Log.Account    = $Account.SamAccountName
                $Log.EmployeeID = $Account.EmployeeID
                $Log.Status     = "Complete"
                $Log.Timestamp  = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                return $Account.UserPrincipalName
            }
            catch {
                $Log.Account     = "None"
                $Log.EmployeeID  = "None"
                $Log.Status      = "Failed"
                $Log.Timestamp   = Get-Date -Format FileDateTime

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
        "Standard" {
            Switch ($PSCmdlet.ParameterSetName) {
                "ID" {
                    $Log.Account     = "Searching"
                    $Log.AccountType = $AccountType
                    $Log.ChangeOrder = $ChangeOrder
                    $Log.DisplayName = "Searching"
                    $Log.Domain      = $Domain
                    $Log.EmployeeID  = $EmployeeID
                    $Log.Status      = "Start"
                    $Log.Step        = $MyInvocation.MyCommand
                    $Log.Timestamp   = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    try {
                        $Account = Get-ADUser -Filter { EmployeeID -eq $EmployeeID } -Properties * -Server $DomainControllers.$Domain.Name
                        $Account = Get-ADUser -Identity $Account.SamAccountName -Properties * -Server $DomainControllers.$Domain.Name
                        if ($Domain -ne "DOMAIN2") {
                            $EmployeeStatus = $Account."$($Domain)EmpStatus"
                        }
                        else {
                            $EmployeeStatus = "N/A"
                        }
                        $Account = [PSCustomObject]@{
                            AccountType                     = $AccountType
                            Department                      = $Account.Department
                            Description                     = $Account.Description
                            DisplayName                     = $Account.DisplayName
                            DistinguishedName               = $Account.DistinguishedName
                            Domain                          = $Domain
                            EmployeeID                      = $Account.EmployeeID
                            EmployeeStatus                  = $EmployeeStatus
                            Enabled                         = $Account.Enabled
                            HomeDirectory                   = $Account.HomeDirectory
                            msNPAllowDialIn                 = $Account.msNPAllowDialIn
                            "msRTCSIP-PrimaryUserAddress"   = $Account."msRTCSIP-PrimaryUserAddress"
                            Name                            = $Account.Name
                            PasswordLastSet                 = $Account.PasswordLastSet
                            SamAccountName                  = $Account.SamAccountName
                            ScriptPath                      = $Account.ScriptPath
                            Title                           = $Account.Title
                            UserPrincipalName               = $Account.UserPrincipalName
                        }
                        $Log.Account     = $Account.SamAccountName
                        $Log.DisplayName = $Account.DisplayName
                        $Log.Status      = "Complete"
                        $Log.Timestamp   = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        return $Account
                    }
                    catch {
                        $Log.Account     = "None"
                        $Log.DisplayName = "None"
                        $Log.Status      = "Failed"
                        $Log.Timestamp   = Get-Date -Format FileDateTime

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
                "Name" {
                    if ($EmployeeName -like "*.*" -and $EmployeeName -notlike "* *") {
                        $Log.Account     = $EmployeeName
                        $Log.AccountType = $AccountType
                        $Log.ChangeOrder = $ChangeOrder
                        $Log.DisplayName = "Searching"
                        $Log.Domain      = $Domain
                        $Log.EmployeeID  = "Searching"
                        $Log.Status      = "Start"
                        $Log.Step        = $MyInvocation.MyCommand
                        $Log.Timestamp   = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        try {
                            $Account = Get-ADUser -Identity $EmployeeName -Properties * -Server $DomainControllers.$Domain.Name
                            if ($Domain -ne "DOMAIN2") {
                                $EmployeeStatus = $Account."$($Domain)EmpStatus"
                            }
                            else {
                                $EmployeeStatus = "N/A"
                            }
                            $Account = [PSCustomObject]@{
                                AccountType                     = $AccountType
                                Department                      = $Account.Department
                                Description                     = $Account.Description
                                DisplayName                     = $Account.DisplayName
                                DistinguishedName               = $Account.DistinguishedName
                                Domain                          = $Domain
                                EmployeeID                      = $Account.EmployeeID
                                EmployeeStatus                  = $EmployeeStatus
                                Enabled                         = $Account.Enabled
                                HomeDirectory                   = $Account.HomeDirectory
                                msNPAllowDialIn                 = $Account.msNPAllowDialIn
                                "msRTCSIP-PrimaryUserAddress"   = $Account."msRTCSIP-PrimaryUserAddress"
                                Name                            = $Account.Name
                                PasswordLastSet                 = $Account.PasswordLastSet
                                SamAccountName                  = $Account.SamAccountName
                                ScriptPath                      = $Account.ScriptPath
                                Title                           = $Account.Title
                                UserPrincipalName               = $Account.UserPrincipalName
                            }
                            $Log.DisplayName = $Account.DisplayName
                            $Log.EmployeeID  = $Account.EmployeeID
                            $Log.Status      = "Complete"
                            $Log.Timestamp   = Get-Date -Format FileDateTime

                            $Main.AddLog($Log)

                            return $Account
                        }
                        catch {
                            $Log.DisplayName = "None"
                            $Log.EmployeeID  = "None"
                            $Log.Status      = "Failed"
                            $Log.Timestamp   = Get-Date -Format FileDateTime

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
                        $Log.Account     = "Searching"
                        $Log.AccountType = $AccountType
                        $Log.ChangeOrder = $ChangeOrder
                        $Log.DisplayName = $EmployeeName
                        $Log.Domain      = $Domain
                        $Log.EmployeeID  = "Searching"
                        $Log.Status      = "Start"
                        $Log.Step        = $MyInvocation.MyCommand
                        $Log.Timestamp   = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        try {
                            $Account = Get-ADUser -Filter { DisplayName -eq $EmployeeName } -Properties * -SearchBase $SearchBase.Default.$Domain -Server $DomainControllers.$Domain.Name
                            if ($Account.Count -gt 1) {
                                $Account = $Account | Select-Object SamAccountName, DisplayName, Department, Title, UserPrincipalName | Out-GridView -Title "Choose Correct Account" -OutputMode Single
                            }
                            else {
                                $Account = Get-ADUser -Identity $Account.SamAccountName -Properties * -Server $DomainControllers.$Domain.Name
                            }
                            if ($Domain -ne "DOMAIN2") {
                                $EmployeeStatus = $Account."$($Domain)EmpStatus"
                            }
                            else {
                                $EmployeeStatus = "N/A"
                            }
                            $Account = [PSCustomObject]@{
                                AccountType                     = $AccountType
                                Department                      = $Account.Department
                                Description                     = $Account.Description
                                DisplayName                     = $Account.DisplayName
                                DistinguishedName               = $Account.DistinguishedName
                                Domain                          = $Domain
                                EmployeeID                      = $Account.EmployeeID
                                EmployeeStatus                  = $EmployeeStatus
                                Enabled                         = $Account.Enabled
                                HomeDirectory                   = $Account.HomeDirectory
                                msNPAllowDialIn                 = $Account.msNPAllowDialIn
                                "msRTCSIP-PrimaryUserAddress"   = $Account."msRTCSIP-PrimaryUserAddress"
                                Name                            = $Account.Name
                                PasswordLastSet                 = $Account.PasswordLastSet
                                SamAccountName                  = $Account.SamAccountName
                                ScriptPath                      = $Account.ScriptPath
                                Title                           = $Account.Title
                                UserPrincipalName               = $Account.UserPrincipalName
                            }

                            $Log.Account    = $Account.SamAccountName
                            $Log.EmployeeID = $Account.EmployeeID
                            $Log.Status     = "Complete"
                            $Log.Timestamp  = Get-Date -Format FileDateTime

                            $Main.AddLog($Log)

                            return $Account
                        }
                        catch {
                            $Log.Account     = "None"
                            $Log.EmployeeID  = "None"
                            $Log.Status      = "Failed"
                            $Log.Timestamp   = Get-Date -Format FileDateTime

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
            }
        }
    }
}