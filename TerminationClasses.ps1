<#
TerminationClasses for Termination Module
#>

class Main {
    # Properties
    [Termination[]] $Terminations
    [Log[]] $Logs = [System.Collections.ArrayList]::New()
    [ErrorLog[]] $ErrorLogs = [System.Collections.ArrayList]::New()
    [PSObject] Hidden $Credential
    [PSObject[]] Hidden $Database
    [PSObject] Hidden $VerifiedAccess
    [String] Hidden $ErrorLogFilePath
    [String] Hidden $GroupFilePath
    [String] Hidden $LogFilePath
    [String] Hidden $ResultFilePath

    # Constructors
    Main () {
        $This.Credential = [PSCustomObject]@{
            DOMAIN1    = $null
            DOMAIN2      = $null
            DOMAIN3         = $null
            O365        = $null
            ServiceDesk = [PSCredential]::New('SvcAcctUsername', $(ConvertTo-SecureString -String "SvcAcctPW" -AsPlainText -Force))
            SKYPE  = $null
        }
        $This.VerifiedAccess = [PSCustomObject]@{
            DOMAIN1    = $null
            DOMAIN2 = $null
            DOMAIN3    = $null
            O365   = $null
        }
    }

    Main (
        [String] $ChangeOrder
    ) {
        $This.Terminations += [Termination]::New($ChangeOrder)
        $This.Credential = [PSCustomObject]@{
            DOMAIN1    = $null
            DOMAIN2 = $null
            DOMAIN3    = $null
            O365   = $null
            ServiceDesk = [PSCredential]::New("SvcAcctUsername", $(ConvertTo-SecureString -String "SvcAcctPW" -AsPlainText -Force))
            SKYPE  = $null
        }
        $This.VerifiedAccess = [PSCustomObject]@{
            DOMAIN1    = $null
            DOMAIN2 = $null
            DOMAIN3    = $null
            O365   = $null
        }
    }

    Main (
        [PSObject] $Object
    ) {
        $This.Credential = [PSCustomObject]@{
            DOMAIN1    = $Object.DOMAIN1
            DOMAIN2 = $Object.DOMAIN2
            DOMAIN3    = $Object.DOMAIN3
            O365   = $Object.O365
            ServiceDesk = [PSCredential]::New("SvcAcctUsername", $(ConvertTo-SecureString -String "SvcAcctPW" -AsPlainText -Force))
            SKYPE  = $Object.SKYPE
        }
        $This.VerifiedAccess = [PSCustomObject]@{
            DOMAIN1    = $null
            DOMAIN2 = $null
            DOMAIN3    = $null
            O365   = $null
        }
    }

    # Methods
    [Void] AddCredentials (
        [PSObject] $Object
    ) {
        $This.Credential.DOMAIN1    = $Object.DOMAIN1
        $This.Credential.DOMAIN2 = $Object.DOMAIN2
        $This.Credential.DOMAIN3    = $Object.DOMAIN3
        $This.Credential.O365   = $Object.O365
        $This.Credential.SKYPE  = $Object.SKYPE
    }

    [Void] AddErrorLog (
        [PSObject] $Object
    ) {
        $This.ErrorLogs += [ErrorLog]::New($Object)
    }
    
    [Void] AddLog (
        [PSObject] $Object
    ) {
        $This.Logs += [Log]::New($Object)
    }

    [Void] AddTermination (
        [PSObject] $Object
    ) {
        if ($Object.ChangeOrderNumber -notin $This.Terminations.ChangeOrder) {
            $ServiceDeskInfo = [ServiceDeskInfo]::New($Object)
            $This.Terminations += [Termination]::New($ServiceDeskInfo)
        }
        else {
            $ServiceDeskInfo = [ServiceDeskInfo]::New($Object)
            ($This.Terminations | Where-Object { $_.ChangeOrder -eq $ServiceDeskInfo.ChangeOrderNumber }).AddServiceDeskInfo($ServiceDeskInfo)
        }
    }

    [Void] AddTermination (
        [String] $ChangeOrder
    ) {
        if ($ChangeOrder -notin $This.Terminations.ChangeOrder) {
            $This.Terminations += [Termination]::New($ChangeOrder)
        }
    }

    [Void] CopyItemsToSharePoint (
        [PSCredential] $Credential
    ) {
        $AccountTypes = "Standard", "Admin"
        $Domains = "DOMAIN1", "DOMAIN2", "DOMAIN3"
        $AllGroups = New-Object System.Collections.Generic.List[System.Object]
        $This.Terminations | ForEach-Object {
            $Termination = $_
            ForEach ($AccountType in $AccountTypes) {
                ForEach ($Domain in $Domains) {
                    if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                        $Termination.AccountData.$AccountType.$Domain.Groups | ForEach-Object {
                            $AllGroups.Add($_)
                        }
                    }
                }
            }
        }
        $AllResults = New-Object System.Collections.Generic.List[System.Object]
        $This.Terminations | ForEach-Object {
            $_.Results | Where-Object { $_.Status -eq "Success" -or $_.Status -eq "Skipped" -or $_.Status -eq "Failed" } | ForEach-Object {
                $AllResults.Add($_)
            }
        }
        Copy-ItemsToSharePoint -Credential $Credential -Groups $AllGroups -Results $AllResults
    }

    [Void] GetErrorLogFile (
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.ErrorLogFilePath = Get-ErrorLogFile -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] GetGroupFile (
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.GroupFilePath = Get-GroupFile -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] GetLogFile (
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.LogFilePath = Get-LogFile -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] GetResultFile (
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.ResultFilePath = Get-ResultFile -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] GetServiceDeskTasks (
        [PSCredential] $Credential
    ) {
        Get-ServiceDeskTasks -Credential $Credential | ForEach-Object {
            $This.AddTermination($_)
        }
    }

    [Void] GetTerminatedDatabase (
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.Database = Get-TerminatedDatabase -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] GetTerminationCredential (
        [String] $Domain,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        if ($Domain -eq "SKYPE") {
            $This.Credential.$Domain = [PSCredential]::New($env:USERNAME, $This.Credential.O365.Password)
        }
        else {
            $NewCredential = Get-TerminationCredential -Domain $Domain -TerminationCount $TerminationCount -CurrentCount $CurrentCount
            if ($null -ne $NewCredential) {
                $This.Credential.$Domain = $NewCredential
            }
        }
    }

    [Void] NewTeamsTerminationNotification () {
        New-TeamsTerminationNotification -Terminations $This.Terminations
    }

    [Void] NewTerminationSession (
        [String] $SessionType,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        New-TerminationSession -SessionType $SessionType -Credential $This.Credential.$SessionType -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] RemoveTerminationSession (
        [String] $SessionType,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        Remove-TerminationSession -SessionType $SessionType -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] TestUserAccountAccess (
        [String] $Domain,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.VerifiedAccess.$Domain = Test-UserAccountAccess -Credential $This.Credential.$Domain -Domain $Domain -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] WriteErrorLogsToFile () {
        $This.ErrorLogs | Export-Csv -Path $This.ErrorLogFilePath -Append -NoClobber -NoTypeInformation
    }

    [Void] WriteGroupsToFile () {
        $AccountTypes = "Standard", "Admin"
        $Domains = "DOMAIN1", "DOMAIN2", "DOMAIN3"
        $AllGroups = New-Object System.Collections.Generic.List[System.Object]
        $This.Terminations | ForEach-Object {
            $Termination = $_
            ForEach ($AccountType in $AccountTypes) {
                ForEach ($Domain in $Domains) {
                    if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                        $Termination.AccountData.$AccountType.$Domain.Groups | ForEach-Object {
                            $AllGroups.Add($_)
                        }
                    }
                }
            }
        }
        $AllGroups | Export-Csv -Path $This.GroupFilePath -Append -NoClobber -NoTypeInformation
    }

    [Void] WriteLogsToFile () {
        $This.Logs | Export-Csv -Path $This.LogFilePath -Append -NoClobber -NoTypeInformation
    }

    [Void] WriteResultsToFile() {
        $AllResults = New-Object System.Collections.Generic.List[System.Object]
        $This.Terminations | ForEach-Object {
            $_.Results | ForEach-Object {
                $AllResults.Add($_)
            }
        }
        $AllResults | Export-Csv -Path $This.ResultFilePath -Append -NoClobber -NoTypeInformation
    }

    [Void] UpdateDatabaseSize(
        [PSObject] $Object
    ) {
        $This.Database[$($Object.Index)].Size = $Object.Size
    }
}

class AdminAccount {
    # Properties
    [String] $AccountType
    [String] $Description
    [String] $DisplayName
    [String] $DistinguishedName
    [String] $Domain
    [String] $EmployeeID
    [Group[]] $Groups = [System.Collections.ArrayList]::New()
    [String] $loginShell
    [String] $msNPAllowDialIn
    [String] $PasswordLastSet
    [String] $SamAccountName
    [String] $unixHomeDirectory
    [String] $UserPrincipalName

    # Constructors
    AdminAccount (
        [PSObject] $Object
    ) {
        $This.AccountType       = $Object.AccountType
        $This.Description       = $Object.Description
        $This.DisplayName       = $Object.DisplayName
        $This.DistinguishedName = $Object.DistinguishedName
        $This.Domain            = $Object.Domain
        $This.EmployeeID        = $Object.EmployeeID
        $This.loginShell        = $Object.loginShell
        $This.msNPAllowDialIn   = $Object.msNPAllowDialIn
        $This.PasswordLastSet   = $Object.PasswordLastSet
        $This.SamAccountName    = $Object.SamAccountName
        $This.unixHomeDirectory = $Object.unixHomeDirectory
        $This.UserPrincipalName = $Object.UserPrincipalName
    }

    # Methods
    [Void] AddGroup (
        [PSObject] $Object
    ) {
        $This.Groups += [Group]::New($Object)
    }

    [Void] DisableDialInAccess (
        [String] $ChangeOrder,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.msNPAllowDialIn = Disable-DialInAccess -Account $This -ChangeOrder $ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] RemoveEmployeeGroups (
        [String] $ChangeOrder,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        Remove-EmployeeGroups -Account $This -ChangeOrder $ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount | ForEach-Object {
            $This.AddGroup($_)
        }
    }

    [Void] RemoveUnixAttributes (
        [String] $ChangeOrder,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $Result = Remove-UnixAttributes -Account $This -ChangeOrder $ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount
        if ($Result) {
            $This.loginShell = $null
            $This.unixHomeDirectory = $null
        }
    }

    [Void] SetRandomPassword (
        [String] $ChangeOrder,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.PasswordLastSet = Set-RandomPassword -Account $This -ChangeOrder $ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] SetTerminatedOrganizationalUnit (
        [String] $ChangeOrder,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.DistinguishedName = Set-TerminatedOrganizationalUnit -Account $This -ChangeOrder $ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }
}

class ErrorLog {
    # Properties
    [String] $Exception
    [String] $ExceptionFullName
    [String] $FullyQualifiedErrorId
    [String] $Line
    [String] $MyCommand
    [String] $ScriptLineNumber
    [String] $ScriptName
    [String] $TimeStamp

    # Constructors
    ErrorLog (
        [PSObject] $Object
    ) {
        $This.Exception             = $Object.Exception
        $This.ExceptionFullName     = $Object.ExceptionFullName
        $This.FullyQualifiedErrorId = $Object.FullyQualifiedErrorId
        $This.Line                  = $Object.Line
        $This.MyCommand             = $Object.MyCommand
        $This.ScriptLineNumber      = $Object.ScriptLineNumber
        $This.ScriptName            = $Object.ScriptName
        $This.TimeStamp             = $Object.TimeStamp
    }
}

class Group {
    # Properties
    [String] $AccountType
    [String] $ChangeOrder
    [String] $Description
    [String] $DistinguishedName
    [String] $Domain
    [String] $FIM
    [String] $GroupCategory
    [String] $GroupResult
    [String] $GroupScope
    [String] $ManagedBy
    [String] $Name
    [String] $SamAccountName
    [String] $TimeStamp

    # Constructors
    Group (
        [PSObject]$Object
    ) {
        $This.AccountType          = $Object.AccountType
        $This.ChangeOrder          = $Object.ChangeOrder
        $This.Description          = $Object.Description
        $This.DistinguishedName    = $Object.DistinguishedName
        $This.Domain               = $Object.Domain
        $This.FIM                  = $Object.FIM
        $This.GroupCategory        = $Object.GroupCategory
        $This.GroupResult          = $Object.GroupResult
        $This.GroupScope           = $Object.GroupScope
        $This.ManagedBy            = $Object.ManagedBy
        $This.Name                 = $Object.Name
        $This.SamAccountName       = $Object.SamAccountName
        $This.TimeStamp            = $Object.TimeStamp
    }
}

class Log {
    # Properties
    [String] $Account
    [String] $AccountType
    [String] $ChangeOrder
    [String] $DisplayName
    [String] $Domain
    [String] $EmployeeID
    [String] $Section
    [String] $Status
    [String] $Step
    [String] $TimeStamp

    # Constructors
    Log (
        [PSObject] $Object
    ) {
        $This.Account     = $Object.Account
        $This.AccountType = $Object.AccountType
        $This.ChangeOrder = $Object.ChangeOrder
        $This.DisplayName = $Object.DisplayName
        $This.Domain      = $Object.Domain
        $This.EmployeeID  = $Object.EmployeeID
        $This.Section     = $Object.Section
        $This.Status      = $Object.Status
        $This.Step        = $Object.Step
        $This.TimeStamp   = $Object.TimeStamp
    }
}

class Mailbox {
    # Properties
    [String] $AccountType
    [String] $Alias
    [String] $AutoReplyState
    [String] $Database
    [String] $DeliverToMailboxAndForward
    [String] $DisplayName
    [String] $DistinguishedName
    [String] $Domain
    [String] $EmployeeID
    [String] $ExchangeGuid
    [String] $ForwardingAddress
    [String] $Guid
    [String] $HiddenFromAddressListsEnabled
    [String] $Id
    [String] $Identity
    [String] $MailboxMoveBatchName
    [String] $MailboxMoveStatus
    [String] $Name
    [String] $SamAccountName
    [String] $UserPrincipalName

    # Constructors
    Mailbox (
        [PSObject] $Object
    ) {
        $This.AccountType                   = $Object.AccountType
        $This.Alias                         = $Object.Alias
        $This.AutoReplyState                = $Object.AutoReplyState
        $This.Database                      = $Object.Database
        $This.DeliverToMailboxAndForward    = $Object.DeliverToMailboxAndForward
        $This.DisplayName                   = $Object.DisplayName
        $This.DistinguishedName             = $Object.DistinguishedName
        $This.Domain                        = $Object.Domain
        $This.EmployeeID                    = $Object.EmployeeID
        $This.ExchangeGuid                  = $Object.ExchangeGuid
        $This.ForwardingAddress             = $Object.ForwardingAddress
        $This.Guid                          = $Object.Guid
        $This.HiddenFromAddressListsEnabled = $Object.HiddenFromAddressListsEnabled
        $This.Id                            = $Object.Id
        $This.Identity                      = $Object.Identity
        $This.MailboxMoveBatchName          = $Object.MailboxMoveBatchName
        $This.MailboxMoveStatus             = $Object.MailboxMoveStatus
        $This.Name                          = $Object.Name
        $This.SamAccountName                = $Object.SamAccountName
        $This.UserPrincipalName             = $Object.UserPrincipalName
    }

    # Methods
    [Void] EnableMailboxForwarding (
        [String] $ChangeOrder,
        [String] $ManagerEmailAddress,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.ForwardingAddress = Enable-MailboxForwarding -ChangeOrder $ChangeOrder -Mailbox $This -ManagerEmailAddress $ManagerEmailAddress -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] MoveMailboxToTerminatedDatabase (
        [String] $ChangeOrder,
        [PSObject[]] $Database,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.MailboxMoveBatchName = Move-MailboxToTerminatedDatabase -ChangeOrder $ChangeOrder -Database $Database -Mailbox $This -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] RemoveInboxRules (
        [String] $ChangeOrder,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        Remove-InboxRules -ChangeOrder $ChangeOrder -Mailbox $This -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] SetContactHiddenFromGAL (
        [String] $ChangeOrder,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        Set-ContactHiddenFromGAL -ChangeOrder $ChangeOrder -Mailbox $This -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] SetMailboxHiddenFromGAL (
        [String] $ChangeOrder,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        Set-MailboxHiddenFromGAL -ChangeOrder $ChangeOrder -Mailbox $This -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] SetTerminatedOutOfOffice (
        [String] $ChangeOrder,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.AutoReplyState = Set-TerminatedOutOfOffice -ChangeOrder $ChangeOrder -Mailbox $This -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }
}

class Result {
    # Properties
    [String] $Account
    [String] $AccountType
    [String] $ChangeOrder
    [String] $DisplayName
    [String] $Domain
    [String] $EmployeeID
    [String] $Section
    [String] $Status
    [String] $Step
    [String] $TimeStamp

    # Constructors
    Result (
        [PSObject] $Object
    ) {
        $This.Account     = $Object.Account
        $This.AccountType = $Object.AccountType
        $This.ChangeOrder = $Object.ChangeOrder
        $This.DisplayName = $Object.DisplayName
        $This.Domain      = $Object.Domain
        $This.EmployeeID  = $Object.EmployeeID
        $This.Section     = $Object.Section
        $This.Status      = $Object.Status
        $This.Step        = $Object.Step
        $This.TimeStamp   = $Object.TimeStamp
    }
}

class ServiceDeskInfo {
    # Properties
    [String] $ChangeOrderCategory
    [String] $ChangeOrderHandle
    [String] $ChangeOrderNumber
    [String] $ChangeOrderUrl
    [String] $DateChangeOrderCreated
    [String] $DateChangeOrderModified
    [String] $DateWorkflowTaskCreated
    [String] $DateWorkflowTaskModified
    [String] $Department
    [String] $EmployeeID
    [String] $EmployeeName
    [String] $EmployeeStatus
    [String] $ManagerName
    [String] $Summary
    [String] $Title
    [String] $WorkflowTaskComment
    [String] $WorkflowTaskDescription
    [String] $WorkflowTaskHandle
    [String] $WorkflowTaskUrl

    # Constructors
    ServiceDeskInfo (
        [PSObject] $Object
    ) {
        $This.ChangeOrderCategory      = $Object.ChangeOrderCategory
        $This.ChangeOrderHandle        = $Object.ChangeOrderHandle
        $This.ChangeOrderNumber        = $Object.ChangeOrderNumber
        $This.ChangeOrderUrl           = $Object.ChangeOrderUrl
        $This.DateChangeOrderCreated   = $Object.DateChangeOrderCreated
        $This.DateChangeOrderModified  = $Object.DateChangeOrderModified
        $This.DateWorkflowTaskCreated  = $Object.DateWorkflowTaskCreated
        $This.DateWorkflowTaskModified = $Object.DateWorkflowTaskModified
        $This.Department               = $Object.Department
        $This.EmployeeID               = $Object.EmployeeID
        $This.EmployeeName             = $Object.EmployeeName
        $This.EmployeeStatus           = $Object.EmployeeStatus
        $This.ManagerName              = $Object.ManagerName
        $This.Summary                  = $Object.Summary
        $This.Title                    = $Object.Title
        $This.WorkflowTaskComment      = $Object.WorkflowTaskComment
        $This.WorkflowTaskDescription  = $Object.WorkflowTaskDescription
        $This.WorkflowTaskHandle       = $Object.WorkflowTaskHandle
        $This.WorkflowTaskUrl          = $Object.WorkflowTaskUrl
    }
}

class StandardAccount {
    #Properties
    [String] $AccountType
    [String] $Department
    [String] $Description
    [String] $DisplayName
    [String] $DistinguishedName
    [String] $Domain
    [String] $EmployeeID
    [String] $EmployeeStatus
    [String] $Enabled
    [Group[]] $Groups = [System.Collections.ArrayList]::New()
    [String] $HomeDirectory
    [String] $ManagerEmailAddress = $null
    [String] $msNPAllowDialIn
    [String] ${msRTCSIP-PrimaryUserAddress}
    [String] $Name
    [String] $PasswordLastSet
    [String] $SamAccountName
    [String] $ScriptPath
    [String] $Title
    [String] $UserPrincipalName

    # Constructors
    StandardAccount (
        [PSObject] $Object
    ) {
        $This.AccountType                   = $Object.AccountType
        $This.Department                    = $Object.Department
        $This.Description                   = $Object.Description
        $This.DisplayName                   = $Object.DisplayName
        $This.DistinguishedName             = $Object.DistinguishedName
        $This.Domain                        = $Object.Domain
        $This.EmployeeID                    = $Object.EmployeeID
        $This.EmployeeStatus                = $Object.EmployeeStatus
        $This.Enabled                       = $Object.Enabled
        $This.HomeDirectory                 = $Object.HomeDirectory
        $This.msNPAllowDialIn               = $Object.msNPAllowDialIn
        $This."msRTCSIP-PrimaryUserAddress" = $Object."msRTCSIP-PrimaryUserAddress"
        $This.Name                          = $Object.Name
        $This.PasswordLastSet               = $Object.PasswordLastSet
        $This.SamAccountName                = $Object.SamAccountName
        $This.ScriptPath                    = $Object.ScriptPath
        $This.Title                         = $Object.Title
        $This.UserPrincipalName             = $Object.UserPrincipalName 
    }

    # Methods
    [Void] AddGroup (
        [PSObject] $Object
    ) {
        $This.Groups += [Group]::New($Object)
    }

    [Void] ClearHomeDirectory (
        [String] $ChangeOrder,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.HomeDirectory = Clear-HomeDirectory -Account $This -ChangeOrder $ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] ClearScriptPath (
        [String] $ChangeOrder,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.ScriptPath = Clear-ScriptPath -Account $This -ChangeOrder $ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] DisableDialInAccess (
        [String] $ChangeOrder,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.msNPAllowDialIn = Disable-DialInAccess -Account $This -ChangeOrder $ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] DisableEmployeeSIP (
        [String] $ChangeOrder,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This."msRTCSIP-PrimaryUserAddress" = Disable-EmployeeSIP -Account $This -ChangeOrder $ChangeOrder -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] MoveBatchFileToArchive (
        [String] $ChangeOrder,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        Move-BatchFileToArchive -Account $This -ChangeOrder $ChangeOrder -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] RemoveEmployeeGroups (
        [String] $ChangeOrder,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        Remove-EmployeeGroups -Account $This -ChangeOrder $ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount | ForEach-Object {
            $This.AddGroup($_)
        }
    }

    [Void] SetRandomPassword (
        [String] $ChangeOrder,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.PasswordLastSet = Set-RandomPassword -Account $This -ChangeOrder $ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] SetTerminatedOrganizationalUnit (
        [String] $ChangeOrder,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $This.DistinguishedName = Set-TerminatedOrganizationalUnit -Account $This -ChangeOrder $ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount
    }

    [Void] UpdateManagerEmailAddress (
        [String] $ManagerEmailAddress
    ) {
        $This.ManagerEmailAddress = $ManagerEmailAddress
    }
}

class Termination {
    # Properties
    [String] $ChangeOrder
    [String] $EmployeeID
    [String] $EmployeeName
    [String] $ManagerName
    [PSObject] $AccountData
    [ServiceDeskInfo[]] $ServiceDeskInfo
    [Result[]] $Results = [System.Collections.ArrayList]::New()
    [PSObject] Hidden $VerifiedStatus

    # Constructors
    Termination (
        [String] $ChangeOrder
    ) {
        $This.ChangeOrder = $ChangeOrder
        $This.AccountData = [PSCustomObject]@{
            Standard = [PSCustomObject]@{
                DOMAIN1    = $null
                DOMAIN2 = $null
                DOMAIN3    = $null
            }
            Admin    = [PSCustomObject]@{
                DOMAIN1    = $null
                DOMAIN2 = $null
                DOMAIN3    = $null
            }
            Mailbox  = [PSCustomObject]@{
                DOMAIN1 = $null
                DOMAIN3 = $null
            }
        }
        $This.VerifiedStatus = [PSCustomObject]@{
            DOMAIN1 = [PSCustomObject]@{
                Continue       = $null
                EmployeeStatus = $null
                Enabled        = $null
            }
            DOMAIN3 = [PSCustomObject]@{
                Continue       = $null
                Description    = $null
                EmployeeStatus = $null
                Enabled        = $null
            }
        }
    }

    Termination (
        [String] $ChangeOrder,
        [String] $EmployeeID
    ) {
        $This.ChangeOrder = $ChangeOrder
        $This.EmployeeID  = $EmployeeID
        $This.AccountData = [PSCustomObject]@{
            Standard = [PSCustomObject]@{
                DOMAIN1    = $null
                DOMAIN2 = $null
                DOMAIN3    = $null
            }
            Admin    = [PSCustomObject]@{
                DOMAIN1    = $null
                DOMAIN2 = $null
                DOMAIN3    = $null
            }
            Mailbox  = [PSCustomObject]@{
                DOMAIN1 = $null
                DOMAIN3 = $null
            }
        }
        $This.VerifiedStatus = [PSCustomObject]@{
            DOMAIN1 = [PSCustomObject]@{
                Continue       = $null
                EmployeeStatus = $null
                Enabled        = $null
            }
            DOMAIN3 = [PSCustomObject]@{
                Continue       = $null
                Description    = $null
                EmployeeStatus = $null
                Enabled        = $null
            }
        }
    }

    Termination (
        [String] $ChangeOrder,
        [String] $EmployeeID,
        [String] $Manager
    ) {
        $This.ChangeOrder = $ChangeOrder
        $This.EmployeeID  = $EmployeeID
        $This.ManagerName = $Manager
        $This.AccountData = [PSCustomObject]@{
            Standard = [PSCustomObject]@{
                DOMAIN1    = $null
                DOMAIN2 = $null
                DOMAIN3    = $null
            }
            Admin    = [PSCustomObject]@{
                DOMAIN1    = $null
                DOMAIN2 = $null
                DOMAIN3    = $null
            }
            Mailbox  = [PSCustomObject]@{
                DOMAIN1 = $null
                DOMAIN3 = $null
            }
        }
        $This.VerifiedStatus = [PSCustomObject]@{
            DOMAIN1 = [PSCustomObject]@{
                Continue       = $null
                EmployeeStatus = $null
                Enabled        = $null
            }
            DOMAIN3 = [PSCustomObject]@{
                Continue       = $null
                Description    = $null
                EmployeeStatus = $null
                Enabled        = $null
            }
        }
    }

    Termination (
        [String] $ChangeOrder,
        [String] $EmployeeID,
        [String] $EmployeeName,
        [String] $Manager
    ) {
        $This.ChangeOrder  = $ChangeOrder
        $This.EmployeeID   = $EmployeeID
        $This.EmployeeName = $EmployeeName
        $This.ManagerName  = $Manager
        $This.AccountData  = [PSCustomObject]@{
            Standard = [PSCustomObject]@{
                DOMAIN1    = $null
                DOMAIN2 = $null
                DOMAIN3    = $null
            }
            Admin    = [PSCustomObject]@{
                DOMAIN1    = $null
                DOMAIN2 = $null
                DOMAIN3    = $null
            }
            Mailbox  = [PSCustomObject]@{
                DOMAIN1 = $null
                DOMAIN3 = $null
            }
        }
        $This.VerifiedStatus = [PSCustomObject]@{
            DOMAIN1 = [PSCustomObject]@{
                Continue       = $null
                EmployeeStatus = $null
                Enabled        = $null
            }
            DOMAIN3 = [PSCustomObject]@{
                Continue       = $null
                Description    = $null
                EmployeeStatus = $null
                Enabled        = $null
            }
        }
    }

    Termination (
        [PSObject] $Object
    ) {
        $This.ChangeOrder  = $Object.ChangeOrderNumber
        $This.EmployeeID   = $Object.EmployeeID
        $This.EmployeeName = $Object.EmployeeName
        $This.ManagerName  = $Object.ManagerName
        $This.AccountData  = [PSCustomObject]@{
            Standard = [PSCustomObject]@{
                DOMAIN1    = $null
                DOMAIN2 = $null
                DOMAIN3    = $null
            }
            Admin    = [PSCustomObject]@{
                DOMAIN1    = $null
                DOMAIN2 = $null
                DOMAIN3    = $null
            }
            Mailbox  = [PSCustomObject]@{
                DOMAIN1 = $null
                DOMAIN3 = $null
            }
        }
        $This.ServiceDeskInfo = [ServiceDeskInfo]::New($Object)
        $This.VerifiedStatus  = [PSCustomObject]@{
            DOMAIN1 = [PSCustomObject]@{
                Continue       = $null
                EmployeeStatus = $null
                Enabled        = $null
            }
            DOMAIN3 = [PSCustomObject]@{
                Continue       = $null
                Description    = $null
                EmployeeStatus = $null
                Enabled        = $null
            }
        }
    }

    # Methods
    [Void] AddAccount (
        [PSObject] $Object
    ) {
        $Domain = $Object.Domain
        Switch ($Object.AccountType) {
            "Standard" {
                $This.AccountData.Standard.$Domain = [StandardAccount]::New($Object)
            }
            "Admin" {
                $This.AccountData.Admin.$Domain = [AdminAccount]::New($Object)
            }
            "Mailbox" {
                $This.AccountData.Mailbox.$Domain = [Mailbox]::New($Object)
            }
        }
    }

    [Void] AddResult (
        [PSObject] $Object
    ) {
        $This.Results += [Result]::New($Object)
    }

    [Void] AddServiceDeskInfo (
        [ServiceDeskInfo] $ServiceDeskInfo
    ) {
        $This.ServiceDeskInfo += $ServiceDeskInfo
    }

    [Void] GetEmployeeAccount (
        [String] $AccountType,
        [String] $ChangeOrder,
        [String] $Domain,
        [String] $SearchType,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        Switch ($AccountType) {
            "Admin" {
                Switch ($SearchType) {
                    "ID" {
                        $Account = Get-EmployeeAccount -AccountType $AccountType -ChangeOrder $ChangeOrder -Domain $Domain -EmployeeID $This.EmployeeID -TerminationCount $TerminationCount -CurrentCount $CurrentCount
                        if ($null -ne $Account) {
                            $This.AddAccount($Account)
                        }
                    }
                    "Name" {
                        $Account = Get-EmployeeAccount -AccountType $AccountType -ChangeOrder $ChangeOrder -Domain $Domain -EmployeeName $This.EmployeeName -TerminationCount $TerminationCount -CurrentCount $CurrentCount
                        if ($null -ne $Account) {
                            $This.AddAccount($Account)
                        }
                    }
                }
            }
            "Manager" {
                if ($null -ne $This.AccountData.Standard.$Domain) {
                    if ($This.ManagerName -ne "" -and $null -ne $This.ManagerName) {
                        $Account = Get-EmployeeAccount -AccountType $AccountType -ChangeOrder $ChangeOrder -Domain $Domain -EmployeeName $This.ManagerName -TerminationCount $TerminationCount -CurrentCount $CurrentCount
                        if ($null -ne $This.AccountData.Standard.$Domain) {
                            if ($null -ne $Account) {
                                $This.AccountData.Standard.$Domain.UpdateManagerEmailAddress($Account)
                            }
                            else {
                                $This.AccountData.Standard.$Domain.UpdateManagerEmailAddress("None")
                            }
                        }
                    }
                    else {
                        $This.AccountData.Standard.$Domain.UpdateManagerEmailAddress("None")
                    }
                }
            }
            "Standard" {
                Switch ($SearchType) {
                    "ID" {
                        $Account = Get-EmployeeAccount -AccountType $AccountType -ChangeOrder $ChangeOrder -Domain $Domain -EmployeeID $This.EmployeeID -TerminationCount $TerminationCount -CurrentCount $CurrentCount
                        if ($null -ne $Account) {
                            $This.AddAccount($Account)
                        }
                    }
                    "Name" {
                        $Account = Get-EmployeeAccount -AccountType $AccountType -ChangeOrder $ChangeOrder -Domain $Domain -EmployeeName $This.EmployeeName -TerminationCount $TerminationCount -CurrentCount $CurrentCount
                        if ($null -ne $Account) {
                            $This.AddAccount($Account)
                        }
                    }
                }
            }
        }
    }

    [Void] GetEmployeeMailbox (
        [String] $ChangeOrder,
        [String] $Domain,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $Mailbox = Get-EmployeeMailbox -Account $This.AccountData.Standard.$Domain -ChangeOrder $ChangeOrder -TerminationCount $TerminationCount -CurrentCount $CurrentCount
        if ($null -ne $Mailbox) {
            $This.AddAccount($Mailbox)
        }
    }

    [Void] GetServiceDeskInfo (
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        $Info = Get-ServiceDeskInfo -ChangeOrder $This.ChangeOrder -Credential $Credential -TerminationCount $TerminationCount -CurrentCount $CurrentCount
        if (@($Info).Count -gt 1) {
            $This.ServiceDeskInfo = [System.Collections.ArrayList]::New()
            $Info | ForEach-Object {
                $This.ServiceDeskInfo += [ServiceDeskInfo]::New($_)
            }
        }
        else {
            $This.ServiceDeskInfo = [ServiceDeskInfo]::New($Info)
        }
    }

    [Void] SetServiceDeskTaskCompleted (
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        if ($This.ServiceDeskInfo.WorkflowTaskHandle.Count -gt 1) {
            $This.ServiceDeskInfo.WorkflowTaskHandle | ForEach-Object {
                Set-ServiceDeskTaskCompleted -Credential $Credential -WorkflowTaskHandle $_ -TerminationCount $TerminationCount -CurrentCount $CurrentCount
            }
        }
        else {
            Set-ServiceDeskTaskCompleted -Credential $Credential -WorkflowTaskHandle $This.ServiceDeskInfo.WorkflowTaskHandle -TerminationCount $TerminationCount -CurrentCount $CurrentCount
        }
    }

    [Void] SetServiceDeskTaskUpdated (
        [String] $Comment,
        [PSCredential] $Credential,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        if ($This.ServiceDeskInfo.Count -gt 1) {
            $This.ServiceDeskInfo | ForEach-Object {
                Set-ServiceDeskTaskUpdated -Comment $Comment -Credential $Credential -WorkflowTaskHandle $_.WorkflowTaskHandle -TerminationCount $TerminationCount -CurrentCount $CurrentCount
            }
        }
        else {
            Set-ServiceDeskTaskUpdated -Comment $Comment -Credential $Credential -WorkflowTaskHandle $This.ServiceDeskInfo.WorkflowTaskHandle -TerminationCount $TerminationCount -CurrentCount $CurrentCount
        }
    }

    [Void] TestAccountStatus (
        [String] $ChangeOrder,
        [String] $Domain,
        [Int32] $TerminationCount,
        [Int32] $CurrentCount
    ) {
        Switch ($Domain) {
            "DOMAIN1" {
                $Continue, $EmployeeStatus, $Enabled = Test-AccountStatus -Account $This.AccountData.Standard.$Domain -ChangeOrder $ChangeOrder -TerminationCount $TerminationCount -CurrentCount $CurrentCount
                $This.VerifiedStatus.$Domain.Continue = $Continue
                $This.VerifiedStatus.$Domain.EmployeeStatus = $EmployeeStatus
                $This.VerifiedStatus.$Domain.Enabled = $Enabled
            }
            "DOMAIN3" {
                $Continue, $Description, $EmployeeStatus, $Enabled = Test-AccountStatus -Account $This.AccountData.Standard.$Domain -ChangeOrder $ChangeOrder -TerminationCount $TerminationCount -CurrentCount $CurrentCount
                $This.VerifiedStatus.$Domain.Continue = $Continue
                $This.VerifiedStatus.$Domain.Description = $Description
                $This.VerifiedStatus.$Domain.EmployeeStatus = $EmployeeStatus
                $This.VerifiedStatus.$Domain.Enabled = $Enabled
            }
        }
    }

    [Void] UpdateEmployeeID (
        [String] $EmployeeID
    ) {
        $This.EmployeeID = $EmployeeID
    }

    [Void] UpdateEmployeeName (
        [String] $EmployeeName
    ) {
        $This.EmployeeName = $EmployeeName
    }

    [Void] UpdateManagerName (
        [String] $ManagerName
    ) {
        $This.ManagerName = $ManagerName
    }
}