<#
.SYNOPSIS
Start employee termination

.DESCRIPTION
Start termination process for single or multiple employees

.PARAMETER ChangeOrder
Change Order number for employee termination

.PARAMETER File
FilePath of CSV to import for batch terminations

.PARAMETER Credentials
Credential object of all required credentials for convenience if user prefers to provide credentials as a parameter

.PARAMETER ServiceDesk
Switch parameter to indicate pulling termination information directly from ServiceDesk

.EXAMPLE
Start-EmployeeTermination "123456"

.EXAMPLE
Start-EmployeeTermination -ChangeOrder "123456" -Credentials $Credentials

.EXAMPLE
Start-EmployeeTermination -File "C:\FileName.csv"

.EXAMPLE
Start-EmployeeTermination -ServiceDesk -Credentials $Credentials

.INPUTS
System.String
System.Object

.OUTPUTS
System.String

.NOTES
Author: Joshua Chase
Written: 30 Apr 2019
#>

Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
Add-Type -Path $(${env:ProgramFiles(x86)} + '\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll')
Import-Module "C:\Program Files\WindowsPowerShell\Modules\SPOMod\SPOMod.psm1"
Import-Module -Name ReportHTML

Function Start-EmployeeTermination {

    [CmdletBinding(DefaultParameterSetName = "Single")]

    Param(

        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = "Single")]
        [ValidateNotNullOrEmpty()]
        [String]
        $ChangeOrder,

        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "File")]
        [ValidateScript({Test-Path -Path $_})]
        [String]
        $File,

        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [PSObject]
        $Credentials,

        [Parameter(Mandatory = $true, ParameterSetName = "ServiceDesk")]
        [Switch]
        $ServiceDesk
    )

    # Set UI variables
    $Host.PrivateData.VerboseForegroundColor = "DarkGray"
    $ErrorActionPreference = "Stop"
    $InformationPreference = "Continue"
    $ProgressPreference = "Continue"
    $WarningPreference = "SilentlyContinue"

    Switch ($PSCmdlet.ParameterSetName) {
        "Single" {
            $global:Main = [Main]::New($ChangeOrder)
            $Domains = "DOMAIN1", "DOMAIN2", "DOMAIN3", "O365", "SKYPE"
            $Section = "PreparationAndSetup"
            $Step = "Main"

            $Log.Account = $env:USERNAME
            $Log.AccountType = "Developer"
            $Log.ChangeOrder = "Single"
            $Log.DisplayName = $env:USERNAME -Replace "\.", " "
            $Log.Domain = "All"
            $Log.EmployeeID = "N/A"
            $Log.Section = $Section
            $Log.Status = "Start"
            $Log.Step = $Step
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $TerminationCount = @($Main.Terminations).Count
            $CurrentCount = 1

            if ($Credentials) {
                Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                $Continue = [System.Management.Automation.Host.ChoiceDescription]::New("&Continue", "Continue with provided credentials")
                $Provide = [System.Management.Automation.Host.ChoiceDescription]::New("&Provide", "Provide new credentials")
                $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($Continue, $Provide)
                $Choice = $Host.UI.PromptForChoice("Select Option", "Provide new credentials?", $Choices, 0)
                Switch ($Choice) {
                    1 {
                        ForEach ($Domain in $Domains) {
                            $Main.GetTerminationCredential($Domain, $TerminationCount, $CurrentCount)
                        }
                    }
                    0 {
                        Write-Information "Continuing..."
                        $Main.AddCredentials($Credentials)
                    }
                }
            }
            else {
                Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                ForEach ($Domain in $Domains) {
                    $Main.GetTerminationCredential($Domain, $TerminationCount, $CurrentCount)
                }
            }

            $Domains | Where-Object { $_ -ne "SKYPE" } | ForEach-Object {
                $Attempts = 0
                $Retry = $true
                do {
                    $Main.TestUserAccountAccess($_, $TerminationCount, $CurrentCount)
                    if ($Main.VerifiedAccess.$_ -ne $true) {
                        
                        Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                        Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount
                        
                        $Provide = [System.Management.Automation.Host.ChoiceDescription]::New("&Provide", "Provide $_ credentials")
                        $Exit = [System.Management.Automation.Host.ChoiceDescription]::New("&Exit", "Exit")
                        $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($Exit, $Provide)
                        $Choice = $Host.UI.PromptForChoice("Unable to verify access for $_", "Provide alternate credentials?", $Choices, 0)
                        Switch ($Choice) {
                            1 {
                                $Main.GetTerminationCredential($_, $TerminationCount, $CurrentCount)
                                $Attempts++
                                if ($_ -eq "O365") {
                                    $Main.GetTerminationCredential("SKYPE", $TerminationCount, $CurrentCount)
                                }
                            }
                            0 {
                                $Retry = $false
                            }
                        }
                    }
                } while ($Main.VerifiedAccess.$_ -ne $true -and $Attempts -lt 3 -and $Retry -eq $true)
            }
            $Domains | Where-Object { $_ -ne "SKYPE" } | ForEach-Object {
                if ($Main.VerifiedAccess.$_ -ne $true) {
                    break
                }
            }

            $Main.GetLogFile($TerminationCount, $CurrentCount)
            if ($null -eq $Main.LogFilePath) {
                break
            }

            $Main.GetErrorLogFile($TerminationCount, $CurrentCount)
            if ($null -eq $Main.ErrorLogFilePath) {
                break
            }

            $Main.GetGroupFile($TerminationCount, $CurrentCount)
            if ($null -eq $Main.GroupFilePath) {
                break
            }

            $Main.GetResultFile($TerminationCount, $CurrentCount)
            if ($null -eq $Main.ResultFilePath) {
                break
            }

            $SessionType = "DOMAIN1"
            $Main.NewTerminationSession($SessionType, $TerminationCount, $CurrentCount)
            $Main.GetTerminatedDatabase($TerminationCount, $CurrentCount)
            if ($null -eq $Main.Database) {
                break
            }
            $Main.RemoveTerminationSession($SessionType, $TerminationCount, $CurrentCount)

            $Log.Domain = "All"
            $Log.Step = "Main"
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            Write-StepProgress -Complete

            $Section = "PreProcessingTasks"
            $Step = "Main"

            $Log.Section = $Section
            $Log.Status = "Start"
            $Log.Step = $Step
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
            Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

            $Termination = $Main.Terminations

            $Termination.GetServiceDeskInfo($Main.Credential.ServiceDesk, $TerminationCount, $CurrentCount)
            $Termination.UpdateEmployeeID($Termination.ServiceDeskInfo.EmployeeID)
            $Termination.UpdateEmployeeName($Termination.ServiceDeskInfo.EmployeeName)
            $Termination.UpdateManagerName($Termination.ServiceDeskInfo.ManagerName)

            $AccountTypes = "Standard", "Admin", "Manager"
            $Domains = "DOMAIN1", "DOMAIN2", "DOMAIN3"

            ForEach ($AccountType in $AccountTypes) {
                if ($AccountType -ne "Manager") {
                    ForEach ($Domain in $Domains) {
                        $Termination.GetEmployeeAccount($AccountType, $ChangeOrder, $Domain, "ID", $TerminationCount, $CurrentCount)
                    }
                }
                else {
                    ForEach ($Domain in $Domains) {
                        if ($Domain -ne "DOMAIN2") {
                            $Termination.GetEmployeeAccount($AccountType, $ChangeOrder, $Domain, "Name", $TerminationCount, $CurrentCount)
                        }
                    }
                }
            }

            $AccountTypes = "Standard", "Admin"

            ForEach ($AccountType in $AccountTypes) {
                ForEach ($Domain in $Domains) {
                    if ($Termination.AccountData.$AccountType.$Domain.SamAccountName -eq "") {
                        $Termination.AccountData.$AccountType.$Domain = $null
                    }
                }
            }

            $AccountType = "Standard"
            $Domains = "DOMAIN1", "DOMAIN3"

            $Log.Status = "Start"
            $Log.Step = "Test-AccountStatus"
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            ForEach ($Domain in $Domains) {
                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Termination.TestAccountStatus($ChangeOrder, $Domain, $TerminationCount, $CurrentCount)
                    if ($Termination.VerifiedStatus.$Domain.Continue -ne $true) {
                        break
                    }
                }
            }

            Write-StepProgress -Complete

            $Log.Account = $env:USERNAME
            $Log.AccountType = "Developer"
            $Log.ChangeOrder = "Single"
            $Log.DisplayName = $env:USERNAME -Replace "\.", " "
            $Log.Domain = "All"
            $Log.EmployeeID = "N/A"
            $Log.Status = "Complete"
            $Log.Step = "Main"
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $Domain = "DOMAIN1"

            if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                $Account = $Termination.AccountData.$AccountType.$Domain

                $Section = "StandardActiveDirectoryDOMAIN1"

                $Log.Account = $Account.SamAccountName
                $Log.AccountType = $Account.AccountType
                $Log.ChangeOrder = $ChangeOrder
                $Log.DisplayName = $Account.DisplayName
                $Log.Domain = $Account.Domain
                $Log.EmployeeID = $Account.EmployeeID
                $Log.Section = $Section
                $Log.Status = "Start"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Section = $Section

                Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                $Account.ClearHomeDirectory($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.ClearScriptPath($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.DisableDialInAccess($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.MoveBatchFileToArchive($ChangeOrder, $TerminationCount, $CurrentCount)
                $Account.RemoveEmployeeGroups($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetRandomPassword($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetTerminatedOrganizationalUnit($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)

                $Log.Status = "Complete"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Status = "Complete"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-StepProgress -Complete

                $Section = "DisableEmployeeSIP"
                $SessionType = "SKYPE"

                $Log.Section = $Section
                $Log.Status = "Start"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Section = $Section
                $Result.Status = "Start"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                $Main.NewTerminationSession($SessionType, $TerminationCount, $CurrentCount)
                $Account.DisableEmployeeSIP($ChangeOrder, $TerminationCount, $CurrentCount)
                $Main.RemoveTerminationSession($SessionType, $TerminationCount, $CurrentCount)

                $Log.Account = $Account.SamAccountName
                $Log.AccountType = $Account.AccountType
                $Log.ChangeOrder = $ChangeOrder
                $Log.DisplayName = $Account.DisplayName
                $Log.Domain = $Account.Domain
                $Log.EmployeeID = $Account.EmployeeID
                $Log.Status = "Complete"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Account = $Account.SamAccountName
                $Result.AccountType = $Account.AccountType
                $Result.ChangeOrder = $ChangeOrder
                $Result.DisplayName = $Account.DisplayName
                $Result.Domain = $Account.Domain
                $Result.EmployeeID = $Account.EmployeeID
                $Result.Status = "Complete"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-StepProgress -Complete

                $AccountType = "Mailbox"
                $Section = "ExchangeDOMAIN1"
                $SessionType = "DOMAIN1"

                $Log.AccountType = $AccountType
                $Log.Section = $Section
                $Log.Status = "Start"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.AccountType = $AccountType
                $Result.Section = $Section
                $Result.Status = "Start"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                $Main.NewTerminationSession($SessionType, $TerminationCount, $CurrentCount)
                $Termination.GetEmployeeMailbox($ChangeOrder, $Domain, $TerminationCount, $CurrentCount)
                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Mailbox = $Termination.AccountData.$AccountType.$Domain

                    $Mailbox.EnableMailboxForwarding($ChangeOrder, $Termination.AccountData.Standard.$Domain.ManagerEmailAddress, $TerminationCount, $CurrentCount)
                    $Mailbox.MoveMailboxToTerminatedDatabase($ChangeOrder, $Main.Database, $TerminationCount, $CurrentCount)
                    $Mailbox.RemoveInboxRules($ChangeOrder, $TerminationCount, $CurrentCount)
                    $Mailbox.SetContactHiddenFromGAL($ChangeOrder, $TerminationCount, $CurrentCount)
                    $Mailbox.SetMailboxHiddenFromGAL($ChangeOrder, $TerminationCount, $CurrentCount)
                    $Mailbox.SetTerminatedOutOfOffice($ChangeOrder, $TerminationCount, $CurrentCount)

                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Mailbox.AccountType
                    $Log.ChangeOrder = $ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Mailbox.AccountType
                    $Result.ChangeOrder = $ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)
                }
                $Main.RemoveTerminationSession($SessionType, $TerminationCount, $CurrentCount)

                $Log.Account = $Account.SamAccountName
                $Log.AccountType = $Account.AccountType
                $Log.ChangeOrder = $ChangeOrder
                $Log.DisplayName = $Account.DisplayName
                $Log.EmployeeID = $Account.EmployeeID
                $Log.Section = "StandardActiveDirectoryDOMAIN1"
                $Log.Status = "Complete"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Account = $Account.SamAccountName
                $Result.AccountType = $Account.AccountType
                $Result.ChangeOrder = $ChangeOrder
                $Result.DisplayName = $Account.DisplayName
                $Result.EmployeeID = $Account.EmployeeID
                $Result.Section = "StandardActiveDirectoryDOMAIN1"
                $Result.Status = "Complete"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-StepProgress -Complete
            }

            $AccountType = "Admin"

            if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                $Account = $Termination.AccountData.$AccountType.$Domain

                $Section = "AdminActiveDirectoryDOMAIN1"

                $Log.Account = $Account.SamAccountName
                $Log.AccountType = $Account.AccountType
                $Log.DisplayName = $Account.DisplayName
                $Log.Section = $Section
                $Log.Status = "Start"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Account = $Account.SamAccountName
                $Result.AccountType = $Account.AccountType
                $Result.DisplayName = $Account.DisplayName
                $Result.Section = $Section
                $Result.Status = "Start"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                $Account.DisableDialInAccess($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.RemoveEmployeeGroups($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.RemoveUnixAttributes($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetRandomPassword($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetTerminatedOrganizationalUnit($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)

                $Log.Status = "Complete"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Status = "Complete"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-StepProgress -Complete
            }

            $AccountType = "Standard"
            $Domain = "DOMAIN2"

            if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                $Account = $Termination.AccountData.$AccountType.$Domain

                $Section = "StandardActiveDirectoryDOMAIN2"

                $Log.Account = $Account.SamAccountName
                $Log.AccountType = $Account.AccountType
                $Log.DisplayName = $Account.DisplayName
                $Log.Domain = $Account.Domain
                $Log.Section = $Section
                $Log.Status = "Start"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Account = $Account.SamAccountName
                $Result.AccountType = $Account.AccountType
                $Result.DisplayName = $Account.DisplayName
                $Result.Domain = $Account.Domain
                $Result.Section = $Section
                $Result.Status = "Start"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                $Account.DisableDialInAccess($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.RemoveEmployeeGroups($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetRandomPassword($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetTerminatedOrganizationalUnit($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)

                $Log.Status = "Complete"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Status = "Complete"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-StepProgress -Complete
            }

            $AccountType = "Admin"

            if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                $Account = $Termination.AccountData.$AccountType.$Domain

                $Section = "AdminActiveDirectoryDOMAIN2"

                $Log.Account = $Account.SamAccountName
                $Log.AccountType = $Account.AccountType
                $Log.DisplayName = $Account.DisplayName
                $Log.Section = $Section
                $Log.Status = "Start"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Account = $Account.SamAccountName
                $Result.AccountType = $Account.AccountType
                $Result.DisplayName = $Account.DisplayName
                $Result.Section = $Section
                $Result.Status = "Start"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                $Account.DisableDialInAccess($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.RemoveEmployeeGroups($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetRandomPassword($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetTerminatedOrganizationalUnit($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)

                $Log.Status = "Complete"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Status = "Complete"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-StepProgress -Complete
            }

            $AccountType = "Standard"
            $Domain = "DOMAIN3"

            if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                $Account = $Termination.AccountData.$AccountType.$Domain

                $Section = "StandardActiveDirectoryDOMAIN3"

                $Log.Account = $Account.SamAccountName
                $Log.AccountType = $Account.AccountType
                $Log.DisplayName = $Account.DisplayName
                $Log.Domain = $Account.Domain
                $Log.Section = $Section
                $Log.Status = "Start"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Account = $Account.SamAccountName
                $Result.AccountType = $Account.AccountType
                $Result.DisplayName = $Account.DisplayName
                $Result.Domain = $Account.Domain
                $Result.Section = $Section
                $Result.Status = "Start"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                $Account.DisableDialInAccess($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.RemoveEmployeeGroups($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetRandomPassword($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetTerminatedOrganizationalUnit($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)

                $Log.Status = "Complete"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Status = "Complete"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-StepProgress -Complete

                $AccountType = "Mailbox"
                $Domain = "DOMAIN3"
                $Section = "ExchangeDOMAIN3"
                $SessionType = "O365"

                $Log.AccountType = $AccountType
                $Log.Domain = $Domain
                $Log.Section = $Section
                $Log.Status = "Start"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.AccountType = $AccountType
                $Result.Domain = $Domain
                $Result.Section = $Section
                $Result.Status = "Start"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                $Main.NewTerminationSession($SessionType, $TerminationCount, $CurrentCount)
                $Termination.GetEmployeeMailbox($ChangeOrder, $Domain)
                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Mailbox = $Termination.AccountData.$AccountType.$Domain

                    $Mailbox.EnableMailboxForwarding($ChangeOrder, $Termination.AccountData.Standard.$Domain.ManagerEmailAddress, $TerminationCount, $CurrentCount)
                    $Mailbox.SetTerminatedOutOfOffice($ChangeOrder, $TerminationCount, $CurrentCount)

                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Mailbox.AccountType
                    $Log.ChangeOrder = $ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Mailbox.AccountType
                    $Result.ChangeOrder = $ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)
                }
                $Main.RemoveTerminationSession($SessionType, $TerminationCount, $CurrentCount)

                $Log.Account = $Account.SamAccountName
                $Log.AccountType = $Account.AccountType
                $Log.ChangeOrder = $ChangeOrder
                $Log.DisplayName = $Account.DisplayName
                $Log.Domain = $Account.Domain
                $Log.EmployeeID = $Account.EmployeeID
                $Log.Section = "StandardActiveDirectoryDOMAIN3"
                $Log.Status = "Complete"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Account = $Account.SamAccountName
                $Result.AccountType = $Mailbox.AccountType
                $Result.ChangeOrder = $ChangeOrder
                $Result.DisplayName = $Account.DisplayName
                $Result.Domain = $Account.Domain
                $Result.EmployeeID = $Account.EmployeeID
                $Result.Section = "StandardActiveDirectoryDOMAIN3"
                $Result.Status = "Complete"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-StepProgress -Complete
            }

            $AccountType = "Admin"

            if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                $Account = $Termination.AccountData.$AccountType.$Domain

                $Section = "AdminActiveDirectoryDOMAIN3"

                $Log.Account = $Account.SamAccountName
                $Log.AccountType = $Account.AccountType
                $Log.DisplayName = $Account.DisplayName
                $Log.Section = $Section
                $Log.Status = "Start"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Account = $Account.SamAccountName
                $Result.AccountType = $Account.AccountType
                $Result.DisplayName = $Account.DisplayName
                $Result.Section = $Section
                $Result.Status = "Start"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                $Account.DisableDialInAccess($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.RemoveEmployeeGroups($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetRandomPassword($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                $Account.SetTerminatedOrganizationalUnit($ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)

                $Log.Status = "Complete"
                $Log.Step = "Main"
                $Log.TimeStamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                $Result.Status = "Complete"
                $Result.Step = "Main"
                $Result.TimeStamp = Get-Date -Format FileDateTime

                $Termination.AddResult($Result)

                Write-StepProgress -Complete
            }
            $Section = "PostProcessingTasks"
            $Step = "WriteLogs"
            
            $Log.Account = $env:USERNAME
            $Log.AccountType = "Developer"
            $Log.ChangeOrder = "Single"
            $Log.DisplayName = $env:USERNAME -Replace "\.", " "
            $Log.Domain = "All"
            $Log.EmployeeID = "N/A"
            $Log.Section = $Section
            $Log.Status = "Complete"
            $Log.Step = $Step
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
            Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

            if ($Termination.ServiceDeskInfo.WorkflowTaskComment -eq "FIM processing has been completed." -and "Failed" -notin $Termination.Result.Status) {
                $Termination.SetServiceDeskTaskCompleted($Main.Credential.ServiceDesk, $TerminationCount, $CurrentCount)
            }
            else {
                $TaskComment = "Needs FIM processing."
                $Termination.SetServiceDeskTaskUpdated($TaskComment, $Main.Credential.ServiceDesk, $TerminationCount, $CurrentCount)
            }

            Write-StepProgress -Complete
            Write-SectionProgress -Complete

            $Main.WriteErrorLogsToFile()
            $Main.WriteGroupsToFile()
            $Main.WriteLogsToFile()
            $Main.WriteResultsToFile()
            $Main.CopyItemsToSharePoint($Main.Credential.O365)
            $Main.NewTeamsTerminationNotification()

            Write-TerminationProgress -Complete
        }

        "File" {
            $TerminationsFromFile = Get-Content -Path $File
            $global:Main = [Main]::New()
            $ChangeOrder = "MultipleFromFile"
            $Domains = "DOMAIN1", "DOMAIN2", "DOMAIN3", "O365", "SKYPE"
            $Section = "PreparationAndSetup"
            $Step = "Main"

            $Log.Account = $env:USERNAME
            $Log.AccountType = "Developer"
            $Log.ChangeOrder = $ChangeOrder
            $Log.DisplayName = $env:USERNAME -Replace "\.", " "
            $Log.Domain = "All"
            $Log.EmployeeID = "N/A"
            $Log.Section = $Section
            $Log.Status = "Start"
            $Log.Step = $Step
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $TerminationCount = 1
            $CurrentCount = 0

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
            Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

            if ($Credentials) {
                $Continue = [System.Management.Automation.Host.ChoiceDescription]::New("&Continue", "Continue with provided credentials")
                $Provide = [System.Management.Automation.Host.ChoiceDescription]::New("&Provide", "Provide new credentials")
                $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($Continue, $Provide)
                $Choice = $Host.UI.PromptForChoice("Select Option", "Provide new credentials?", $Choices, 0)
                Switch ($Choice) {
                    1 {
                        ForEach ($Domain in $Domains) {
                            $Main.GetTerminationCredential($Domain, $TerminationCount, $CurrentCount)
                        }
                    }
                    0 {
                        Write-Information "Continuing..."
                        $Main.AddCredentials($Credentials)
                    }
                }
            }
            else {
                ForEach ($Domain in $Domains) {
                    $Main.GetTerminationCredential($Domain, $TerminationCount, $CurrentCount)
                }
            }
            $Domains | Where-Object { $_ -ne "SKYPE" } | ForEach-Object {
                $Attempts = 0
                $Retry = $true
                do {
                    $Main.TestUserAccountAccess($_, $TerminationCount, $CurrentCount)
                    if ($Main.VerifiedAccess.$_ -ne $true) {
                        $Provide = [System.Management.Automation.Host.ChoiceDescription]::New("&Provide", "Provide $_ credentials")
                        $Exit = [System.Management.Automation.Host.ChoiceDescription]::New("&Exit", "Exit")
                        $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($Exit, $Provide)
                        $Choice = $Host.UI.PromptForChoice("Unable to verify access for $_", "Provide alternate credentials?", $Choices, 0)
                        Switch ($Choice) {
                            1 {
                                $Main.GetTerminationCredential($_, $TerminationCount, $CurrentCount)
                                $Attempts++
                                if ($_ -eq "O365") {
                                    $Main.GetTerminationCredential("SKYPE", $TerminationCount, $CurrentCount)
                                }
                            }
                            0 {
                                $Retry = $false
                            }
                        }
                    }
                } while ($Main.VerifiedAccess.$_ -ne $true -and $Attempts -lt 3 -and $Retry -eq $true)
            }
            $Domains | Where-Object { $_ -ne "SKYPE" } | ForEach-Object {
                if ($Main.VerifiedAccess.$_ -ne $true) {
                    break
                }
            }
            $Main.GetLogFile($TerminationCount, $CurrentCount)
            if ($null -eq $Main.LogFilePath) {
                break
            }
            $Main.GetErrorLogFile($TerminationCount, $CurrentCount)
            if ($null -eq $Main.ErrorLogFilePath) {
                break
            }
            $Main.GetGroupFile($TerminationCount, $CurrentCount)
            if ($null -eq $Main.GroupFilePath) {
                break
            }
            $Main.GetResultFile($TerminationCount, $CurrentCount)
            if ($null -eq $Main.ResultFilePath) {
                break
            }
            $SessionType = "DOMAIN1"
            $Main.NewTerminationSession($SessionType, $TerminationCount, $CurrentCount)
            $Main.GetTerminatedDatabase($TerminationCount, $CurrentCount)
            if ($null -eq $Main.Database) {
                break
            }
            $Main.RemoveTerminationSession($SessionType, $TerminationCount, $CurrentCount)

            $Log.Domain = "All"
            $Log.Step = "Main"
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $Section = "PreProcessingTasks"
            $Step = "Main"

            $Log.Section = $Section
            $Log.Status = "Start"
            $Log.Step = $Step
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $TerminationsFromFile | ForEach-Object {
                $Main.AddTermination($_)
            }

            $TerminationCount = $Main.Terminations.Count
            $CurrentCount = 1

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section

            $Main.Terminations | ForEach-Object {
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount
                $Termination = $_
                $Termination.GetServiceDeskInfo($Main.Credential.ServiceDesk, $TerminationCount, $CurrentCount)
                $Termination.UpdateEmployeeID($Termination.ServiceDeskInfo.EmployeeID)
                $Termination.UpdateEmployeeName($Termination.ServiceDeskInfo.EmployeeName)
                $Termination.UpdateManagerName($Termination.ServiceDeskInfo.ManagerName)
                $CurrentCount++
            }

            $AccountTypes = "Standard", "Admin", "Manager"
            $Domains = "DOMAIN1", "DOMAIN2", "DOMAIN3"

            $CurrentCount = 1

            $Main.Terminations | ForEach-Object {
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount
                $Termination = $_
                ForEach ($AccountType in $AccountTypes) {
                    if ($AccountType -ne "Manager") {
                        ForEach ($Domain in $Domains) {
                            $Termination.GetEmployeeAccount($AccountType, $Termination.ChangeOrder, $Domain, "ID", $TerminationCount, $CurrentCount)
                        }
                    }
                    else {
                        ForEach ($Domain in $Domains) {
                            if ($Domain -ne "DOMAIN2") {
                                $Termination.GetEmployeeAccount($AccountType, $Termination.ChangeOrder, $Domain, "Name", $TerminationCount, $CurrentCount)
                            }
                        }
                    }
                }
                $CurrentCount++
            }

            $AccountTypes = "Standard", "Admin"

            $Main.Terminations | ForEach-Object {
                $Termination = $_
                ForEach ($AccountType in $AccountTypes) {
                    ForEach ($Domain in $Domains) {
                        if ($Termination.AccountData.$AccountType.$Domain.SamAccountName -eq "") {
                            $Termination.AccountData.$AccountType.$Domain = $null
                        }
                    }
                }
            }

            $Log.Account = $env:USERNAME
            $Log.AccountType = "Developer"
            $Log.ChangeOrder = $ChangeOrder
            $Log.DisplayName = $env:USERNAME -Replace "\.", " "
            $Log.Domain = "All"
            $Log.EmployeeID = "N/A"
            $Log.Status = "Complete"
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $AccountType = "Standard"
            $Domains = "DOMAIN1", "DOMAIN3"

            $Log.Status = "Start"
            $Log.Step = "Test-AccountStatus"
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $CurrentCount = 1

            $Main.Terminations | ForEach-Object {
                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount
                $Termination = $_
                ForEach ($Domain in $Domains) {
                    if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                        $Termination.TestAccountStatus($Termination.ChangeOrder, $Domain, $TerminationCount, $CurrentCount)
                        if ($Termination.VerifiedStatus.$Domain.Continue -ne $true) {
                            $Termination.AccountData = $null
                            $CurrentCount--
                            $TerminationCount--
                        }
                    }
                }
                $CurrentCount++
            }

            $Log.Account = $env:USERNAME
            $Log.AccountType = "Developer"
            $Log.ChangeOrder = $ChangeOrder
            $Log.DisplayName = $env:USERNAME -Replace "\.", " "
            $Log.Domain = "All"
            $Log.EmployeeID = "N/A"
            $Log.Status = "Complete"
            $Log.Step = "Main"
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $Domain = "DOMAIN1"
            $Section = "StandardActiveDirectoryDOMAIN1"

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section

            $CurrentCount = 1

            $Main.Terminations | ForEach-Object {
                $Termination = $_

                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain

                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)

                    $Account.ClearHomeDirectory($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.ClearScriptPath($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.MoveBatchFileToArchive($Termination.ChangeOrder, $TerminationCount, $CurrentCount)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)

                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)
                }
                $CurrentCount++
            }

            $Section = "DisableEmployeeSIP"
            $SessionType = "SKYPE"

            $Log.Section = $Section

            $CurrentCount = 1

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
            Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

            $Main.NewTerminationSession($SessionType, $TerminationCount, $CurrentCount)

            $Main.Terminations | ForEach-Object {
                $Termination = $_

                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain

                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)

                    $Account.DisableEmployeeSIP($Termination.ChangeOrder, $TerminationCount, $CurrentCount)

                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)
                }
                $CurrentCount++
            }

            $Main.RemoveTerminationSession($SessionType, $TerminationCount, $CurrentCount)

            $AccountType = "Mailbox"
            $Section = "ExchangeDOMAIN1"
            $SessionType = "DOMAIN1"
            $Step = "Main"

            $Log.Section = $Section

            $CurrentCount = 1

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
            Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

            $Main.NewTerminationSession($SessionType, $TerminationCount, $CurrentCount)

            $Main.Terminations | ForEach-Object {
                $Termination = $_

                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                if ($null -ne $Termination.AccountData.Standard.$Domain) {
                    $Account = $Termination.AccountData.Standard.$Domain

                    $Termination.GetEmployeeMailbox($Termination.ChangeOrder, $Domain, $TerminationCount, $CurrentCount)
                    if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                        $Mailbox = $Termination.AccountData.$AccountType.$Domain

                        $Log.Account = $Account.SamAccountName
                        $Log.AccountType = $Mailbox.AccountType
                        $Log.ChangeOrder = $Termination.ChangeOrder
                        $Log.DisplayName = $Account.DisplayName
                        $Log.Domain = $Account.Domain
                        $Log.EmployeeID = $Account.EmployeeID
                        $Log.Section = $Section
                        $Log.Status = "Start"
                        $Log.Step = $Step
                        $Log.TimeStamp = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        $Result.Account = $Account.SamAccountName
                        $Result.AccountType = $Mailbox.AccountType
                        $Result.ChangeOrder = $Termination.ChangeOrder
                        $Result.DisplayName = $Account.DisplayName
                        $Result.Domain = $Account.Domain
                        $Result.EmployeeID = $Account.EmployeeID
                        $Result.Section = $Section
                        $Result.Status = "Start"
                        $Result.Step = $Step
                        $Result.TimeStamp = Get-Date -Format FileDateTime

                        $Termination.AddResult($Result)

                        $Mailbox.EnableMailboxForwarding($Termination.ChangeOrder, $Account.ManagerEmailAddress, $TerminationCount, $CurrentCount)
                        $Mailbox.MoveMailboxToTerminatedDatabase($Termination.ChangeOrder, $Main.Database, $TerminationCount, $CurrentCount)
                        $Main.UpdateDatabaseSize($DatabaseUpdate)
                        $Mailbox.RemoveInboxRules($Termination.ChangeOrder, $TerminationCount, $CurrentCount)
                        $Mailbox.SetContactHiddenFromGAL($Termination.ChangeOrder, $TerminationCount, $CurrentCount)
                        $Mailbox.SetMailboxHiddenFromGAL($Termination.ChangeOrder, $TerminationCount, $CurrentCount)
                        $Mailbox.SetTerminatedOutOfOffice($Termination.ChangeOrder, $TerminationCount, $CurrentCount)

                        $Log.Status = "Complete"
                        $Log.Step = "Main"
                        $Log.TimeStamp = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        $Result.Status = "Complete"
                        $Result.Step = "Main"
                        $Result.TimeStamp = Get-Date -Format FileDateTime

                        $Termination.AddResult($Result)
                    }
                }
                $CurrentCount++
            }

            $Main.RemoveTerminationSession($SessionType, $TerminationCount, $CurrentCount)

            $Section = "AdminActiveDirectoryDOMAIN1"
            $AccountType = "Admin"

            $CurrentCount = 1

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section

            $Main.Terminations | ForEach-Object {
                $Termination = $_

                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain

                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)

                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.RemoveUnixAttributes($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)

                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)
                }
                $CurrentCount++
            }

            $Section = "StandardActiveDirectoryDOMAIN2"
            $AccountType = "Standard"
            $Domain = "DOMAIN2"

            $CurrentCount = 1

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section

            $Main.Terminations | ForEach-Object {
                $Termination = $_

                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain

                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)

                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    
                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)
                }
                $CurrentCount++
            }

            $Section = "AdminActiveDirectoryDOMAIN2"
            $AccountType = "Admin"

            $CurrentCount = 1

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section

            $Main.Terminations | ForEach-Object {
                $Termination = $_

                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain

                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)

                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)

                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)
                }
                $CurrentCount++
            }

            $Section = "StandardActiveDirectoryDOMAIN3"
            $AccountType = "Standard"
            $Domain = "DOMAIN3"

            $CurrentCount = 1

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section

            $Main.Terminations | ForEach-Object {
                $Termination = $_

                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain

                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)

                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)

                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)
                }
                $CurrentCount++
            }

            $AccountType = "Mailbox"
            $Section = "ExchangeDOMAIN3"
            $SessionType = "O365"

            $Log.Section = $Section

            $CurrentCount = 1

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section
            Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

            $Main.NewTerminationSession($SessionType, $TerminationCount, $CurrentCount)

            $Main.Terminations | ForEach-Object {
                $Termination = $_

                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                if ($null -ne $Termination.AccountData.Standard.$Domain) {
                    $Account = $Termination.AccountData.Standard.$Domain

                    $Termination.GetEmployeeMailbox($Termination.ChangeOrder, $Domain, $TerminationCount, $CurrentCount)
                    if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                        $Mailbox = $Termination.AccountData.$AccountType.$Domain

                        $Log.Account = $Account.SamAccountName
                        $Log.AccountType = $Mailbox.AccountType
                        $Log.ChangeOrder = $Termination.ChangeOrder
                        $Log.DisplayName = $Account.DisplayName
                        $Log.Domain = $Account.Domain
                        $Log.EmployeeID = $Account.EmployeeID
                        $Log.Section = $Section
                        $Log.Status = "Start"
                        $Log.Step = "Main"
                        $Log.TimeStamp = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        $Result.Account = $Account.SamAccountName
                        $Result.AccountType = $Mailbox.AccountType
                        $Result.ChangeOrder = $Termination.ChangeOrder
                        $Result.DisplayName = $Account.DisplayName
                        $Result.Domain = $Account.Domain
                        $Result.EmployeeID = $Account.EmployeeID
                        $Result.Section = $Section
                        $Result.Status = "Start"
                        $Result.Step = "Main"
                        $Result.TimeStamp = Get-Date -Format FileDateTime

                        $Termination.AddResult($Result)

                        $Mailbox.EnableMailboxForwarding($Termination.ChangeOrder, $Account.ManagerEmailAddress, $TerminationCount, $CurrentCount)
                        $Mailbox.SetTerminatedOutOfOffice($Termination.ChangeOrder, $TerminationCount, $CurrentCount)

                        $Log.Status = "Complete"
                        $Log.Step = "Main"
                        $Log.TimeStamp = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        $Result.Status = "Complete"
                        $Result.Step = "Main"
                        $Result.TimeStamp = Get-Date -Format FileDateTime

                        $Termination.AddResult($Result)
                    }
                }
                $CurrentCount++
            }

            $Main.RemoveTerminationSession($SessionType, $TerminationCount, $CurrentCount)

            $Section = "AdminActiveDirectoryDOMAIN3"
            $AccountType = "Admin"

            $CurrentCount = 1

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section

            $Main.Terminations | ForEach-Object {
                $Termination = $_

                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain

                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)

                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)

                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)

                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime

                    $Termination.AddResult($Result)
                }
                $CurrentCount++
            }

            $Section = "PostProcessingTasks"
            $Step = "WriteLogs"

            $Log.Account = $env:USERNAME
            $Log.AccountType = "Developer"
            $Log.ChangeOrder = "N/A"
            $Log.DisplayName = $env:USERNAME -Replace "\.", " "
            $Log.Domain = "All"
            $Log.EmployeeID = "N/A"
            $Log.Section = $Section
            $Log.Status = "Complete"
            $Log.Step = $Step
            $Log.TimeStamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $CurrentCount = 1

            Write-TerminationProgress -ParameterSet $PSCmdlet.ParameterSetName -Section $Section

            $Main.Terminations | ForEach-Object {
                $Termination = $_

                Write-SectionProgress -Section $Section -TerminationCount $TerminationCount -CurrentCount $CurrentCount

                if ($Termination.ServiceDeskInfo.WorkflowTaskComment -eq "FIM processing has been completed." -and "Failed" -notin $Termination.Result.Status) {
                    $Termination.SetServiceDeskTaskCompleted($Main.Credential.ServiceDesk, $TerminationCount, $CurrentCount)
                }
                else {
                    $TaskComment = "Needs FIM processing."
                    $Termination.SetServiceDeskTaskUpdated($TaskComment, $Main.Credential.ServiceDesk, $TerminationCount, $CurrentCount)
                }
                $CurrentCount++
            }

            $BoxDatabase | ConvertTo-Json -Depth 99 | Out-File .\QA\SingleTerminationLogs\BoxDatabaseTesting.json

            $Main.WriteErrorLogsToFile()
            $Main.WriteGroupsToFile()
            $Main.WriteLogsToFile()
            $Main.WriteResultsToFile()
            $Main.CopyItemsToSharePoint($Main.Credential.O365)
            $Main.NewTeamsTerminationNotification()
        }

        "ServiceDesk" {
            $global:Main = [Main]::New()
            $ChangeOrder = "MultipleFromServiceDesk"
            $Domains = "DOMAIN1", "DOMAIN2", "DOMAIN3", "O365", "SKYPE"
            $Section = "PreparationAndSetup"
            $Step = "Main"
    
            $Log.Account = $env:USERNAME
            $Log.AccountType = "Developer"
            $Log.ChangeOrder = $ChangeOrder
            $Log.DisplayName = $env:USERNAME -Replace "\.", " "
            $Log.Domain = "All"
            $Log.EmployeeID = "N/A"
            $Log.Section = $Section
            $Log.Status = "Start"
            $Log.Step = $Step
            $Log.TimeStamp = Get-Date -Format FileDateTime
    
            $Main.AddLog($Log)
    
            if ($Credentials) {
                $Continue = [System.Management.Automation.Host.ChoiceDescription]::New("&Continue", "Continue with provided credentials")
                $Provide = [System.Management.Automation.Host.ChoiceDescription]::New("&Provide", "Provide new credentials")
                $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($Continue, $Provide)
                $Choice = $Host.UI.PromptForChoice("Select Option", "Provide new credentials?", $Choices, 0)
                Switch ($Choice) {
                    1 {
                        ForEach ($Domain in $Domains) {
                            $Main.GetTerminationCredential($Domain)
                        }
                    }
                    0 {
                        Write-Information "Continuing..."
                        $Main.AddCredentials($Credentials)
                    }
                }
            }
            else {
                ForEach ($Domain in $Domains) {
                    $Main.GetTerminationCredential($Domain)
                }
            }
            $Domains | Where-Object { $_ -ne "SKYPE" } | ForEach-Object {
                $Attempts = 0
                $Retry = $true
                do {
                    $Main.TestUserAccountAccess($_)
                    if ($Main.VerifiedAccess.$_ -ne $true) {
                        $Provide = [System.Management.Automation.Host.ChoiceDescription]::New("&Provide", "Provide $_ credentials")
                        $Exit = [System.Management.Automation.Host.ChoiceDescription]::New("&Exit", "Exit")
                        $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($Exit, $Provide)
                        $Choice = $Host.UI.PromptForChoice("Unable to verify access for $_", "Provide alternate credentials?", $Choices, 0)
                        Switch ($Choice) {
                            1 {
                                $Main.GetTerminationCredential($_)
                                $Attempts++
                                if ($_ -eq "O365") {
                                    $Main.GetTerminationCredential("SKYPE")
                                }
                            }
                            0 {
                                $Retry = $false
                            }
                        }
                    }
                } while ($Main.VerifiedAccess.$_ -ne $true -and $Attempts -lt 3 -and $Retry -eq $true)
            }
            $Domains | Where-Object { $_ -ne "SKYPE" } | ForEach-Object {
                if ($Main.VerifiedAccess.$_ -ne $true) {
                    break
                }
            }
            $Main.GetLogFile()
            if ($null -eq $Main.LogFilePath) {
                break
            }
            $Main.GetErrorLogFile()
            if ($null -eq $Main.ErrorLogFilePath) {
                break
            }
            $Main.GetGroupFile()
            if ($null -eq $Main.GroupFilePath) {
                break
            }
            $Main.GetResultFile()
            if ($null -eq $Main.ResultFilePath) {
                break
            }
            $SessionType = "DOMAIN1"
            $Main.NewTerminationSession($SessionType)
            $Main.GetTerminatedDatabase()
            if ($null -eq $Main.Database) {
                break
            }
            $Main.RemoveTerminationSession($SessionType)
    
            $Section = "PreProcessingTasks"
            $Step = "Main"
    
            $Log.Section = $Section
            $Log.Status = "Start"
            $Log.Step = $Step
            $Log.TimeStamp = Get-Date -Format FileDateTime
    
            $Main.AddLog($Log)
    
            $Main.GetServiceDeskTasks($Main.Credential.ServiceDesk)
    
            $AccountTypes = "Standard", "Admin", "Manager"
            $Domains = "DOMAIN1", "DOMAIN2", "DOMAIN3"
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                ForEach ($AccountType in $AccountTypes) {
                    if ($AccountType -ne "Manager") {
                        ForEach ($Domain in $Domains) {
                            $Termination.GetEmployeeAccount($AccountType, $Termination.ChangeOrder, $Domain, "ID")
                        }
                    }
                    else {
                        ForEach ($Domain in $Domains) {
                            if ($Domain -ne "DOMAIN2") {
                                $Termination.GetEmployeeAccount($AccountType, $Termination.ChangeOrder, $Domain, "Name")
                            }
                        }
                    }
                }
            }
    
            $AccountTypes = "Standard", "Admin"
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                ForEach ($AccountType in $AccountTypes) {
                    ForEach ($Domain in $Domains) {
                        if ($Termination.AccountData.$AccountType.$Domain.SamAccountName -eq "") {
                            $Termination.AccountData.$AccountType.$Domain = $null
                        }
                    }
                }
            }
    
            $Log.Account = $env:USERNAME
            $Log.AccountType = "Developer"
            $Log.ChangeOrder = $ChangeOrder
            $Log.DisplayName = $env:USERNAME -Replace "\.", " "
            $Log.Domain = "All"
            $Log.EmployeeID = "N/A"
            $Log.Status = "Complete"
            $Log.TimeStamp = Get-Date -Format FileDateTime
    
            $Main.AddLog($Log)
    
            $AccountType = "Standard"
            $Domains = "DOMAIN1", "DOMAIN3"
    
            $Log.Status = "Start"
            $Log.Step = "Test-AccountStatus"
            $Log.TimeStamp = Get-Date -Format FileDateTime
    
            $Main.AddLog($Log)
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                ForEach ($Domain in $Domains) {
                    if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                        $Termination.TestAccountStatus($Termination.ChangeOrder, $Domain)
                        if ($Termination.VerifiedStatus.$Domain.Continue -ne $true) {
                            $Termination.AccountData = $null
                        }
                    }
                }
            }
    
            $Log.Account = $env:USERNAME
            $Log.AccountType = "Developer"
            $Log.ChangeOrder = $ChangeOrder
            $Log.DisplayName = $env:USERNAME -Replace "\.", " "
            $Log.Domain = "All"
            $Log.EmployeeID = "N/A"
            $Log.Status = "Complete"
            $Log.Step = "Main"
            $Log.TimeStamp = Get-Date -Format FileDateTime
    
            $Main.AddLog($Log)
    
            $Domain = "DOMAIN1"
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain
    
                    $Section = "StandardActiveDirectoryDOMAIN1"
    
                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
    
                    $Account.ClearHomeDirectory($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.ClearScriptPath($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.MoveBatchFileToArchive($Termination.ChangeOrder)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain)
    
                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
                }
            }
    
            $Section = "DisableEmployeeSIP"
            $SessionType = "SKYPE"
    
            $Log.Section = $Section
    
            $Main.NewTerminationSession($SessionType)
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain
    
                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
    
                    $Account.DisableEmployeeSIP($Termination.ChangeOrder)
    
                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
                }
            }
    
            $Main.RemoveTerminationSession($SessionType)
    
            $AccountType = "Mailbox"
            $Section = "ExchangeDOMAIN1"
            $SessionType = "DOMAIN1"
            $Step = "Main"
    
            $Log.Section = $Section
    
            $Main.NewTerminationSession($SessionType)
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                if ($null -ne $Termination.AccountData.Standard.$Domain) {
                    $Account = $Termination.AccountData.Standard.$Domain
    
                    $Termination.GetEmployeeMailbox($Termination.ChangeOrder, $Domain)
                    if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                        $Mailbox = $Termination.AccountData.$AccountType.$Domain
    
                        $Log.Account = $Account.SamAccountName
                        $Log.AccountType = $Mailbox.AccountType
                        $Log.ChangeOrder = $Termination.ChangeOrder
                        $Log.DisplayName = $Account.DisplayName
                        $Log.Domain = $Account.Domain
                        $Log.EmployeeID = $Account.EmployeeID
                        $Log.Section = $Section
                        $Log.Status = "Start"
                        $Log.Step = $Step
                        $Log.TimeStamp = Get-Date -Format FileDateTime
    
                        $Main.AddLog($Log)
    
                        $Result.Account = $Account.SamAccountName
                        $Result.AccountType = $Mailbox.AccountType
                        $Result.ChangeOrder = $Termination.ChangeOrder
                        $Result.DisplayName = $Account.DisplayName
                        $Result.Domain = $Account.Domain
                        $Result.EmployeeID = $Account.EmployeeID
                        $Result.Section = $Section
                        $Result.Status = "Start"
                        $Result.Step = $Step
                        $Result.TimeStamp = Get-Date -Format FileDateTime
    
                        $Termination.AddResult($Result)
    
                        $Mailbox.EnableMailboxForwarding($Termination.ChangeOrder, $Account.ManagerEmailAddress)
                        $Mailbox.MoveMailboxToTerminatedDatabase($Termination.ChangeOrder, $Main.Database)
                        $Mailbox.RemoveInboxRules($Termination.ChangeOrder)
                        $Mailbox.SetContactHiddenFromGAL($Termination.ChangeOrder)
                        $Mailbox.SetMailboxHiddenFromGAL($Termination.ChangeOrder)
                        $Mailbox.SetTerminatedOutOfOffice($Termination.ChangeOrder)
    
                        $Log.Status = "Complete"
                        $Log.Step = "Main"
                        $Log.TimeStamp = Get-Date -Format FileDateTime
    
                        $Main.AddLog($Log)
    
                        $Result.Status = "Complete"
                        $Result.Step = "Main"
                        $Result.TimeStamp = Get-Date -Format FileDateTime
    
                        $Termination.AddResult($Result)
                    }
                }
            }
            $Main.RemoveTerminationSession($SessionType)
    
            $AccountType = "Admin"
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain
    
                    $Section = "AdminActiveDirectoryDOMAIN1"
    
                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
    
                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.RemoveUnixAttributes($Termination.ChangeOrder, $Main.Credential.$Domain, $TerminationCount, $CurrentCount)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain)
    
                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
                }
            }
    
            $AccountType = "Standard"
            $Domain = "DOMAIN2"
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain
    
                    $Section = "StandardActiveDirectoryDOMAIN2"
    
                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
    
                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain)
                        
                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
                }
            }
    
            $AccountType = "Admin"
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain
    
                    $Section = "AdminActiveDirectoryDOMAIN2"
    
                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
    
                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain)
    
                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
                }
            }
    
            $AccountType = "Standard"
            $Domain = "DOMAIN3"
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain
    
                    $Section = "StandardActiveDirectoryDOMAIN3"
    
                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
    
                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain)
    
                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
                }
            }
    
            $AccountType = "Mailbox"
            $Section = "ExchangeDOMAIN3"
            $SessionType = "O365"
    
            $Log.Section = $Section
    
            $Main.NewTerminationSession($SessionType)
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                if ($null -ne $Termination.AccountData.Standard.$Domain) {
                    $Account = $Termination.AccountData.Standard.$Domain
    
                    $Termination.GetEmployeeMailbox($Termination.ChangeOrder, $Domain)
                    if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                        $Mailbox = $Termination.AccountData.$AccountType.$Domain
    
                        $Log.Account = $Account.SamAccountName
                        $Log.AccountType = $Mailbox.AccountType
                        $Log.ChangeOrder = $Termination.ChangeOrder
                        $Log.DisplayName = $Account.DisplayName
                        $Log.Domain = $Account.Domain
                        $Log.EmployeeID = $Account.EmployeeID
                        $Log.Section = $Section
                        $Log.Status = "Start"
                        $Log.Step = "Main"
                        $Log.TimeStamp = Get-Date -Format FileDateTime
    
                        $Main.AddLog($Log)
    
                        $Result.Account = $Account.SamAccountName
                        $Result.AccountType = $Mailbox.AccountType
                        $Result.ChangeOrder = $Termination.ChangeOrder
                        $Result.DisplayName = $Account.DisplayName
                        $Result.Domain = $Account.Domain
                        $Result.EmployeeID = $Account.EmployeeID
                        $Result.Section = $Section
                        $Result.Status = "Start"
                        $Result.Step = "Main"
                        $Result.TimeStamp = Get-Date -Format FileDateTime
    
                        $Termination.AddResult($Result)
    
                        $Mailbox.EnableMailboxForwarding($Termination.ChangeOrder, $Account.ManagerEmailAddress)
                        $Mailbox.SetTerminatedOutOfOffice($Termination.ChangeOrder)
    
                        $Log.Status = "Complete"
                        $Log.Step = "Main"
                        $Log.TimeStamp = Get-Date -Format FileDateTime
    
                        $Main.AddLog($Log)
    
                        $Result.Status = "Complete"
                        $Result.Step = "Main"
                        $Result.TimeStamp = Get-Date -Format FileDateTime
    
                        $Termination.AddResult($Result)
                    }
                }
            }
            $Main.RemoveTerminationSession($SessionType)
    
            $AccountType = "Admin"
    
            $Main.Terminations | ForEach-Object {
                $Termination = $_
                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $Account = $Termination.AccountData.$AccountType.$Domain
    
                    $Section = "AdminActiveDirectoryDOMAIN3"
    
                    $Log.Account = $Account.SamAccountName
                    $Log.AccountType = $Account.AccountType
                    $Log.ChangeOrder = $Termination.ChangeOrder
                    $Log.DisplayName = $Account.DisplayName
                    $Log.Domain = $Account.Domain
                    $Log.EmployeeID = $Account.EmployeeID
                    $Log.Section = $Section
                    $Log.Status = "Start"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Account = $Account.SamAccountName
                    $Result.AccountType = $Account.AccountType
                    $Result.ChangeOrder = $Termination.ChangeOrder
                    $Result.DisplayName = $Account.DisplayName
                    $Result.Domain = $Account.Domain
                    $Result.EmployeeID = $Account.EmployeeID
                    $Result.Section = $Section
                    $Result.Status = "Start"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
    
                    $Account.DisableDialInAccess($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.RemoveEmployeeGroups($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.SetRandomPassword($Termination.ChangeOrder, $Main.Credential.$Domain)
                    $Account.SetTerminatedOrganizationalUnit($Termination.ChangeOrder, $Main.Credential.$Domain)
    
                    $Log.Status = "Complete"
                    $Log.Step = "Main"
                    $Log.TimeStamp = Get-Date -Format FileDateTime
    
                    $Main.AddLog($Log)
    
                    $Result.Status = "Complete"
                    $Result.Step = "Main"
                    $Result.TimeStamp = Get-Date -Format FileDateTime
    
                    $Termination.AddResult($Result)
                }
            }
            $Section = "PostProcessingTasks"
            $Step = "WriteLogs"
    
            $Log.Account = $env:USERNAME
            $Log.AccountType = "Developer"
            $Log.ChangeOrder = "N/A"
            $Log.DisplayName = $env:USERNAME -Replace "\.", " "
            $Log.Domain = "All"
            $Log.EmployeeID = "N/A"
            $Log.Section = $Section
            $Log.Status = "Complete"
            $Log.Step = $Step
            $Log.TimeStamp = Get-Date -Format FileDateTime
    
            $Main.AddLog($Log)

            $Main.Terminations | ForEach-Object {
                $Termination = $_
                if ($Termination.ServiceDeskInfo.WorkflowTaskComment -eq "FIM processing has been completed." -and "Failed" -notin $Termination.Result.Status) {
                    $Termination.SetServiceDeskTaskCompleted($Main.Credential.ServiceDesk)
                }
                else {
                    $TaskComment = "Needs FIM processing."
                    $Termination.SetServiceDeskTaskUpdated($TaskComment, $Main.Credential.ServiceDesk)
                }
            }
    
            $Main.WriteErrorLogsToFile()
            $Main.WriteGroupsToFile()
            $Main.WriteLogsToFile()
            $Main.WriteResultsToFile()
            $Main.CopyItemsToSharePoint($Main.Credential.O365)
            $Main.NewTeamsTerminationNotification()
        }
    }
}