<#
Termination Module
#>

##################################################
#                                                #
#              DOT-SOURCED SCRIPTS               #
#                                                #
##################################################

. $PSScriptRoot\ClearHomeDirectory.ps1
. $PSScriptRoot\ClearScriptPath.ps1
. $PSScriptRoot\CopyItemsToSharePoint.ps1
. $PSScriptRoot\DisableDialInAccess.ps1
. $PSScriptRoot\DisableEmployeeSIP.ps1
. $PSScriptRoot\EnableMailboxForwarding.ps1
. $PSScriptRoot\GetEmployeeAccount.ps1
. $PSScriptRoot\GetEmployeeMailbox.ps1
. $PSScriptRoot\GetErrorLogFile.ps1
. $PSScriptRoot\GetGroupFile.ps1
. $PSScriptRoot\GetLogFile.ps1
. $PSScriptRoot\GetResultFile.ps1
. $PSScriptRoot\GetServiceDeskInfo.ps1
. $PSScriptRoot\GetServiceDeskTasks.ps1
. $PSScriptRoot\GetTerminatedDatabase.ps1
. $PSScriptRoot\GetTerminationCredential.ps1
. $PSScriptRoot\MoveBatchFileToArchive.ps1
. $PSScriptRoot\MoveMailboxToTerminatedDatabase.ps1
. $PSScriptRoot\NewTeamsTerminationNotification.ps1
. $PSScriptRoot\NewTerminationSession.ps1
. $PSScriptRoot\RemoveEmployeeGroups.ps1
. $PSScriptRoot\RemoveInboxRules.ps1
. $PSScriptRoot\RemoveTerminationSession.ps1
. $PSScriptRoot\RemoveUnixAttributes.ps1
. $PSScriptRoot\SetContactHiddenFromGAL.ps1
. $PSScriptRoot\SetMailboxHiddenFromGAL.ps1
. $PSScriptRoot\SetRandomPassword.ps1
. $PSScriptRoot\SetServiceDeskTaskCompleted.ps1
. $PSScriptRoot\SetServiceDeskTaskUpdated.ps1
. $PSScriptRoot\SetTerminatedOrganizationalUnit.ps1
. $PSScriptRoot\SetTerminatedOutOfOffice.ps1
. $PSScriptRoot\StartEmployeeTermination.ps1
. $PSScriptRoot\TerminationClasses.ps1
. $PSScriptRoot\TestAccountStatus.ps1
. $PSScriptRoot\TestUserAccountAccess.ps1
. $PSScriptRoot\WriteTerminationProgressLevels.ps1

##################################################
#                                                #
#                   VARIABLES                    #
#                                                #
##################################################

# AD Groups that either should not or cannot be removed through AD by NOC
New-Variable -Name ExcludeGroups -Value @{ } -Scope Script -Description "AD Groups Excluded from Removal"

$ExcludeGroups["Standard"] = @{ }
$ExcludeGroups["Admin"] = @{ }

$ExcludeGroups.Standard["Domain1"] = @(
    "Domain Users"
)
$ExcludeGroups.Standard["Domain2"] = @()
$ExcludeGroups.Standard["Domain3"] = @(
    "Domain Users"
)

$ExcludeGroups.Admin["Domain1"] = @()
$ExcludeGroups.Admin["Domain2"] = @()
$ExcludeGroups.Admin["Domain3"] = @()

# Messages used to set OOO for terminated employees
New-Variable -Name Message -Value @{ } -Scope Script -Description "Approved Message for Terminated Out of Office"

$Message["Domain1"] = "Hello, and thank you for your email."

$Message["Domain3"] = "Hello, and thank you for your email."

# Hash tables of relevant Organizational Units
New-Variable -Name SearchBase -Value @{ } -Scope Script -Description "Organization Units for Employees & Admin Accounts"

$SearchBase["Standard"] = @{ }
$SearchBase["Admin"] = @{ }
$SearchBase["Default"] = @{ }

$SearchBase.Standard["Active"] = @{ }
$SearchBase.Standard["Terminated"] = @{ }

$SearchBase.Admin["Active"] = @{ }
$SearchBase.Admin["Terminated"] = @{ }

$SearchBase.Standard.Active["Domain1"] = "OU=Active,OU=Employee,OU=Managed,OU=Production,DC=Domain1,DC=com"
$SearchBase.Standard.Active["Domain2"] = "OU=Active,OU=Employee,OU=Managed,OU=Production,DC=Domain2,DC=com"
$SearchBase.Standard.Active["Domain3"] = "OU=Active,OU=Employee,OU=Managed,OU=Production,DC=Domain3,DC=com"

$SearchBase.Standard.Terminated["Domain1"] = "OU=Terminated,OU=Employee,OU=Managed,OU=Production,DC=Domain1,DC=com"
$SearchBase.Standard.Terminated["Domain2"] = "OU=Terminated,OU=Employee,OU=Managed,OU=Production,DC=Domain2,DC=com"
$SearchBase.Standard.Terminated["Domain3"] = "OU=Terminated,OU=Employee,OU=Managed,OU=Production,DC=Domain3,DC=com"

$SearchBase.Admin.Active["Domain1"] = "OU=Elevated Accounts,OU=Managed,OU=Production,DC=Domain1,DC=com"
$SearchBase.Admin.Active["Domain2"] = "OU=Elevated Accounts,OU=Managed,OU=Production,DC=Domain2,DC=com"
$SearchBase.Admin.Active["Domain3"] = "OU=Elevated Accounts,OU=Managed,OU=Production,DC=Domain3,DC=com"

$SearchBase.Admin.Terminated["Domain1"] = "OU=Accounts,OU=Inactive Resources,DC=Domain1,DC=com"
$SearchBase.Admin.Terminated["Domain2"] = "OU=Accounts,OU=Inactive Resources,DC=Domain2,DC=com"
$SearchBase.Admin.Terminated["Domain3"] = "OU=Accounts,OU=Inactive Resources,DC=Domain3,DC=com"

$SearchBase.Default["Domain1"] = "OU=Employee,OU=Managed,OU=Production,DC=Domain1,DC=com"
$SearchBase.Default["Domain2"] = "OU=Employee,OU=Managed,OU=Production,DC=Domain2,DC=com"
$SearchBase.Default["Domain3"] = "OU=Employee,OU=Managed,OU=Production,DC=Domain3,DC=com"

# Get Domain Controllers
New-Variable -Name DomainControllers -Value @{ } -Scope Script -Description "Domain Controllers"

$DomainControllers["Domain1"] = Get-ADDomainController -Discover -DomainName "Domain1"
$DomainControllers["Domain2"] = Get-ADDomainController -Discover -DomainName "Domain2"
$DomainControllers["Domain3"] = Get-ADDomainController -Discover -DomainName "Domain3"

# Initialize $ErrorLog object
New-Variable -Name ErrorLog -Value (New-Object -TypeName PSCustomObject) -Scope Script -Description "PSCustomObject used to create error logs"

Add-Member -InputObject $ErrorLog -NotePropertyName Exception -NotePropertyValue $null
Add-Member -InputObject $ErrorLog -NotePropertyName ExceptionFullName -NotePropertyValue $null
Add-Member -InputObject $ErrorLog -NotePropertyName FullyQualifiedErrorId -NotePropertyValue $null
Add-Member -InputObject $ErrorLog -NotePropertyName Line -NotePropertyValue $null
Add-Member -InputObject $ErrorLog -NotePropertyName MyCommand -NotePropertyValue $null
Add-Member -InputObject $ErrorLog -NotePropertyName OffsetInLine -NotePropertyValue $null
Add-Member -InputObject $ErrorLog -NotePropertyName ScriptLineNumber -NotePropertyValue $null
Add-Member -InputObject $ErrorLog -NotePropertyName ScriptName -NotePropertyValue $null
Add-Member -InputObject $ErrorLog -NotePropertyName TimeStamp -NotePropertyValue $null

# Initialize $Log object
New-Variable -Name Log -Value (New-Object -TypeName PSCustomObject) -Scope Script -Description "PSCustomObject used to create logs"

Add-Member -InputObject $Log -NotePropertyName Account -NotePropertyValue $null
Add-Member -InputObject $Log -NotePropertyName AccountType -NotePropertyValue $null
Add-Member -InputObject $Log -NotePropertyName ChangeOrder -NotePropertyValue $null
Add-Member -InputObject $Log -NotePropertyName Detail -NotePropertyValue $null
Add-Member -InputObject $Log -NotePropertyName DisplayName -NotePropertyValue $null
Add-Member -InputObject $Log -NotePropertyName Domain -NotePropertyValue $null
Add-Member -InputObject $Log -NotePropertyName EmployeeID -NotePropertyValue $null
Add-Member -InputObject $Log -NotePropertyName Section -NotePropertyValue $null
Add-Member -InputObject $Log -NotePropertyName Status -NotePropertyValue $null
Add-Member -InputObject $Log -NotePropertyName Step -NotePropertyValue $null
Add-Member -InputObject $Log -NotePropertyName TimeStamp -NotePropertyValue $null

# Initialize $Result object
New-Variable -Name Result -Value (New-Object -TypeName PSCustomObject) -Scope Script -Description "PSCustomObject used to log results"

Add-Member -InputObject $Result -NotePropertyName Account -NotePropertyValue $null
Add-Member -InputObject $Result -NotePropertyName AccountType -NotePropertyValue $null
Add-Member -InputObject $Result -NotePropertyName ChangeOrder -NotePropertyValue $null
Add-Member -InputObject $Result -NotePropertyName Detail -NotePropertyValue $null
Add-Member -InputObject $Result -NotePropertyName DisplayName -NotePropertyValue $null
Add-Member -InputObject $Result -NotePropertyName Domain -NotePropertyValue $null
Add-Member -InputObject $Result -NotePropertyName EmployeeID -NotePropertyValue $null
Add-Member -InputObject $Result -NotePropertyName Section -NotePropertyValue $null
Add-Member -InputObject $Result -NotePropertyName Status -NotePropertyValue $null
Add-Member -InputObject $Result -NotePropertyName Step -NotePropertyValue $null
Add-Member -InputObject $Result -NotePropertyName TimeStamp -NotePropertyValue $null

# Get list of scripts to parse
$FileList = Get-ChildItem -Path $PSScriptRoot -Filter *.ps1 | Where-Object { $_.Name -notlike "*Classes*" -and $_.Name -notlike "Write*" }

# Parse script content for Write-Progress helpers
$global:ParseArray = New-Object System.Collections.Generic.List[System.Object]

$FileList | ForEach-Object {
    $File = $_
    $Parse = [System.Management.Automation.PSParser]::Tokenize((Get-Content "$PSScriptRoot\$File"), [ref]$null)
    $Command = ($Parse | Where-Object { $_.Type -eq "CommandArgument" -and $_.Content -like "*-*" }).Content
    $StartLine = ($Parse | Where-Object { $_.Type -eq "GroupStart" } | Sort-Object StartLine | Select-Object -First 1).StartLine
    $EndLine = ($Parse | Where-Object { $_.Type -eq "GroupEnd" } | Sort-Object StartLine | Select-Object -Last 1).StartLine
    $obj = [PSCustomObject]@{
        Command = $Command -Replace "^.*\:", ""
        Tokens = $Parse | Where-Object { $_.StartLine -ge $StartLine -and $_.EndLine -le $EndLine }
    }
    $ParseArray.Add($obj)
}

$global:Tokens = [PSCustomObject]@{}

$ParseArray | ForEach-Object {
    $obj = [PSCustomObject]@{
        ScriptName = ($_.Command -Replace "\-", "") + ".ps1"
        Command = $_.Command
        Parsed = $_.Tokens
    }
    Add-Member -InputObject $Tokens -NotePropertyName $_.Command -NotePropertyValue $obj
}

$EndLine = ($Tokens."Start-EmployeeTermination".Parsed | Where-Object { $_.Content -eq "[Main]" })[1].StartLine

# Arrays for Progress counts
$global:Sections = @(
    "PreparationAndSetup"
    "PreProcessingTasks"
    "StandardActiveDirectoryDOMAIN1"
    "DisableEmployeeSIP"
    "ExchangeDOMAIN1"
    "AdminActiveDirectoryDOMAIN1"
    "StandardActiveDirectoryDOMAIN2"
    "AdminActiveDirectoryDOMAIN2"
    "StandardActiveDirectoryDOMAIN3"
    "ExchangeDOMAIN3"
    "AdminActiveDirectoryDOMAIN3"
    "PostProcessingTasks"
)

$global:Methods = $FileList.Name -Replace "(.ps1)", ""

# Count for Section Progress
$global:SectionDetails = [PSCustomObject]@{}

$SectionCount = 0

$Tokens."Start-EmployeeTermination".Parsed | Where-Object { $_.Content -in $Sections -and $_.Type -eq "String" -and $_.StartLine -lt $EndLine } | Select-Object Content, StartLine | ForEach-Object {
    $SectionCount++
    $obj = [PSCustomObject]@{
        SectionNumber = $SectionCount
        StartLine = $_.StartLine
        EndLine = $null
        StepDetails = [PSCustomObject]@{}
    }
    Add-Member -InputObject $SectionDetails -NotePropertyName $_.Content -NotePropertyValue $obj
}

$SectionCount = 0

do {
    if ($Sections[$($SectionCount + 1)]) {
        $SectionDetails.$($Sections[$SectionCount]).EndLine = $SectionDetails.$($Sections[$($SectionCount + 1)]).StartLine - 1
        $SectionCount++
    }
    else {
        $SectionDetails.$($Sections[$SectionCount]).EndLine = $EndLine - 1
        $SectionCount++
    }
} while ($SectionCount -lt $Sections.Count)

ForEach ($Section in $Sections) {
    $StepCount = ($Tokens."Start-EmployeeTermination".Parsed | Where-Object { $_.Type -eq "Member" -and $_.Content -in $Methods -and $_.StartLine -gt $SectionDetails.$Section.StartLine -and $_.StartLine -lt $SectionDetails.$Section.EndLine } | Select-Object Content -Unique).Count
    Add-Member -InputObject $SectionDetails.$Section -NotePropertyName StepCount -NotePropertyValue $StepCount
    $StepNames = ($Tokens."Start-EmployeeTermination".Parsed | Where-Object { $_.Type -eq "Member" -and $_.Content -in $Methods -and $_.StartLine -gt $SectionDetails.$Section.StartLine -and $_.StartLine -lt $SectionDetails.$Section.EndLine } | Select-Object Content -Unique).Content
    $StepNameCount = 0
    ForEach ($StepName in $StepNames) {
        $StepNameCount++
        Add-Member -InputObject $SectionDetails.$Section.StepDetails -NotePropertyName $StepName -NotePropertyValue $StepNameCount
    }
}

# Array of characters used to create randomized passwords
New-Variable -Name CharacterArray -Value @() -Scope Script -Description "Character array for creating randomized passwords"

$CharacterArray = 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '-', '_', '=', '+', '\', '|', '[', '{', ']', '}', '"', "'", ';', ':', ',', '<', '.', '>', '/', '?'

# PSObject used to track/update mailbox database size
$DatabaseUpdate = [PSCustomObject]@{
    Name = $null
    Index = $null
    Size = $null
}

# Various Uris used
New-Variable -Name TeamsUri -Value $TeamsWebhookURI -Scope Script -Description "Uri for sending Teams notifications"

New-Variable -Name ServiceDeskUri -Value $TicketSystemAPI -Scope Script -Description "Uri for Ticket System API"

New-Variable -Name SharePointUrl -Value $SharePointURI -Scope Script -Description "Url for SharePoint"

# SharePoint List Titles
New-Variable -Name GroupListTitle -Value "Group Termination Results" -Scope Script -Description "Title for SharePoint list of group termination results"
New-Variable -Name ResultsListTitle -Value "Daily Termination Results" -Scope Script -Description "Title for SharePoint list of all non-group termination results"

##################################################
#                                                #
#                EXPORTED ITEMS                  #
#                                                #
##################################################

Export-ModuleMember -Function Start-EmployeeTermination