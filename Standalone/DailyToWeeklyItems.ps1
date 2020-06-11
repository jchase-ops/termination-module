<#
Daily script to move all Termination results from Daily to Weekly lists.
#>

Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
Add-Type -Path $(${env:ProgramFiles(x86)} + '\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll')
Import-Module "C:\Program Files\WindowsPowerShell\Modules\SPOMod\SPOMod.psm1"
Import-Module -Name ReportHTML

$Url = "https://domain1.sharepoint.com/sites/Developer"
$DailyListTitle = "Daily Termination Results"
$GroupListTitle = "Group Termination Results"
$WeeklyListTitle = "Weekly Termination Results"
$DocLib = "Termination Records"
$Credential = [PSCredential]::New("SvcAcctUsername@domain1.com", $(ConvertTo-SecureString -String "SvcAcctPW" -AsPlainText -Force))
$TemporaryDirectory = New-Item -Name TermReports -Path $env:TEMP -ItemType Directory

$CurrentDate = New-Object -Type DateTime -ArgumentList (Get-Date).Year, (Get-Date).Month, (Get-Date).Day, 6, 0, 0, 0

Connect-SPOCSOM -Url $Url -Credential $Credential

$DailyListItems = Get-SPOListItems -ListTitle $DailyListTitle -IncludeAllProperties $true | Where-Object { $_.TimeStamp.ToDateTime($null) -lt $CurrentDate }
$GroupListItems = Get-SPOListItems -ListTitle $GroupListTitle -IncludeAllProperties $true | Where-Object { $_.Title -in $DailyListItems.Title }

$Report = @()
$TitleText = "Termination Results for " + $CurrentDate.AddDays(-1).ToString("ddMMMyyyy")
$LogoPath = "$PSScriptRoot\Resources"
$LeftLogoName = "Patch"
$RightLogoName = "Logo"
$TabArray = "Steps", "Groups", "All"

$Report += Get-HTMLOpenPage -TitleText $TitleText -LogoPath $LogoPath -LeftLogoName $LeftLogoName -RightLogoName $RightLogoName
$Report += Get-HTMLTabHeader -TabNames $TabArray
$TabArray | ForEach-Object {
    $Tab = $_
    $Report += Get-HTMLTabContentOpen -TabName $Tab -TabHeading "Termination Results: $Tab"
    Switch ($Tab) {
        "Steps" {
            $DailyListItems.Title | Sort-Object -Unique | ForEach-Object {
                $ChangeOrder = $_
                $TableData = $DailyListItems | Where-Object { $_.Title -eq $ChangeOrder } | Select-Object Step, Status, Domain, AccountType, Account, EmployeeID
                $Report += Get-HTMLContentOpen -HeaderText $ChangeOrder -IsHidden
                $Report += Get-HTMLContentDataTable -ArrayOfObjects $TableData -HideFooter -PagingOptions '5,10,15,'
                $Report += Get-HTMLContentClose
            }
        }
        "Groups" {
            $GroupListItems.Title | Sort-Object -Unique | ForEach-Object {
                $ChangeOrder = $_
                $TableData = $GroupListItems | Where-Object { $_.Title -eq $ChangeOrder } | Select-Object GroupName, Result, Category, FIMManaged, Domain, AccountType
                $Report += Get-HTMLContentOpen -HeaderText $ChangeOrder -IsHidden
                $Report += Get-HTMLContentDataTable -ArrayOfObjects $TableData -HideFooter -PagingOptions '5,10,15,'
                $Report += Get-HTMLContentClose
            }
        }
        "All" {
            $DailyListItems.Title | Sort-Object -Unique | ForEach-Object {
                $ChangeOrder = $_
                $StepTableData = $DailyListItems | Where-Object { $_.Title -eq $ChangeOrder } | Select-Object Step, Status, Domain, AccountType, Account, EmployeeID, CompletedBy, TimeStamp
                $GroupTableData = $GroupListItems | Where-Object { $_.Title -eq $ChangeOrder } | Select-Object GroupName, Result, Category, FIMManaged, Domain, AccountType, CompletedBy, TimeStamp
                $Report += Get-HTMLContentOpen -HeaderText $ChangeOrder -IsHidden
                $Report += Get-HTMLContentOpen -HeaderText "Steps" -IsHidden
                $Report += Get-HTMLContentDataTable -ArrayOfObjects $StepTableData -HideFooter -PagingOptions '5,10,15,'
                $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentOpen -HeaderText "Groups" -IsHidden
                $Report += Get-HTMLContentDataTable -ArrayOfObjects $GroupTableData -HideFooter -PagingOptions '5,10,15,'
                $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentClose
            }
        }
    }
    $Report += Get-HTMLTabContentClose
}
$Report += Get-HTMLClosePage
$SavedReport = Save-HTMLReport -ReportContent $Report -ReportName $CurrentDate.AddDays(-1).ToString("ddMMMyyyy") -ReportPath $TemporaryDirectory

$Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
$Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.Username, $Credential.Password)
$List = $Context.Web.Lists.GetByTitle($DocLib)
$Context.Load($List.RootFolder)
$Context.ExecuteQuery()

$FileStream = New-Object IO.FileStream($SavedReport, [System.IO.FileMode]::Open)
$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
$FileCreationInfo.ContentStream = $FileStream
$FileCreationInfo.Url = $List.RootFolder.ServerRelativeUrl + "/Daily/" + $($SavedReport -Replace "^.*\\", "")
$UploadFile = $List.RootFolder.Files.Add($FileCreationInfo)
$Context.Load($UploadFile)
$Context.ExecuteQuery()

$UploadItems = New-Object System.Collections.Generic.List[System.Object]

$DailyListItems | Where-Object { $_.AccountType -eq "Mailbox" } | ForEach-Object {
    $_.AccountType = "Standard"
}

$DailyListItems.Title | Sort-Object -Unique | ForEach-Object {
    $ChangeOrder = $_
    $DailyItems = $DailyListItems | Where-Object { $_.Title -eq $ChangeOrder }
    $GroupItems = $GroupListItems | Where-Object { $_.Title -eq $ChangeOrder }
    $EmployeeID = $DailyItems.EmployeeID | Sort-Object -Unique
    $DomainAccounts = $DailyItems | Sort-Object Account, Domain, AccountType | Select-Object Account, Domain, AccountType -Unique
    $DomainAccounts | ForEach-Object {
        $Account = $_
        if ($Account.AccountType -eq "Admin") {
            $StepsPerformed = $DailyItems | Where-Object { $_.AccountType -eq $Account.AccountType -and $_.Domain -eq $Account.Domain }
            $GroupsProcessed = $GroupItems | Where-Object { $_.AccountType -eq $Account.AccountType -and $_.Domain -eq $Account.Domain }
        }
        else {
            $StepsPerformed = $DailyItems | Where-Object { $_.AccountType -ne "Admin" -and $_.Domain -eq $Account.Domain }
            $GroupsProcessed = $GroupItems | Where-Object { $_.AccountType -ne "Admin" -and $_.Domain -eq $Account.Domain }
        }
        $Hyperlink = New-Object Microsoft.SharePoint.Client.FieldURLValue
        $Hyperlink.Description = $CurrentDate.AddDays(-1).ToString("ddMMMyyyy")
        $Hyperlink.Url = "https://company.sharepoint.com/sites/Developer/Termination Records/Daily/$($Hyperlink.Description).html"
        $obj = [PSCustomObject]@{
            ChangeOrder = $ChangeOrder
            EmployeeID = $EmployeeID
            Account = $_.Account
            AccountType = $_.AccountType
            Domain = $_.Domain
            StepsPerformed = @($StepsPerformed).Count
            SuccessfulSteps = @($StepsPerformed | Where-Object { $_.Status -eq "Success" }).Count
            SkippedSteps = @($StepsPerformed | Where-Object { $_.Status -eq "Skipped" }).Count
            FailedSteps = @($StepsPerformed | Where-Object { $_.Status -eq "Failed" }).Count
            GroupsProcessed = @($GroupsProcessed).Count
            RemovedGroups = @($GroupsProcessed | Where-Object { $_.Result -eq "Removed" }).Count
            ExcludedGroups = @($GroupsProcessed | Where-Object { $_.Result -eq "Excluded" }).Count
            FailedGroups = @($GroupsProcessed | Where-Object { $_.Result -eq "Failed" }).Count
            GroupResults = $Hyperlink
            CompletedBy = $_.CompletedBy
            TimeStamp = $_.TimeStamp
        }
        $UploadItems.Add($obj)
    }
}

$List = $Context.Web.Lists.GetByTitle($WeeklyListTitle)
$Context.Load($List)
$Context.ExecuteQuery()

$UploadItems | ForEach-Object {
    $ListItemCreationInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
    $NewListItem = $List.AddItem($ListItemCreationInfo)
    $NewListItem["Title"] = $_.ChangeOrder
    $NewListItem["EmployeeID"] = $_.EmployeeID
    $NewListItem["Account"] = $_.Account
    $NewListItem["AccountType"] = $_.AccountType
    $NewListItem["Domain"] = $_.Domain
    $NewListItem["StepsPerformed"] = $_.StepsPerformed
    $NewListItem["SuccessfulSteps"] = $_.SuccessfulSteps
    $NewListItem["SkippedSteps"] = $_.SkippedSteps
    $NewListItem["FailedSteps"] = $_.FailedSteps
    $NewListItem["GroupsProcessed"] = $_.GroupsProcessed
    $NewListItem["Removed"] = $_.RemovedGroups
    $NewListItem["Excluded"] = $_.ExcludedGroups
    $NewListItem["Failed"] = $_.FailedGroups
    $NewListItem["GroupResults"] = [Microsoft.SharePoint.Client.FieldUrlValue]$($_.GroupResults)
    $NewListItem["CompletedBy"] = $_.CompletedBy
    $NewListItem["TimeStamp"] = $_.TimeStamp
    $NewListItem.Update()
    $Context.ExecuteQuery()
}

$DailyListItems | ForEach-Object {
    Remove-SPOListItem -ListTitle $DailyListTitle -ItemID $_.ID | Out-Null
}

$GroupListItems | ForEach-Object {
    Remove-SPOListItem -ListTitle $GroupListTitle -ItemID $_.ID | Out-Null
}

Remove-Item -Path $TemporaryDirectory -Recurse