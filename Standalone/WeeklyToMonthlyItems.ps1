<#
Weekly to Monthly SharePoint Items for Termination Module
#>

Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
Add-Type -Path $(${env:ProgramFiles(x86)} + '\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll')
Import-Module "C:\Program Files\WindowsPowerShell\Modules\SPOMod\SPOMod.psm1"
Import-Module -Name ReportHTML

$Url = "https://domain1.sharepoint.com/sites/Developer"
$WeeklyListTitle = "Weekly Termination Results"
$MonthlyListTitle = "Monthly Termination Results"
$DocLib = "Termination Records"
$Credential = [PSCredential]::New("SvcAcctUsername@domain1.com", $(ConvertTo-SecureString -String "SvcAcctPW" -AsPlainText -Force))
$TemporaryDirectory = New-Item -Name TermReports -Path $env:TEMP -ItemType Directory

$CurrentDate = New-Object -Type DateTime -ArgumentList (Get-Date).Year, (Get-Date).Month, (Get-Date).Day, 6, 0, 0, 0

if ($CurrentDate.AddDays(-8).Year -ne $CurrentDate.Year -or $CurrentDate.AddDays(-7).Month -ne $CurrentDate.Month) {
    if ($CurrentDate.Day -gt 1) {
        $StartDate = New-Object -Type DateTime -ArgumentList $CurrentDate.Year, $CurrentDate.Month, 1, 6, 0, 0, 0
        $EndDate = $CurrentDate.AddDays(-1)
    }
    else {
        $StartDate = $CurrentDate.AddDays(-8)
        $EndDate = New-Object -Type DateTime -ArgumentList $CurrentDate.Year, $CurrentDate.Month, 1, 6, 0, 0, 0
    }
}
else {
    $StartDate = $CurrentDate.AddDays(-8)
    $EndDate = $CurrentDate.AddDays(-1)
}

$ReportTitleDate = $StartDate.ToString("ddMMMyyyy") + "-" + $EndDate.ToString("ddMMMyyyy")

Connect-SPOCSOM -Url $Url -Credential $Credential

$WeeklyListItems = Get-SPOListItems -ListTitle $WeeklyListTitle -IncludeAllProperties $true | Where-Object { $_.TimeStamp.ToDateTime($null) -lt $EndDate -or $_.Created.ToLocalTime() -lt $EndDate }

$DailyReports = Get-SPOListItems -ListTitle $DocLib -IncludeAllProperties $true -Recursive | Where-Object { $_.FSObjType -eq "0" -and $_.FileDirRef -like "*Daily" }

$WebClient = New-Object System.Net.WebClient
$WebClient.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.Username, $Credential.Password)
$WebClient.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")

ForEach ($Report in $DailyReports) {
    $SourceFileName = $Report.FileLeafRef
    [System.Uri]$SourceFileUrl = "https://domain1.sharepoint.com" + $Report.FileRef
    $DownloadPath = $TemporaryDirectory.FullName + "\" + $SourceFileName
    $WebClient.DownloadFile($SourceFileUrl.AbsoluteUri, $DownloadPath)
}

$WebClient.Dispose()

$ItemsFromRecord = New-Object System.Collections.Generic.List[System.Object]

Get-ChildItem -Path $TemporaryDirectory | ForEach-Object {
    $File = $_
    $HTML = New-Object -ComObject "HTMLFile"
    $HTML.IHTMLDocument2_write($(Get-Content -Path $File.FullName -Raw))
    $All = $HTML.GetElementById("All")
    ($All.All.Tags("A") | Where-Object { $_.InnerText -Match "\d{6}" } | Sort-Object InnerText -Unique).parentElement.parentElement | ForEach-Object {
        $ChangeOrderHTML = $_
        $ChangeOrder = ($ChangeOrderHTML.All.Tags("A") | Where-Object { $_.InnerText -Match "\d{6}" }).InnerText
        $ChangeOrderObj = [PSCustomObject]@{
            ChangeOrder   = $ChangeOrder
            DateProcessed = $File.Name -Replace "\..*$", ""
            StepResults   = New-Object System.Collections.Generic.List[System.Object]
            GroupResults  = New-Object System.Collections.Generic.List[System.Object]
        }
        ($ChangeOrderHTML.All.Tags("A") | Where-Object { $_.InnerText -eq "Steps" }).parentElement.parentElement.All.Tags("TBODY") | ForEach-Object {
            ($_.OuterHTML) -Split "\r?\n" | Select-String -Pattern "<TR>" -Context (0, 8) | ForEach-Object {
                $StepArray = @()
                $_ -Split "\r?\n" | Where-Object { $_ -like "*TD*" } | ForEach-Object {
                    $StepArray += (($_ -Replace "^.*(<TD>)", "") -Replace "<.*$", "").Trim()
                }
                $StepObj = [PSCustomObject]@{
                    Step        = $StepArray[0]
                    Status      = $StepArray[1]
                    Domain      = $StepArray[2]
                    AccountType = $StepArray[3]
                    Account     = $StepArray[4]
                    EmployeeID  = $StepArray[5]
                    CompletedBy = $StepArray[6]
                    TimeStamp   = $StepArray[7]
                }
                $ChangeOrderObj.StepResults.Add($StepObj)
            }
        }
        ($ChangeOrderHTML.All.Tags("A") | Where-Object { $_.InnerText -eq "Groups" }).parentElement.parentElement.All.Tags("TBODY") | ForEach-Object {
            ($_.OuterHTML) -Split "\r?\n" | Select-String -Pattern "<TR>" -Context (0, 8) | ForEach-Object {
                $GroupArray = @()
                $_ -Split "\r?\n" | Where-Object { $_ -like "*TD*" } | ForEach-Object {
                    $GroupArray += (($_ -Replace "^.*(<TD>)", "") -Replace "<.*$", "").Trim()
                }
                $GroupObj = [PSCustomObject]@{
                    GroupName   = $GroupArray[0]
                    Result      = $GroupArray[1]
                    Category    = $GroupArray[2]
                    FIMManaged  = $GroupArray[3]
                    Domain      = $GroupArray[4]
                    AccountType = $GroupArray[5]
                    CompletedBy = $GroupArray[6]
                    TimeStamp   = $GroupArray[7]
                }
                $ChangeOrderObj.GroupResults.Add($GroupObj)
            }
        }
        $ItemsFromRecord.Add($ChangeOrderObj)
    }
}

$TabArray = @()
$DateCount = 0
do {
    $TabArray += $StartDate.AddDays($DateCount).ToString("ddMMMyyyy")
    $DateCount++
} while ($StartDate.AddDays($DateCount) -le $EndDate)

$Report = @()
$TitleText = "Termination Results for $ReportTitleDate"
$LogoPath = "$PSScriptRoot\Resources"
$LeftLogoName = "Patch"
$RightLogoName = "Logo"

$Report += Get-HTMLOpenPage -TitleText $TitleText -LogoPath $LogoPath -LeftLogoName $LeftLogoName -RightLogoName $RightLogoName
$Report += Get-HTMLTabHeader -TabNames $TabArray
$TabArray | ForEach-Object {
    $Tab = $_
    $Items = $ItemsFromRecord | Where-Object { $_.DateProcessed -eq $Tab }
    $Report += Get-HTMLTabContentOpen -TabName $Tab -TabHeading "Termination Results: $Tab"
    $Items | Sort-Object ChangeOrder | ForEach-Object {
        $StepTableData = $_.StepResults
        $GroupTableData = $_.GroupResults
        $Report += Get-HTMLContentOpen -HeaderText $_.ChangeOrder -IsHidden
        $Report += Get-HTMLContentOpen -HeaderText "Steps" -IsHidden
        $Report += Get-HTMLContentDataTable -ArrayOfObjects $StepTableData -HideFooter -PagingOptions "5,10,15,"
        $Report += Get-HTMLContentClose
        $Report += Get-HTMLContentOpen -HeaderText "Groups" -IsHidden
        $Report += Get-HTMLContentDataTable -ArrayOfObjects $GroupTableData -HideFooter -PagingOptions "5,10,15,"
        $Report += Get-HTMLContentClose
        $Report += Get-HTMLContentClose
    }
    $Report += Get-HTMLTabContentClose
}
$Report += Get-HTMLClosePage
$SavedReport = Save-HTMLReport -ReportContent $Report -ReportName $ReportTitleDate -ReportPath $TemporaryDirectory

$Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
$Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.Username, $Credential.Password)
$List = $Context.Web.Lists.GetByTitle($DocLib)
$Context.Load($List.RootFolder)
$Context.ExecuteQuery()

$Filestream = New-Object IO.FileStream($SavedReport, [System.IO.FileMode]::Open)
$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
$FileCreationInfo.ContentStream = $FileStream
$FileCreationInfo.Url = $List.RootFolder.ServerRelativeUrl + "/Weekly/" + $($SavedReport -Replace "^.*\\", "")
$UploadFile = $List.RootFolder.Files.Add($FileCreationInfo)
$Context.Load($UploadFile)
$Context.ExecuteQuery()

$UploadItems = New-Object System.Collections.Generic.List[System.Object]

$WeeklyListItems.Title | Sort-Object -Unique | ForEach-Object {
    $ChangeOrder = $_
    $WeeklyItems = $WeeklyListItems | Where-Object { $_.Title -eq $ChangeOrder }
    $EmployeeID = $WeeklyItems.EmployeeID | Sort-Object -Unique
    $Accounts = ($WeeklyItems | Sort-Object Domain, Account -Unique).Count
    $Domains = ($WeeklyItems | Sort-Object Domain -Unique).Count
    $SuccessfulSteps = 0
    $WeeklyItems.SuccessfulSteps | ForEach-Object {
        $SuccessfulSteps = $SuccessfulSteps + $_
    }
    $SkippedSteps = 0
    $WeeklyItems.SkippedSteps | ForEach-Object {
        $SkippedSteps = $SkippedSteps + $_
    }
    $FailedSteps = 0
    $WeeklyItems.FailedSteps | ForEach-Object {
        $FailedSteps = $FailedSteps + $_
    }
    $RemovedGroups = 0
    $WeeklyItems.Removed | ForEach-Object {
        $RemovedGroups = $RemovedGroups + $_
    }
    $ExcludedGroups = 0
    $WeeklyItems.Excluded | ForEach-Object {
        $ExcludedGroups = $ExcludedGroups + $_
    }
    $FailedGroups = 0
    $WeeklyItems.Failed | ForEach-Object {
        $FailedGroups = $FailedGroups + $_
    }
    $Hyperlink = New-Object Microsoft.SharePoint.Client.FieldURLValue
    $Hyperlink.Description = $ReportTitleDate
    $Hyperlink.Url = "https://domain1.sharepoint.com/sites/Developer/Termination Records/Weekly/$($Hyperlink.Description).html"
    $CompletedBy = $WeeklyItems.CompletedBy | Sort-Object -Unique
    $TimeStamp = $WeeklyItems.TimeStamp | Sort-Object -Unique | Select-Object -First 1
    $obj = [PSCustomObject]@{
        ChangeOrder     = $ChangeOrder
        EmployeeID      = $EmployeeID
        Accounts        = $Accounts
        Domains         = $Domains
        SuccessfulSteps = $SuccessfulSteps
        SkippedSteps    = $SkippedSteps
        FailedSteps     = $FailedSteps
        RemovedGroups   = $RemovedGroups
        ExcludedGroups  = $ExcludedGroups
        FailedGroups    = $FailedGroups
        Hyperlink       = $Hyperlink
        CompletedBy     = $CompletedBy
        TimeStamp       = $TimeStamp
    }
    $UploadItems.Add($obj)
}

$List = $Context.Web.Lists.GetByTitle($MonthlyListTitle)
$Context.Load($List)
$Context.ExecuteQuery()

$UploadItems | ForEach-Object {
    $ListItemCreationInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
    $NewListItem = $List.AddItem($ListItemCreationInfo)
    $NewListItem["Title"] = $_.ChangeOrder
    $NewListItem["EmployeeID"] = $_.EmployeeID
    $NewListItem["Account"] = $_.Accounts
    $NewListItem["Domains"] = $_.Domains
    $NewListItem["SuccessfulSteps"] = $_.SuccessfulSteps
    $NewListItem["SkippedSteps"] = $_.SkippedSteps
    $NewListItem["FailedSteps"] = $_.FailedSteps
    $NewListItem["RemovedGroups"] = $_.RemovedGroups
    $NewListItem["ExcludedGroups"] = $_.ExcludedGroups
    $NewListItem["FailedGroups"] = $_.FailedGroups
    $NewListItem["Report"] = [Microsoft.SharePoint.Client.FieldUrlValue]$($_.Hyperlink)
    $NewListItem["CompletedBy"] = $_.CompletedBy
    $NewListItem["TimeStamp"] = $_.TimeStamp
    $NewListItem.Update()
    $Context.ExecuteQuery()
}

$DailyReports | ForEach-Object {
    $File = $Context.Web.GetFileByServerRelativeUrl($_.ServerRelativeUrl)
    $Context.Load($File)
    $Context.ExecuteQuery()
    $File.DeleteObject()
    $Context.ExecuteQuery()
}

$WeeklyListItems | ForEach-Object {
    Remove-SPOListItem -ListTitle $WeeklyListTitle -ItemID $_.ID | Out-Null
}

Remove-Item -Path $TemporaryDirectory -Recurse