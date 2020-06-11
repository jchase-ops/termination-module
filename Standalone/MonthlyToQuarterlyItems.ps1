<#
Retrieve all Monthly Items and Weekly Records from SharePoint and combine into Quarterly report for Termination Module
#>

Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
Add-Type -Path $(${env:ProgramFiles(x86)} + '\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll')
Import-Module "C:\Program Files\WindowsPowerShell\Modules\SPOMod\SPOMod.psm1"
Import-Module -Name ReportHTML

$Url = "https://domain1.sharepoint.com/sites/Developer"
$MonthlyListTitle = "Monthly Termination Results"
$DocLib = "Termination Records"
$Credential = [PSCredential]::New("SvcAcctUsername@domain1.com", $(ConvertTo-SecureString -String "SvcAcctPW" -AsPlainText -Force))
$TemporaryDirectory = New-Item -Name TermReports -Path $env:TEMP -ItemType Directory

$CurrentDate = New-Object -Type DateTime -ArgumentList (Get-Date).Year, (Get-Date).Month, (Get-Date).Day, 6, 0, 0, 0

$StartDate = $CurrentDate.AddMonths(-1)
$EndDate = $CurrentDate.AddDays(-1)

$ReportTitleDate = $StartDate.ToString("ddMMMyyyy") + "-" + $EndDate.ToString("ddMMMyyyy")

Connect-SPOCSOM -Url $Url -Credential $Credential

$MonthlyListItems = Get-SPOListItems -ListTitle $MonthlyListTitle -IncludeAllProperties $true | Where-Object { $_.TimeStamp.ToDateTime($null) -lt $EndDate -or $_.Created.ToLocalTime() -lt $EndDate }

$WeeklyReports = Get-SPOListItems -ListTitle $DocLib -IncludeAllProperties $true -Recursive | Where-Object { $_.FsObjType -eq "0" -and $_.FileDirRef -like "*Weekly" }

$WebClient = New-Object System.Net.WebClient
$WebClient.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.Username, $Credential.Password)
$WebClient.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")

ForEach ($Report in $WeeklyReports) {
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
    $HTML.All.Tags("LI") | ForEach-Object {
        $DateProcessed = $_.InnerText.Trim()
        $TabData = $HTML.GetElementById($DateProcessed)
        if ($DateProcessed -Match "^\d{2}[a-zA-Z]{3}\d{2}$") { $DateProcessed = $DateProcessed -Replace "\d{2}$", "2019" }
        if (($TabData.InnerText -Split "\r?\n").Count -gt 1) {
            ($TabData.All.Tags("A") | Where-Object { $_.InnerText -Match "\d{6}" } | Sort-Object InnerText -Unique).parentElement.parentElement | ForEach-Object {
                $ChangeOrderHTML = $_
                $ChangeOrder = ($ChangeOrderHTML.All.Tags("A") | Where-Object { $_.InnerText -Match "\d{6}" }).InnerText
                $ChangeOrderObj = [PSCustomObject]@{
                    ChangeOrder   = $ChangeOrder
                    WeekProcessed = $File.Name -Replace "\..*$", ""
                    DateProcessed = $DateProcessed
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
    }
}

$WeekArray = @()

$ItemsFromRecord | Select-Object WeekProcessed -Unique | ForEach-Object {
    $obj = [PSCustomObject]@{
        WeekStart = [DateTime]::ParseExact($($_.WeekProcessed -Replace "\-.*$", ""), "ddMMMyyyy", $null).AddHours(6)
        WeekEnd = [DateTime]::ParseExact($($_.WeekProcessed -Replace "^.*\-", ""), "ddMMMyyyy", $null).AddHours(6)
    }
    $WeekArray += $obj
}

$WeekArray = $WeekArray | Sort-Object WeekStart

$TabArray = @()

$WeekArray | ForEach-Object {
    $TabArray += $_.WeekStart.ToString("ddMMMyyyy") + "-" + $_.WeekEnd.ToString("ddMMMyyyy")
}

$Report = @()
$TitleText = "Termination Results for $ReportTitleDate"
$LogoPath = "$PSScriptRoot\Resources"
$LeftLogoName = "Patch"
$RightLogoName = "Logo"

$Report += Get-HTMLOpenPage -TitleText $TitleText -LogoPath $LogoPath -LeftLogoName $LeftLogoName -RightLogoName $RightLogoName
$Report += Get-HTMLTabHeader -TabNames $TabArray
$TabArray | ForEach-Object {
    $Tab = $_
    $WeekItems = $ItemsFromRecord | Where-Object { $_.WeekProcessed -eq $Tab }
    $Days = @()
    $WeekItems | ForEach-Object {
        $Days += [DateTime]::ParseExact($_.DateProcessed, "ddMMMyyyy", $null).AddHours(6)
    }
    $Days = $Days | Sort-Object -Unique
    $Report += Get-HTMLTabContentOpen -TabName $Tab -TabHeading "Termination Results: $Tab"
    ForEach ($Day in $Days) {
        $DayItems = $WeekItems | Where-Object { $_.DateProcessed -eq $Day.ToString("ddMMMyyyy") }
        $Report += Get-HTMLContentOpen -HeaderText $Day.ToString("ddMMMyyyy") -IsHidden
        $DayItems | Sort-Object ChangeOrder | ForEach-Object {
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
        $Report += Get-HTMLContentClose
    }
    $Report += Get-HTMLTabContentClose
}
$Report += Get-HTMLClosePage

$SavedReport = Save-HTMLReport -ReportContent $Report -ReportName $ReportTitleDate -ReportPath $TemporaryDirectory

$Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
$Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName, $Credential.Password)
$List = $Context.Web.Lists.GetByTitle($DocLib)
$Context.Load($List.RootFolder)
$Context.ExecuteQuery()

$FileStream = New-Object IO.FileStream($SavedReport, [System.IO.FileMode]::Open)
$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
$FileCreationInfo.ContentStream = $FileStream
$FileCreationInfo.Url = $List.RootFolder.ServerRelativeUrl + "/Monthly/" + $($SavedReport -Replace "^.*\\", "")
$UploadFile = $List.RootFolder.Files.Add($FileCreationInfo)
$Context.Load($UploadFile)
$Context.ExecuteQuery()

$MonthlyListItems | ForEach-Object {
    Remove-SPOListItem -ListTitle $MonthlyListTitle -ItemID $_.ID | Out-Null
}

Remove-Item -Path $TemporaryDirectory -Recurse