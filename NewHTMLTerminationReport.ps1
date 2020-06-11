<#
NewHTMLTerminationReport method for Termination Module
#>

Function global:New-HTMLTerminationReport {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Array]
        $Groups,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.Array]
        $Results
    )

    $AccountTypes = @(
        "Standard"
        "Admin"
        "Mailbox"
    )

    $Domains = @(
        "DOMAIN1"
        "DOMAIN2"
        "DOMAIN3"
    )

    $StepTypes = @(
        "Clear-HomeDirectory"
        "Clear-ScriptPath"
        "Disable-DialInAccess"
        "Disable-EmployeeSIP"
        "Enable-MailboxForwarding"
        "Move-BatchFileToArchive"
        "Move-MailboxToTerminatedDatabase"
        "Remove-EmployeeGroups"
        "Remove-InboxRules"
        "Set-ContactHiddenFromGAL"
        "Set-MailboxHiddenFromGAL"
        "Set-RandomPassword"
        "Set-TerminatedOrganizationalUnit"
        "Set-TerminatedOutOfOffice"
    )

    $global:Report = @()

    $global:TitleText = "Termination Results"
    $global:LogoPath = "$PSScriptRoot\Resources\"
    $global:LeftLogoName = "Patch"
    $global:RightLogoName = "Logo"
    $global:TabArray = @("Summary", "Steps", "Accounts", "Domains", "Groups", "Change Orders", "All")

    $global:ResultCounts = [PSCustomObject]@{
        Accounts = [PSCustomObject]@{
            Total = @($Results | Where-Object { $_.AccountType -eq "Standard" -or $_.AccountType -eq "Admin" } | Sort-Object Account, Domain -Unique).Count
            Standard = [PSCustomObject]@{
                Total = @($Results | Where-Object { $_.AccountType -eq "Standard" } | Sort-Object Account, Domain -Unique).Count
                DOMAIN1 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN1" } | Sort-Object Account -Unique).Count
                DOMAIN2 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN2" } | Sort-Object Account -Unique).Count
                DOMAIN3 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN3" } | Sort-Object Account -Unique).Count
            }
            Admin = [PSCustomObject]@{
                Total = @($Results | Where-Object { $_.AccountType -eq "Admin" } | Sort-Object Account, Domain -Unique).Count
                DOMAIN1 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN1" } | Sort-Object Account -Unique).Count
                DOMAIN2 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN2" } | Sort-Object Account -Unique).Count
                DOMAIN3 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN3" } | Sort-Object Account -Unique).Count
            }
        }
        ChartDataSets = [PSCustomObject]@{
            SummaryPieDataOne = $Results | Where-Object { $_.AccountType -eq "Standard" -or $_.AccountType -eq "Admin" } | Sort-Object Account -Unique | Group-Object AccountType -NoElement | Sort-Object Count -Descending
            SummaryPieDataTwo = $Results | Where-Object { $_.AccountType -eq "Standard" } | Sort-Object Account, Domain -Unique | Group-Object Domain -NoElement | Sort-Object Count -Descending
            SummaryPieDataThree = $Results | Where-Object { $_.AccountType -eq "Admin" } | Sort-Object Account, Domain -Unique | Group-Object Domain -NoElement | Sort-Object Count -Descending
            StepPieDataOne = $Results | Where-Object { $_.Status -ne "Start" -and $_.Status -ne "Complete" } | Group-Object Status -NoElement | Sort-Object Count -Descending
            StepPieDataTwo = $Results | Where-Object { $_.Status -ne "Start" -and $_.Status -ne "Complete" -and $_.AccountType -eq "Standard" } | Group-Object Status -NoElement | Sort-Object Count -Descending
            StepPieDataThree = $Results | Where-Object { $_.Status -ne "Start" -and $_.Status -ne "Complete" -and $_.AccountType -eq "Admin" } | Group-Object Status -NoElement | Sort-Object Count -Descending
            StepPieDataFour = $Results | Where-Object { $_.Status -ne "Start" -and $_.Status -ne "Complete" -and $_.Domain -eq "DOMAIN1" } | Group-Object Status -NoElement | Sort-Object Count -Descending
            StepPieDataFive = $Results | Where-Object { $_.Status -ne "Start" -and $_.Status -ne "Complete" -and $_.Domain -eq "DOMAIN2" } | Group-Object Status -NoElement | Sort-Object Count -Descending
            StepPieDataSix = $Results | Where-Object { $_.Status -ne "Start" -and $_.Status -ne "Complete" -and $_.Domain -eq "DOMAIN3" } | Group-Object Status -NoElement | Sort-Object Count -Descending
        }
        Groups = [PSCustomObject]@{
            Total = @($Groups | Sort-Object DistinguishedName -Unique).Count
            Standard = @($Groups | Where-Object { $_.AccountType -eq "Standard" } | Sort-Object DistinguishedName -Unique).Count
            Admin = @($Groups | Where-Object { $_.AccountType -eq "Admin" } | Sort-Object DistinguishedName -Unique).Count
            Removed = [PSCustomObject]@{
                Total = @($Groups | Where-Object { $_.GroupResult -eq "Removed" }).Count
                DOMAIN1 = @($Groups | Where-Object { $_.GroupResult -eq "Removed" -and $_.Domain -eq "DOMAIN1" }).Count
                DOMAIN2 = @($Groups | Where-Object { $_.GroupResult -eq "Removed" -and $_.Domain -eq "DOMAIN2" }).Count
                DOMAIN3 = @($Groups | Where-Object { $_.GroupResult -eq "Removed" -and $_.Domain -eq "DOMAIN3" }).Count
                Standard = [PSCustomObject]@{
                    Total = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.GroupResult -eq "Removed" }).Count
                    DOMAIN1 = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN1" -and $_.GroupResult -eq "Removed" }).Count
                    DOMAIN2 = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN2" -and $_.GroupResult -eq "Removed" }).Count
                    DOMAIN3 = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN3" -and $_.GroupResult -eq "Removed" }).Count
                }
                Admin = [PSCustomObject]@{
                    Total = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.GroupResult -eq "Removed" }).Count
                    DOMAIN1 = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN1" -and $_.GroupResult -eq "Removed" }).Count
                    DOMAIN2 = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN2" -and $_.GroupResult -eq "Removed" }).Count
                    DOMAIN3 = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN3" -and $_.GroupResult -eq "Removed" }).Count
                }
            }
            Excluded = [PSCustomObject]@{
                Total = @($Groups | Where-Object { $_.GroupResult -eq "Excluded" }).Count
                DOMAIN1 = @($Groups | Where-Object { $_.GroupResult -eq "Excluded" -and $_.Domain -eq "DOMAIN1" }).Count
                DOMAIN2 = @($Groups | Where-Object { $_.GroupResult -eq "Excluded" -and $_.Domain -eq "DOMAIN2" }).Count
                DOMAIN3 = @($Groups | Where-Object { $_.GroupResult -eq "Excluded" -and $_.Domain -eq "DOMAIN3" }).Count
                Standard = [PSCustomObject]@{
                    Total = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.GroupResult -eq "Excluded" }).Count
                    DOMAIN1 = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN1" -and $_.GroupResult -eq "Excluded" }).Count
                    DOMAIN2 = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN2" -and $_.GroupResult -eq "Excluded" }).Count
                    DOMAIN3 = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN3" -and $_.GroupResult -eq "Excluded" }).Count
                }
                Admin = [PSCustomObject]@{
                    Total = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.GroupResult -eq "Excluded" }).Count
                    DOMAIN1 = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN1" -and $_.GroupResult -eq "Excluded" }).Count
                    DOMAIN2 = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN2" -and $_.GroupResult -eq "Excluded" }).Count
                    DOMAIN3 = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN3" -and $_.GroupResult -eq "Excluded" }).Count
                }
            }
            Failed = [PSCustomObject]@{
                Total = @($Groups | Where-Object { $_.GroupResult -eq "Failed" }).Count
                DOMAIN1 = @($Groups | Where-Object { $_.GroupResult -eq "Failed" -and $_.Domain -eq "DOMAIN1" }).Count
                DOMAIN2 = @($Groups | Where-Object { $_.GroupResult -eq "Failed" -and $_.Domain -eq "DOMAIN2" }).Count
                DOMAIN3 = @($Groups | Where-Object { $_.GroupResult -eq "Failed" -and $_.Domain -eq "DOMAIN3" }).Count
                Standard = [PSCustomObject]@{
                    Total = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.GroupResult -eq "Failed" }).Count
                    DOMAIN1 = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN1" -and $_.GroupResult -eq "Failed" }).Count
                    DOMAIN2 = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN2" -and $_.GroupResult -eq "Failed" }).Count
                    DOMAIN3 = @($Groups | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN3" -and $_.GroupResult -eq "Failed" }).Count
                }
                Admin = [PSCustomObject]@{
                    Total = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.GroupResult -eq "Failed" }).Count
                    DOMAIN1 = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN1" -and $_.GroupResult -eq "Failed" }).Count
                    DOMAIN2 = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN2" -and $_.GroupResult -eq "Failed" }).Count
                    DOMAIN3 = @($Groups | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN3" -and $_.GroupResult -eq "Failed" }).Count
                }
            }
        }
        Steps = [PSCustomObject]@{
            Total = @($Results | Where-Object { $_.Status -eq "Start" }).Count
            Successful = [PSCustomObject]@{
                Total = @($Results | Where-Object { $_.Status -eq "Success" }).Count
                DOMAIN1 = @($Results | Where-Object { $_.Status -eq "Success" -and $_.Domain -eq "DOMAIN1" }).Count
                DOMAIN2 = @($Results | Where-Object { $_.Status -eq "Success" -and $_.Domain -eq "DOMAIN2" }).Count
                DOMAIN3 = @($Results | Where-Object { $_.Status -eq "Success" -and $_.Domain -eq "DOMAIN3" }).Count
                Standard = [PSCustomObject]@{
                    Total = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Status -eq "Success" }).Count
                    DOMAIN1 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN1" -and $_.Status -eq "Success" }).Count
                    DOMAIN2 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN2" -and $_.Status -eq "Success" }).Count
                    DOMAIN3 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN3" -and $_.Status -eq "Success" }).Count
                }
                Admin = [PSCustomObject]@{
                    Total = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Status -eq "Success" }).Count
                    DOMAIN1 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN1" -and $_.Status -eq "Success" }).Count
                    DOMAIN2 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN2" -and $_.Status -eq "Success" }).Count
                    DOMAIN3 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN3" -and $_.Status -eq "Success" }).Count
                }
            }
            Skipped = [PSCustomObject]@{
                Total = @($Results | Where-Object { $_.Status -eq "Skipped" }).Count
                DOMAIN1 = @($Results | Where-Object { $_.Status -eq "Skipped" -and $_.Domain -eq "DOMAIN1" }).Count
                DOMAIN2 = @($Results | Where-Object { $_.Status -eq "Skipped" -and $_.Domain -eq "DOMAIN2" }).Count
                DOMAIN3 = @($Results | Where-Object { $_.Status -eq "Skipped" -and $_.Domain -eq "DOMAIN3" }).Count
                Standard = [PSCustomObject]@{
                    Total = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Status -eq "Skipped" }).Count
                    DOMAIN1 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN1" -and $_.Status -eq "Skipped" }).Count
                    DOMAIN2 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN2" -and $_.Status -eq "Skipped" }).Count
                    DOMAIN3 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN3" -and $_.Status -eq "Skipped" }).Count
                }
                Admin = [PSCustomObject]@{
                    Total = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Status -eq "Skipped" }).Count
                    DOMAIN1 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN1" -and $_.Status -eq "Skipped" }).Count
                    DOMAIN2 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN2" -and $_.Status -eq "Skipped" }).Count
                    DOMAIN3 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN3" -and $_.Status -eq "Skipped" }).Count
                }
            }
            Failed = [PSCustomObject]@{
                Total = @($Results | Where-Object { $_.Status -eq "Failed" }).Count
                DOMAIN1 = @($Results | Where-Object { $_.Status -eq "Failed" -and $_.Domain -eq "DOMAIN1" }).Count
                DOMAIN2 = @($Results | Where-Object { $_.Status -eq "Failed" -and $_.Domain -eq "DOMAIN2" }).Count
                DOMAIN3 = @($Results | Where-Object { $_.Status -eq "Failed" -and $_.Domain -eq "DOMAIN3" }).Count
                Standard = [PSCustomObject]@{
                    Total = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Status -eq "Failed" }).Count
                    DOMAIN1 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN1" -and $_.Status -eq "Failed" }).Count
                    DOMAIN2 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN2" -and $_.Status -eq "Failed" }).Count
                    DOMAIN3 = @($Results | Where-Object { $_.AccountType -eq "Standard" -and $_.Domain -eq "DOMAIN3" -and $_.Status -eq "Failed" }).Count
                }
                Admin = [PSCustomObject]@{
                    Total = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Status -eq "Failed" }).Count
                    DOMAIN1 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN1" -and $_.Status -eq "Failed" }).Count
                    DOMAIN2 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN2" -and $_.Status -eq "Failed" }).Count
                    DOMAIN3 = @($Results | Where-Object { $_.AccountType -eq "Admin" -and $_.Domain -eq "DOMAIN3" -and $_.Status -eq "Failed" }).Count
                }
            }
        }
        Terminations = @($Results | Sort-Object ChangeOrder -Unique).Count
    }

    $TableData = [PSCustomObject]@{
        Summary = [PSCustomObject]@{
            AccountSummaryTableDataByDomain = New-Object System.Collections.Generic.List[System.Object]
            Step = [PSCustomObject]@{
                StepSummaryTableDataByAccount = New-Object System.Collections.Generic.List[System.Object]
                StepSummaryTableDataByAccountByDomain = New-Object System.Collections.Generic.List[System.Object]
            }
            Group = [PSCustomObject]@{
                GroupTableDataByAccount = New-Object System.Collections.Generic.List[System.Object]
                GroupTableDataByAccountByDomain = New-Object System.Collections.Generic.List[System.Object]
            }
        }
        Step = [PSCustomObject]@{
            DetailedTableDataForSuccessfulSteps = $Results | Where-Object { $_.Status -eq "Success" } | Select-Object Account, ChangeOrder, Domain, EmployeeID, Section, Step
            DetailedTableDataForSkippedSteps = $Results | Where-Object { $_.Status -eq "Skipped" } | Select-Object Account, ChangeOrder, Domain, EmployeeID, Section, Step
            DetailedTableDataForFailedSteps = $Results | Where-Object { $_.Status -eq "Failed" } | Select-Object Account, ChangeOrder, Domain, EmployeeID, Section, Step
            StepStatusTableData = New-Object System.Collections.Generic.List[System.Object]
            StepStatusTableDataByAccount = New-Object System.Collections.Generic.List[System.Object]
            StepStatusTableDataByDomain = New-Object System.Collections.Generic.List[System.Object]
            StepStatusTableDataByAccountByDomain = New-Object System.Collections.Generic.List[System.Object]
        }
    }
    
    ForEach ($AccountType in $AccountTypes) {
        $obj = [PSCustomObject]@{
            AccountType = $AccountType
            DOMAIN1 = $ResultCounts.Accounts.$AccountType.DOMAIN1
            DOMAIN2 = $ResultCounts.Accounts.$AccountType.DOMAIN2
            DOMAIN3 = $ResultCounts.Accounts.$AccountType.DOMAIN3
        }
        $TableData.Summary.AccountSummaryTableDataByDomain.Add($obj)
        $obj = [PSCustomObject]@{
            AccountType = $AccountType
            Successful = $ResultCounts.Steps.Successful.$AccountType.Total
            Skipped = $ResultCounts.Steps.Skipped.$AccountType.Total
            Failed = $ResultCounts.Steps.Failed.$AccountType.Total
        }
        $TableData.Summary.Step.StepSummaryTableDataByAccount.Add($obj)
        $obj = [PSCustomObject]@{
            AccountType = $AccountType
            Removed = $ResultCounts.Groups.Removed.$AccountType.Total
            Excluded = $ResultCounts.Groups.Excluded.$AccountType.Total
            Failed = $ResultCounts.Groups.Excluded.$AccountType.Total
        }
        $TableData.Summary.Group.GroupTableDataByAccount.Add($obj)
        ForEach ($Domain in $Domains) {
            $obj = [PSCustomObject]@{
                AccountType = $AccountType
                Domain = $Domain
                Successful = $ResultCounts.Steps.Successful.$AccountType.$Domain
                Skipped = $ResultCounts.Steps.Skipped.$AccountType.$Domain
                Failed =  $ResultCounts.Steps.Failed.$AccountType.$Domain
            }
            $TableData.Summary.Step.StepSummaryTableDataByAccountByDomain.Add($obj)
            $obj = [PSCustomObject]@{
                AccountType = $AccountType
                Domain = $Domain
                Removed = $ResultCounts.Groups.Removed.$AccountType.$Domain
                Excluded = $ResultCounts.Groups.Excluded.$AccountType.$Domain
                Failed = $ResultCounts.Groups.Failed.$AccountType.$Domain
            }
            $TableData.Summary.Group.GroupTableDataByAccountByDomain.Add($obj)
        }
        ForEach ($StepType in $StepTypes) {
            if ($StepType -notin $TableData.Step.StepStatusTableData.StepType) {
                $obj = [PSCustomObject]@{
                    Step = $StepType
                    Successful = @($Results | Where-Object { $_.Step -eq $StepType -and $_.Status -eq "Success" }).Count
                    Skipped = @($Results | Where-Object { $_.Step -eq $StepType -and $_.Status -eq "Skipped" }).Count
                    Failed = @($Results | Where-Object { $_.Step -eq $StepType -and $_.Status -eq "Failed" }).Count
                }
                $TableData.Step.StepStatusTableData.Add($obj)
            }
            $obj = [PSCustomObject]@{
                AccountType = $AccountType
                Step = $StepType
                Successful = @($Results | Where-Object { $_.AccountType -eq $AccountType -and $_.Step -eq $StepType -and $_.Status -eq "Success" }).Count
                Skipped = @($Results | Where-Object { $_.AccountType -eq $AccountType -and $_.Step -eq $StepType -and $_.Status -eq "Skipped" }).Count
                Failed = @($Results | Where-Object { $_.AccountType -eq $AccountType -and $_.Step -eq $StepType -and $_.Status -eq "Failed" }).Count
            }
            $TableData.Step.StepStatusTableDataByAccount.Add($obj)
            ForEach ($Domain in $Domains) {
                $obj = [PSCustomObject]@{
                    Domain = $Domain
                    Step = $StepType
                    Successful = @($Results | Where-Object { $_.Domain -eq $Domain -and $_.Step -eq $StepType -and $_.Status -eq "Success" }).Count
                    Skipped = @($Results | Where-Object { $_.Domain -eq $Domain -and $_.Step -eq $StepType -and $_.Status -eq "Skipped" }).Count
                    Failed = @($Results | Where-Object { $_.Domain -eq $Domain -and $_.Step -eq $StepType -and $_.Status -eq "Failed" }).Count
                }
                $TableData.Step.StepStatusTableDataByDomain.Add($obj)
                $obj = [PSCustomObject]@{
                    AccountType = $AccountType
                    Domain = $Domain
                    Step = $StepType
                    Successful = @($Results | Where-Object { $_.AccountType -eq $AccountType -and $_.Domain -eq $Domain -and $_.Step -eq $StepType -and $_.Status -eq "Success" }).Count
                    Skipped = @($Results | Where-Object { $_.AccountType -eq $AccountType -and $_.Domain -eq $Domain -and $_.Step -eq $StepType -and $_.Status -eq "Skipped" }).Count
                    Failed = @($Results | Where-Object { $_.AccountType -eq $AccountType -and $_.Domain -eq $Domain -and $_.Step -eq $StepType -and $_.Status -eq "Failed" }).Count
                }
                $TableData.Step.StepStatusTableDataByAccountByDomain.Add($obj)
            }
        }
    }

    $Report += Get-HTMLOpenPage -TitleText $TitleText -LogoPath $LogoPath -LeftLogoName $LeftLogoName -RightLogoName $RightLogoName
        $Report += Get-HTMLTabHeader -TabNames $TabArray
        $Report += Get-HTMLTabContentOpen -TabName $TabArray[0] -TabHeading ("Termination Results: " + $TabArray[0] + " Page")
            $Report += Get-HTMLContentOpen -HeaderText ($TabArray[0] + ": Charts")
                $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
                    $SummaryPieObjectOne = Get-HTMLPieChartObject
                    $SummaryPieObjectOne.Title = "Account Types: Total"
                    $SummaryPieObjectOne.Size.Height = 400
                    $SummaryPieObjectOne.Size.Width = 400
                    $SummaryPieObjectOne.ChartStyle.ChartType = "doughnut"
                    $SummaryPieObjectOne.ChartStyle.ColorSchemeName = "Random"
                    $SummaryPieObjectOne.DataDefinition.DataNameColumnName = "Name"
                    $SummaryPieObjectOne.DataDefinition.DataValueColumnName = "Count"
                    $Report += Get-HTMLPieChart -ChartObject $SummaryPieObjectOne -DataSet $ResultCounts.ChartDataSets.SummaryPieDataOne
                $Report += Get-HTMLColumnClose
                $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
                    $SummaryPieObjectTwo = Get-HTMLPieChartObject
                    $SummaryPieObjectTwo.Title = "Standard Accounts: By Domain"
                    $SummaryPieObjectTwo.Size.Height = 400
                    $SummaryPieObjectTwo.Size.Width = 400
                    $SummaryPieObjectTwo.ChartStyle.ChartType = "doughnut"
                    $SummaryPieObjectTwo.ChartStyle.ColorSchemeName = "Random"
                    $SummaryPieObjectTwo.DataDefinition.DataNameColumnName = "Name"
                    $SummaryPieObjectTwo.DataDefinition.DataValueColumnName = "Count"
                    $Report += Get-HTMLPieChart -ChartObject $SummaryPieObjectTwo -DataSet $ResultCounts.ChartDataSets.SummaryPieDataTwo
                $Report += Get-HTMLColumnClose
                $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
                    $SummaryPieObjectThree = Get-HTMLPieChartObject
                    $SummaryPieObjectThree.Title = "Admin Accounts: By Domain"
                    $SummaryPieObjectThree.Size.Height = 400
                    $SummaryPieObjectThree.Size.Width = 400
                    $SummaryPieObjectThree.ChartStyle.ChartType = "doughnut"
                    $SummaryPieObjectThree.ChartStyle.ColorSchemeName = "Random"
                    $SummaryPieObjectThree.DataDefinition.DataNameColumnName = "Name"
                    $SummaryPieObjectThree.DataDefinition.DataValueColumnName = "Count"
                    $Report += Get-HTMLPieChart -ChartObject $SummaryPieObjectThree -DataSet $ResultCounts.ChartDataSets.SummaryPieDataThree
                $Report += Get-HTMLColumnClose
            $Report += Get-HTMLContentClose
            $Report += Get-HTMLContentOpen -HeaderText ($TabArray[0] + ": Highlights")
                $Report += Get-HTMLContentOpen -HeaderText "Totals"
                    $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 4
                        $Report += Get-HTMLContentText -Heading "Terminations" -Detail $ResultCounts.Terminations
                    $Report += Get-HTMLColumnClose
                    $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 4
                        $Report += Get-HTMLContentText -Heading "Accounts" -Detail $ResultCounts.Accounts.Total
                    $Report += Get-HTMLColumnClose
                    $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 4
                        $Report += Get-HTMLContentText -Heading "Steps Performed" -Detail $ResultCounts.Steps.Total
                    $Report += Get-HTMLColumnClose
                    $Report += Get-HTMLColumnOpen -ColumnNumber 4 -ColumnCount 4
                        $Report += Get-HTMLContentText -Heading "Groups Processed" -Detail $ResultCounts.Groups.Total
                    $Report += Get-HTMLColumnClose
                $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentOpen -HeaderText "Accounts" -IsHidden
                    $Report += Get-HTMLContentOpen -HeaderText "Totals per Account Type"
                        $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2
                            $Report += Get-HTMLContentText -Heading "Standard Accounts" -Detail $ResultCounts.Accounts.Standard.Total
                        $Report += Get-HTMLColumnClose
                        $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2
                            $Report += Get-HTMLContentText -Heading "Admin Accounts" -Detail $ResultCounts.Accounts.Admin.Total
                        $Report += Get-HTMLColumnClose
                    $Report += Get-HTMLContentClose
                    $Report += Get-HTMLContentOpen -HeaderText "Accounts per Domain"
                        $Report += Get-HTMLContentTable -ArrayOfObjects $TableData.Summary.AccountSummaryTableDataByDomain -ColumnTotals DOMAIN1, DOMAIN2, DOMAIN3 -Fixed
                    $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentOpen -HeaderText "Steps" -IsHidden
                    $Report += Get-HTMLContentOpen -HeaderText "Totals per Step Status"
                        $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
                            $Report += Get-HTMLContentText -Heading "Successful" -Detail $ResultCounts.Steps.Successful.Total
                        $Report += Get-HTMLColumnClose
                        $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
                            $Report += Get-HTMLContentText -Heading "Skipped" -Detail $ResultCounts.Steps.Skipped.Total
                        $Report += Get-HTMLColumnClose
                        $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
                            $Report += Get-HTMLContentText -Heading "Failed" -Detail $ResultCounts.Steps.Failed.Total
                        $Report += Get-HTMLColumnClose
                    $Report += Get-HTMLContentClose
                    $Report += Get-HTMLContentOpen -HeaderText "Step Status per Account Type" -IsHidden
                        $Report += Get-HTMLContentTable -ArrayOfObjects $TableData.Summary.Step.StepSummaryTableDataByAccount -ColumnTotals Successful, Skipped, Failed -Fixed
                    $Report += Get-HTMLContentClose
                    $Report += Get-HTMLContentOpen -HeaderText "Step Status per Domain" -IsHidden
                        $Report += Get-HTMLContentOpen -HeaderText "DOMAIN1"
                            $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Successful" -Detail $ResultCounts.Steps.Successful.DOMAIN1
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Skipped" -Detail $ResultCounts.Steps.Skipped.DOMAIN1
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Failed" -Detail $ResultCounts.Steps.Failed.DOMAIN1
                            $Report += Get-HTMLColumnClose
                        $Report += Get-HTMLContentClose
                        $Report += Get-HTMLContentOpen -HeaderText "DOMAIN2"
                            $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Successful" -Detail $ResultCounts.Steps.Successful.DOMAIN2
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Skipped" -Detail $ResultCounts.Steps.Skipped.DOMAIN2
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Failed" -Detail $ResultCounts.Steps.Failed.DOMAIN2
                            $Report += Get-HTMLColumnClose
                        $Report += Get-HTMLContentClose
                        $Report += Get-HTMLContentOpen -HeaderText "DOMAIN3"
                            $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Successful" -Detail $ResultCounts.Steps.Successful.DOMAIN3
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Skipped" -Detail $ResultCounts.Steps.Skipped.DOMAIN3
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Failed" -Detail $ResultCounts.Steps.Failed.DOMAIN3
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLContentOpen -HeaderText "Step Status per Account Type per Domain" -IsHidden
                                $Report += Get-HTMLContentTable -ArrayOfObjects $TableData.Summary.Step.StepSummaryTableDataByAccountByDomain -ColumnTotals Successful, Skipped, Failed -Fixed -GroupBy AccountType
                            $Report += Get-HTMLContentClose
                        $Report += Get-HTMLContentClose
                    $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentOpen -HeaderText "Groups" -IsHidden
                    $Report += Get-HTMLContentOpen -HeaderText "Totals per Group Account Type"
                        $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2
                            $Report += Get-HTMLContentText -Heading "Standard" -Detail $ResultCounts.Groups.Standard
                        $Report += Get-HTMLColumnClose
                        $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2
                            $Report += Get-HTMLContentText -Heading "Admin" -Detail $ResultCounts.Groups.Admin
                        $Report += Get-HTMLColumnClose
                    $Report += Get-HTMLContentClose
                    $Report += Get-HTMLContentOpen -HeaderText "Totals per Group Removal Status"
                        $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
                            $Report += Get-HTMLContentText -Heading "Removed" -Detail $ResultCounts.Groups.Removed.Total
                        $Report += Get-hTMLColumnClose
                        $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
                            $Report += Get-HTMLContentText -Heading "Excluded" -Detail $ResultCounts.Groups.Excluded.Total
                        $Report += Get-HTMLColumnClose
                        $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
                            $Report += Get-HTMLContentText -Heading "Failed" -Detail $ResultCounts.Groups.Failed.Total
                        $Report += Get-HTMLColumnClose
                    $Report += Get-HTMLContentClose
                    $Report += Get-HTMLContentOpen -HeaderText "Group Removal Status per Account Type" -IsHidden
                        $Report += Get-HTMLContentTable -ArrayOfObjects $TableData.Summary.Group.GroupTableDataByAccount -ColumnTotals Removed, Excluded, Failed -Fixed
                    $Report += Get-HTMLContentClose
                    $Report += Get-HTMLContentOpen -HeaderText "Status per Domain" -IsHidden
                        $Report += Get-HTMLContentOpen -HeaderText "DOMAIN1"
                            $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Removed" -Detail $ResultCounts.Groups.Removed.DOMAIN1
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Excluded" -Detail $ResultCounts.Groups.Excluded.DOMAIN1
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Failed" -Detail $ResultCounts.Groups.Failed.DOMAIN1
                            $Report += Get-HTMLColumnClose
                        $Report += Get-HTMLContentClose
                        $Report += Get-HTMLContentOpen -HeaderText "DOMAIN2"
                            $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Removed" -Detail $ResultCounts.Groups.Removed.DOMAIN2
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Excluded" -Detail $ResultCounts.Groups.Excluded.DOMAIN2
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Failed" -Detail $ResultCounts.Groups.Failed.DOMAIN2
                            $Report += Get-HTMLColumnClose
                        $Report += Get-HTMLContentClose
                        $Report += Get-HTMLContentOpen -HeaderText "DOMAIN3"
                            $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Removed" -Detail $ResultCounts.Groups.Removed.DOMAIN3
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Excluded" -Detail $ResultCounts.Groups.Excluded.DOMAIN3
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
                                $Report += Get-HTMLContentText -Heading "Failed" -Detail $ResultCounts.Groups.Failed.DOMAIN3
                            $Report += Get-HTMLColumnClose
                            $Report += Get-HTMLContentOpen -HeaderText "Group Removal Status per Account Type per Domain" -IsHidden
                                $Report += Get-HTMLContentTable -ArrayOfObjects $TableData.Summary.Group.GroupTableDataByAccountByDomain -ColumnTotals Removed, Excluded, Failed -Fixed -GroupBy AccountType
                            $Report += Get-HTMLContentClose
                        $Report += Get-HTMLContentClose
                    $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentClose
            $Report += Get-HTMLContentClose
        $Report += Get-HTMLTabContentClose
        $Report += Get-HTMLTabContentOpen -TabName $TabArray[1] -TabHeading ("Termination Results: " + $TabArray[1] + " Page")
            $Report += Get-HTMLContentOpen -HeaderText ($TabArray[1] + ": Charts")
                $Report += Get-HTMLContentOpen -HeaderText "Step Status Totals: All Steps & Steps per Account Type"
                    $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
                        $StepPieObjectOne = Get-HTMLPieChartObject
                        $StepPieObjectOne.Title = "All Steps: Total"
                        $StepPieObjectOne.Size.Height = 400
                        $StepPieObjectOne.Size.Width = 400
                        $StepPieObjectOne.ChartStyle.ChartType = "doughnut"
                        $StepPieObjectOne.ChartStyle.ColorSchemeName = "Random"
                        $StepPieObjectOne.DataDefinition.DataNameColumnName = "Name"
                        $StepPieObjectOne.DataDefinition.DataValueColumnName = "Count"
                        $Report += Get-HTMLPieChart -ChartObject $StepPieObjectOne -DataSet $ResultCounts.ChartDataSets.StepPieDataOne
                    $Report += Get-HTMLColumnClose
                    $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
                        $StepPieObjectTwo = Get-HTMLPieChartObject
                        $StepPieObjectTwo.Title = "Standard Account Step Status: Total"
                        $StepPieObjectTwo.Size.Height = 400
                        $StepPieObjectTwo.Size.Width = 400
                        $StepPieObjectTwo.ChartStyle.ChartType = "doughnut"
                        $StepPieObjectTwo.ChartStyle.ColorSchemeName = "Random"
                        $StepPieObjectTwo.DataDefinition.DataNameColumnName = "Name"
                        $StepPieObjectTwo.DataDefinition.DataValueColumnName = "Count"
                        $Report += Get-HTMLPieChart -ChartObject $StepPieObjectTwo -DataSet $ResultCounts.ChartDataSets.StepPieDataTwo
                    $Report += Get-HTMLColumnClose
                    $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
                        $StepPieObjectThree = Get-HTMLPieChartObject
                        $StepPieObjectThree.Title = "Admin Account Step Status: Total"
                        $StepPieObjectThree.Size.Height = 400
                        $StepPieObjectThree.Size.Width = 400
                        $StepPieObjectThree.ChartStyle.ChartType = "doughnut"
                        $StepPieObjectThree.ChartStyle.ColorSchemeName = "Random"
                        $StepPieObjectThree.DataDefinition.DataNameColumnName = "Name"
                        $StepPieObjectThree.DataDefinition.DataValueColumnName = "Count"
                        $Report += Get-HTMLPieChart -ChartObject $StepPieObjectThree -DataSet $ResultCounts.ChartDataSets.StepPieDataThree
                    $Report += Get-HTMLColumnClose
                $Report += Get-HTMLContentCLose
                $Report += Get-HTMLContentOpen -HeaderText "Step Status Totals per Domain"
                    $Report += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
                        $StepPieObjectFour = Get-HTMLPieChartObject
                        $StepPieObjectFour.Title = "DOMAIN1"
                        $StepPieObjectFour.Size.Height = 400
                        $StepPieObjectFour.Size.Width = 400
                        $StepPieObjectFour.ChartStyle.ChartType = "doughnut"
                        $StepPieObjectFour.ChartStyle.ColorSchemeName = "Random"
                        $StepPieObjectFour.DataDefinition.DataNameColumnName = "Name"
                        $StepPieObjectFour.DataDefinition.DataValueColumnName = "Count"
                        $Report += Get-HTMLPieChart -ChartObject $StepPieObjectFour -DataSet $ResultCounts.ChartDataSets.StepPieDataFour
                    $Report += Get-HTMLColumnClose
                    $Report += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
                        $StepPieObjectFive = Get-HTMLPieChartObject
                        $StepPieObjectFive.Title = "DOMAIN2"
                        $StepPieObjectFive.Size.Height = 400
                        $StepPieObjectFive.Size.Width = 400
                        $StepPieObjectFive.ChartStyle.ChartType = "doughnut"
                        $StepPieObjectFive.ChartStyle.ColorSchemeName = "Random"
                        $StepPieObjectFive.DataDefinition.DataNameColumnName = "Name"
                        $StepPieObjectFive.DataDefinition.DataValueColumnName = "Count"
                        $Report += Get-HTMLPieChart -ChartObject $StepPieObjectFive -DataSet $ResultCounts.ChartDataSets.StepPieDataFive
                    $Report += Get-HTMLColumnClose
                    $Report += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
                        $StepPieObjectSix = Get-HTMLPieChartObject
                        $StepPieObjectSix.Title = "DOMAIN3"
                        $StepPieObjectSix.Size.Height = 400
                        $StepPieObjectSix.Size.Width = 400
                        $StepPieObjectSix.ChartStyle.ChartType = "doughnut" 
                        $StepPieObjectSix.ChartStyle.ColorSchemeName = "Random"
                        $StepPieObjectSix.DataDefinition.DataNameColumnName = "Name"
                        $StepPieObjectSix.DataDefinition.DataValueColumnName = "Count"
                        $Report += Get-HTMLPieChart -ChartObject $StepPieObjectSix -DataSet $ResultCounts.ChartDataSets.StepPieDataSix
                    $Report += Get-HTMLColumnClose
                $Report += Get-HTMLContentClose
            $Report += Get-HTMLContentClose
            $Report += Get-HTMLContentOpen -HeaderText ($TabArray[1] + ": Highlights") -IsHidden
                $Report += Get-HTMLContentOpen -HeaderText "Step Status Totals" -IsHidden
                    $Report += Get-HTMLContentDataTable -ArrayOfObjects $TableData.Step.StepStatusTableData -HideFooter -PagingOptions "15,25,50,"
                $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentOpen -HeaderText "Step Status Totals per Account Type" -IsHidden
                    $Report += Get-HTMLContentDataTable -ArrayOfObjects $TableData.Step.StepStatusTableDataByAccount -HideFooter -PagingOptions "15,25,50,"
                $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentOpen -HeaderText "Step Status Totals per Domain" -IsHidden
                    $Report += Get-HTMLContentDataTable -ArrayOfObjects $TableData.Step.StepStatusTableDataByDomain -HideFooter -PagingOptions "15,25,50,"
                $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentOpen -HeaderText "Step Status Totals per Account Type per Domain" -IsHidden
                    $Report += Get-HTMLContentDataTable -ArrayOfObjects $TableData.Step.StepStatusTableDataByAccountByDomain -HideFooter -PagingOptions "15,25,50,"
                $Report += Get-HTMLContentClose
            $Report += Get-HTMLContentClose
            $Report += Get-HTMLContentOpen -HeaderText ($TabArray[1] + ": Details") -IsHidden
                $Report += Get-HTMLContentOpen -HeaderText "Successful Step Details" -IsHidden
                    $Report += Get-HTMLContentDataTable -ArrayOfObjects $TableData.Step.DetailedTableDataForSuccessfulSteps -HideFooter -PagingOptions "15,25,50,"
                $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentOpen -HeaderText "Skipped Step Details" -IsHidden
                    $Report += Get-HTMLContentDataTable -ArrayOfObjects $TableData.Step.DetailedTableDataForSkippedSteps -HideFooter -PagingOptions "15,25,50,"
                $Report += Get-HTMLContentClose
                $Report += Get-HTMLContentOpen -HeaderText "Failed Step Details" -IsHidden
                    $Report += Get-HTMLContentDataTable -ArrayOfObjects $TableData.Step.DetailedTableDataForFailedSteps -HideFooter -PagingOptions "15,25,50,"
                $Report += Get-HTMLContentClose
            $Report += Get-HTMLContentClose
        $Report += Get-HTMLTabContentClose
    $Report += Get-HTMLClosePage
    Save-HTMLReport -ReportContent $Report -ReportName "TermResultsSingleTermination" -ReportPath $HOME -ShowReport
}