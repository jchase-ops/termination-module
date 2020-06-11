<#
NewTeamsTerminationNotification method for Termination Module
#>

Function New-TeamsTerminationNotification {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Array]
        $Terminations
    )

    $AccountTypes = "Standard", "Admin", "Mailbox"
    $Domains = "DOMAIN1", "DOMAIN2", "DOMAIN3"
    $GroupResults = "Removed", "Excluded", "Failed"
    $StepResults = "Success", "Skipped", "Failed"
    $AllTerminationResultsPrefix = "https://$SharePointURL/sites/Developer/Lists/Daily%20Termination%20Results/AllItems.aspx"
    $GroupsUriPrefix = "https://$SharePointURL/sites/Developer/Lists/Group%20Termination%20Results/AllItems.aspx"

    $CountObject = [PSCustomObject]@{
        Accounts     = [PSCustomObject]@{
            Standard = [PSCustomObject]@{
                DOMAIN1    = 0
                DOMAIN2 = 0
                DOMAIN3    = 0
            }
            Admin    = [PSCustomObject]@{
                DOMAIN1    = 0
                DOMAIN2 = 0
                DOMAIN3    = 0
            }
            Mailbox  = [PSCustomObject]@{
                DOMAIN1    = 0
                DOMAIN2 = 0
                DOMAIN3    = 0
            }
        }
        Groups       = [PSCustomObject]@{
            Removed  = [PSCustomObject]@{
                DOMAIN1    = 0
                DOMAIN2 = 0
                DOMAIN3    = 0
            }
            Excluded = [PSCustomObject]@{
                DOMAIN1    = 0
                DOMAIN2 = 0
                DOMAIN3    = 0
            }
            Failed   = [PSCustomObject]@{
                DOMAIN1    = 0
                DOMAIN2 = 0
                DOMAIN3    = 0
            }
        }
        Steps        = [PSCustomObject]@{
            Success = [PSCustomObject]@{
                DOMAIN1    = 0
                DOMAIN2 = 0
                DOMAIN3    = 0
            }
            Skipped = [PSCustomObject]@{
                DOMAIN1    = 0
                DOMAIN2 = 0
                DOMAIN3    = 0
            }
            Failed  = [PSCustomObject]@{
                DOMAIN1    = 0
                DOMAIN2 = 0
                DOMAIN3    = 0
            }
        }
        Terminations = 0
    }

    $CountObject.Terminations = @($Terminations).Count
    $Terminations | ForEach-Object {
        $Termination = $_
        ForEach ($AccountType in $AccountTypes) {
            ForEach ($Domain in $Domains) {
                if ($null -ne $Termination.AccountData.$AccountType.$Domain) {
                    $CountObject.Accounts.$AccountType.$Domain++
                    ForEach ($GroupResult in $GroupResults) {
                        $CountObject.Groups.$GroupResult.$Domain = $CountObject.Groups.$GroupResult.$Domain + @($Termination.AccountData.$AccountType.$Domain.Groups | Where-Object { $_.GroupResult -eq $GroupResult }).Count
                    }
                    ForEach ($StepResult in $StepResults) {
                        $CountObject.Steps.$StepResult.$Domain = $CountObject.Steps.$StepResult.$Domain + @($Termination.Results | Where-Object { $_.AccountType -eq $AccountType -and $_.Domain -eq $Domain -and $_.Status -eq $StepResult }).Count
                    }
                }
            }
        }
    }
    $AccountFacts = @()
    $AccountFacts += [ordered]@{ name = "Standard"; value = $($($CountObject.Accounts.Standard.DOMAIN1) + $($CountObject.Accounts.Standard.DOMAIN2) + $($CountObject.Accounts.Standard.DOMAIN3)) }
    $AccountFacts += [ordered]@{ name = "Admin"; value = $($($CountObject.Accounts.Admin.DOMAIN1) + $($CountObject.Accounts.Admin.DOMAIN2) + $($CountObject.Accounts.Admin.DOMAIN3)) }

    $StepFacts = @()
    $StepFacts += [ordered]@{ name = "Successful"; value = "<font style=`"color:lime`">$($($CountObject.Steps.Success.DOMAIN1) + $($CountObject.Steps.Success.DOMAIN2) + $($CountObject.Steps.Success.DOMAIN3))</font>" }
    $StepFacts += [ordered]@{ name = "Skipped"; value = "<font style=`"color:gold`">$($($CountObject.Steps.Skipped.DOMAIN1) + $($CountObject.Steps.Skipped.DOMAIN2) + $($CountObject.Steps.Skipped.DOMAIN3))</font>" }
    $StepFacts += [ordered]@{ name = "Failed"; value = "<font style=`"color:red`">$($($CountObject.Steps.Failed.DOMAIN1) + $($CountObject.Steps.Failed.DOMAIN2) + $($CountObject.Steps.Failed.DOMAIN3))</font>" }

    $GroupFacts = @()
    $GroupFacts += [ordered]@{ name = "Removed"; value = "<font style=`"color:lime`">$($($CountObject.Groups.Removed.DOMAIN1) + $($CountObject.Groups.Removed.DOMAIN2) + $($CountObject.Groups.Removed.DOMAIN3))</font>" }
    $GroupFacts += [ordered]@{ name = "Excluded"; value = "<font style=`"color:gold`">$($($CountObject.Groups.Excluded.DOMAIN1) + $($CountObject.Groups.Excluded.DOMAIN2) + $($CountObject.Groups.Excluded.DOMAIN3))</font>" }
    $GroupFacts += [ordered]@{ name = "Failed"; value = "<font style=`"color:red`">$($($CountObject.Groups.Failed.DOMAIN1) + $($CountObject.Groups.Failed.DOMAIN2) + $($CountObject.Groups.Failed.DOMAIN3))</font>" }

    $PotentialAction = @()
            
    $Termination = $Terminations

    $PotentialAction += [ordered]@{
        "@context" = "http://schema.org"
        "@type"    = "OpenUri"
        name       = "View All Results"
        targets    = @(
            [ordered]@{ os = "default"; uri = $AllTerminationResultsPrefix }
        )
    }
    $PotentialAction += [ordered]@{
        "@context" = "http://schema.org"
        "@type"    = "OpenUri"
        name       = "View All Groups"
        targets    = @(
            [ordered]@{ os = "default"; uri = $GroupsUriPrefix }
        )
    }
    $PotentialAction += [ordered]@{
        "@context" = "http://schema.org"
        "@type"    = "OpenUri"
        name       = "View Failed Steps"
        targets    = @(
            [ordered]@{ os = "default"; uri = $AllTerminationResultsPrefix + "?FilterField1=Status&FilterValue1=Failed&FilterType1=Choice" }
        )
    }
    $PotentialAction += [ordered]@{
        "@context" = "http://schema.org"
        "@type"    = "OpenUri"
        name       = "View Failed Groups"
        targets    = @(
            [ordered]@{ os = "default"; uri = $GroupsUriPrefix + "?FilterField1=Status&FilterValue1=Failed&FilterType1=Choice" }
        )
    }
            
    $Notification = [PSCustomObject]@{
        "@type"         = "MessageCard"
        "@context"      = "http://schema.org/extensions"
        themeColor      = "A02AE0"
        summary         = "Termination Results"
        title           = "PS Notifications"
        sections        = @(
            [ordered]@{
                activityTitle    = "Termination Results"
                activitySubtitle = "Processed at: " + (Get-Date).ToShortTimeString()
                activityText     = "Terminations: " + $CountObject.Terminations
            }
            [ordered]@{
                title = "Accounts Processed"
                facts = $AccountFacts
            }
            [ordered]@{
                title = "Step Results"
                facts = $StepFacts
            }
            [ordered]@{
                title = "Group Results"
                facts = $GroupFacts
            }
        )
        potentialAction = $PotentialAction
    }

    $Body = $Notification | ConvertTo-Json -Depth 99

    Invoke-RestMethod -Method Post -Body $Body -Uri $TeamsUri -ContentType "application/json"
}