<#
CopyItemsToSharePoint method for Termination Module
#>

Function Copy-ItemsToSharePoint {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.Array]
        $Groups,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.Array]
        $Results
    )

    Connect-SPOCSOM -Credential $Credential -Url $SharePointUrl

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SharePointUrl)
    $SharePointCredential = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName, $Credential.Password)
    $Context.Credentials = $SharePointCredential
    $List = $Context.Web.Lists.GetByTitle($GroupListTitle)
    $Context.Load($List)
    $Context.ExecuteQuery()

    $GroupUploadCount = 1

    $Groups | ForEach-Object {
        Write-Progress -Activity " " -Status "$GroupUploadCount out of $(@($Groups).Count) groups uploaded" -CurrentOperation $_.Name -Id 2 -ParentId 1 -PercentComplete (($GroupUploadCount / @($Groups).Count) * 100)
        $ListItemCreationInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $NewListItem = $List.AddItem($ListItemCreationInfo)
        $NewListItem["Title"] = $_.ChangeOrder
        $NewListItem["AccountType"] = $_.AccountType
        $NewListItem["Description"] = $_.Description
        $NewListItem["DistinguishedName"] = $_.DistinguishedName
        $NewListItem["Domain"] = $_.Domain
        $NewListItem["FIMManaged"] = $_.FIM
        $NewListItem["Category"] = $_.GroupCategory
        $NewListItem["Result"] = $_.GroupResult
        $NewListItem["Scope"] = $_.GroupScope
        $NewListItem["DisplayedOwner"] = $_.ManagedBy
        $NewListItem["GroupName"] = $_.Name
        $NewListItem["SamAccountName"] = $_.SamAccountName
        $NewListItem["CompletedBy"] = $env:USERNAME
        $NewListItem["TimeStamp"] = [DateTime]::ParseExact($_.Timestamp, "yyyyMMddTHHmmssffff", $null).ToString("%M/dd/yyyy %h:mm tt")
        $NewListItem.Update()
        $Context.ExecuteQuery() | Out-Null
        $GroupUploadCount++
    }

    Write-Progress -Activity " " -Id 2 -ParentId 1 -Completed

    $List = $Context.Web.Lists.GetByTitle($ResultsListTitle)
    $Context.Load($List)
    $Context.ExecuteQuery()

    $ResultUploadCount = 1

    $Results | ForEach-Object {
        Write-Progress -Activity " " -Status "$ResultUploadCount out of $(@($Results).Count) results uploaded" -CurrentOperation $_.ChangeOrder -Id 2 -ParentId 1 -PercentComplete (($ResultUploadCount / @($Results).Count) * 100)
        $ListItemCreationInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $NewListItem = $List.AddItem($ListItemCreationInfo)
        $NewListItem["Title"] = $_.ChangeOrder
        $NewListItem["EmployeeID"] = $_.EmployeeID
        $NewListItem["Account"] = $_.Account
        $NewListItem["AccountType"] = $_.AccountType
        $NewListItem["DisplayName"] = $_.DisplayName
        $NewListItem["Domain"] = $_.Domain
        $NewListItem["Section"] = $_.Section
        $NewListItem["Status"] = $_.Status
        $NewListItem["Step"] = $_.Step
        $NewListItem["TimeStamp"] = [DateTime]::ParseExact($_.Timestamp, "yyyyMMddTHHmmssffff", $null).ToString("%M/dd/yyyy %h:mm tt")
        $NewListItem["CompletedBy"] = $env:USERNAME
        $NewListItem.Update()
        $Context.ExecuteQuery()
        $ResultUploadCount++
    }
    Write-Progress -Activity " " -Id 2 -ParentId 1 -Completed
}