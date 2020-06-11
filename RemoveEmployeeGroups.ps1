<#
RemoveEmployeeGroups method for Termination Module
#>

Function Remove-EmployeeGroups {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [PSObject]
        $Account,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ChangeOrder,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Mandatory = $true, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 4)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $CurrentCount
    )
    
    $AccountType = $Account.AccountType
    $Domain = $Account.Domain

    $Log.Account     = $Account.SamAccountName
    $Log.AccountType = $Account.AccountType
    $Log.ChangeOrder = $ChangeOrder
    $Log.DisplayName = $Account.DisplayName
    $Log.Domain      = $Account.Domain
    $Log.EmployeeID  = $Account.EmployeeID
    $Log.Status      = "Start"
    $Log.Step        = $MyInvocation.MyCommand
    $Log.Timestamp   = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    $Result.Account     = $Account.SamAccountName
    $Result.AccountType = $Account.AccountType
    $Result.ChangeOrder = $ChangeOrder
    $Result.DisplayName = $Account.DisplayName
    $Result.Domain      = $Account.Domain
    $Result.EmployeeID  = $Account.EmployeeID
    $Result.Status      = "Start"
    $Result.Step        = $MyInvocation.MyCommand
    $Result.Timestamp   = Get-Date -Format FileDateTime

    $Termination.AddResult($Result)

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount

    $Groups = New-Object System.Collections.Generic.List[System.Object]

    $TempGroups = Get-ADUser -Identity $Account.SamAccountName -Property MemberOf -Server $DomainControllers.$Domain.Name | Select-Object -ExpandProperty MemberOf | Get-ADGroup -Properties * -Server $DomainControllers.$Domain.Name

    $GroupCount = 1

    $TempGroups | ForEach-Object {
        $Group = [PSCustomObject]@{
            AccountType       = $AccountType
            ChangeOrder       = $ChangeOrder
            Description       = $_.Description
            DistinguishedName = $_.DistinguishedName
            Domain            = $Domain
            GroupCategory     = $_.GroupCategory
            GroupScope        = $_.GroupScope
            ManagedBy         = $_.ManagedBy
            Name              = $_.Name
            SamAccountName    = $_.SamAccountName
            TimeStamp         = Get-Date -Format FileDateTime
        }
        if ($_.extensionAttribute3 -eq "FIM_Managed") {
            Add-Member -InputObject $Group -NotePropertyName "FIM" -NotePropertyValue $true
        }
        else {
            Add-Member -InputObject $Group -NotePropertyName "FIM" -NotePropertyValue $false
        }
        try {
            if ($_.Name -notin $ExcludeGroups.$AccountType.$Domain) {
                Write-Progress -Activity " " -Status "$GroupCount out of $($TempGroups.Count) groups" -CurrentOperation "Removing $($Group.Name)" -Id 4 -ParentId 3 -PercentComplete (($GroupCount / $TempGroups.Count) * 100)
                Remove-ADGroupMember -Identity $_ -Members $Account.SamAccountName -Server $DomainControllers.$Domain.Name -Credential $Credential -Confirm:$false
                Add-Member -InputObject $Group -NotePropertyName "GroupResult" -NotePropertyValue "Removed"
                $GroupCount++
            }
            else {
                Write-Progress -Activity " " -Status "$GroupCount out of $($TempGroups.Count) groups" -CurrentOperation "Excluding $($Group.Name)" -Id 4 -ParentId 3 -PercentComplete (($GroupCount / $TempGroups.Count) * 100)
                Add-Member -InputObject $Group -NotePropertyName "GroupResult" -NotePropertyValue "Excluded"
                $GroupCount++
            }
        }
        catch {
            Write-Progress -Activity " " -Status "$GroupCount out of $($TempGroups.Count) groups" -CurrentOperation "Failed $($Group.Name)" -Id 4 -ParentId 3 -PercentComplete (($GroupCount / $TempGroups.Count) * 100)
            Add-Member -InputObject $Group -NotePropertyName "GroupResult" -NotePropertyValue "Failed"

            $ErrorLog.Exception             = $_.Exception
            $ErrorLog.ExceptionFullName     = $_.Exception.GetType().FullName
            $ErrorLog.FullyQualifiedErrorId = $_.FullyQualifiedErrorId
            $ErrorLog.Line                  = $_.InvocationInfo.Line
            $ErrorLog.MyCommand             = $_.InvocationInfo.MyCommand
            $ErrorLog.ScriptLineNumber      = $_.InvocationInfo.ScriptLineNumber
            $ErrorLog.ScriptName            = $_.InvocationInfo.ScriptName
            $ErrorLog.Timestamp             = Get-Date -Format FileDateTime

            $Main.AddErrorLog($ErrorLog)
            $GroupCount++
        }
        finally {
            if ($Group.GroupResult -eq "Removed" -or $Group.GroupResult -eq "Excluded") {
                $Log.Status    = "Success"
                $Result.Status = "Success"
            }
            else {
                $Log.Status    = "Failed"
                $Result.Status = "Failed"
            }

            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            $Result.Timestamp = Get-Date -Format FileDateTime

            $Termination.AddResult($Result)

            $Groups.Add($Group)
        }
    }

    $Log.Status    = "Complete"
    $Log.Timestamp = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    return $Groups
}