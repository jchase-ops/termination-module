<#
GetLogFile method for Termination Module
#>

Function Get-LogFile {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $CurrentCount
    )

    $Log.Account     = $env:USERNAME
    $Log.AccountType = "Developer"
    $Log.ChangeOrder = "N/A"
    $Log.DisplayName = $env:USERNAME -Replace "\.", " "
    $Log.Domain      = "All"
    $Log.EmployeeID  = "N/A"
    $Log.Status      = "Start"
    $Log.Step        = $MyInvocation.MyCommand
    $Log.Timestamp   = Get-Date -Format FileDateTime

    $Main.AddLog($Log)

    Write-StepProgress -Section $Log.Section -Step ($MyInvocation.MyCommand -Replace "\-", "") -TerminationCount $TerminationCount -CurrentCount $CurrentCount

    $FileName = $env:USERNAME + "-" + "TerminationScript.log"
    $Folder   = $env:USERNAME
    $Date     = Get-Date -UFormat %Y%m%d
    $Month    = Get-Date -Format MMM
    $Year     = Get-Date -Format yyyy
    $Root     = "\\$NetworkShare\Termination\logs"
    
    if (Test-Path -Path "$Root\$Year") {
        if (Test-Path -Path "$Root\$Year\$Month") {
            if (Test-Path -Path "$Root\$Year\$Month\$Date") {
                if (Test-Path -Path "$Root\$Year\$Month\$Date\$Folder") {
                    if (Test-Path -Path "$Root\$Year\$Month\$Date\$Folder\$FileName") {
                        $Path = "$Root\$Year\$Month\$Date\$Folder\$FileName"

                        $Log.Status    = "Complete"
                        $Log.Timestamp = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        return $Path
                    }
                    else {
                        try {
                            $Path = (New-Item -Name $FileName -Path "$Root\$Year\$Month\$Date\$Folder" -ItemType "File" -Force).FullName

                            $Log.Status    = "Complete"
                            $Log.Timestamp = Get-Date -Format FileDateTime

                            $Main.AddLog($Log)

                            return $Path
                        }
                        catch {
                            $Log.Status    = "Failed"
                            $Log.Timestamp = Get-Date -Format FileDateTime

                            $Main.AddLog($Log)

                            return $null
                        }
                    }
                }
                else {
                    try {
                        $PersonalFolder = (New-Item -Name $Folder -Path "$Root\$Year\$Month\$Date" -ItemType "Directory" -Force).FullName
                        $Path = (New-Item -Name $FileName -Path $PersonalFolder -ItemType "File" -Force).FullName

                        $Log.Status    = "Complete"
                        $Log.Timestamp = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)

                        return $Path
                    }
                    catch {
                        $Log.Status    = "Failed"
                        $Log.Timestamp = Get-Date -Format FileDateTime

                        $Main.AddLog($Log)
                    
                        return $null
                    }
                }
            }
            else {
                try {
                    $DateFolder = (New-Item -Name $Date -Path "$Root\$Year\$Month" -ItemType "Directory" -Force).FullName
                    $PersonalFolder = (New-Item -Name $Folder -Path $DateFolder -ItemType "Directory" -Force).FullName
                    $Path = (New-Item -Name $FileName -Path $PersonalFolder -ItemType "File" -Force).FullName

                    $Log.Status    = "Complete"
                    $Log.Timestamp = Get-Date -Format FileDateTime
                
                    $Main.AddLog($Log)

                    return $Path
                }
                catch {
                    $Log.Status    = "Failed"
                    $Log.Timestamp = Get-Date -Format FileDateTime

                    $Main.AddLog($Log)
                    
                    return $null
                }
            }
        }
        else {
            try {
                $MonthFolder = (New-Item -Name $Month -Path "$Root\$Year" -ItemType "Directory" -Force).FullName
                $DateFolder = (New-Item -Name $Date -Path $MonthFolder -ItemType "Directory" -Force).FullName
                $PersonalFolder = (New-Item -Name $Folder -Path $DateFolder -ItemType "Directory" -Force).FullName
                $Path = (New-Item -Name $FileName -Path $PersonalFolder -ItemType "File" -Force).FullName

                $Log.Status    = "Complete"
                $Log.Timestamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)

                return $Path
            }
            catch {
                $Log.Status    = "Failed"
                $Log.Timestamp = Get-Date -Format FileDateTime

                $Main.AddLog($Log)
                    
                return $null
            }
        }
    }
    else {
        try {
            $YearFolder = (New-Item -Name $Year -Path $Root -ItemType "Directory" -Force).FullName
            $MonthFolder = (New-Item -Name $Month -Path $YearFolder -ItemType "Directory" -Force).FullName
            $DateFolder = (New-Item -Name $Date -Path $MonthFolder -ItemType "Directory" -Force).FullName
            $PersonalFolder = (New-Item -Name $Folder -Path $DateFolder -ItemType "Directory" -Force).FullName
            $Path = (New-Item -Name $FileName -Path $PersonalFolder -ItemType "File" -Force).FullName

            $Log.Status    = "Complete"
            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)

            return $Path
        }
        catch {
            $Log.Status    = "Failed"
            $Log.Timestamp = Get-Date -Format FileDateTime

            $Main.AddLog($Log)
                    
            return $null
        }
    }
}