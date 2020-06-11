<#
Improved Write-Progress functions for Termination Module
#>


Function Global:Write-TerminationProgress {

    [CmdletBinding(DefaultParameterSetName = "Processing")]

    Param(

        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = "Processing")]
        [ValidateNotNullOrEmpty()]
        [String]
        $ParameterSet,

        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "Processing")]
        [ValidateNotNullOrEmpty()]
        [String]
        $Section,

        [Parameter(Mandatory = $true, ParameterSetName = "Completed")]
        [Switch]
        $Complete
    )

    Switch ($PSCmdlet.ParameterSetName) {
        "Processing" {
            $SectionCount = $Sections.Count
            $SectionNumber = $SectionDetails.$Section.SectionNumber
            Write-Progress -Activity "Termination Type: $ParameterSet" -Status "$SectionNumber out of $SectionCount sections" -CurrentOperation $Section -Id 1 -ParentId -1 -PercentComplete (($SectionNumber / $SectionCount) * 100)
        }
        "Completed" {
            Write-Progress -Activity "Termination Type: $ParameterSet" -Id 1 -Completed
        }
    }
}


Function Global:Write-SectionProgress {

    [CmdletBinding(DefaultParameterSetName = "Processing")]

    Param(

        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = "Processing")]
        [ValidateNotNullOrEmpty()]
        [String]
        $Section,

        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "Processing")]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = "Processing")]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $CurrentCount,

        [Parameter(Mandatory = $true, ParameterSetName = "Completed")]
        [Switch]
        $Complete
    )

    Switch ($PSCmdlet.ParameterSetName) {
        "Processing" {
            Write-Progress -Activity " " -Status "$CurrentCount out of $TerminationCount Terminations" -CurrentOperation " " -Id 2 -ParentId 1 -PercentComplete (($CurrentCount / $TerminationCount) * 100)
        }
        "Completed" {
            Write-Progress -Activity " " -Id 2 -ParentId 1 -Completed
        }
    }
}


Function Global:Write-StepProgress {

    [CmdletBinding(DefaultParameterSetName = "Processing")]

    Param(

        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = "Processing")]
        [ValidateNotNullOrEmpty()]
        [String]
        $Section,

        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "Processing")]
        [ValidateNotNullOrEmpty()]
        [String]
        $Step,

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = "Processing")]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $TerminationCount,

        [Parameter(Mandatory = $true, Position = 3, ParameterSetName = "Processing")]
        [ValidateNotNullOrEmpty()]
        [Int32]
        $CurrentCount,

        [Parameter(Mandatory = $true, ParameterSetName = "Completed")]
        [Switch]
        $Complete
    )

    Switch ($PSCmdlet.ParameterSetName) {
        "Processing" {
            $Steps = $SectionDetails.$Section.StepCount
            $StepNumber = $SectionDetails.$Section.StepDetails.$Step
            Write-Progress -Activity " " -Status "$StepNumber out of $Steps steps" -CurrentOperation $Step -Id 3 -ParentId 2 -PercentComplete (($StepNumber / $Steps) * 100)
        }
        "Completed" {
            Write-Progress -Activity " " -Id 3 -ParentId 2 -Completed
        }
    }
}