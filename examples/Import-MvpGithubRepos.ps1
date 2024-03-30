#requires -version 7 -Modules MicrosoftMvp
<#
.SYNOPSIS
Updates your MVP profile with your Github repositories. Supports -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
	[Parameter(Mandatory)]
	$User,

	[Parameter(Mandatory)]
	$TechnologyFocusArea,

	#A regex filter for repo titles, can be useful to select certain repos to update for different technology focus areas.
	$Filter,

	[ValidateNotNullOrEmpty()]
	[int]$MinimumStars = 5,

	[int]$ThrottleLimit = 20
)
$apiUrl = "https://api.github.com/users/$User/repos?per_page=100&sort=pushed"
$response = Invoke-RestMethod -FollowRelLink -Uri $apiUrl

#Flatten the multi-valued response
$repos = foreach ($i in $response) { foreach ($y in $i) { $y } }

$existingActivities = Search-MvpActivitySummary -First 10000
| Where-Object type -Like 'Open Source*'

$targetRepos = $repos
| Where-Object stargazers_count -GT $MinimumStars
| Where-Object title -Match $Filter


#TODO: This "upsert" logic should probably be available in the main module
$mvpModulePath = (Get-Module MicrosoftMvp).path
$InformationPreference = 'Continue'

$targetRepos | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
	Import-Module -Force $USING:mvpModulePath
	$VerbosePreference = $USING:VerbosePreference
	$WhatIfPreference = $USING:WhatIfPreference
	$DebugPreference = $USING:DebugPreference

	$repo = $PSItem
	$activityTitle = "GitHub: $($repo.name)"
	$existingActivity = $USING:existingActivities
	| Where-Object title -EQ $activityTitle
	| Get-MvpActivity
	| Where-Object url -EQ $repo.url

	if ($existingActivity.count -gt 1) {
		Write-Warning "Multiple activities found for '$activityTitle' with url $($repo.url). Remove one from your MVP profile to update this."
		continue
	}

	$activity = $existingActivity ? ($existingActivity | Get-MvpActivity) : $(
		$newMvpActivityParams = @{
			Title               = $activityTitle
			Type                = 'Open Source/Project/Sample code/Tools'
			TechnologyFocusArea = $USING:TechnologyFocusArea
			TargetAudience      = 'Developer', 'IT Pro'
			Description         = $repo.description ?? 'No description provided.'
			Date                = $repo.pushed_at
			EndDate             = $repo.pushed_at
			Quantity            = 1
			Reach               = $repo.stargazers_count
		}
		$newActivity = New-MvpActivity @newMvpActivityParams
		$newActivity.url = $repo.url
		$newActivity
	)

	if ($existingActivity) {
		#Workaround for whatif not working as it is supposed to in parallel
		if ($WhatIfPreference) {
			Write-Host "WhatIf: Updating repository '$activityTitle' with url $($repo.url)"
			return
		}
		Write-Verbose "Updating repository '$activityTitle' with url $($repo.url)"
		#Update the activity with the latest data
		$activity.Date = $repo.created_at
		$activity.DateEnd = $repo.updated_at
		$activity.Reach = $repo.stargazers_count
		$activity.Description = $repo.description
		Set-MvpActivity $activity -WhatIf:$WhatIfPreference
	} else {
		if ($WhatIfPreference) {
			Write-Host "WhatIf: Adding repository '$activityTitle' with url $($repo.url)"
			return
		}
		Write-Verbose "Adding new repository '$activityTitle' with url $($repo.url)"
		Add-MvpActivity $activity -WhatIf:$WhatIfPreference
	}
}