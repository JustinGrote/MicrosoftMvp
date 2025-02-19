#requires -module MicrosoftMvp,PowerHTML

using namespace HtmlAgilityPack

<#
.SYNOPSIS
Sessionize doesn't provide a Speaker API, so we must scrape the result of the 'https://sessionize.com/app/speaker/events' page. THIS IS A SCREENSCRAPER AND MAY BREAK IF SESSIONIZE CHANGES THEIR PAGE LAYOUT.
.EXAMPLE
Import-MvpSessionize -RequestHeaders (Get-Clipboard -Raw)

#Log into sessionize, hit F12 to get into devtools, go to your events, right click the "events" item, and choose "copy -> copy request headers", an interactive login using webview2 is on the TODO list
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(

	[Parameter(Mandatory)]$RequestHeaders,
	#Which focus area to import events for. Example: "PowerShell". Script will fail if this is not exact.
	[Parameter(Mandatory)]$TechnologyFocusArea,
	#A regex filter for event titles, can be useful to select certain events to update for different focus areas.
	$Filter,
	$ThrottleLimit = 30
)

if ($RequestHeaders -notmatch '.AspNet.ApplicationCookie=([^;]*)') {
	throw 'Invalid Request Headers, it should have an aspnet.applicationcookie in the content.'
}

$session = [Microsoft.PowerShell.Commands.WebRequestSession]::new()
$authCookie = [Net.Cookie]::new('.AspNet.ApplicationCookie', $matches[1], '/', 'sessionize.com')
$session.Cookies.Add($authCookie)
$PSDefaultParameterValues['Invoke-RestMethod:WebSession'] = $session
$PSDefaultParameterValues['Invoke-RestMethod:Verbose'] = $false

$ErrorActionPreference = 'stop'
Write-Warning 'This is a screenscraper and may break if Sessionize changes their events page layout.'

#region Utility
filter Assert-ScreenScrapeData {
	param(
		[string]$Target
	)

	if (-not $PSItem) { throw "Invalid Data Detected: $Target. this is either not a valid Sessionize events response, or Sessionize has changed their format and broken this script" }
	return $PSItem
}

filter ConvertFrom-HtmlEncoding {
	[System.Net.WebUtility]::HtmlDecode($PSItem)
}
#endregion

function ConvertFrom-SessionizeSessionTab {
	<#
	.SYNOPSIS
	Screenscrapes the data from a Sessionize event tab.
	#>
	param(
		[Parameter(Mandatory)][HtmlNode]$Tab
	)

	$rows = $Tab.SelectNodes('.//div[@class="table-responsive"]/table/tbody/tr')

	$currentEvent = $null
	$sessions = foreach ($row in $rows) {
		#Events are specified in the row with an event class, so determine if this row is an event or a session.
		try {
			if ($row.GetClasses() -contains 'event') {
				$newEvent = $row.SelectNodes('.//h4/a').InnerText.Trim()
				| Assert-ScreenScrapeData 'Table Event Name'
				$currentEvent = $newEvent

				#The first row should be an event, so if it is not, something is wrong with the format.
				if (-not $currentEvent) { throw 'Event not found in first row of the table, this is either not a valid Sessionize events response, or Sessionize has changed their format and broken this script' }
				continue
			}

			#Otherwise it should be a session entry
			$session = [ordered]@{}

			$session.Event = $currentEvent

			$session.Title = $row.SelectSingleNode('.//td/a[starts-with(@href, "/app/speaker/session")]').InnerText.Trim()
			| ConvertFrom-HtmlEncoding
			| Assert-ScreenScrapeData 'Session Title'

			$session.Name = "$($session.Event): $($session.Title)"
			if ($Filter -and ($session.Name -notmatch $Filter)) {
				Write-Verbose "Skipping $($session.Name) due to not matching filter"
				continue
			}

			$session.Url = 'https://sessionize.com' + $row.SelectSingleNode('.//td/a[starts-with(@href, "/app/speaker/session")]').GetAttributeValue('href', '')


			$session.Status = $row.SelectSingleNode('.//td/span').InnerText.Trim()
			| ConvertFrom-HtmlEncoding
			| Assert-ScreenScrapeData 'Event Acceptance Status'

			if ($session.Status -ne 'Accepted') {
				Write-Verbose "Skipping unaccepted session '$($session.Title)' for '$currentEvent'"
				#Unaccepted sessions dont have a date which will crash the below line.
				continue
			}

			$session.Date = [datetime](
				$row.SelectSingleNode('.//td[3]').InnerText.Trim()
				| ConvertFrom-HtmlEncoding
				| Assert-ScreenScrapeData 'Session Date'
			)

			$session.TechFocusArea = $TechnologyFocusArea

			$eventDetailHtml = ConvertFrom-Html (Invoke-RestMethod $session.Url)
			$descriptionXPath = './/div[@class="ibox-content"]/h4[contains(text(),"Description")]/following-sibling::p'
			$session.Description = $eventDetailHtml.SelectSingleNode($descriptionXPath).InnerText.Trim()
			| ConvertFrom-HtmlEncoding
			| Assert-ScreenScrapeData 'Session Description'

			[PSCustomObject]$Session
		} catch {
			Write-Warning "Invalid session data detected: $PSItem. Skipping this session."
		}
	}

	return $sessions
}


#region Main

#Retrieve the data we care about from the HTML response.
$eventHtml = Invoke-RestMethod 'https://sessionize.com/app/speaker/events'

$page = ConvertFrom-Html $eventHtml

$currentTab = $page.SelectSingleNode('//div[@id="tab-current"]')
| Assert-ScreenScrapeData 'Current Sessions Tab'

$currentSessions = ConvertFrom-SessionizeSessionTab $currentTab

$archiveTab = $page.SelectSingleNode('//div[@id="tab-archived"]')
| Assert-ScreenScrapeData 'Past Sessions Tab'

$archiveSessions = ConvertFrom-SessionizeSessionTab $archiveTab

Write-Verbose "Found $($currentSessions.Count) current Sessions and $($archiveSessions.Count) past Sessions."

$sessionsToUpdate = $currentSessions + $archiveSessions
| Where-Object Name -Match $Filter

#TODO: This "upsert" logic should probably be available in the main module
$mvpModulePath = (Get-Module MicrosoftMvp).path
$InformationPreference = 'Continue'

$existingActivities = Search-MvpActivitySummary -First 10000
| Where-Object type -Like 'Speaker*'

$sessionsToUpdate | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
	Import-Module $USING:mvpModulePath
	$VerbosePreference = $USING:VerbosePreference
	$WhatIfPreference = $USING:WhatIfPreference
	$DebugPreference = $USING:DebugPreference

	$session = $PSItem
	$activityTitle = $session.Name
	$existingActivity = $USING:existingActivities
	| Where-Object title -EQ $activityTitle
	| Get-MvpActivity
	| Where-Object url -EQ $session.Url

	if ($existingActivity.count -gt 1) {
		Write-Warning "Multiple activities found for '$activityTitle' with url $($session.url). Remove one from your MVP profile to update this."
		continue
	}

	$activity = $existingActivity ? ($existingActivity | Get-MvpActivity) : $(
		$activityParams = @{
			Title = $activityTitle
			Type = 'Speaker/Presenter at Third-party event'
			Date = $session.date
			Description = $session.description.Substring(0, [Math]::Min($session.description.Length, 1000))
			TechnologyFocusArea = $session.TechFocusArea
			TargetAudience = 'Developer', 'IT Pro'
		}
		$newActivity = New-MvpActivity @activityParams
		$newActivity.url = $session.url
		$newActivity
	)

	if ($existingActivity) {
		#Workaround for whatif not working as it is supposed to in parallel
		if ($WhatIfPreference) {
			Write-Host "WhatIf: Updating speaking event '$activityTitle' with url $($session.url)"
			return
		}
		Write-Verbose "Updating speaking event '$activityTitle' with url $($session.url)"
		#Update the activity with the latest data
		$activity.Date = $session.date
		$activity.Description = $session.description
		Set-MvpActivity $activity -WhatIf:$WhatIfPreference
	} else {
		if ($WhatIfPreference) {
			Write-Host "WhatIf: Adding speaking event '$activityTitle' with url $($session.url)"
			return
		}
		Write-Verbose "Adding new speaking event '$activityTitle' with url $($session.url)"
		Add-MvpActivity $activity -WhatIf:$WhatIfPreference
	}
}
#endregion