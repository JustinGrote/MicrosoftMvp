#Requires -Module PSAuthClient, ThreadJob
using namespace Microsoft.PowerShell.Commands
using namespace System.Management.Automation
$ErrorActionPreference = 'Stop'

#region Model
class MvpActivity {
	[int]$id
	[int]$userProfileId = (Get-MvpContext).MvpProfile.id
	[string]$tenant = (Get-MvpContext).Tenant
	[string]$title
	[string]$description
	[string]$privateDescription
	[Nullable[datetime]]$date
	[Nullable[datetime]]$dateEnd
	[string]$url
	[string]$imageUrl
	[string]$activityTypeName
	[string]$activityTypeLocKey
	[string]$technologyFocusArea
	[string[]]$targetAudience
	[string[]]$additionalTechnologyAreas
	[int]$quantity
	[int]$reach
	[string]$role
	[string]$contributionRoleLocKey
	[string]$companySize
	[string]$companyName
	[string]$microsoftEvent
	[string]$microsoftEventOther
}
Update-TypeData -TypeName 'MvpActivity' -DefaultDisplayPropertySet 'id', 'title', 'activityTypeName', 'date' -Force

class MvpSearch {
	[int]$pageIndex = 1
	[int]$pageSize = 1000
	[string]$searchKey
	[string]$tenant = (Get-MvpContext).Tenant
	[string]$userProfileIdentifier = (Get-MvpContext).MvpProfile.userProfileIdentifier
	[string[]]$contributionTargetAudience = @()
	[string[]]$technologyFocusArea = @()
	[string[]]$type = @()
}

class MvpSearchResult {
	[int]$searchScore
	[int]$id
	[string]$title
	[string]$description
	[Nullable[datetime]]$date
	[string]$imageUrl
	[bool]$isFirstParty
	[bool]$isHighImpact
	[string]$type
	[string]$typeLocKey
	[string]$tenantName
	[bool]$deletable
	[bool]$editable
	[bool]$highImpactToggleable
}
Update-TypeData -TypeName 'MvpSearchResult' -DefaultDisplayPropertySet 'id', 'date', 'title', 'type' -Force

class ActivityTypes: IValidateSetValuesGenerator {
	[string[]] GetValidValues() {
		return $((Get-MvpActivityData).activityTypes.name)
	}
}
class TechnologyArea: IValidateSetValuesGenerator {
	[string[]] GetValidValues() {
		return (Get-MvpActivityData).technologyArea.technologyName
	}
}
class TargetAudience: IValidateSetValuesGenerator {
	[string[]] GetValidValues() {
		return (Get-MvpActivityData).targetAudience.name
	}
}


#endregion Model


#region Engine

#Defaults
$SCRIPT:Tenant = 'MVP'
[string]$SCRIPT:BaseUri = 'https://mavenapi-prod.azurewebsites.net/api/'

#This is used to track the context of the logged in user across runspaces to enable easy parallelization

if (-not [type]::GetType('MicrosoftMvp.UserProfile')) {
	Add-Type -TypeDefinition '
		namespace MicrosoftMvp;
		public static class UserProfile {
			public static object Context { get; set; }
		}
	'
}

function Get-MvpContext {
	[MicrosoftMvp.UserProfile]::Context
}

function Connect-Mvp {
	[CmdletBinding(DefaultParameterSetName = 'Interactive')]
	param(
		[ValidateNotNullOrEmpty()]
		[string]$BaseUri = $SCRIPT:BaseUri,
		[ValidateNotNullOrEmpty()]
		[string]$Tenant = $SCRIPT:Tenant,
		[switch]$Force
	)

	if ($Force) {
		Disconnect-Mvp
	}

	if ((Get-MvpContext).MvpProfile.userProfileIdentifier) {
		Write-Warning "You are already connected as $((Get-MvpContext).GraphUser.UserPrincipalName). Use -Force or Disconnect-Mvp first to reconnect."
		return
	}

	$code = Invoke-OAuth2AuthorizationEndpoint -uri 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize' -client_id 'e83f495c-dfa2-48e2-b1d9-3680b16e74e4' -scope 'openid profile User.Read offline_access' -redirect_uri 'https://mvp.microsoft.com' -response_type 'code' -response_mode 'fragment'

	#Use the authorization code to fetch a graph token. Origin is important here since we are impersonating an SPA, so we cannot use the typical method to get this token.
	$graphContext = Invoke-RestMethod -Uri 'https://login.microsoftonline.com/common/oauth2/v2.0/token' -Method Post -Body @{
		client_id     = $code.client_id
		redirect_uri  = $code.redirect_uri
		scope         = 'openid profile User.Read offline_access'
		code          = $code.code
		code_verifier = $code.code_verifier
		grant_type    = 'authorization_code'
		client_info   = 1
	} -Headers @{
		Origin  = 'https://mvp.microsoft.com'
		Referer = 'https://mvp.microsoft.com'
	}


	$me = Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/me' -Authentication Bearer -Token ($graphContext.access_token | ConvertTo-SecureString -AsPlainText)

	$mvpToken = Invoke-RestMethod -Uri 'https://login.microsoftonline.com/common/oauth2/v2.0/token' -Method Post -Body @{
		client_id     = $code.client_id
		refresh_token = $graphContext.refresh_token
		scope         = 'api://6dabb447-da84-4b4c-b68f-99f5215b2ca7/User.All openid profile offline_access'
		grant_type    = 'refresh_token'
		client_info   = 1
	} -Headers @{
		Origin  = 'https://mvp.microsoft.com'
		Referer = 'https://mvp.microsoft.com'
	}

	[MicrosoftMvp.UserProfile]::Context = @{
		Graph       = $graphContext
		GraphUser   = $me
		GraphExpire = (Get-Date).AddSeconds($graphContext.expires_in - 60)
		Mvp         = $mvpToken
		Tenant      = $Tenant
		Data        = @{}
	}
	(Get-MvpContext).MvpProfile = (Invoke-MvpRestMethod ('UserStatus/' + $me.userPrincipalName)).userStatusModel
	#Pre-Populate the data, have seen some race conditions if this is lazily evaluated
	[void](Get-MvpActivityData)
	Write-Verbose "Connected as $($me.UserName) to $BaseUri"
}

function Assert-MvpConnection {
	if (-not (Get-MvpContext)) {
		throw 'You must connect to the MVP API first using Connect-Mvp'
	}
}

function Disconnect-Mvp {
	[MicrosoftMvp.UserProfile]::Context = $null
}

function Invoke-MvpRestMethod {
	[CmdletBinding()]
	param(
		[Parameter(Position = 0, Mandatory, ParameterSetName = 'Endpoint')]
		[string]$Endpoint,

		$Body,

		[WebRequestMethod]$Method = 'GET',

		[ValidateNotNullOrEmpty()]
		[Parameter(ParameterSetName = 'Uri')]
		[string]$Uri = $BaseUri
	)
	Assert-MvpConnection
	if ($Endpoint) {
		$Uri = $BaseUri + $Endpoint
	}

	$irmParams = @{
		Uri            = $Uri
		Method         = $Method
		Body           = $Body
		Authentication = 'Bearer'
		Token          = (Get-MvpContext).Mvp.access_token | ConvertTo-SecureString -AsPlainText
		Debug          = $false
		Verbose        = $false
	}
	#Use the JSON content type for all methods except GET
	if ($Method -ne 'GET') {
		$irmParams['ContentType'] = 'application/json'
		$irmParams['Body'] = if ($irmParams['Body']) { $irmParams['Body'] | ConvertTo-Json -Depth 5 }
	}

	Write-Debug "MVP: $("$Method".ToUpper()) $Uri$(if ($irmParams['Body']) {' with body ' + $irmParams['Body']})"
	try {
		$response = Invoke-RestMethod @irmParams
	} catch {
		$jsonError = $($PSItem.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue)
		if ($jsonError.error) {
			$PSItem | ThrowCmdletError "$($jsonError.error.code): $($jsonError.error.message) $($jsonError.error.Details)"
			return
		}
		if ($jsonError.errors) {
			$PSItem | ThrowCmdletError "$($jsonError.title)`n$($jsonError.errors | Format-List | Out-String)"
		}
		throw $PSItem
	}
	return $response
}

function Get-MvpActivityData {
	<#
	.SYNOPSIS
	Fetches the activity data for the current tenant. Used for dynamic intellisense
	#>
	if (-not (Get-MvpContext).Data.Activity) {
		Write-Verbose 'MVP: Fetching Activity Data'
		$response = Invoke-MvpRestMethod 'SiteContent/Activity/Common/Data' -Body @{tenant = $SCRIPT:Tenant }
		(Get-MvpContext).Data.Activity = $response.data
	}
	return (Get-MvpContext).Data.Activity
}

#endregion Engine

function Search-MvpActivitySummary {
	[OutputType('MvpSearchResult')]
	param(
		[string]$Filter,
		[int]$First = 100,
		[int]$Skip = 0
	)

	$Search = [MvpSearch]@{
		pageIndex = $Skip + 1
		pageSize  = $First
	}
	if ($Filter) {
		$Search.SearchKey = $Filter
	}

	$response = (Invoke-MvpRestMethod -Endpoint 'Contributions/CommunityLeaderActivities/search' -Body $Search -Method 'POST')

	return [MvpSearchResult[]]$response.CommunityLeaderActivities
}

filter Get-MvpActivity {
	[OutputType('MvpActivity')]
	[CmdletBinding(DefaultParameterSetName = 'Id')]
	param(
		[Parameter(Position = 0, ParameterSetName = 'Filter')][string]$Filter,
		[int]$First = 100,
		[int]$Skip = 0,
		[Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Id')][int]$Id
	)
	if (-not $Id) {
		Search-MvpActivitySummary -Filter $Filter -First $First -Skip $Skip | Get-MvpActivity
		return
	}
	try {
		[MvpActivity](Invoke-MvpRestMethod "Activities/$Id")
	} catch {
		$PSItem | ThrowCmdletError
	}
}

filter Set-MvpActivity {
	[OutputType('MvpActivity')]
	[CmdletBinding(SupportsShouldProcess)]
	param(
		[Parameter(Mandatory, ValueFromPipeline)][MvpActivity]$Activity
	)

	if (-not $PSCmdlet.ShouldProcess("$($Activity.title): $($Activity | ConvertTo-Json -Depth 2)", 'Update Activity')) {
		return
	}
	try {
		$response = Invoke-MvpRestMethod -Endpoint 'Activities' -Body @{activity = $Activity } -Method 'PUT'
		#HACK: Workaround for the fact the returned info is not the same type. Slight performance impact
		Get-MvpActivity -Id $response.id
	} catch {
		$PSItem | ThrowCmdletError
	}
}

filter Add-MvpActivity {
	[OutputType('MvpActivity')]
	[CmdletBinding(SupportsShouldProcess)]
	param(
		[Parameter(Mandatory, ValueFromPipeline)][MvpActivity]$Activity
	)

	if (-not $PSCmdlet.ShouldProcess("$($Activity.title): $($Activity | ConvertTo-Json -Depth 2)", 'Create Activity')) {
		return
	}
	try {
		$response = Invoke-MvpRestMethod -Endpoint 'Activities' -Body @{activity = $Activity } -Method 'POST'
		#HACK: Workaround for the fact the returned info is not the same type. Slight performance impact
		Get-MvpActivity -Id $response.contributionId
	} catch {
		$PSItem | ThrowCmdletError
	}
}

filter New-MvpActivity {
	param(
		[Parameter(Mandatory)][string]$Title,
		[Parameter(Mandatory)][string]$Description,
		[ValidateSet([ActivityTypes])]
		[Parameter(Mandatory)][string]$Type,
		[ValidateSet([TechnologyArea])]
		[Parameter(Mandatory)][string]$TechnologyFocusArea,
		[ValidateSet([TargetAudience])]
		[Parameter(Mandatory)][string[]]$TargetAudience,
		[DateTime]$Date,
		[DateTime]$EndDate,
		[int]$Quantity,
		[int]$Reach
	)
	return [MvpActivity]@{
		userProfileId       = (Get-MvpContext).MvpProfile.id
		tenant              = (Get-MvpContext).Tenant
		title               = $Title
		description         = $Description
		date                = $Date ?? (Get-Date)
		dateEnd             = $EndDate ?? (Get-Date)
		activityTypeName    = $Type
		technologyFocusArea = $TechnologyFocusArea
		quantity            = $Quantity ?? 1
		reach               = $Reach ?? 0
		url                 = 'https://mvp.microsoft.com'
		targetAudience      = $TargetAudience
	}
}

filter Remove-MvpActivity {
	[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'Activity')]
	param(
		[Parameter(ParameterSetName = 'Activity', Mandatory, ValueFromPipeline)]
		[MvpActivity]$MvpActivity,

		[Parameter(ParameterSetName = 'SearchResult', Mandatory, ValueFromPipeline)]
		[MvpSearchResult]$SearchResult
	)

	#TODO: Should probably use an interface or a base class
	$Activity = $SearchResult ?? $MvpActivity

	if (-not $PSCmdlet.ShouldProcess("$($Activity.id): $($Activity.title)", 'Delete Activity')) {
		return
	}

	try {
		$response = Invoke-MvpRestMethod -Endpoint "Activities/$($Activity.id)" -Method 'DELETE'
		if ($response -notmatch $Activity.id) {
			throw "Expected $($Activity.id) to be deleted but $response was returned. This is probably a bug."
		}
	} catch {
		$PSItem | ThrowCmdletError
	}
}

filter ThrowCmdletError {
	param(
		[string]$Message,
		[Parameter(ValueFromPipeline)][ErrorRecord]$ErrorRecord,
		$ThisCmdlet = $(Get-Variable -Scope 1 -Name PSCmdlet -ValueOnly)
	)
	if ($Message) {
		$ErrorRecord.ErrorDetails = $Message
	}
	$ThisCmdlet.ThrowTerminatingError($ErrorRecord)
}