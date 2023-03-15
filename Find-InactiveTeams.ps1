# Script gets M365 Groups activity details and filters based on last activity date
# for inactive Teams.
# Inactive M365 groups are written to a SharePoint Online list
#
# created by Thorsten Pickhan
# Initial script created on 01.12.2022 (12/01/2022)
#
# Version 1.0

# PNP PowerShell is required in Version

# Install-Module -Name PnPOnline
# Import-Module PNPOnline

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

# Define the SharePoint teamsite Url where the list is located
$RootURL = "https://<Your Tenant name>.sharepoint.com/teams/TeamsAutomation/"

# Define the list Url
$SharePointListName = "Lists/InactiveTeams"

# Define the period of report [D7, D30, D90, D180]
$PeriodOfReport = "D30"

# Define the CSV file path for usage report data
$csvfilename = ".\report.csv"

# Connect to SharePoint Online
$RootConnection = Connect-PnPOnline -Url $RootUrl -Interactive -ReturnConnection -ErrorAction Stop

# Define Tenant Id for access token
$TenantId ="<Your Tenant Id>"

$Scope = "https://graph.microsoft.com/.default"
$Url = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

# Define AppId and App Secret for access token request
$AppId = "<Your App Id>"
$AppSecret = "<Your App Secret>"


# Add System.Web for urlencode
Add-Type -AssemblyName System.Web

# Create body
$Body = @{
    client_id = $AppId
	client_secret = $AppSecret
	scope = $Scope
	grant_type = 'client_credentials'
}

# Splat the parameters for Invoke-Restmethod for cleaner code
$PostSplat = @{
    ContentType = 'application/x-www-form-urlencoded'
    Method = 'POST'
    Body = $Body
    Uri = $Url
}

# Request the access token
$Request = Invoke-RestMethod @PostSplat

# Create token header
$Header = @{
    Authorization = "$($Request.token_type) $($Request.access_token)"
}

# Define Graph APU Url to get O365 Group Activity report

$GraphApiUrl = "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='$($PeriodOfReport)')"

$Report = Invoke-WebRequest -UseBasicParsing -Headers $Header -Uri $GraphApiUrl -outfile $csvfilename

# Import generated data
$UsageData = Import-Csv $csvfilename
$Count = $UsageData.count
$Counter = 1

# Set check date to compare with the last activity date
$CheckDate = (Get-Date).adddays(-30)

# Check and validate each M365 Group
ForEach ($UsageRecord in $UsageData) {
	Write-Output "Proceed list entry $($Counter) from $($Count)..."
	if ($UsageRecord.'Is Deleted' -eq "True") {
		$Counter++
		continue
	}
    # Set columne values for SharePoint list entry
	$ReportRefreshDate = $UsageRecord."Report Refresh Date"
	$GroupDisplayName = $UsageRecord."Group Display Name"
	$IsDeleted = $UsageRecord."Is Deleted"
	$OwnerPrincipalName = $UsageRecord."Owner Principal Name"
	$LastActivityDate = $UsageRecord."Last Activity Date"
	$GroupType = $UsageRecord."Group Type"
	$MemberCount = $UsageRecord."Member Count"
	$ExternalMemberCount = $UsageRecord."External Member Count"
	$ExchangeReceivedEmailCount = $UsageRecord."Exchange Received Email Count"
	$SharePointActiveFileCount = $UsageRecord."SharePoint Active File Count"
	$YammerPostedMessageCount  = $UsageRecord."Yammer Posted Message Count"
	$YammerReadMessageCount = $UsageRecord."Yammer Read Message Count"
	$YammerLikedMessageCount = $UsageRecord."Yammer Liked Message Count"
	$ExchangeMailboxTotalItemCount = $UsageRecord."Exchange Mailbox Total Item Count"
	$ExchangeMailboxStorageUsedByte = $UsageRecord."Exchange Mailbox Storage Used (Byte)"
	$SharePointTotalFileCount = $UsageRecord."SharePoint Total File Count"
	$SharePointSiteStorageUsedByte = $UsageRecord."SharePoint Site Storage Used (Byte)"
	$GroupId = $UsageRecord."Group Id"
	$ReportPeriod = $UsageRecord."Report Period"

    # Check if there is a last activity date in CSV
	if ($LastActivityDate) {
		$LastActiveDate = Get-Date $LastActivityDate
	}
	else {
		$LastActiveDate = Get-Date "01.01.1900"
		try {
			# No activity Teams should be archived
			$AddSPListPerm = Add-PnPListItem -List $SharePointListName -Values @{"ReportRefreshDate" = $ReportRefreshDate; "GroupDisplayName" = $GroupDisplayName; "Title" = $GroupDisplayName; "IsDeleted" = $IsDeleted; "GroupOwners" = $OwnerPrincipalName; "LastActivityDate" = $LastActiveDate; "GroupType" = $GroupType; "MemberCount" = $MemberCount; "ExternalMemberCount" = $ExternalMemberCount; "ExchangeReceivedEmailCount" = $ExchangeReceivedEmailCount; "SharePointActiveFileCount" = $SharePointActiveFileCount; "ExchangeMailboxTotalItemCount" = $ExchangeMailboxTotalItemCount; "ExchangeMailboxStorageUsedByte" = $ExchangeMailboxStorageUsedByte; "SharePointSiteStorageUsedByte" = $SharePointSiteStorageUsedByte; "GroupId" = $GroupId; "ReportPeriod" = $ReportPeriod; "ShouldBeArchived" ="True"; "IsArchived" ="False" } -Connection $RootConnection -ErrorAction Stop
		}
		catch {
			Write-Output $LastActiveDate
			Write-Output "Could not add entry to SharePoint List - $($_.Exception.Message)"
			break
		}
		$Counter++
		continue   
	}

    # if Last Activity date exists, check M365 group activity over the last 30 days
	if ($LastActiveDate -lt $CheckDate){
		try {
			# No activity -> Teams should be archived
			$AddSPListPerm = Add-PnPListItem -List $SharePointListName -Values @{"ReportRefreshDate" = $ReportRefreshDate; "GroupDisplayName" = $GroupDisplayName; "Title" = $GroupDisplayName; "IsDeleted" = $IsDeleted; "GroupOwners" = $OwnerPrincipalName; "LastActivityDate" = $LastActiveDate; "GroupType" = $GroupType; "MemberCount" = $MemberCount; "ExternalMemberCount" = $ExternalMemberCount; "ExchangeReceivedEmailCount" = $ExchangeReceivedEmailCount; "SharePointActiveFileCount" = $SharePointActiveFileCount; "ExchangeMailboxTotalItemCount" = $ExchangeMailboxTotalItemCount; "ExchangeMailboxStorageUsedByte" = $ExchangeMailboxStorageUsedByte; "SharePointSiteStorageUsedByte" = $SharePointSiteStorageUsedByte; "GroupId" = $GroupId; "ReportPeriod" = $ReportPeriod; "ShouldBeArchived" ="True"; "IsArchived" ="False" } -Connection $RootConnection -ErrorAction Stop
		}
		catch {
			Write-Output $LastActiveDate
			Write-Output $GroupDisplayName
			Write-Output "Could not add entry to SharePoint List - $($_.Exception.Message)"
			break
		}
	}

	$Counter++
}


