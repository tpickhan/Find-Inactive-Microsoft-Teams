# Script gets M365 Groups activity details and filters based on last activity date
# for inactive Teams.
# Inactive M365 groups are written to a SharePoint Online list
#
# created by Thorsten Pickhan
# Initial script created on 01.12.2022 (12/01/2022)
# 20231004 - Rewrite script to Managed Identity mode
# 20240415 - Filter M365 groups to Teams enabled only
# 20240714 - Edited routine to filter for Teams Enabled M365 Groups
#            Added Filter skipping Teams which are already archived
# 20260619 - Updated script to use Graph API for enabling/disabling concealed reports and updated Graph API calls to latest version of Microsoft Graph PowerShell SDK

# Version 1.3

# PowerShell 7.2 is required
# PNP PowerShell is NOT required anymore
# Microsoft Graph PowerShell in version 2.6.1 required
# Microsoft.Graph.Authtentication
# Microsoft.Graph.Reports
# Microsoft.Graph.Teams

# https://mmsharepoint.wordpress.com/2023/05/04/authentication-in-azure-automation-with-managed-identity-on-sharepoint-and-microsoft-graph/
# https://thesysadminchannel.com/graph-api-using-a-managed-identity-in-an-automation-runbook/

# Should M365 Report data be concealed?
# more information can be found here
# https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/miscellaneous/reports-show-anonymous-user-name
#
#####
$DisConcealedDisplayName = "True"

try {
    # Logging in to Azure.
    # Connect-AzAccount -Identity | Out-Null
    # Get token and connect to MgGraph
    Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop
    $MGContext = Get-MgContext
    Write-Output "Connected to Microsoft Graph with tenant $($MGContext.Tenant.Id)"
}
catch {
    Write-Error -Message $_.Exception
    throw $_.Exception
}

# Function to disable or enable concealed M365 Reports
# Attention: Graph API endpoint is still beta (Oct 2023)
#
#####
function SetM365ReportSettings ([bool]$action){
    Write-Output "Setting M365 Report settings..."
    # Get current state for DisplayConcealedNames in M365 Usage Reports
    #
    ###
    $FeatureEnabled = $action
    $ReportSettingsGet = Get-MgAdminReportSetting
    $CurrentStatus = $ReportSettingsGet.displayConcealedNames
    Write-Output "Current State: $($CurrentStatus)"
    Write-Output "Requested State: $($FeatureEnabled)"

    if ($CurrentStatus -ne $FeatureEnabled) {
        # if DisplayConcealedNames is false, enable it
        # else disable it
        #
        ###
        
        if ($FeatureEnabled -eq $true) {
            # Update-MgAdminReportSetting
            $ReportSettingsSet = Update-MgAdminReportSetting -DisplayConcealedNames
            $ReportSettingsGet = Get-MgAdminReportSetting
            Write-Output "Setting after update: $($ReportSettingsGet.displayConcealedNames)"
        }

        if ($FeatureEnabled -eq $false) {
            $ReportSettingsSet = Update-MgAdminReportSetting
            $ReportSettingsGet = Get-MgAdminReportSetting
            Write-Output "Setting after update: $($ReportSettingsGet.displayConcealedNames)"

        }
    }
    else {
        Write-Output "no change required"
    }

}

# Define the SharePoint teamsite Url where the list is located - please edit
#
#####
$RootURL = "https://xxxxx.sharepoint.com/teams/TeamsAutomation/"

# Define the SharePoint Online List Id - please edit
#
#####
$SharePointList = "InactiveTeams"


# Connect to SharePoint Online
#
#####
Write-Output "Start connecting to SharePoint Online.."
try {
    $SPOSiteUrl = Get-AutomationVariable -Name "SPO-SiteURL" -ErrorAction SilentlyContinue
    }
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Error in getting Automation Variable SPo-SiteURL"
    break
}

#region RampUp Connection Details
########################################################
##             Block 1 - Create Base URL
##          
########################################################
# Capture the current timestamp in the format of Year-Month-Day Hour:Minute:Second
# $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
# Output the timestamp with a message indicating the start of Azure Variables ramp-up
# Write-Output "$TimeStamp - RampUp - Azure Variables"


# Get SharePoint Root Site Id
try {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Get SharePoint Root Site ID"
    $SPORoot = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/root" -Method GET -ErrorAction Stop
    $SPORootId = $SPORoot.Id
    #Write-Output $SPORootId
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - SharePoint Root Site ID retrieved successfully"
}
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Error "$TimeStamp - RampUp - Failed to get SharePoint Root Site Id!"
    throw "$TimeStamp - RampUp - Failed to get SharePoint Root Site Id!"
    Write-Output $Error
    break
}

# Get SharePoint Sub Site Id and WebUrl from SPO List

try {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Get SharePoint Sub Site ID"
    #Write-Output "SPOSiteUrl: $($SPOSiteUrl)"
    #Write-Output "SPORootId: $($SPORootId)"
    $SPOSiteURLTemp = $SPORootId+":"+$SPOSiteUrl
    $SPOSub = Get-MgSite -SiteId $SPOSiteURLTemp
    $SPOSubId = $SPOSub.Id
}
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Failed to get SharePoint Sub Site Id!"
    Write-Output $Error
    break
}

# Get SharePoint Root Sub Site Lists
try { 
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Get required SharePoint Lists"
    $SPOSubLists = Get-MgSiteList -SiteId $SPOSubId
}
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Error "$TimeStamp - RampUp - Failed to get SharePoint Site Lists!"
    throw "$TimeStamp - RampUp - Could not connect to SharePoint Site Lists!"
    Write-Output $Error
    break
}

# Get SharePoint Teams Requests List Id and WebUrl
try { 
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Get SharePoint Teams Request List"
    $SPOTeamsInactiveList = $SPOSubLists | Where-Object {$_.name -eq $SharePointList}
    $SPOTeamsInactiveListId = $SPOTeamsInactiveList.Id
}
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Error "$TimeStamp - RampUp - Failed to get SharePoint Teams Request List!"
    throw "$TimeStamp - RampUp - Could not connect to SharePoint Teams Request List!"
    Write-Output $Error
    break
}


# Define the period of report [D7, D30, D90, D180]
#
#####
$PeriodOfReport = "D30"

# Define the CSV file path for usage report data
#
#####
$CsvFileName = ".\report.csv"

# If Concelead Reports should be disbaled, run function
#
#####
if ($DisConcealedDisplayName -eq "true") {
    SetM365ReportSettings -action $False
}

try {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Start retrieving M365 Groups activity data for period $($PeriodOfReport)"
    $TempReports = Get-MgReportOffice365GroupActivityDetail -Period $PeriodOfReport -OutFile $CsvFileName -ErrorAction Stop
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - M365 Groups activity data retrieved successfully"
    # If Concelead Report should be disbaled, run function to re-eanble it
    #
    #####
    if ($DisConcealedDisplayName -eq "true") {
        SetM365ReportSettings -action $True
    }
}
catch {
    # If Concelead Reports should be disbaled, run function to reeanble it
    #
    #####
    if ($DisConcealedDisplayName -eq "true") {
        SetM365ReportSettings -action $True
    }
    Write-Error -Message $_.Exception
}



# Import M365 Usage report data
#
#####
$UsageData = Import-Csv $CsvFileName

$Count = $UsageData.count
$Counter = 1

# Set check date to compare with the last activity date
# Checkdate is set to today minus 30 days
# please customize to fit your business needs
#
#####
$CheckDate = (Get-Date).adddays(-30)

# Get all Microsoft Teams
#
######
try {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Start retrieving all Microsoft Teams in the tenant"
    $AllTeams = Get-MgTeam -All -ErrorAction Stop
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - All Microsoft Teams retrieved successfully"
}
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Error "$TimeStamp - RampUp - Failed to retrieve Microsoft Teams in the tenant!"
    throw "$TimeStamp - RampUp - Failed to retrieve Microsoft Teams in the tenant!"
    Write-Output $Error
    break
}

# Check and validate each M365 Group
#
#####
ForEach ($UsageRecord in $UsageData) {
	Write-Output "Proceed list entry $($Counter) from $($Count)..."
	if ($UsageRecord.'Is Deleted' -eq "True") {
		$Counter++
		continue
	}

   # Get Group ID and validate if it is Teams enabled
	$GroupId = $UsageRecord."Group Id"
    $TeamsEnabled = $AllTeams | Where-Object { $_.Id -eq $GroupId }
    if ($TeamsEnabled -eq $null) {
        Write-Output "M365 Group with Id $($GroupId) is not Teams enabled - skip this record"
        Write-Output "Skipping this M365Group"
        $Counter++
        #break
        continue
    }
    
    # Check if Team is already archived
    if ($TeamsEnabled.IsArchived -eq $True) {
        Write-Output "Teams $($TeamsEnabled.Id) is already archived - skip this record"
        $Counter++
        continue
    }

    # Set columne values for SharePoint list entry
	$ReportRefreshDate = $UsageRecord."Report Refresh Date"
    $ReportRefreshDate = Get-Date $ReportRefreshDate
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
	$ReportPeriod = $UsageRecord."Report Period"

    # Check if there is a last activity date in CSV
	if ($LastActivityDate) {
		$LastActiveDate = Get-Date $LastActivityDate
	}
	else {
		$LastActiveDate = Get-Date "01.01.1900"
    }
    <#
		try {
			# No activity Teams should be archived
            $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
            Write-Output "$($TimeStamp) - Inactive Teams - Adding $($GroupDisplayName) to the list..."
            $Fields = @{
                fields = @{
                    TeamsDeploymentStatus = "Provisioning"
                    ReportRefreshDate = $ReportRefreshDate
                    GroupDisplayName = $GroupDisplayName
                    Title = $GroupDisplayName
                    IsDeleted = $IsDeleted
                    GroupOwners = $OwnerPrincipalName
                    LastActivityDate = $LastActiveDate
                    GroupType = $GroupType
                    MemberCount = $MemberCount
                    ExternalMemberCount = $ExternalMemberCount
                    ExchangeReceivedEmailCount = $ExchangeReceivedEmailCount
                    SharePointActiveFileCount = $SharePointActiveFileCount
                    ExchangeMailboxTotalItemCount = $ExchangeMailboxTotalItemCount
                    ExchangeMailboxStorageUsedByte = $ExchangeMailboxStorageUsedByte
                    SharePointSiteStorageUsedByte = $SharePointSiteStorageUsedByte
                    GroupId = $GroupId
                    ReportPeriod = $ReportPeriod
                    ShouldBeArchived = "True"
                    IsArchived = "False"
                    ApprovedToArchive = "False"
                }
            }
            #Write-Output "Adding $($GroupDisplayName) to the list..."
			#$AddSPListPerm = Add-PnPListItem -List $SharePointList -Values @{"ReportRefreshDate" = $ReportRefreshDate; "GroupDisplayName" = $GroupDisplayName; "Title" = $GroupDisplayName; "IsDeleted" = $IsDeleted; "GroupOwners" = $OwnerPrincipalName; "LastActivityDate" = $LastActiveDate; "GroupType" = $GroupType; "MemberCount" = $MemberCount; "ExternalMemberCount" = $ExternalMemberCount; "ExchangeReceivedEmailCount" = $ExchangeReceivedEmailCount; "SharePointActiveFileCount" = $SharePointActiveFileCount; "ExchangeMailboxTotalItemCount" = $ExchangeMailboxTotalItemCount; "ExchangeMailboxStorageUsedByte" = $ExchangeMailboxStorageUsedByte; "SharePointSiteStorageUsedByte" = $SharePointSiteStorageUsedByte; "GroupId" = $GroupId; "ReportPeriod" = $ReportPeriod; "ShouldBeArchived" ="True"; "IsArchived" ="False"; "ApprovedToArchive" ="False" } -Connection $RootConnection -ErrorAction Stop
            $AddSPListPerm = Add-MgSiteListItem -SiteId $SPOSubId -ListId $SPOTeamsInactiveListId -BodyParameter $Fields -ErrorAction Stop
		}
		catch {
			Write-Output $LastActiveDate
			Write-Output "Could not add entry to SharePoint List - $($_.Exception.Message)"
			break
		}
		$Counter++
		continue   
	}
    #>
    # if Last Activity date exists, check M365 group activity over the last 30 days
	if ($LastActiveDate -lt $CheckDate){
		try {
			# No activity -> Teams should be archived
            Write-Output "Adding $($GroupDisplayName) to the list..."
            $Fields = @{
                fields = @{
                    ReportRefreshDate = $ReportRefreshDate
                    GroupDisplayName = $GroupDisplayName
                    Title = $GroupDisplayName
                    IsDeleted = $IsDeleted
                    GroupOwners = $OwnerPrincipalName
                    LastActivityDate = $LastActiveDate
                    GroupType = $GroupType
                    MemberCount = $MemberCount
                    ExternalMemberCount = $ExternalMemberCount
                    ExchangeReceivedEmailCount = $ExchangeReceivedEmailCount
                    SharePointActiveFileCount = $SharePointActiveFileCount
                    ExchangeMailboxTotalItemCount = $ExchangeMailboxTotalItemCount
                    ExchangeMailboxStorageUsedByte = $ExchangeMailboxStorageUsedByte
                    SharePointSiteStorageUsedByte = $SharePointSiteStorageUsedByte
                    GroupId = $GroupId
                    ReportPeriod = $ReportPeriod
                    ShouldBeArchived = 1
                    IsArchived = 0
                    ApprovedToArchive = 0
                }     
            }
            <#
            $Fields = @{
>>                 fields = @{
>>                     GroupDisplayName = $GroupDisplayName
>>                     Title = $GroupDisplayName
>>                     IsDeleted = $IsDeleted
>>                     GroupOwners = $OwnerPrincipalName
>>                     GroupType = $GroupType
>>                     MemberCount = $MemberCount
>>                     GroupId = $GroupId
>>                     ReportPeriod = $ReportPeriod
>>                 }
>>             }
#>
            $AddSPListPerm = New-MgSiteListItem -SiteId $SPOSubId -ListId $SPOTeamsInactiveListId -BodyParameter $Fields -ErrorAction Stop
		}
		catch {
			Write-Output $LastActiveDate
			Write-Output $GroupDisplayName
			Write-Output "Could not add entry to SharePoint List"# - $($_.Exception.Message)"
			break
		}
	}

	$Counter++
}
