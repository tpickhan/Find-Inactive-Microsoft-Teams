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
#region RampUp Variables
########################################################
##             Block 0 - Define Variables
##          
########################################################
# Capture the current timestamp in the format of Year-Month-Day Hour:Minute:Second
# $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
# Output the timestamp with a message indicating the start of Azure Variables ramp-up
# Write-Output "$TimeStamp - RampUp - Azure Variables"

$TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
Write-Output "$TimeStamp - RampUp - Define Variables and Functions"
$DisConcealedDisplayName = "True"

# Define the SharePoint Online List Id - please edit
#
#####
$SharePointListName = "InactiveTeams"

# Define the period of report [D7, D30, D90, D180]
#
#####
$PeriodOfReport = "D30"

# Define the CSV file path for usage report data
#
#####
$CsvFileName = ".\report.csv"

# Set check date to compare with the last activity date
# Checkdate is set to today minus 30 days
# please customize to fit your business needs
#
#####
$CheckDate = (Get-Date).adddays(-30)

# Get Automation Variable for SharePoint Online Site URL
#
#####
$TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
Write-Output "$TimeStamp - RampUp - Getting Automation Variable SPOSiteURL..."
try {
    $SPOSiteUrl = Get-AutomationVariable -Name "SPOSiteURL" -ErrorAction Stop
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Automation Variable SPOSiteURL retrieved successfully: $SPOSiteUrl"
    }
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Error in getting Automation Variable SPOSiteURL"
    break
}

# Function to disable or enable concealed M365 Reports
# Attention: Graph API endpoint is still beta (Oct 2023)
#
#####
function SetM365ReportSettings ([bool]$action){
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Setting M365 Report settings..."
    # Get current state for DisplayConcealedNames in M365 Usage Reports
    #
    ###
    $FeatureEnabled = $action
    $ReportSettingsGet = Get-MgAdminReportSetting
    $CurrentStatus = $ReportSettingsGet.displayConcealedNames
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Current State: $($CurrentStatus)"
    Write-Output "$TimeStamp - RampUp - Requested State: $($FeatureEnabled)"

    if ($CurrentStatus -ne $FeatureEnabled) {
        # if DisplayConcealedNames is false, enable it
        # else disable it
        #
        ###
        
        if ($FeatureEnabled -eq $true) {
            # Update-MgAdminReportSetting
            $ReportSettingsSet = Update-MgAdminReportSetting -DisplayConcealedNames
            $ReportSettingsGet = Get-MgAdminReportSetting
            $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
            Write-Output "$TimeStamp - RampUp - Setting after update: $($ReportSettingsGet.displayConcealedNames)"
        }

        if ($FeatureEnabled -eq $false) {
            $ReportSettingsSet = Update-MgAdminReportSetting
            $ReportSettingsGet = Get-MgAdminReportSetting
            $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
            Write-Output "$TimeStamp - RampUp - Setting after update: $($ReportSettingsGet.displayConcealedNames)"

        }
    }
    else {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$TimeStamp - RampUp - Admin Reporting Settings ok - no change required"
    }

}


#region RampUp Connection Details
########################################################
##             Block 1 - Connect MG Graph with Managed Identity
##          
########################################################


try {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Connect to Microsoft Graph with Managed Identity"
    Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop
    $MGContext = Get-MgContext
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Connected to Microsoft Graph with tenant $($MGContext.TenantId)"
}
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Error Connecting to Microsoft Graph: $_.Exception.Message"
    throw $_.Exception
}

#region RampUp Connection Details
########################################################
##             Block 2 - Create Base URL
##          
########################################################


# Get SharePoint Root Site Id
try {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Get SharePoint Root Site ID"
    $SPORoot = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/root" -Method GET -ErrorAction Stop
    $SPORootId = $SPORoot.Id
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
if ($null -eq $SPOSiteUrl) {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - No SharePoint Site URL provided, please check Automation Variable SPOSiteURL"
    break
}

try {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Get SharePoint Sub Site ID for URL $($SPOSiteUrl)"
    #Write-Output "SPOSiteUrl: $($SPOSiteUrl)"
    #Write-Output "SPORootId: $($SPORootId)"
    $SPOSiteURLTemp = $SPORootId+":"+$SPOSiteUrl
    $SPOSub = Get-MgSite -SiteId $SPOSiteURLTemp
    $SPOSubId = $SPOSub.Id
}
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Failed to get SharePoint Sub Site Id for URL $($SPOSiteUrl)!"
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
    Write-Output "$TimeStamp - RampUp - Get SharePoint Teams Inactive List"
    $SPOTeamsInactiveList = $SPOSubLists | Where-Object {$_.name -eq $SharePointListName}
    $SPOTeamsInactiveListId = $SPOTeamsInactiveList.Id
}
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Error "$TimeStamp - RampUp - Failed to get SharePoint Teams Inactive List!"
    throw "$TimeStamp - RampUp - Could not connect to SharePoint Teams Inactive List!"
    Write-Output $Error
    break
}

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

#region RampUp Connection Details
########################################################
##             Block 3 - Get Office 365 Groups Activity Details and write to SharePoint List
##          
########################################################


# If Concelead Reports should be disbaled, run function
#
#####
if ($DisConcealedDisplayName -eq "true") {
    SetM365ReportSettings -action $False
}

try {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - Action - Start retrieving M365 Groups activity data for period $($PeriodOfReport)"
    $TempReports = Get-MgReportOffice365GroupActivityDetail -Period $PeriodOfReport -OutFile $CsvFileName -ErrorAction Stop
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - Action - M365 Groups activity data retrieved successfully"
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
$TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
Write-Output "$TimeStamp - Action - Importing M365 Groups activity data from CSV file $($CsvFileName)"
$UsageData = Import-Csv $CsvFileName

$Count = $UsageData.count
$Counter = 1



# Check and validate each M365 Group
#
#####
ForEach ($UsageRecord in $UsageData) {
	$TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - Action - Proceed list entry $($Counter) from $($Count)..."
	if ($UsageRecord.'Is Deleted' -eq "True") {
		$Counter++
		continue
	}

   # Get Group ID and validate if it is Teams enabled
	$GroupId = $UsageRecord."Group Id"
    $TeamsEnabled = $AllTeams | Where-Object { $_.Id -eq $GroupId }
    if ($null -eq $TeamsEnabled) {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$TimeStamp - Action - M365 Group with Id $($GroupId) is not Teams enabled - skip this record"
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - Action - Skipping this M365Group"
        $Counter++
        #break
        continue
    }
    
    # Check if Team is already archived
    if ($TeamsEnabled.IsArchived -eq $True) {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$TimeStamp - Action - Teams $($TeamsEnabled.Id) is already archived - skip this record"
        $Counter++
        continue
    }

    # Set columne values for SharePoint list entry
	$ReportRefreshDate = (Get-Date $UsageRecord."Report Refresh Date").ToString("o")
	$GroupDisplayName = $UsageRecord."Group Display Name"
	$IsDeleted = $UsageRecord."Is Deleted"
	$OwnerPrincipalName = $UsageRecord."Owner Principal Name"
	$GroupType = $UsageRecord."Group Type"
	$MemberCount = $UsageRecord."Member Count"
	$ExternalMemberCount = $UsageRecord."External Member Count"
    if ($UsageRecord.'Exchange Received Email Count'){
	    $ExchangeReceivedEmailCount = $UsageRecord."Exchange Received Email Count"
    } else {
        $ExchangeReceivedEmailCount = 0
    }
	if ($UsageRecord."SharePoint Active File Count"){
	    $SharePointActiveFileCount = $UsageRecord."SharePoint Active File Count"
    } else {
        $SharePointActiveFileCount = 0
    }
	if ($UsageRecord."Exchange Mailbox Total Item Count"){
        $ExchangeMailboxTotalItemCount = $UsageRecord."Exchange Mailbox Total Item Count"
    } else {
        $ExchangeMailboxTotalItemCount = 0
    }
	if ($UsageRecord."Exchange Mailbox Storage Used (Byte)"){
        $ExchangeMailboxStorageUsedByte = $UsageRecord."Exchange Mailbox Storage Used (Byte)"
    } else {
        $ExchangeMailboxStorageUsedByte = 0
    }
	if ($UsageRecord."SharePoint Site Storage Used (Byte)"){
        $SharePointSiteStorageUsedByte = $UsageRecord."SharePoint Site Storage Used (Byte)"
    } else {
        $SharePointSiteStorageUsedByte = 0
    }
	$ReportPeriod = $UsageRecord."Report Period"

    # Check if there is a last activity date in CSV
    if ($UsageRecord."Last Activity Date") {
        $LastActiveDate = (Get-Date $UsageRecord."Last Activity Date").ToString("o")
    } else {
		$LastActiveDate = (Get-Date "01.01.1900").ToString("o")
    }
 
    # if Last Activity date exists, check M365 group activity over the last 30 days
	if ($LastActiveDate -lt $CheckDate){
		try {
			# No activity -> Teams should be archived
            $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
            Write-Output "$TimeStamp - Action - Adding $($GroupDisplayName) to the list..."
            $Fields = @{
                fields = @{
                    GroupDisplayName = $GroupDisplayName
                    Title = $GroupDisplayName
                    IsDeleted = $IsDeleted
                    GroupOwners = $OwnerPrincipalName
                    GroupType = $GroupType
                    MemberCount = $MemberCount
                    ExternalMemberCount = $ExternalMemberCount
                    GroupId = $GroupId
                    ReportPeriod = $ReportPeriod
                    ReportRefreshDate = $ReportRefreshDate
                    LastActivityDate = $LastActiveDate
                    ExchangeReceivedEmailCount = $ExchangeReceivedEmailCount
                    SharePointActiveFileCount = $SharePointActiveFileCount
                    ExchangeMailboxTotalItemCount = $ExchangeMailboxTotalItemCount
                    ExchangeMailboxStorageUsedByte = $ExchangeMailboxStorageUsedByte
                    SharePointSiteStorageUsedByte = $SharePointSiteStorageUsedByte
                    ShouldBeArchived = $True
                    IsArchived = $False
                    ApprovedToArchive = $False
                }     
            }
            $AddSPListPerm = New-MgSiteListItem -SiteId $SPOSubId -ListId $SPOTeamsInactiveListId -BodyParameter $Fields -ErrorAction Stop
		}
		catch {
			Write-Output $LastActiveDate
			Write-Output $GroupDisplayName
			$TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
            Write-Output "$TimeStamp - Action - Could not add entry to SharePoint List"# - $($_.Exception.Message)"
			break
		}
	}

	$Counter++
}
