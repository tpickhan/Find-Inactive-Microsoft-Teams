# Script gets M365 Groups activity details and filters based on last activity date
# for inactive Teams.
# Inactive M365 groups are written to a SharePoint Online list
#
# created by Thorsten Pickhan
# Initial script created on 01.12.2022 (12/01/2022)
# 20231004 - Rewrite script to Managed Identity mode
#
# Version 1.1

# PowerShell 7.2 is required
# PNP PowerShell is required in Version
# Microsoft Graph PowerShell in version 2.6.1 required
# Microsoft.Graph.Authtentication
# Microsoft.Graph.Reports

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
    Connect-AzAccount -Identity | Out-Null
    # Get token and connect to MgGraph
    Connect-MgGraph -Identity -NoWelcome
} catch {
    Write-Error -Message $_.Exception
    throw $_.Exception
}

# Function to disable or enable concealed M365 Reports
# Attention: Graph API endpoint is still beta (Oct 2023)
#
#####
function SetM365ReportSettings ([bool]$action){
    Write-Output "Setting M365 Report settings..."
    $HeaderApp = @{
        Authorization = "$((Get-AzAccessToken -ResourceTypeName MSGraph).type) $((Get-AzAccessToken -ResourceTypeName MSGraph).token)"
    }

    # GraphAPI endpoint for report settings
    #
    ###

    $GraphApiUrl = "https://graph.microsoft.com/beta/admin/reportSettings"

    # Get current state for DisplayConcealedNames in M365 Usage Reports
    #
    ###
    $FeatureEnabled = $action
    $ReportSettingsGet = Invoke-RestMethod -Headers $HeaderApp -Uri $GraphApiUrl -ContentType 'application/json' -Method Get
    $CurrentStatus = $ReportSettingsGet.displayConcealedNames
    Write-Output "Current State: $($CurrentStatus)"
    Write-Output "Requested State: $($FeatureEnabled)"

    if ($CurrentStatus -ne $FeatureEnabled) {
        # if DisplayConcealedNames is false, enable it
        # else disable it
        #
        ###
        
        if ($FeatureEnabled -eq $true) {
            #Write-Output "Enabling displayConcealedNames to hidde display Name from M365 Groups, Owner and member"
            $GraphApiBody = "{
                ""displayConcealedNames"": ""true""
            }"
            $ReportSettingsSet = Invoke-RestMethod -Headers $HeaderApp -Uri $GraphApiUrl -Body $GraphApiBody -ContentType 'application/json' -Method Patch
            $ReportSettingsGet = Invoke-RestMethod -Headers $HeaderApp -Uri $GraphApiUrl -ContentType 'application/json' -Method Get
            #Write-Output "Setting after update: $($ReportSettingsGet.displayConcealedNames)"
        }

        if ($FeatureEnabled -eq $false) {
            #Write-Output "Disabling displayConcealedNames to display Name from M365 Groups, Owner and member"
            $GraphApiBody = "{
                ""displayConcealedNames"": ""false""
            }"
            $ReportSettingsSet = Invoke-RestMethod -Headers $HeaderApp -Uri $GraphApiUrl -Body $GraphApiBody -ContentType 'application/json' -Method Patch
            $ReportSettingsGet = Invoke-RestMethod -Headers $HeaderApp -Uri $GraphApiUrl -ContentType 'application/json' -Method Get
            #Write-Output "Setting after update: $($ReportSettingsGet.displayConcealedNames)"

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
$SharePointList = "Lists/InactiveTeams"


# Connect to SharePoint Online
#
#####
Write-Output "Start connecting to SharePoint Online.."
Connect-PnPOnline -ManagedIdentity -Url $RootUrl | Out-Null
$RootConnection = Get-PnPConnection

if ($RootConnection) {
    Write-Output "SharePoint Online succesfully connected!"
    }
else {
    Write-Output "SharePoint Online connection failed!"
    Write-Output "Script will stop now"
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
    $TempReports = Get-MgReportOffice365GroupActivityDetail -Period $PeriodOfReport -OutFile $CsvFileName -ErrorAction Stop
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

# If Concelead Report should be disbaled, run function to re-eanble it
#
#####
if ($DisConcealedDisplayName -eq "true") {
    SetM365ReportSettings -action $True
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

# Check and validate each M365 Group
#
#####
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
            Write-Output "Adding $($GroupDisplayName) to the list..."
			$AddSPListPerm = Add-PnPListItem -List $SharePointList -Values @{"ReportRefreshDate" = $ReportRefreshDate; "GroupDisplayName" = $GroupDisplayName; "Title" = $GroupDisplayName; "IsDeleted" = $IsDeleted; "GroupOwners" = $OwnerPrincipalName; "LastActivityDate" = $LastActiveDate; "GroupType" = $GroupType; "MemberCount" = $MemberCount; "ExternalMemberCount" = $ExternalMemberCount; "ExchangeReceivedEmailCount" = $ExchangeReceivedEmailCount; "SharePointActiveFileCount" = $SharePointActiveFileCount; "ExchangeMailboxTotalItemCount" = $ExchangeMailboxTotalItemCount; "ExchangeMailboxStorageUsedByte" = $ExchangeMailboxStorageUsedByte; "SharePointSiteStorageUsedByte" = $SharePointSiteStorageUsedByte; "GroupId" = $GroupId; "ReportPeriod" = $ReportPeriod; "ShouldBeArchived" ="True"; "IsArchived" ="False"; "ApprovedToArchive" ="False" } -Connection $RootConnection -ErrorAction Stop
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
            Write-Output "Adding $($GroupDisplayName) to the list..."
			$AddSPListPerm = Add-PnPListItem -List $SharePointList -Values @{"ReportRefreshDate" = $ReportRefreshDate; "GroupDisplayName" = $GroupDisplayName; "Title" = $GroupDisplayName; "IsDeleted" = $IsDeleted; "GroupOwners" = $OwnerPrincipalName; "LastActivityDate" = $LastActiveDate; "GroupType" = $GroupType; "MemberCount" = $MemberCount; "ExternalMemberCount" = $ExternalMemberCount; "ExchangeReceivedEmailCount" = $ExchangeReceivedEmailCount; "SharePointActiveFileCount" = $SharePointActiveFileCount; "ExchangeMailboxTotalItemCount" = $ExchangeMailboxTotalItemCount; "ExchangeMailboxStorageUsedByte" = $ExchangeMailboxStorageUsedByte; "SharePointSiteStorageUsedByte" = $SharePointSiteStorageUsedByte; "GroupId" = $GroupId; "ReportPeriod" = $ReportPeriod; "ShouldBeArchived" ="True"; "IsArchived" ="False"; "ApprovedToArchive" ="False" } -Connection $RootConnection -ErrorAction Stop
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
