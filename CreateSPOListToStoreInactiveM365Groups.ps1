# Script will create a SharePoint Online list to store ownerless M365 groups
#
# created by Thorsten Pickhan
# Initial script created on 07.06.2022 (06/07/2022)
#
# Version 1.0

# PNP PowerShell is required in Version

# Install-Module -Name PnP.Powershell
# Import-Module PNP.Powershell

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

# Define the SharePoint teamsite Url where the list should be created
$RootURL = "https://<Your Tenant name>.sharepoint.com/teams/TeamsAutomation/"

# Define the list Url
$SharePointListName = "Lists/InactiveTeams"

# Define the list display name 
$SharePointListDisplayName = "Inactive Teams"

# Connect to SharePoint Online
$RootConnection = Connect-PnPOnline -Url $RootUrl -Interactive -ReturnConnection -ErrorAction Stop

# Create a new generic List
$item = New-PnPList -Title $SharePointListDisplayName -Template GenericList -Url $SharePointListName -EnableVersioning -OnQuickLaunch -Connection $RootConnection

# Create required columns
$item = Add-PnPField -List $SharePointListName -DisplayName "Report Refresh Date" -InternalName "ReportRefreshDate" -Type DateTime -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Report Period" -InternalName "ReportPeriod" -Type Number -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Group Id" -InternalName "GroupId" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Group Display Name" -InternalName "GroupDisplayName" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Group Owners" -InternalName "GroupOwners" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Group Type" -InternalName "GroupType" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Is Deleted" -InternalName "IsDeleted" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Last Activity Date" -InternalName "LastActivityDate" -Type DateTime -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Member Count" -InternalName "MemberCount" -Type Number -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "External Member Count" -InternalName "ExternalMemberCount" -Type Number -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Exchange Received Email Count" -InternalName "ExchangeReceivedEmailCount" -Type Number -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Exchange Mailbox Total Item Count" -InternalName "ExchangeMailboxTotalItemCount" -Type Number -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Exchange Mailbox Storage Used Byte" -InternalName "ExchangeMailboxStorageUsedByte" -Type Number -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "SharePoint Active File Count" -InternalName "SharePointActiveFileCount" -Type Number -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "SharePoint Site Storage Used Byte" -InternalName "SharePointSiteStorageUsedByte" -Type Number -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Should be Archived" -InternalName "ShouldBeArchived" -Type Boolean -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Approved to Archive" -InternalName "ApprovedToArchive" -Type Boolean -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Approved by" -InternalName "ApprovedBy" -Type Text -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Is Archived" -InternalName "IsArchived" -Type Boolean -AddToDefaultView -Connection $RootConnection
$item = Add-PnPField -List $SharePointListName -DisplayName "Archive Azure Runbook Status" -InternalName "ArchiveAzureRunbookStatus" -Type Text -AddToDefaultView -Connection $RootConnection
