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

# https://mmsharepoint.wordpress.com/2023/05/04/authentication-in-azure-automation-with-managed-identity-on-sharepoint-and-microsoft-graph/
# https://thesysadminchannel.com/graph-api-using-a-managed-identity-in-an-automation-runbook/


try {
    # Logging in to Azure.
    $AzCon = Connect-AzAccount -Identity
    # Connect to MgGraph - disabled because Microsoft Graph commands do not work consistently in Oct 2023
    # Connect-MgGraph -Identity -NoWelcome
} catch {
    Write-Error -Message $_.Exception
}

# If you are using Automation Account Variables
#
#$RootUrl = Get-AutomationVariable -Name "RootUrl"
#$SharePointList = Get-AutomationVariable -Name "ActivityList"
#
#####

# Define the SharePoint teamsite Url where the list is located - please edit
#
#####
$RootURL = "https://xxxxxxx.sharepoint.com/teams/TeamsAutomation/"

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

# Build Auth Token from AzConnect
# for Invoke-RestMethod
#
#####
$HeaderApp = @{
    Authorization = "$((Get-AzAccessToken -ResourceTypeName MSGraph).type) $((Get-AzAccessToken -ResourceTypeName MSGraph).token)"
}

# Query SPO List with inactive Teams
#
#####
$SiteItems = Get-PnPListItem -List $SharePointList -Connection $RootConnection  


ForEach ($SiteItem in $SiteItems) {
    $TeamsGroupId = $SiteItem["GroupId"]
    $ArchiveResponseByOwner = $SiteItem["ApprovedBy"]
    $ShouldBeArchived =  $SiteItem["ShouldBeArchived"]
    $ApprovedToArchive = $SiteItem["ApprovedToArchive"]
    $IsArchived = $SiteItem["IsArchived"]
    $GroupDisplayName = $SiteItem["GroupDisplayName"]

    # if Teams is archived, Should not be archived or
    # approval to archive is missing, skip entry
    #
    #####
    if (($ApprovedToArchive -ne $true) -Or ($IsArchived -eq $true) -Or ($ShouldBeArchived -eq $false))
    {
        #Write-Output "Teams $($GroupDisplayName) won't be archived..."
        continue
    }
    Write-Output "Team should be archived..."
    $TeamsArchiveApiUrl = "https://graph.microsoft.com/v1.0/teams/"+$TeamsGroupId+"/archive"
    try {
        $TeamsArchiveRequest = Invoke-RestMethod -Uri $TeamsArchiveApiUrl -Headers $HeaderApp -Method Post -ContentType 'application/json' #-ErrorAction $Stop
        # Microsoft Graph command - Oct 2023 permission error
        #$TeamsArchiveRequest = Invoke-MgGraphRequest -Uri $TeamsArchiveApiUrl -Method Post -ContentType 'application/json' #-ErrorAction $Stop
        #$TeamsArchiveRequest = Invoke-MgArchiveTeam -TeamId $TeamsGroupId
        $SetSiteItem = Set-PnPListItem -List $SharePointList -Identity $SiteItem.Id -Values @{"IsArchived" = "True"} -UpdateType SystemUpdate -Connection $RootConnection
        $SetSiteItem = Set-PnPListItem -List $SharePointList -Identity $SiteItem.Id -Values @{"ArchiveAzureRunbookStatus" = "Success"} -UpdateType SystemUpdate -Connection $RootConnection
    }
    catch {
        Write-Error -Message $_.Exception
        $SetSiteItem = Set-PnPListItem -List $SharePointList -Identity $SiteItem.Id -Values @{"ArchiveAzureRunbookStatus" = "Failed - Could not archive Team"} -UpdateType SystemUpdate -Connection $RootConnection

    }
}

