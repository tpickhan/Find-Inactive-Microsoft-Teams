# Script will create a SharePoint Online list to store ownerless M365 groups
#
# created by Thorsten Pickhan
# Initial script created on 07.06.2022 (06/07/2022)
#
# Version 1.0

# PNP PowerShell is required in Version

# Install-Module -Name PnP.Powershell
# Import-Module PNP.Powershell

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


# Define the list Url
# $SharePointListUrl = "Lists/InactiveTeams"

# Define the list display name 
#
######
$SharePointListDisplayName = "Inactive Teams"

# Define the SharePoint Online List Id - please edit
#
#####
$SharePointListName = $SharePointListDisplayName -replace " ",""


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
##             Block 2 - Find Base URL and check for existing SharePoint List
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

# Check if SPO List for Inactive Teams already exists
$TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
Write-Output "$TimeStamp - RampUp - Get SharePoint Teams Inactive List"
$SPOTeamsInactiveList = $SPOSubLists | Where-Object {$_.name -eq $SharePointListName}
if ($SPOTeamsInactiveList) {
    $SPOTeamsInactiveListId = $SPOTeamsInactiveList.Id
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - SharePoint Teams Inactive List retrieved successfully with Id $($SPOTeamsInactiveListId)"
    Write-Output "$TimeStamp - RampUp - No need to create the list, it already exists!"
    Write-Output "$TimeStamp - RampUp - Please check the list and delete it if you want to create a new one with the script!"
    break
}


#region RampUp Connection Details
########################################################
##             Block 3 - Create New SharePoint List to store Inactive Teams
##          
########################################################

# Create a new generic List
# $item = New-PnPList -Title $SharePointListDisplayName -Template GenericList -Url $SharePointListName -EnableVersioning -OnQuickLaunch -Connection $RootConnection

# Create SPO List
#
#########
$TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
Write-Output "$($TimeStamp) - RampUp - Creating SharePoint Online List for managing Inactive Teams"
$params = @{
    displayName = $SharePointListName
    description = "List of Inactive Teams identified by the Find-InactiveTeams script. This list is used to track Teams that are candidates for archiving based on their inactivity."
    list = @{
        displayName = $SharePointListName
        template = "genericList"
    }
}
try {
    $SPOListInactiveTeams = New-MgSiteList -SiteId $SPOSubId -BodyParameter $params -ErrorAction Stop
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$($TimeStamp) - RampUp - SharePoint Online List for managing Inactive Teams created."

} catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$($TimeStamp) - Error creating SharePoint Online List for managing Inactive Teams"
    Write-Output $Error[0]
    break
}

# Create columns
#
#########

if ($SPOListInactiveTeams){
    $SPOListId = $SPOListInactiveTeams.Id

    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$($TimeStamp) - RampUp - Creating required columns for SharePoint Online List"

    <#
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
    #>

    # Column Group Display Name
    $params = @{
        description = "Group Display Name"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "GroupDisplayName"
        displayName = "Group Display Name"
        text = @{
            allowMultipleLines = $false
            appendChangesToExistingText = $false
            linesForEditing = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Group Display Name'"
        Write-Output $Error[0]
        break
    }

    # Column Is Deleted
    $params = @{
        description = "Is Deleted"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "IsDeleted"
        displayName = "Is Deleted"
        text = @{
            allowMultipleLines = $false
            appendChangesToExistingText = $false
            linesForEditing = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Is Deleted'"
        Write-Output $Error[0]
        break
    }

    # Column Group Owners
    $params = @{
        description = "Group Owners"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "GroupOwners"
        displayName = "Group Owners"
        text = @{
            allowMultipleLines = $false
            appendChangesToExistingText = $false
            linesForEditing = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Group Owners'"
        Write-Output $Error[0]
        break
    }

    # Column Group Type
    $params = @{
        description = "Group Type"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "GroupType"
        displayName = "Group Type"
        text = @{
            allowMultipleLines = $false
            appendChangesToExistingText = $false
            linesForEditing = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Group Type'"
        Write-Output $Error[0]
        break
    }

    # Column Member Count
    $params = @{
        description = "Member Count"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "MemberCount"
        displayName = "Member Count"
        number = @{
            decimalPlaces = 0
        }

    }
    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Member Count'"
        Write-Output $Error[0]
        break
    }

    # Column External Member Count
    $params = @{
        description = "External Member Count"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "ExternalMemberCount"
        displayName = "External Member Count"
        number = @{
            decimalPlaces = 0
        }

    }
    
    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'External Member Count'"
        Write-Output $Error[0]
        break
    }

    # Column Group Id
    $params = @{
        description = "Group Id"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "GroupId"
        displayName = "Group Id"
        text = @{
            allowMultipleLines = $false
            appendChangesToExistingText = $false
            linesForEditing = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Group Id'"
        Write-Output $Error[0]
        break
    }

    #Column Report Period
    $params = @{
        description = "Report Period"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "ReportPeriod"
        displayName = "Report Period"
        number = @{
            decimalPlaces = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Report Period'"
        Write-Output $Error[0]
        break
    }

    # Column Report Refresh Date
    $params = @{
        description = "Report Refresh Date"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "ReportRefreshDate"
        displayName = "Report Refresh Date"
        dateTime = @{
            displayAs = "dateOnly"
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Report Refresh Date'"
        Write-Output $Error[0]
        break
    }

    # column Last Activity Date
    $params = @{
        description = "Last Activity Date"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "LastActivityDate"
        displayName = "Last Activity Date"
        dateTime = @{
            displayAs = "dateOnly"
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Last Activity Date'"
        Write-Output $Error[0]
        break
    }

    #Column Should be Archived
    $params = @{
        name = "ShouldBeArchived"
        displayName = "Should Be Archived"
        description = "Should be Archived"
        boolean = [pscustomobject]@{}
        required = $false
        hidden              = $false
        indexed             = $false
        enforceUniqueValues = $false
    } | ConvertTo-Json -Depth 10

    
    try {
        $NewSPOColumn = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$SPOSubId/lists/$SPOListId/columns" -Method POST -Body $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Should Be Archived'"
        Write-Output $Error[0]
        break
    }

    # Column Approved to Archive
    $params = @{
        name = "ApprovedToArchive"
        displayName = "Approved to Archive"
        description = "Approved to Archive"
        boolean = [pscustomobject]@{}
        required = $false
        hidden              = $false
        indexed             = $false
        enforceUniqueValues = $false
    } | ConvertTo-Json -Depth 10

    
    try {
        $NewSPOColumn = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$SPOSubId/lists/$SPOListId/columns" -Method POST -Body $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Approved to Archive'"
        Write-Output $Error[0]
        break
    }

    #Column Approved by
    $params = @{
        description = "Approved by"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "ApprovedBy"
        displayName = "Approved by"
        text = @{
            allowMultipleLines = $false
            appendChangesToExistingText = $false
            linesForEditing = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Approved by'"
        Write-Output $Error[0]
        break
    }

    # Column Is Archived
    $params = @{
        name = "IsArchived"
        displayName = "Is Archived"
        description = "Is Archived"
        boolean = [pscustomobject]@{}
        required = $false
        hidden              = $false
        indexed             = $false
        enforceUniqueValues = $false
    } | ConvertTo-Json -Depth 10

    
    try {
        $NewSPOColumn = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$SPOSubId/lists/$SPOListId/columns" -Method POST -Body $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Is Archived'"
        Write-Output $Error[0]
        break
    }

    # Column ExchangeReceivedEmailCount Number
    $params = @{
        description = "Exchange Received Email Count"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "ExchangeReceivedEmailCount"
        displayName = "Exchange Received Email Count"
        number = @{
            decimalPlaces = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Exchange Received Email Count'"
        Write-Output $Error[0]
        break
    }

    # Column Exchange Mailbox Total Item Count Number
    $params = @{
        description = "Exchange Mailbox Total Item Count"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "ExchangeMailboxTotalItemCount"
        displayName = "Exchange Mailbox Total Item Count"
        number = @{
            decimalPlaces = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Exchange Mailbox Total Item Count'"
        Write-Output $Error[0]
        break
    }

    # Column Exchange Mailbox Storage Used Byte Number
    $params = @{
        description = "Exchange Mailbox Storage Used Byte"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "ExchangeMailboxStorageUsedByte"
        displayName = "Exchange Mailbox Storage Used Byte"
        number = @{
            decimalPlaces = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Exchange Mailbox Storage Used Byte'"
        Write-Output $Error[0]
        break
    }

    # Column SharePoint Active File Count Number
    $params = @{
        description = "SharePoint Active File Count"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "SharePointActiveFileCount"
        displayName = "SharePoint Active File Count"
        number = @{
            decimalPlaces = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'SharePoint Active File Count'"
        Write-Output $Error[0]
        break
    }

    # Column SharePoint Site Storage Used Byte Number
    $params = @{
        description = "SharePoint Site Storage Used Byte"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "SharePointSiteStorageUsedByte"
        displayName = "SharePoint Site Storage Used Byte"
        number = @{
            decimalPlaces = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'SharePoint Site Storage Used Byte'"
        Write-Output $Error[0]
        break
    }


    # Column Archive Azure Runbook Status
    $params = @{
        description = "Archive Azure Runbook Status"
        enforceUniqueValues = $false
        hidden = $false
        indexed = $false
        name = "ArchiveAzureRunbookStatus"
        displayName = "Archive Azure Runbook Status"
        text = @{
            allowMultipleLines = $false
            appendChangesToExistingText = $false
            linesForEditing = 0
        }

    }

    try {
        $NewSPOColumn = New-MgSiteListColumn -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
    } catch {
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        Write-Output "$($TimeStamp) - Error creating column 'Archive Azure Runbook Status'"
        Write-Output $Error[0]
        break
    }

    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$($TimeStamp) - Run - New SPO List created for Inactive Teams management with all required columns."
}

try {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$($TimeStamp) - Run - Updating SPO List Display Name for Inactive Teams Management"
    $params = @{
        displayName = "Inactive Teams"
    }
    $UpdateSPOList = Update-MgSiteList -SiteId $SPOSubId -ListId $SPOListId -BodyParameter $params -ErrorAction Stop
}
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$($TimeStamp) - Run - Error: Error updating SPO List Display Name for Inactive Teams Management"
    Write-Output "$($TimeStamp) - Run - Error: $($_.Exception.Message)"
    Write-Output "$($TimeStamp) - Run - Error: $($_.Exception.StackTrace)"
    Write-Output "$($TimeStamp) - Run - Error: $($_.Exception.InnerException)"
}

#####
#region Closing connections and clean up
########################################################
##             Block 4 - Closing connections and clean up
##          
########################################################
try {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Disconnecting from Microsoft Graph"
    $DisconnectMG = Disconnect-MgGraph
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Disconnected from Microsoft Graph successfully"
}
catch {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - RampUp - Error disconnecting from Microsoft Graph: $($_.Exception.Message)"
}



