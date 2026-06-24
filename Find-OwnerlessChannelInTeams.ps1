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
$SharePointListName = "OwnerlessChannelsInTeams"

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

# Process
# going through each Teams
$OwnerlessChannelsInTeams = @()
$TeamsCounter=0
ForEach ($Team in $AllTeams) {
    $TeamsCounter++
    $GroupId = $Team.Id
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - Run - Analyzing Team $($TeamsCounter) of $($AllTeams.Count)"
    $AllTeamsChannels = Get-MgTeamChannel -TeamId $GroupId
    $ChannelCount = 0
    ForEach ($Channel in $AllTeamsChannels) {
        $ChannelCount++
        $ChannelId = $Channel.Id
        $ChannelName = $Channel.DisplayName
        $ChannelType = $Channel.MembershipType
        $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
        #Write-Output "$TimeStamp - Run - Getting Members for Channel: $ChannelName in Team: $($Team.DisplayName) - Channel $($ChannelCount) of $($AllTeamsChannels.Count)"
        $ChannelMembers = Get-MgTeamChannelMember -TeamId $GroupId -ChannelId $ChannelId
        $ChannelOwners = $ChannelMembers | Where-Object {$_.Roles -contains "owner"}

        if ($ChannelOwners.Count -eq 0) {
            $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
            Write-Output "$TimeStamp - Run - Channel: $ChannelName in Team: $($Team.DisplayName) has no owners"
            $OwnerlessChannelsInTeams += [PSCustomObject]@{
                TeamName = $Team.DisplayName
                TeamId = $GroupId
                ChannelName = $ChannelName
                ChannelId = $ChannelId
                ChannelType = $ChannelType
            }
            $param = @{
                TeamName = $Team.DisplayName
                TeamId = $GroupId
                ChannelName = $ChannelName
                ChannelId = $ChannelId
                ChannelType = $ChannelType
            }
            $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
            Write-Output "$TimeStamp - Run - Writting Ownerless Channel to SharePoint List: $SharePointListName"
            try {
                $SPOTeamsInactiveListItem = New-MgSiteListItem -SiteId $SPOSubId -ListId $SPOTeamsInactiveListId -BodyParameter $param -ErrorAction Stop
                $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
                Write-Output "$TimeStamp - Run - Ownerless Channel written to SharePoint List successfully"
            }
            catch {
                $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
                Write-Output "$TimeStamp - Run - Error writing Ownerless Channel to SharePoint List: $_.Exception.Message"
            }

        }
    }
}
Write-Output "$TimeStamp - Run - Total Ownerless Channels Found: $($OwnerlessChannelsInTeams.Count)"
ForEach ($OwnerlessChannel in $OwnerlessChannelsInTeams) {
    $TimeStamp = ([datetime]::now).tostring("yyyy-MM-dd HH:mm:ss")
    Write-Output "$TimeStamp - Run - Ownerless Channel Found: Team: $($OwnerlessChannel.TeamName), Channel: $($OwnerlessChannel.ChannelName)"
}
# Get Channels of each Teams
# Get Members of Each Teams
# if Owner is missing, put Teams ID and Channel ID in array
# optional: write information in SPO List
## Teams Name
## Teams ID
## Teams Owner
## Channel Name
## Channel Type

