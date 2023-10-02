# Requires PowerShell 7.2 and latest PnP.Online
# Script will grant SharePoint Online permissions to a Managed Identity account
#
# created by Thorsten Pickhan
# Initial script created on 07.06.2022 (06/07/2022)
#
# Version 1.0

# PNP PowerShell is required in Version 2.2.0

# Install-Module -Name PnP.Powershell
# Import-Module PNP.Powershell

# Define DisplayName of the Managed Identity - please edit
#
#####
$ManagedIdentityAccount = "AA-TeamsAutomation"

# Connect to SharePoint Online - please edit
#
#####
Connect-PnpOnline -Interactive -Url https://xxxxxx.sharepoint.com/Teams/TeamsAutomation

# Grant required permissions
#
#####
Add-PnPAzureADServicePrincipalAppRole -Principal $ManagedIdentityAccount -AppRole "Sites.FullControl.All" -BuiltInType SharePointOnline