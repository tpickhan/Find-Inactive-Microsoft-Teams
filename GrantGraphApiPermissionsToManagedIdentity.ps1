# Requires PowerShell 7.2 and latest Microsoft Graph modules
# Script will grant Report.Read.All and Reports.Settings.ReadWrite.All permissions to a Managed Identity account
#
# created by Thorsten Pickhan
# Initial script created on 02.10.2023 (10/02/2023)
#
# Version 1.0
#

# Verify if required Microsoft Graph Modules are installed
#
#####
if (!(Get-InstalledModule -Name Microsoft.Graph.Authentication -ErrorAction SilentlyContinue )) {
    Write-Host "Microsoft.Graph.Authentication not installed, installing..." -ForegroundColor Yellow
    Install-Module -Name  Microsoft.Graph.Authentication -Scope CurrentUser
    Write-Host "Import Module Microsoft.Graph.Authentication"
    Import-Module Microsoft.Graph.Authentication
}
else {
    if (!(Get-Module -Name Microsoft.Graph.Authentication)) {
        Write-Host "Import Microsoft.Graph.Authentication"
        Import-Module Microsoft.Graph.Authentication 
    }
    else {
        Write-Host "Microsoft.Graph.Authentication already imported"
    }
}

if (!(Get-InstalledModule -Name Microsoft.Graph.Applications -ErrorAction SilentlyContinue )) {
    Write-Host "Microsoft.Graph.Applications not installed, installing..." -ForegroundColor Yellow
    Install-Module -Name  Microsoft.Graph.Applications -Scope CurrentUser
    Write-Host "Import Module Microsoft.Graph.Applications"
    Import-Module Microsoft.Graph.Applications
}
else {
    if (!(Get-Module -Name Microsoft.Graph.Applications)) {
        Write-Host "Import Microsoft.Graph.Applications"
        Import-Module Microsoft.Graph.Applications 
    }
    else {
        Write-Host "Microsoft.Graph.Applications already imported"
    }
}

# DisplayName of our Managed Identity - please edit
#
#####
$MsiName = "AA-TeamsAutomation" 

# Your Tenant Id - please edit
#
#####
$TenantID="xxxxxxxxxxxx"


# GraphApi Permissions required
#
#####
$oPermissions = @(
  "Reports.Read.All"
  "ReportSettings.ReadWrite.All"
)


# Microsoft Graph API App Id - don't change it
#
#####
$GraphAppId = "00000003-0000-0000-c000-000000000000"

# Connect to Microsof Graph API
#
#####
Connect-MgGraph -Scopes 'Application.Read.All','AppRoleAssignment.ReadWrite.All' -TenantId $TenantID

# Get our Managed Identity Account
#
#####
$ManagedIdentity = Get-MgServicePrincipal -Filter "displayName eq '$MsiName'"

# Get Microsoft Graph Api App Id
#
#####
$oGraphSpn = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"

# Get Graph API permissions
#
#####
$oAppRoles = $oGraphSpn.AppRoles | Where-Object {($_.Value -in $oPermissions) -and ($_.AllowedMemberTypes -contains "Application")}

# Loop through Roles and assign it to your Managed Identity Account
ForEach($AppRole in $oAppRoles)
{
  $oAppRoleAssignment = @{
    "PrincipalId" = $ManagedIdentity.Id
    "ResourceId" = $oGraphSpn.Id
    "AppRoleId" = $AppRole.Id
  }
  
  New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $oAppRoleAssignment.PrincipalId -BodyParameter $oAppRoleAssignment -Verbose
}

