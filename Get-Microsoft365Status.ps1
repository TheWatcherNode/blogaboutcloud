﻿Clear-Host
 <#Information
 
    Author: thewatchernode
    Contact: author@blogabout.cloud
    Published: 30th Ocotber 2020

    .DESCRIPTION
    Tool to gather Microsoft 365 Health Status.

    # IMPORTANT 
    Office 365 Service Communications API needs to be configured with your Tenant. http://www.blogabout.cloud/2020/10/1884/

    Version Changes            
    
    : 0.1 Initial Script Build
    : 1.0 Initial Build Release
     
    Credit:
     

    .EXAMPLE
    .\Get-Microsoft365Status.ps1

    Description
    -----------
    Runs script with default values.


    .INPUTS
    None. You cannot pipe objects to this script.
#>
 #region To be configured by the script runner

# Objects
$tenantId = ''
$client_id = ''
$client_secret = ''
$path = "$env:USERPROFILE\desktop\M365Status_Log.txt"

#endregion

 #region Shortnames
 $Red = 'Red'
 $Green = 'Green'
 $DarkRed = 'DarkRed'
 $White = 'White'
 $DarkCyan = 'DarkCyan'
 $DarkGray = 'DarkGray'
 #endregion

Function Get-M365Status {
# Construct URI for OAuth Token
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Construct Body for OAuth Token
$body = @{
    client_id     = $client_id
    scope         = "https://manage.office.com/.default"
    client_secret = $client_secret
    grant_type    = "client_credentials"
}

# Get OAuth 2.0 Token
$tokenRequest = try {

    Invoke-RestMethod -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop

}
catch [System.Net.WebException] {

    Write-Warning "Exception was caught: $($_.Exception.Message)"
    
}

$token = $tokenRequest.access_token

# Get Office 365 Status
$m365status = try {

    Invoke-RestMethod -Method Get -Uri "https://manage.office.com/api/v1.0/$tenantid/ServiceComms/CurrentStatus" -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop

}
catch [System.Net.WebException] {

    Write-Warning "Exception was caught: $($_.Exception.Message)" 
    
} 

# List service overview status
$m365status.Value | Format-Table WorkloadDisplayName, StatusDisplayName, Status, IncidentIds
}
Write-host 'Version information - You are running script version 1.0' -ForegroundColor $White -BackgroundColor $DarkGray
  @'
  ┌─────────────────────────────────────────────────────────────┐
           Gather the status of Microsoft 365 Service Health

               Follow @thewatchernode on Twitter                               
  └─────────────────────────────────────────────────────────────┘
'@
Start-Transcript -Path $path
Get-M365Status
Stop-Transcript


