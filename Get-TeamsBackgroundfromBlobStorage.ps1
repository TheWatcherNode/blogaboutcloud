<#Information
 
    Author: thewatchernode
    Contact: author@blogabout.cloud
    Published: 5th March 2020

    .DESCRIPTION
    Tool to assist with application delivery

    Version Changes            
    : 0.1 Initial Script Build
    : 1.0 Initial Build Release
 
    .EXAMPLE
    .\get-teamsbackgroundfromblobstorage.ps1

    Description
    -----------
    Runs script with default values.


    .INPUTS
    None. You cannot pipe objects to this script.
#>

# The following strings are set to configure the downloads required to complete the deployment. #
param(
  [String] $InstallDir = "$env:HOMEDRIVE\_TeamsImages",
  [string] $TeamsBackground1 = 'https://blogaboutcloud.blob.core.windows.net/intuneblogaboutcloud/Images/BullRing-Desktop.jpg',
  [string] $TeamsBackground2 = 'https://blogaboutcloud.blob.core.windows.net/intuneblogaboutcloud/Images/BullRing.jpg')

  Write-host 'Version information - You are running script version 1.0' -ForegroundColor White -BackgroundColor DarkGray
  @'
  ┌─────────────────────────────────────────────────────────────┐
               Custom Corporate Teams Background

               Follow @thewatchernode on Twitter                               
  └─────────────────────────────────────────────────────────────┘
'@

#region Download Application Installers
# Download the files required. #
New-Item -ItemType Directory -Path "$env:HOMEDRIVE\_TeamsImages"
$wc = New-Object Net.webclient
$wc.DownloadFile($TeamsBackground1, ('{0}\BullRing-Desktop.jpg' -f $InstallDir))
$wc.DownloadFile($TeamsBackground2, ('{0}\BullRing.jpg' -f $InstallDir))
Start-Sleep 30
#endregion

#region Move Images to User Profile

#Robocopy
$source = $installdir
$destination = Join-Path $env:Appdata 'Microsoft\Teams\Backgrounds\Uploads'

C:\Windows\System32\Robocopy.exe $source $destination /mov /e /fp /njh /njs /x /eta /r:0 /copy:dat

#endregion

#region Cleanup
#Delete TeamImages Folder. #

Remove-Item "$env:HOMEDRIVE\_TeamsImages" -Recurse -Force

#endregion

