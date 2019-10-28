<#Information
 
    Author: thewatchernode
    Contact: author@blogabout.cloud
    Published: 3rd January 2018

    .DESCRIPTION
    Tool to assist with application delivery

    Version Changes            
    : 0.1 Initial Script Build
    : 1.0 Initial Build Release


     
    .EXAMPLE
    .\get-appsfromblobstorage.ps1

    Description
    -----------
    Runs script with default values.


    .INPUTS
    None. You cannot pipe objects to this script.
#>

# The following strings are set to configure the downloads required to complete the deployment. #
param(
  [String] $InstallDir = "c:\_build",
  [string] $chromeinstaller = "https://intunestorage2019.blob.core.windows.net/intunestorage/Intune/GoogleChromeStandaloneEnterprise64.msi")

  Write-host 'Version information - You are running script version 1.0' -ForegroundColor White -BackgroundColor DarkGray
  @'
  ┌─────────────────────────────────────────────────────────────┐
           Installing Applications from Blob Storage

               Follow @thewatchernode on Twitter                               
  └─────────────────────────────────────────────────────────────┘
'@

#region Download Application Installers
# Download the files required. #
New-Item -ItemType Directory -Path "c:\_build"
$wc = New-Object Net.webclient
$wc.DownloadFile($chromeinstaller, "$InstallDir\chrome.msi")
$wc.DownloadFile($CleanWindows, "$InstallDir\Windows_Clean.ps1")
Start-Sleep 360
#endregion


#region Install Applications

#Install Chrome. #
Start-Process "c:\_build\chrome.msi" -ArgumentList "/quiet /norestart" -wait
#endregion

#region Additional Scripts
#Run Windows DeClutter. #
#& C:\_build\Windows_Clean.ps1 -ClearStart
#Start-Sleep 90
#endregion

#region Cleanup
#Delete Build Folder. #
Remove-Item "c:\_build" -Recurse -Force
#endregion

#region Reg Modifications (Example)
# Set Reg Key
#Set-ItemProperty -Path "HKLM:\Software\Wow6432Node\Javasoft\Java Update\Policy" -Name "EnableJavaUpdate" -Value 0
#endregion