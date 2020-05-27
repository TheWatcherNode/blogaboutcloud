﻿<#Information
 
    Author: thewatchernode
    Contact: author@blogabout.cloud
    Published: 28th October 2019

    .DESCRIPTION
    Tool to assist with application delivery

    Version Changes            
    : 0.1 Initial Script Build
    : 1.0 Initial Build Release
    : 1.1 Updates


     
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
  [string] $wincleaner = "https://blogaboutcloud.blob.core.windows.net/intuneblogaboutcloud/Tools/Windows_Clean.ps1")

  Write-host 'Version information - You are running script version 1.1' -ForegroundColor White -BackgroundColor DarkGray
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
$wc.DownloadFile($wincleaner, "$InstallDir\Windows_Clean.ps1")
Start-Sleep 90
#endregion

#region Install Applications

#Install 

#endregion

#region Additional Scripts
#Run Windows DeClutter. 
& C:\_build\Windows_Clean.ps1 -ClearStart
#endregion

#region Cleanup
#Delete Build Folder. #
Remove-Item "c:\_build" -Recurse -Force

# Rename and Disable Local Admin Account
Write-Host "Rename local Admin Account to #NotInUse#" -ForegroundColor Green
Rename-LocalUser -Name "Administrator" -NewName "#NotInUse#" -ErrorAction SilentlyContinue
Get-LocalUser "#NotInUse#" | Disable-LocalUser -ErrorAction SilentlyContinue

#endregion

#region Reg Modifications (Example)
# Set Reg Key
#Set-ItemProperty -Path "HKLM:\Software\Wow6432Node\Javasoft\Java Update\Policy" -Name "EnableJavaUpdate" -Value 0
#endregion
