<#Information
 
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
[CmdletBinding()]
param(
  [String] $InstallDir = "$env:HOMEDRIVE\_build",
  [String] $BoxDrive = 'https://blogaboutcloud.blob.core.windows.net/intuneblogaboutcloud/Tools/Box-x64.msi',
  [String] $wincleaner = 'https://blogaboutcloud.blob.core.windows.net/intuneblogaboutcloud/Tools/Windows_Clean.ps1')

  Write-host 'Version information - You are running script version 1.2' -ForegroundColor White -BackgroundColor DarkGray
  @'
  ┌─────────────────────────────────────────────────────────────┐
           Installing Applications from Blob Storage

               Follow @thewatchernode on Twitter                               
  └─────────────────────────────────────────────────────────────┘
'@

#region Download Application Installers


# Download the files required. #
  New-Item -ItemType Directory -Path "$env:HOMEDRIVE\_build"
  $wc = New-Object -TypeName Net.webclient
  $wc.DownloadFile($wincleaner, ('{0}\Windows_Clean.ps1' -f $InstallDir))
  $wc.DownloadFile($BoxDrive, ('{0}\Box-x64.msi' -f $InstallDir))
  Start-Sleep -Seconds 15
  
#endregion

#region Install Applications
Function Get-InstallApps {
# Install Box Drive. 
  Write-Host 'Installing BoxDrive' -ForegroundColor Green
  Start-Process -FilePath "$env:homedrive\_build\Box-x64.msi" -ArgumentList '/quiet' -Wait
  }
#endregion

#region Additional Scripts
Function Get-AddScripts {
# Script 1 
 & C:\_build\Windows_Clean.ps1 -ClearStart
   Start-Sleep 20
   }

#endregion

#region Cleanup

Function Get-CleanUp {
# Delete Build Folder
  Remove-Item -Path "$env:HOMEDRIVE\_build" -Recurse -Force

# Rename and Disable Local Admin Account
  #Write-Host "Rename local Admin Account to #NotInUse#" -ForegroundColor Green
  #Rename-LocalUser -Name "Administrator" -NewName "#NotInUse#" -ErrorAction SilentlyContinue
  #Get-LocalUser "#NotInUse#" | Disable-LocalUser -ErrorAction SilentlyContinue
  }

#endregion

#region Reg Modifications (Example)

Function Get-RegMod {
# Set Reg Key
  #Set-ItemProperty -Path "HKLM:\Software\Wow6432Node\Javasoft\Java Update\Policy" -Name "EnableJavaUpdate" -Value 0
  }
#endregion

Start-Transcript $env:userprofile\desktop\log.txt

Get-InstallApps
Get-AddScripts
Get-CleanUp
Get-RegMod
Stop-Transcript