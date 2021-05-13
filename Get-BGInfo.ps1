<#Information
 
    Author: thewatchernode
    Contact: author@blogabout.cloud
    Published: 07th May 2021

    .DESCRIPTION
    Tool to assist with application delivery

    Version Changes            
    : 0.1 Initial Script Build
    : 1.0 Initial Build Release


    .EXAMPLE
    .\Get-BGInfo.ps1

    Description
    -----------
    Runs script with default values.


    .INPUTS
    None. You cannot pipe objects to this script.
#>
[CmdletBinding()]
param(

  [String] $InstallDir = "$env:ProgramW6432\BGInfo",
  [String] $bginfo64 = 'https://XX/Bginfo64.exe',
  [String] $layout = 'https://XX/Layout.bgi')
  Write-host 'Version information - You are running script version 1.0' -ForegroundColor White -BackgroundColor DarkGray
  @'
  ┌─────────────────────────────────────────────────────────────┐
       Install BGInfo for all enrolled Windows 10 devivces

               Follow @thewatchernode on Twitter                               
  └─────────────────────────────────────────────────────────────┘
'@

#region Variables
$bgInfoRegPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
$bgInfoRegKey = 'BgInfo'
$bgInfoRegType = 'String'
$regKeyExists = (Get-Item -Path $bgInfoRegPath -ErrorAction Ignore).Property -contains $bgInfoRegkey
$bgInfoRegKeyValue = ('"{0}\Bginfo64.exe" "{0}\layout.bgi" /timer:0 /nolicprompt' -f $InstallDir)
$writeEmptyLine = "`n"
$writeSeperator = ' - '
$time = Get-Date -UFormat '%A %m/%d/%Y %R'
$Green = 'Green'
#endregion

#region Download Application Installers


# Download the files required. #
  
  $Path = Test-Path -Path $InstallDir

  if ($Path -eq 'True'){
  Write-Host ('INFO: {0} path exists' -f $InstallDir)
  }
  else
 {
  Write-Host ('INFO: Path not detected, creating {0}' -f $InstallDir)
  New-Item -ItemType Directory -Path "$env:ProgramW6432\BGInfo" -ErrorAction SilentlyContinue
}
 
  $wc = New-Object -TypeName Net.webclient
  $wc.DownloadFile($bginfo64, ('{0}\Bginfo64.exe' -f $InstallDir))
  $wc.DownloadFile($layout, ('{0}\Layout.bgi' -f $InstallDir))
  Start-Sleep -Seconds 15
  
#endregion
#region Install Applications
Function Get-InstallApps {
  <#
      .SYNOPSIS
      Describe purpose of "Get-InstallApps" in 1-2 sentences.

      .DESCRIPTION
      Add a more complete description of what the function does.

      .EXAMPLE
      Get-InstallApps
      Describe what this call does

      .NOTES
      Place additional notes here.

      .LINK
      URLs to related sites
      The first link is opened by Get-Help -Online Get-InstallApps

      .INPUTS
      List of input types that are accepted by this function.

      .OUTPUTS
      List of output types produced by this function.
  #>
  ##--------------------------------------------------------------------------
  ## Create BgInfo Registry Key to AutoStart
 
  If ($regKeyExists -eq $True)
  {
  Write-Host ($writeEmptyLine + 'BgInfo regkey exists, script wil go on' + $writeSeperator + $time) -ForegroundColor $Green $writeEmptyLine
  }
  Else
  {
  New-ItemProperty -Path $bgInfoRegPath -Name $bgInfoRegkey -PropertyType $bgInfoRegType -Value $bgInfoRegkeyValue
  Write-Host ($writeEmptyLine + '# BgInfo regkey added' + $writeSeperator + $time) -ForegroundColor $Green $writeEmptyLine
  }
 
  ## --------------------------------------------------------------------------
  
  ## Run BGInfo 
  Write-Host 'Run BGInfo' -ForegroundColor $Green
  Start-Process -FilePath "$env:ProgramW6432\BGInfo\Bginfo64.exe" -WorkingDirectory "$env:ProgramW6432\BGInfo" -ArgumentList 'layout.bgi /timer:0 /silent /nolicprompt'
  
#region Script Actions

Start-Transcript -Path $env:ProgramW6432\BGInfo\InstallationLog.txt
Get-InstallApps
Stop-Transcript
#endregion


