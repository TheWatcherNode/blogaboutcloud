Clear-Host
<#Information
 
    Author: thewatchernode
    Contact: author@blogabout.cloud
    Published: 12th November 2019

    .DESCRIPTION
    Obtain details about your ATP Configuration

    Version Changes            
    : 0.1 Initial Script Build
    : 1.0 Initial Build Release

    .LINK
     

    .EXAMPLE
    .\get-atpreport

    Description
    -----------
    Runs script with default values.


    .INPUTS
    None. You cannot pipe objects to this script.
#>
#Requires -Version 3.0
#region Shortnames

$DarkCyan = 'DarkCyan'
$DarkRed = 'DarkRed'
$DarkGray = 'DarkGray'
$Green = 'Green'
$Red = 'Red'
$Yellow = 'Yellow'
$White = 'White'

$Azure = 'Az'
$AzureAD = 'AzureAD'
$MicrosoftTeams = 'MicrosoftTeams'
$MSOnline = 'MSOnline'
$SharePointOnline = 'Microsoft.Online.SharePoint.PowerShell'
$CloudConnector = 'CloudConnect'
$ATP = 'ORCA'

#endregion Shortnames

#region Functions
function Test-IsAdmin {
  <#
      .SYNOPSIS
      Describe purpose of "Test-IsAdmin" in 1-2 sentences.

      .DESCRIPTION
      Add a more complete description of what the function does.

      .EXAMPLE
      Test-IsAdmin
      Describe what this call does

      .NOTES
      Place additional notes here.

      .LINK
      URLs to related sites
      The first link is opened by Get-Help -Online Test-IsAdmin

      .INPUTS
      List of input types that are accepted by this function.

      .OUTPUTS
      List of output types produced by this function.
  #>

  ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')

}
if (!(Test-IsAdmin)){
  throw 'Please Note: You are trying to run this script without administative priviliges. In order to run this script you will required PowerShell running in Administrator Mode'
}
else {
  Write-Verbose -Message 'Are you running as an Administator' -verbose
}

# Connect to Exchange
Function Connect-ExOnline {

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

}

# ATP
Function Get-ATPPSVersion {
  # MSOL PowerShell Version
  
  $ModuleVersion = Get-InstalledModule -Name $ATP | Select-Object -Property name,version
  Write-Host 'Your client machine is running the following version of Office ATP Recommended Configuration Analyzer Module' -ForegroundColor $White -BackgroundColor $DarkCyan
  $moduleversion
}
Function Get-ATP {
  $ModuleCheck = Get-InstalledModule -name $ATP -ErrorAction SilentlyContinue   

  if ($ModuleCheck) {
    Write-Host 'Info: Detected an installation of the Office ATP Recommended Configuration Analyzer Module' -ForegroundColor $Green
    $Module = Get-Module -Name $ATP -ListAvailable
    # Identify modules with multiple versions installed
    $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
    # Check Module from PSGallery
    Write-Host 'Checking Office ATP Recommended Configuration Analyzer module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
    $gallery = $module | Where-Object {$_.repositorysourcelocation}

    Write-Host 'Comparing installed version against online version of Office ATP Recommended Configuration Analyzer module' -ForegroundColor $White -BackgroundColor $DarkCyan
    foreach ($module in $gallery) {

      #find the current version in the gallery
      Try {
        $online = Find-Module -Name $module.name -Repository PSGallery -ErrorAction Stop
      }
      Catch {
        Write-Warning -Message ('Module {0} was not found in the PSGallery' -f $module.name)
      }

      #compare versions
      if ($online.version -gt $module.version) {
        $UpdateAvailable = 'Version removed'
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Warning: Legacy Version of Office ATP Recommended Configuration Analyzer Module detected. Starting removing process"
        Uninstall-Module -Name $MSOnline -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of Office ATP Recommended Configuration Analyzer Module now removed"
        Install-Module -Name $MSOnline -RequiredVersion $online.Version -Force
      }
      else {
        $UpdateAvailable = 'No update required'
      }

      #write a custom object to the pipeline
      [pscustomobject]@{
        Name = $module.name
        MultipleVersions = ($g.name -contains $module.name)
        InstalledVersion = $module.version
        OnlineVersion = $online.version
        Update = $UpdateAvailable
        Path = $module.modulebase
      }
 
    } 
    # Microsoft Office ATP Recommended Configuration Analyzer
    Get-ATPPSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the Office ATP Recommended Configuration Analyzer Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the Office ATP Recommended Configuration Analyzer Module' -ForegroundColor $Green
    Install-Module -Name $ATP
   
    # Microsoft Office ATP Recommended Configuration Analyzer
    Get-ATPPSVersion

  }
}

 #endregion
#region Script Block
 cls
 Write-host 'Version information - You are running script version 1.0' -ForegroundColor $White -BackgroundColor $DarkGray
 Test-IsAdmin
 @'
  ┌─────────────────────────────────────────────────────────────┐
     Office Advanced Threat Recommended Configuration Analyzer

               Follow @thewatchernode on Twitter                               
  └─────────────────────────────────────────────────────────────┘
'@
 Start-Transcript -Path $env:userprofile\Desktop\ATPLog.txt
 Connect-ExOnline
 Get-ATP
 Get-ORCAReport
 Stop-Transcript
 #endregion
