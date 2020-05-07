Clear-Host
<#Information
 
    Author: thewatchernode
    Contact: author@blogabout.cloud
    Published: 3rd January 2018

    .DESCRIPTION
    Tool to assist with removal of legacy Microsoft Team Module

    Version Changes            
    : 0.1 Initial Script Build
    : 1.0 Initial Build Release
    : 1.1 Minor Modification
    : 1.2 Minor Modification
    : 1.3 Tested Microsoft Teams Module 1.0.0
    : 1.4 Improved Functions for Get-TeamsPSVersion and Get-TeamsClientCheck
    : 1.5 Included CloudConnector Module
    : 1.6 Included ATP Module
    : 1.7 Included ImportExcel Module
    : 1.8 Microsoft Graph Module  

    .LINK
     

    .EXAMPLE
    .\

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
$importexcel = 'ImportExcel'
$graph = 'Microsoft.Graph.Intune'


 

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
# Microsoft Teams
Function Get-TeamsClientCheck {
  # Microsoft Teams Client Check
  
  $teamclientcheck = "$env:UserProfile\AppData\Roaming\Microsoft\Teams\settings.json"
  if (Test-path -Path $teamclientcheck)
  {
    Write-Host 'Your client machine is running the following version, ring and environment of Microsoft Teams Client' -ForegroundColor $White -BackgroundColor $DarkCyan
    Get-Content -Path $env:UserProfile"\AppData\Roaming\Microsoft\Teams\settings.json" | ConvertFrom-Json | Select-Object -Property Version, Ring, Environment
  }
  else
  {
    Write-Host 'Error: Unable to Microsoft Teams Client installation for versioning information' -ForegroundColor $White -BackgroundColor $DarkRed
  }
}
Function Get-TeamsPSVersion {
  # Microsoft Teams PowerShell Version
  
  $ModuleVersion = Get-InstalledModule -Name $MicrosoftTeams | Select-Object -Property name,version
  Write-Host 'Your client machine is running the following version of Microsoft Teams Module' -ForegroundColor $White -BackgroundColor $DarkCyan
  $moduleversion
}
Function Get-MicrosoftTeams {
  $ModuleCheck = Get-InstalledModule -name MicrosoftTeams -ErrorAction SilentlyContinue   

  if ($ModuleCheck) {
    Write-Host 'Info: Detected an installation of the Microsoft Teams Module' -ForegroundColor $Green
    $Module = Get-Module -Name $MicrosoftTeams -ListAvailable
    # Identify modules with multiple versions installed
    $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
    # Check Module from PSGallery
    Write-Host 'Checking Microsoft Teams module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
    $gallery = $module | Where-Object {$_.repositorysourcelocation}

    Write-Host 'Comparing installed version against online version of Microsoft Teams module' -ForegroundColor $White -BackgroundColor $DarkCyan
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Warning: Legacy Version of Microsoft Teams Module detected. Starting removing process"
        Uninstall-Module -Name $MicrosoftTeams -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of Microsoft Teams Module now removed"
        Install-Module -Name $MicrosoftTeams -RequiredVersion $online.Version -Force
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
    # Microsoft Teams PowerShell Version
    Get-TeamsPSVersion
   
    # Microsoft Teams Client Check
    Get-TeamsClientCheck
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the Microsoft Teams Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the Microsoft Teams Module' -ForegroundColor $Green
    Install-Module -Name $MicrosoftTeams
   
    # Microsoft Teamd PowerShell Check
    Get-TeamsPSVersion
    # Microsoft Teams Client Check
    Get-TeamsClientCheck
  }
}
# SharePoint
Function Get-SPOPSVersion {
  # SharePoint PowerShell Version
  
  $ModuleVersion = Get-InstalledModule -Name $SharePointOnline | Select-Object -Property name,version
  Write-Host 'Your client machine is running the following version of SharePoint Online Module' -ForegroundColor $White -BackgroundColor $DarkCyan
  $moduleversion
}
Function Get-SharePointOnline {
  $ModuleCheck = Get-InstalledModule -name $SharePointOnline -ErrorAction SilentlyContinue   

  if ($ModuleCheck) {
    Write-Host 'Info: Detected an installation of the SharePoint Module' -ForegroundColor $Green
    $Module = Get-Module -Name $SharePointOnline -ListAvailable
    # Identify modules with multiple versions installed
    $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
    # Check Module from PSGallery
    Write-Host 'Checking SharePoint Online module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
    $gallery = $module | Where-Object {$_.repositorysourcelocation}

    Write-Host 'Comparing installed version against online version of SharePoint Online module' -ForegroundColor $White -BackgroundColor $DarkCyan
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Warning: Legacy Version of SharePoint Online Module detected. Starting removing process"
        Uninstall-Module -Name $SharePointOnline -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of SharePoint Online Module now removed"
        Install-Module -Name $SharePointOnline -RequiredVersion $online.Version -Force
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
    # Microsoft Teams PowerShell Version
    Get-SPOPSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the SharePoint Online Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the SharePoint Online Module' -ForegroundColor $Green
    Install-Module -Name $SharePointOnline
   
    # Microsoft Teamd PowerShell Check
    Get-SPOPSVersion

  }
}
# MSOnline
Function Get-MSOLPSVersion {
  # MSOL PowerShell Version
  
  $ModuleVersion = Get-InstalledModule -Name $MSOnline | Select-Object -Property name,version
  Write-Host 'Your client machine is running the following version of MSOnline Module' -ForegroundColor $White -BackgroundColor $DarkCyan
  $moduleversion
}
Function Get-MSOnline {
  $ModuleCheck = Get-InstalledModule -name $MSOnline -ErrorAction SilentlyContinue   

  if ($ModuleCheck) {
    Write-Host 'Info: Detected an installation of the MSOnline Module' -ForegroundColor $Green
    $Module = Get-Module -Name $MSOnline -ListAvailable
    # Identify modules with multiple versions installed
    $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
    # Check Module from PSGallery
    Write-Host 'Checking SharePoint Online module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
    $gallery = $module | Where-Object {$_.repositorysourcelocation}

    Write-Host 'Comparing installed version against online version of MSOnline module' -ForegroundColor $White -BackgroundColor $DarkCyan
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Warning: Legacy Version of MSOnline Module detected. Starting removing process"
        Uninstall-Module -Name $MSOnline -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of MSOnline Module now removed"
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
    # Microsoft MSOL Version
    Get-MSOLPSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the MSOnline Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the MSOnline Module' -ForegroundColor $Green
    Install-Module -Name $MSOnline
   
    # Microsoft Microsoft MSOL Version
    Get-MSOLPSVersion

  }
}
# Azure
Function Get-AzurePSVersion {
  # SharePoint PowerShell Version
  
  $ModuleVersion = Get-InstalledModule -Name $Azure | Select-Object -Property name,version
  Write-Host 'Your client machine is running the following version of Azure Module' -ForegroundColor $White -BackgroundColor $DarkCyan
  $moduleversion
}
Function Get-Azure {
  $ModuleCheck = Get-InstalledModule -name $Azure -ErrorAction SilentlyContinue   

  if ($ModuleCheck) {
    Write-Host 'Info: Detected an installation of the Azure Module' -ForegroundColor $Green
    $Module = Get-Module -Name $Azure -ListAvailable
    # Identify modules with multiple versions installed
    $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
    # Check Module from PSGallery
    Write-Host 'Checking Azure module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
    $gallery = $module | Where-Object {$_.repositorysourcelocation}

    Write-Host 'Comparing installed version against online version of Azure module' -ForegroundColor $White -BackgroundColor $DarkCyan
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Warning: Legacy Version of Azure Module detected. Starting removing process"
        Uninstall-Module -Name $Azure -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of Azure Module now removed"
        Install-Module -Name $Azure -RequiredVersion $online.Version -Force
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
    # Microsoft Teams PowerShell Version
    Get-AzurePSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the Azure Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the Azure Module' -ForegroundColor $Green
    Install-Module -Name $Azure -allowclobber
   
    # Microsoft Teamd PowerShell Check
    Get-AzurePSVersion

  }
}
Function Get-AzureADPSVersion {
  # SharePoint PowerShell Version
  
  $ModuleVersion = Get-InstalledModule -Name $AzureAD | Select-Object -Property name,version
  Write-Host 'Your client machine is running the following version of AzureAD Module' -ForegroundColor $White -BackgroundColor $DarkCyan
  $moduleversion
}
Function Get-AzureAD {
  $ModuleCheck = Get-InstalledModule -name $AzureAD -ErrorAction SilentlyContinue   

  if ($ModuleCheck) {
    Write-Host 'Info: Detected an installation of the AzureAD Module' -ForegroundColor $Green
    $Module = Get-Module -Name $AzureAD -ListAvailable
    # Identify modules with multiple versions installed
    $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
    # Check Module from PSGallery
    Write-Host 'Checking AzureAD module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
    $gallery = $module | Where-Object {$_.repositorysourcelocation}

    Write-Host 'Comparing installed version against online version of AzureAD module' -ForegroundColor $White -BackgroundColor $DarkCyan
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Warning: Legacy Version of AzureAD Module detected. Starting removing process"
        Uninstall-Module -Name $AzureAD -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of AzureAD Module now removed"
        Install-Module -Name $AzureAD -RequiredVersion $online.Version -Force
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
    # Microsoft Teams PowerShell Version
    Get-SPOPSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the AzureAD Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the AzureAD Module' -ForegroundColor $Green
    Install-Module -Name $AzureAD
   
    # Microsoft Teamd PowerShell Check
    Get-AzureADPSVersion

  }
 }
# Skype for Business
Function Get-SfBOModule {
  Start-BitsTransfer -
}
# Exchange
Function Get-CCPSVersion {
  # MSOL PowerShell Version
  
  $ModuleVersion = Get-InstalledModule -Name $CloudConnector | Select-Object -Property name,version
  Write-Host 'Your client machine is running the following version of CloudConnector Module' -ForegroundColor $White -BackgroundColor $DarkCyan
  $moduleversion
}
Function Get-CloudConnector {
  $ModuleCheck = Get-InstalledModule -name $CloudConnector -ErrorAction SilentlyContinue   

  if ($ModuleCheck) {
    Write-Host 'Info: Detected an installation of the CloudConnector Module' -ForegroundColor $Green
    $Module = Get-Module -Name $CloudConnector -ListAvailable
    # Identify modules with multiple versions installed
    $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
    # Check Module from PSGallery
    Write-Host 'Checking CloudConnector module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
    $gallery = $module | Where-Object {$_.repositorysourcelocation}

    Write-Host 'Comparing installed version against online version of CloudConnector module' -ForegroundColor $White -BackgroundColor $DarkCyan
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Warning: Legacy Version of CloudConnect Module detected. Starting removing process"
        Uninstall-Module -Name $MSOnline -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of CloudConnect Module now removed"
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
    # Microsoft MSOL Version
    Get-CCPSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the CloudConnect Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the CloudConnect Module' -ForegroundColor $Green
    Install-Module -Name $CloudConnector
   
    # Microsoft Microsoft MSOL Version
    Get-CCPSVersion

  }
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
        Uninstall-Module -Name $ztp -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of Office ATP Recommended Configuration Analyzer Module now removed"
        Install-Module -Name $atp-RequiredVersion $online.Version -Force
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
# ImportExcel
Function Get-ImportExcelPSVersion {
  # MSOL PowerShell Version
  
  $ModuleVersion = Get-InstalledModule -Name $ImportExcel | Select-Object -Property name,version
  Write-Host 'Your client machine is running the following version of ImportExcel Module' -ForegroundColor $White -BackgroundColor $DarkCyan
  $moduleversion
}
Function Get-ImportExcel {
  $ModuleCheck = Get-InstalledModule -name $ImportExcel -ErrorAction SilentlyContinue   

  if ($ModuleCheck) {
    Write-Host 'Info: Detected an installation of the ImportExcel Module' -ForegroundColor $Green
    $Module = Get-Module -Name $ImportExcel -ListAvailable
    # Identify modules with multiple versions installed
    $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
    # Check Module from PSGallery
    Write-Host 'Checking ImportExcel module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
    $gallery = $module | Where-Object {$_.repositorysourcelocation}

    Write-Host 'Comparing installed version against online version of ImportExcel module' -ForegroundColor $White -BackgroundColor $DarkCyan
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Warning: Legacy Version of ImportExcel Module detected. Starting removing process"
        Uninstall-Module -Name $ImportExcel -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of ImportExcel Analyzer Module now removed"
        Install-Module -Name $ImportExcel -RequiredVersion $online.Version -Force
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
    # ImportExcel
    Get-ImportExcelPSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the ImportExcel Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the ImportExcel Module' -ForegroundColor $Green
    Install-Module -Name $ImportExcel
   
    # ImportExcel
    Get-ImportExcelPSVersion

  }
}
 
 #endregion
#region Script Block
 cls
 Write-host 'Version information - You are running script version 1.7' -ForegroundColor $White -BackgroundColor DarkGray
 Test-IsAdmin
 @'
  ┌─────────────────────────────────────────────────────────────┐
            Common PowerShell modules of the IT Pro

               Follow @thewatchernode on Twitter                               
  └─────────────────────────────────────────────────────────────┘
'@
 Start-Transcript -Path $env:userprofile\Desktop\CommonModulesLog.txt
 Get-MicrosoftTeams
 Get-SharePointOnline
 Get-MSOnline
 Get-Azure
 Get-CloudConnector
 Get-ATP
 Get-ImportExcel
 Stop-Transcript
 #endregion
