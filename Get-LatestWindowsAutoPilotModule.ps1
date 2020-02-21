#Requires -Version 3.0
Clear-Host
<# Get Latest Version of the Windows AutoPilot Module
    
    Author: The Watcher
    Contact: thewatcher@blogabout.cloud
    Published: 2020

    .DESCRIPTION
    Tool to assist with removal of legacy Windows AutoPilot Module

    Version Changes            
    : 0.1 Initial Script Build
    : 1.0 Initial Build Release

    .LINK
     

    .EXAMPLE
    .\get-latestwindowsautopilotmodule

    Description
    -----------
    Runs script with default values.


    .INPUTS
    None. You cannot pipe objects to this script.

#>
#region Banner
[string] $Root = @'
 
  ┌─────────────────────────────────────────────────────────────┐
               Latest Windows AuotPilot Module Check              
           
              Follow me @thewatchernode on Twitter               
  └─────────────────────────────────────────────────────────────┘
This tool is designed to detect legacy version of the Windows
AutoPilot installed on any host machine and replace with the
latest from the PowerShell Galley

'@
#endregion Banner
#region Shortnames

$Cyan = 'Cyan'
$DarkCyan = 'DarkCyan'
$DarkGray = 'DarkGray'
$DarkRed = 'DarkRed'
$Green = 'Green'
$Red = 'Red'
$Yellow = 'Yellow'
$White = 'White'

$MWCC = 'modernworkplaceclientcenter'
$APMOD = 'WindowsAutoPilotIntune'
$Quit = 'Q'

#endregion Shortnames
#region Functions
Function Test-IsAdmin {
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
Function Get-APPSVersion {
    # Windows AutoPilot PowerShell Version
  
      $ModuleVersion = Get-InstalledModule -Name $APMOD | Select-Object -Property name,version
      Write-Host 'Your client machine is running the following version of Windows AutoPilot Module' -ForegroundColor $White -BackgroundColor $DarkCyan
      $moduleversion
}
Function Get-WinAutoPilotModule {
    
    $ModuleCheck = Get-InstalledModule -name $APMOD -ErrorAction SilentlyContinue   

  if ($ModuleCheck) {
    Write-Host 'Info: Detected an installation of the Windows AutoPilot Module' -ForegroundColor $Green
    $Module = Get-Module -Name $APMOD -ListAvailable
    # Identify modules with multiple versions installed
    $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
    # Check Module from PSGallery
    Write-Host 'Checking Windows AutoPilot module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
    $gallery = $module | Where-Object {$_.repositorysourcelocation}

    Write-Host 'Comparing installed version against online version of Windows AutoPilot module' -ForegroundColor $White -BackgroundColor $DarkCyan
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Warning: Legacy Version of Windows AutoPilot Module detected. Starting removing process"
        Uninstall-Module -Name $APMOD -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of Windows AutoPilot Module now removed"
        Install-Module -Name $APMOD -RequiredVersion $online.Version -Force
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
    # Windows AutoPilot PowerShell Check
    Get-APPSVersion
  
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the Windows AutoPilot Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the Windows AutoPilot Module' -ForegroundColor $Green
    Install-Module -Name $APMOD
   
    # Windows AutoPilot PowerShell Check
    Get-APPSVersion
  }
}
#endregion Functions

Test-IsAdmin
Write-host 'Version information - You are running script version 1.0' -ForegroundColor $White -BackgroundColor $DarkGray
$Root
Get-WinAutoPilotModule