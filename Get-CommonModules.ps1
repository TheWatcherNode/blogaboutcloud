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

#region Elevate Script
# Original Script located at:
# http://blogs.msdn.com/b/virtual_pc_guy/archive/2010/09/23/a-self-elevating-powershell-script.aspx

# Get the ID and security principal of the current user account
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)

# Get the security principal for the Administrator role
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator

# Check to see if we are currently running "as Administrator"
if ($myWindowsPrincipal.IsInRole($adminRole))

   {
   # We are running "as Administrator" - so change the title and background color to indicate this
   $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
   $Host.UI.RawUI.BackgroundColor = "DarkBlue"
   clear-host

   }
else
   {
   # We are not running "as Administrator" - so relaunch as administrator

   # Create a new process object that starts PowerShell
   $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";

   # Specify the current script path and name as a parameter
   $newProcess.Arguments = $myInvocation.MyCommand.Definition;

   # Indicate that the process should be elevated
   $newProcess.Verb = "runas";

   # Start the new process
   [System.Diagnostics.Process]::Start($newProcess);

   # Exit from the current, unelevated, process
   exit

   }
#endregion
#region Shortnames

$DarkCyan = 'DarkCyan'
$DarkRed = 'DarkRed'
$Green = 'Green'
$Red = 'Red'
$Yellow = 'Yellow'
$White = 'White'

$ATP = 'ORCA'
$Azure = 'Az'
$AzureAD = 'AzureAD'
$CloudConnector = 'CloudConnect'
$EXO = 'ExchangeOnlineManagement'
$graph = 'Microsoft.Graph.Intune'
$importexcel = 'ImportExcel'
$MicrosoftTeams = 'MicrosoftTeams'
$MSOnline = 'MSOnline'
$SharePointOnline = 'Microsoft.Online.SharePoint.PowerShell'

#endregion Shortnames

#region Functions
#region Test-IsAdmin
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
#endregion
#region Elevate Script
     # Self-elevate the script if required
    if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
     if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
      $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
      Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
      Exit
     }
    }
#endregion
#region PSGallery
Function Get-TrustedRepo {

Install-PackageProvider -name NuGet -Force
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

}
#endregion

#region ATP
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
        Uninstall-Module -Name $ATP -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of Office ATP Recommended Configuration Analyzer Module now removed"
        Install-Module -Name $atp -RequiredVersion $online.Version -Force
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
#region Azure
Function Get-AzurePSVersion {
  # Azure PowerShell Version
  
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
    # Microsoft Azure PowerShell Version
    Get-AzurePSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the Azure Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the Azure Module' -ForegroundColor $Green
    Install-Module -Name $Azure -allowclobber
   
    # Microsoft Azure PowerShell Check
    Get-AzurePSVersion

  }
}
Function Get-AzureADPSVersion {
  # Azure PowerShell Version
  
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
    # Microsoft Azure PowerShell Version
    Get-AzureADPSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the AzureAD Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the AzureAD Module' -ForegroundColor $Green
    Install-Module -Name $AzureAD
   
    # Microsoft Azure PowerShell Check
    Get-AzureADPSVersion

  }
 }
#endregion
#region Exchange
Function Get-EXOPSVersion {
  # Exchange Online PowerShell Version
  
  $ModuleVersion = Get-InstalledModule -Name $EXO | Select-Object -Property name,version
  Write-Host 'Your client machine is running the following version of Exchange Online Module' -ForegroundColor $White -BackgroundColor $DarkCyan
  $moduleversion
}
Function Get-ExoMod {
  $ModuleCheck = Get-InstalledModule -name $EXO -ErrorAction SilentlyContinue   

  if ($ModuleCheck) {
    Write-Host 'Info: Detected an installation of the Exchange Online Module' -ForegroundColor $Green
    $Module = Get-Module -Name $EXO -ListAvailable
    # Identify modules with multiple versions installed
    $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
    # Check Module from PSGallery
    Write-Host 'Checking Exchange Online module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
    $gallery = $module | Where-Object {$_.repositorysourcelocation}

    Write-Host 'Comparing installed version against online version of Exchange Online module' -ForegroundColor $White -BackgroundColor $DarkCyan
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Warning: Legacy Version of Exchange Online Module detected. Starting removing process"
        Uninstall-Module -Name $MSOnline -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of Exchange Online Module now removed"
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
    Get-EXOPSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the Exchange Online Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the Exchange Online Module' -ForegroundColor $Green
    Install-Module -Name $EXO
   
    # Exchange Version
    Get-EXOPSVersion

  }
}
#endregion
#region Graph
Function Get-GraphPSVersion {
  # MSOL PowerShell Version
  
  $ModuleVersion = Get-InstalledModule -Name $graph | Select-Object -Property name,version
  Write-Host 'Your client machine is running the following version of GraphlModule' -ForegroundColor $White -BackgroundColor $DarkCyan
  $moduleversion
}
Function Get-Graph {
  $ModuleCheck = Get-InstalledModule -name $graph -ErrorAction SilentlyContinue   

  if ($ModuleCheck) {
    Write-Host 'Info: Detected an installation of the Graph Module' -ForegroundColor $Green
    $Module = Get-Module -Name $graph -ListAvailable
    # Identify modules with multiple versions installed
    $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
    # Check Module from PSGallery
    Write-Host 'Checking Graph module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
    $gallery = $module | Where-Object {$_.repositorysourcelocation}

    Write-Host 'Comparing installed version against online version of Graph module' -ForegroundColor $White -BackgroundColor $DarkCyan
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Warning: Legacy Version of Graph Module detected. Starting removing process"
        Uninstall-Module -Name $MSOnline -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White "Info: Legacy Version of Graph Module now removed"
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
    # Microsoft Graph
    Get-GraphPSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the Graph Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the Graph Module' -ForegroundColor $Green
    Install-Module -Name $graph
   
    # Microsoft Microsoft Graph
    Get-GraphPSVersion

  }
}


#endregion
#region ImportExcel
Function Get-ImportExcelPSVersion {
  # ImportExcel PowerShell Version
  
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
#region Microsoft Teams
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
#endregion
#region MSOnline
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
#endregion
#region SharePoint
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
    # Microsoft SharePoint PowerShell Version
    Get-SPOPSVersion
   
  }
  else
  {
    Write-Host 'Error: Failed to detect an installation of the SharePoint Online Module' -ForegroundColor $Red
    Write-Host 'Info: Attempting an installation of the SharePoint Online Module' -ForegroundColor $Green
    Install-Module -Name $SharePointOnline
   
    # Microsoft SharePoint PowerShell Check
    Get-SPOPSVersion

  }
}
#endregion

 #region Legacy Modules
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

 #endregion

 #endregion

#region Script Block
 cls
 Write-host 'Version information - You are running script version 1.7' -ForegroundColor $White -BackgroundColor DarkGray
 #Test-IsAdmin
 @'
  ┌─────────────────────────────────────────────────────────────┐
            Common PowerShell modules of the IT Pro

               Follow @thewatchernode on Twitter                               
  └─────────────────────────────────────────────────────────────┘
'@
 Start-Transcript -Path $env:HOMEDRIVE\PowerShellModulesInstallLog.txt
 Get-TrustedRepo

 Get-ATP
 Get-Azure
 #Get-CloudConnector
 Get-ExoMod
 Get-ImportExcel
 Get-MicrosoftTeams
 Get-MSOnline
 Get-SharePointOnline
 
 Stop-Transcript
 #endregion
