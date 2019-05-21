﻿ Clear-Host
 #region Shortnames
 $Red = 'Red'
 $Green = 'Green'
 $DarkRed = 'DarkRed'
 $White = 'White'
 $DarkCyan = 'DarkCyan'
 $DarkGray = 'DarkGray'
 #endregion
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
 #endregion
 Test-IsAdmin
 Write-host 'Version information - You are running script version 1.1' -ForegroundColor $White -BackgroundColor $DarkGray
  @'
  ┌─────────────────────────────────────────────────────────────┐
           Updating your PSGallery PowerShell Modules

               Follow @thewatchernode on Twitter                               
  └─────────────────────────────────────────────────────────────┘
'@

 $Array = @(Get-InstalledModule)

 Foreach ($Module in $Array)
 {
    $ModuleCheck = Get-InstalledModule -name $Module.Name -ErrorAction SilentlyContinue   

   if ($ModuleCheck) {
     Write-Host 'Info: Detected an installation of the',$Module.Name,'Module' -ForegroundColor $Green
     $Module = Get-Module -Name $Module.Name -ListAvailable
     # Identify modules with multiple versions installed
     $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
     # Check Module from PSGallery
     Write-Host 'Checking',$Module.Name,' module from the PSGallery' -ForegroundColor $White -BackgroundColor $DarkCyan
     $gallery = $module | Where-Object {$_.repositorysourcelocation}

     Write-Host 'Comparing installed version against online version of',$Module.Name,'module' -ForegroundColor $White -BackgroundColor $DarkCyan
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
         Write-Host -BackgroundColor $DarkRed -ForegroundColor $White 'Warning: Legacy Version of',$Module.name,'module detected. Starting removing process'
         Uninstall-Module -Name $Module.Name -RequiredVersion $module.version 
         Write-Host -BackgroundColor $DarkRed -ForegroundColor $White 'Info: Legacy Version of',$Module.name,'module now removed'
         Install-Module -Name $Module.Name -RequiredVersion $online.Version -Force
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
     $ModuleVersion = Get-InstalledModule -Name $Module.Name | Select-Object -Property name,version
     Write-Host 'Your client machine is running the following version of',$Module.Name,'Module' -ForegroundColor $White -BackgroundColor $DarkCyan
     $moduleversion
   
   }
   else
   {
     Write-Host 'Error: Failed to detect an installation of the',$Module.name,'Azure Module' -ForegroundColor $Red
   }
 }


