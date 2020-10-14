 Clear-Host
 <#Information
 
    Author: thewatchernode
    Contact: author@blogabout.cloud
    Published: 3rd January 2018

    .DESCRIPTION
    Tool to assist with removal of legacy installed PowerShell Module from PSGallery

    Version Changes            
    
    : 0.1 Initial Script Build
    : 1.0 Initial Build Release
    : 1.1 Minor Modification
    : 1.2 Minor Modification
    : 1.3 Included logging
    : 1.4 Bug fixes 
    : 1.5 Handle modules with dependencies

     
    Credit:
     http://www.maxtblog.com/2018/11/custom-powershell-function-to-remove-azure-module/

    .EXAMPLE
    .\get-installedmoduleupdate.ps1

    Description
    -----------
    Runs script with default values.


    .INPUTS
    None. You cannot pipe objects to this script.
#>
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
  Function Get-Modules {
   $Array = @(Get-InstalledModule)

   Foreach ($Module in $Array)
   {
     $ModuleCheck = Get-InstalledModule -name $Module.Name -ErrorAction SilentlyContinue   

     if ($ModuleCheck) {
       Write-Host 'Info: Detected an installation of the',$Module.Name,'Module' -ForegroundColor $Green
       $Module = Get-InstalledModule -Name $Module.Name #-ListAvailable
       # Identify modules with multiple versions installed
       $g = $module | Group-Object -Property name -NoElement | Where-Object count -gt 1
       # Check Module from PSGallery
       Write-Host 'Checking',$Module.Name,'module from the PSGallery... Please wait' -ForegroundColor $White -BackgroundColor $DarkCyan
       $gallery = $module | Where-Object {$_.repositorysourcelocation}

       Write-Host 'Comparing installed version against online version of',$Module.Name,'module... Please wait' -ForegroundColor $White -BackgroundColor $DarkCyan
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
           Uninstall-AllModules -TargetModule $Module.Name -Force
           #Uninstall-Module -Name $Module.Name  -RequiredVersion $module.version 
           Write-Host -BackgroundColor $DarkRed -ForegroundColor $White 'Info: Legacy Version of',$Module.name,'module now removed'
           Install-Module -Name $Module.Name -RequiredVersion $online.Version -Force -allowclobber
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
 }
  Function Uninstall-AllModules {
[CmdletBinding(SupportsShouldProcess)]
param (
[Parameter(Mandatory = $true)]
[string]
$TargetModule,
[Parameter(Mandatory = $false)]
[string]
$Version,
[switch]
$Force
)

$AllModules = @()

Write-host 'Info: Checking for list of dependencies... Please wait' -ForegroundColor $Green
$target = Find-Module $TargetModule
$target.Dependencies | ForEach-Object {
$AllModules += New-Object -TypeName psobject -Property @{ name = $_.name}
}
$AllModules += New-Object -TypeName psobject -Property @{ name = $TargetModule}

$cnt = 1;
foreach ($module in $AllModules)
{
Write-Host ("[$cnt] - " + 'Info: Uninstalling the following dependant modules {0} ' -f $module.name) ;
$cnt++;
try
{
if ($PSCmdlet.ShouldProcess($module.name, 'Uninstall'))
{
Uninstall-Module -Name $module.name -Force:$Force -ErrorAction Stop;
};
}
catch
{
Write-Host ("`t" + $_.Exception.Message)
}
}
}

 #endregion
 Test-IsAdmin
 Write-host 'Version information - You are running script version 1.5' -ForegroundColor $White -BackgroundColor $DarkGray
  @'
    ┌─────────────────────────────────────────────────────────────┐
            Updating your Installed PowerShell Modules

               Follow @thewatchernode on Twitter                               
  └─────────────────────────────────────────────────────────────┘
'@
Start-Transcript -Path $env:USERPROFILE\desktop\InstalledModuleUpdate_Log.txt
Get-Modules
Stop-Transcript
