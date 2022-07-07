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
    : 1.6 Code cleanup

    Credit:
     http://www.maxtblog.com/2018/11/custom-powershell-function-to-remove-azure-module/

    .EXAMPLE
    .\Get-InstalledModulesUpdate.ps1

    Description
    -----------
    Runs script with default values.

    .INPUTS
    None. You cannot pipe objects to this script.
#>


$Script:Transcribing = $false
$Script:Version = "1.6"
$Script:TranscriptLogFolder = "$($env:HOMEDRIVE)\_PowerShellInstalledModulesUpdate" 
$Script:UpdatedComponents = @()


function Test-IsAdmin
{
  ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')
}


function Uninstall-AllModules
{
  [CmdletBinding(SupportsShouldProcess)]
  param (
    [Parameter(Mandatory = $true)]
    [string] $TargetModule,

    [Parameter(Mandatory = $false)]
    [string] $Version,

    [switch] $Force
  )

  $AllModules = @()
  (Find-Module -Name $TargetModule).Dependencies | ForEach-Object {
    $AllModules += $_.Name
  }
  $AllModules += $TargetModule

  foreach ($Module in $AllModules)
  {
    try
    {
      # avoid updating the same component twice
      if ( -not ($Script:UpdatedComponents -contains $Module))
      {
        $Type = if ($Module -eq $TargetModule) { "" } else { "dependent " }

        Write-Host "Info: uninstalling the $($Type)module '$($Module)'" -ForegroundColor White -BackgroundColor DarkRed
        if ($PSCmdlet.ShouldProcess($Module, 'Uninstall'))
        {
          Uninstall-Module -Name $Module -Force:$Force -ErrorAction Stop
        }
        $Script:UpdatedComponents += $Module
      }
    }
    catch
    {
      Write-Host "`t$($_.Exception.Message)"
    }
  }
}


function Get-Modules
{
  [CmdletBinding(SupportsShouldProcess)]
  param ()

  # Iterate over all installed modules
  $InstalledModules = @(Get-InstalledModule)
  foreach ($InstalledModule in $InstalledModules)
  {
    # Check module in PSGallery
    Write-Host "Checking '$($InstalledModule.Name)' module in PSGallery..." -ForegroundColor White -BackgroundColor DarkCyan

    $ModulesWithSoureLocations = $InstalledModule | Where-Object { $_.RepositorySourceLocation }
    foreach ($LocalModule in $ModulesWithSoureLocations) {
      # Find the current version in the gallery
      try
      {
        $OnlineModule = Find-Module -Name $LocalModule.Name -Repository PSGallery -ErrorAction Stop
      }
      catch
      {
        Write-Host "Warning: unable to find the '$($LocalModule.Name)' module in PSGallery" -ForegroundColor Yellow
      }

      # Compare versions
      if ($OnlineModule.Version -gt $LocalModule.Version)
      {
        Write-Host "Info: removing legacy version of '$($LocalModule.name)' module and it's dependencies" -BackgroundColor DarkRed -ForegroundColor White
        Uninstall-AllModules -TargetModule $LocalModule.Name -Force

        Write-Host "Info: updating '$($LocalModule.name)' module to new version" -ForegroundColor White -BackgroundColor DarkCyan
        if ($PSCmdlet.ShouldProcess("$($LocalModule.Name) $($OnlineModule.Version)", 'Install'))
        {
          Install-Module -Name $LocalModule.Name -RequiredVersion $OnlineModule.Version -Force -AllowClobber
        }
        $UpdatedVersion = $OnlineModule.Version
      }
      else
      {
        $UpdatedVersion = '(no update available)'
      }

      # Identify modules with multiple versions installed
      $ModulesWithMultipleVersions = $InstalledModule | Group-Object -Property Name -NoElement | Where-Object { $_.count -gt 1 }

      # Write a custom object to the pipeline
      [PSCustomObject]@{
        Name = $LocalModule.Name
        MultipleVersions = ($ModulesWithMultipleVersions.Name -contains $LocalModule.Name)
        InstalledVersion = $LocalModule.Version
        UpdatedVersion = $UpdatedVersion
      } | Format-List
    } 
  }
}


Clear-Host

if ( -not (Test-IsAdmin))
{
  throw 'Please note: you are trying to run this script without administative priviliges. In order to run this script you will required running PowerShell in Administrator Mode!'
}
else
{
  Write-Verbose 'Are you running this script as an Administator'
}

Write-host "Version information - You are running script version $($Script:Version)" -ForegroundColor White -BackgroundColor DarkGray
@'
  ┌─────────────────────────────────────────────────────────────┐
            Updating your Installed PowerShell Modules

               Follow @thewatchernode on Twitter
  └─────────────────────────────────────────────────────────────┘
'@

try
{
  Start-Transcript -Path "$($Script:TranscriptLogFolder)\InstalledModuleUpdate_Log.txt"
  $Script:Transcribing = $true
  Get-Modules
}
finally
{
  if ($Script:Transcribing)
  {
    Stop-Transcript
  }
}