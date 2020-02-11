<#Information
 
    Author: thewatchernode
    Contact: author@blogabout.cloud
    Published: 4th October 2011

    .DESCRIPTION
    Tool to assist with application delivery

    Version Changes            
    : 0.1 Initial Script Build
    : 1.0 Initial Build Release
    : 1.1 Resolved Option 2 issues not collecting Immuta

    .EXAMPLE
    .\set-immutableid.ps1

    Description
    -----------
    Runs script with default values. It 


    .INPUTS
    None. You cannot pipe objects to this script.
#>

#region Shortname

$DarkCyan = 'DarkCyan'
$DarkRed = 'DarkRed'
$Green = 'Green'
$Red = 'Red'
$Yellow = 'Yellow'
$Cyan = 'Cyan'
$White = 'White'

$AzureAD = 'AzureAD'
$Quit = 'Q'

#endregion

#region Banner
[string] $Root = @'
 
  ┌─────────────────────────────────────────────────────────────┐
            Gather ImmutableID in Bulk using PowerShell              
           
                Follow me @thewatchernode on Twitter 
                
                
   This script gathers the ImmutableID from Active Directory
   and Azure Active Directory.                   
  └─────────────────────────────────────────────────────────────┘
   
  1)  Connect to Azure AD                             -->

  2)  Get ImmutableID for AD                          -->
  3)  Get ImmutableID for AAD                         -->
  4)  Merge AD and AAD outputs (Coming Soon)          --> 
  5)  Set ImmutableID using Option 4 (Coming Soon)    --> 
     
  Q) Quit

  Select an option.. [1-99]?
'@
[string] $ADRoot = @'
 
  ┌─────────────────────────────────────────────────────────────┐
            Gather ImmutableID in Bulk using PowerShell              
           
                Follow me @thewatchernode on Twitter 
                
                
   This section will gather the UserPrincipleName,SamAccountName 
   and ImmutableID for all Azure AD Users..                   
  └─────────────────────────────────────────────────────────────┘
'@
[string] $AADRoot = @'
 
  ┌─────────────────────────────────────────────────────────────┐
            Gather ImmutableID in Bulk using PowerShell              
           
                Follow me @thewatchernode on Twitter 
                
                
   This section will gather the UserPrincipleName,ObjectID and
   ImmutableID for all Azure AD Users.                 
  └─────────────────────────────────────────────────────────────┘
'@
[string] $ImportExcel = @'
 
  ┌─────────────────────────────────────────────────────────────┐
            Gather ImmutableID in Bulk using PowerShell              
           
                Follow me @thewatchernode on Twitter 
                
                
   Please Note: The worksheet in both files MUST be called 
                Sheet1                
  └─────────────────────────────────────────────────────────────┘
'@
#endregion Banner
#region Menu Prompt
function Get-Root    {
  # Menu Prompt
  
    Do {
      $MenuOption = Read-Host -Prompt $Root
      Clear-Host
      switch ($MenuOption){
        1 { # Connect to AzureAD
             Get-AADConnect
          } 
        2 { # Get ImmutableID for AD
             Get-ADImmutableID
          }
        3 { # Get ImmutableID for AAD
             Get-ADDImmutableID
          }
        4 { # Merge Worksheets
            Get-ImportExcel
            Get-MergeFiles
          }
        5 { # Set ImmutableID in Azure
            
          }

        $Quit {return} 
      }
    
    }  until ($Root -eq {$Quit})
  }
#endregion Menu Prompt
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White 'Warning: Legacy Version of AzureAD Module detected. Starting removing process'
        Uninstall-Module -Name $AzureAD -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White 'Info: Legacy Version of AzureAD Module now removed'
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
Function Get-AADConnect {

Get-AzureADPSVersion
Get-AzureAD

Write-Host 'INFO: Connecting to Azure Active Directory, prompting for relevant administrative credentials' -BackgroundColor $Green
Connect-AzureAD

}
Function Get-ADImmutableID {
$ADRoot
$reportoutput=@()
$users = Get-ADUser -Filter * -Properties *
$users | Foreach-Object {

    $user = $_
    $immutableid = [System.Convert]::ToBase64String($user.ObjectGUID.tobytearray())
    $userid = $user | Select-Object @{Name='Access Rights';Expression={[string]::join(', ', $immutableid)}}

    $report = New-Object -TypeName PSObject
    $report | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $user.UserPrincipalName
    $report | Add-Member -MemberType NoteProperty -Name 'SamAccountName' -Value $user.samaccountname
    $report | Add-Member -MemberType NoteProperty -Name 'ImmutableID' -Value $immutableid
    $reportoutput += $report
}
 # Report
$reportoutput | Export-Csv -Path $env:USERPROFILE\desktop\ImmutableID4AD.csv -NoTypeInformation -Encoding UTF8 }
Function Get-ADDImmutableID {
$AADRoot
$reportoutput=@()
$users = Get-ADUser -Filter * -Properties *
$users | Foreach-Object {

    $user = $_
    $immutableid = [System.Convert]::ToBase64String($user.ObjectGUID.tobytearray())
    $userid = $user | Select-Object @{Name='Access Rights';Expression={[string]::join(', ', $immutableid)}}

    $report = New-Object -TypeName PSObject
    $report | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $user.UserPrincipalName
    $report | Add-Member -MemberType NoteProperty -Name 'SamAccountName' -Value $user.samaccountname
    $report | Add-Member -MemberType NoteProperty -Name 'ImmutableID' -Value $immutableid
    $reportoutput += $report
}
 # Report
$reportoutput | Export-Csv -Path $env:USERPROFILE\desktop\ImmutableID4AAD.csv -NoTypeInformation -Encoding UTF8 }
Function Get-RenameCSVtoXLSX {

    $proj_files = Get-ChildItem | Where-Object {$_.Extension -ne '.jpg'}
    ForEach ($file in $proj_files) {
    $filenew = $file.Name + '.xlsx'
    Rename-Item $file $filenew
    }
}
Function Get-RenameXLSXtoCSV {

    $proj_files = Get-ChildItem | Where-Object {$_.Extension -ne '.jpg'}
    ForEach ($file in $proj_files) {
    $filenew = $file.Name + '.csv'
    Rename-Item $file $filenew
    }
}
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
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White 'Warning: Legacy Version of ImportExcel Module detected. Starting removing process'
        Uninstall-Module -Name $ImportExcel -RequiredVersion $module.version 
        Write-Host -BackgroundColor $DarkRed -ForegroundColor $White 'Info: Legacy Version of ImportExcel Analyzer Module now removed'
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
Function Get-MergeFiles {
   
   Clear-Host

   $ref = Read-Host -Prompt 'Specify your Reference File for example (c:\ref.xlsx)'
   Write-Host -Foreground $Cyan ('You have specified {0}' -f $ref)
   $dif = Read-Host -Prompt 'Specify your Difference File for example (c:\dif.xlsx)'
   Write-Host -Foreground $Cyan ('You have specified {0}' -f $dif)
   $out = Read-Host -Prompt 'Specify your Reference File for example (c:\output.xlsx)'
   Write-Host -Foreground $Cyan ('You have specified {0}' -f $out)

   Merge-Worksheet -Referencefile $ref -Differencefile $dif -OutputFile $out -WorksheetName Sheet1 -Startrow 1 -OutputSheetName Sheet1 -NoHeader
}
Function Set-ImmutableID {}
#endregion


#region Code Launch
clear-host
#Requires -Version 5.0
Write-Output "I'm version 5.0 or above"
$PSVersionTable
Test-IsAdmin
Write-host 'Version information - You are running script version 1.1' -ForegroundColor White -BackgroundColor DarkGray
Get-Root
#endregion Code Launch

