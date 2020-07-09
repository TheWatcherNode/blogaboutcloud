 <#Information
     Office Deployment Tool Version officedeploymenttool_8529.3600

     Date Published: 11/12/2018 (UK Format)
     File Size: 2.65 MB

     Configures the necessary registry key for Office 365 ProPlus

     .DESCRIPTION
     Toolto assist with the deployment of Office 365 ProPlus

     .NOTES
     Version				: 1.1
     Wish list			    : Better error trapping
                            : 
     Rights Required	    : Local administrator on Workstation
                            : 
                            : Powershell in Administrator Mode
     Sched Task Required	: No
     Windows Server Version	: 
     Author/Copyright	    : © The Watcher - All Rights Reserved
     Email/Blog/Twitter	    : thewatcher@blogabout.cloud @thewatchernode
     Dedicated Post	        : 
 
     Disclaimer              : You running this script means you won't blame me if this breaks your stuff. This script is provided AS IS without warranty of any kind. 
                              I disclaim all implied warranties including,
                              without limitation, any implied warranties of merchantability or of fitness
                              .SYNOPSIS for a particular
                              purpose. The entire risk arising out of the use or performance of the sample scripts and
                              documentation remains with you. In no event shall I be liable for any damages whatsoever
                              (including, without limitation, damages for loss of business profits, business interruption,
                              loss of business information, or other pecuniary loss) arising out of the use of or inability
                              to use the script or documentation.

     Acknowledgements 	        : 
                                : 
                                :
     Assumptions		            : ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
                                :  
     Limitations			          : 
     Known issues				        :
  
     Version Changes            : 0.1 Initial Script Build
                                : 1.0 Initial Build
                                : 1.1 Microsoft Apps for Enterprise Channel Name Updates
     .LINK
     

     .EXAMPLE
     .\get-officeproplustoolkit.ps1

     Description
     -----------
     Runs script with default values.


     .INPUTS
     None. You cannot pipe objects to this script.
 #>
 <# Future Updates
  OS Version check for Telemetry Dashboar
 #>
 #region Admin Test
 function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')
    }
    if (!(Test-IsAdmin)){
    throw 'Please Note: You are trying to run this script without evalated Administator Priviliges. In order to run this script you will required Powershell running in Administrator Mode'
    }
    else {
    Write-Verbose -Message 'Are you running as evalated Administator' -verbose
    } 
 #endregion Admin Test
 #region Shortnames
  #File Paths
  $path = 'HKLM:\Software\Policies\Microsoft\office\16.0\common\officeupdate\'
  $path1 = 'HKLM:\Software\Policies\Microsoft\'
  $path2 = 'HKLM:\Software\Policies\Microsoft\office\'
  $path3 = 'HKLM:\Software\Policies\Microsoft\office\16.0\'
  $path4 = 'HKLM:\Software\Policies\Microsoft\office\16.0\common\'
  $path5 = 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.5'
  $DownloadC2R = "$env:USERPROFILE\Downloads\Office365 ProPlus\OfficeC2R"
  $DownloadDir = "$env:USERPROFILE\Downloads\Office365 ProPlus"
  $C2RPath = "$env:CommonProgramW6432\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"

  # Shortened Name
  $key1 = 'office'
  $key2 = '16.0'
  $key3 = 'common'
  $key4 = 'officeupdate'
  $DisplaySetting = '/update user displaylevel=true'
  $download = '/download'
  $updatebranch = 'updatebranch'
  $Quit = 'Q'
  
  # Channel Names 
  $CurrentPreview = 'CurrentPreview'
  $Current = 'Current'
  $MonthlyEnterprise = 'MonthlyEnterprise'
  $SemiAnnual = 'SemiAnnual'
  $SemiAnnualPreview = 'SemiAnnualPreview'
  $SilentlyContinue = 'SilentlyContinue'
  
  # Colours
  $White = 'White'
  $Red = 'Red'
  $Green = 'Green'
  $DarkGreen = 'DarkGreen'
  $DarkBlue = 'DarkBlue'
  $DarkGray = 'DarkGray'
  $DarkRed = 'DarkRed'
  $DarkCyan = 'DarkCyan'
  
  #Readiness Tool
  $Readiness = 'https://download.microsoft.com/download/E/C/7/EC7CDC27-3750-4F12-8DB4-ABFDFDF7DA5C/ReadinessToolkitForOffice.msi'
  $ReadinessFile = "$env:USERPROFILE\Downloads\Office365 ProPlus\ReadinessToolkitForOffice.msi"

  # FixIT Files
  $FixIT = 'https://aka.ms/diag_officeuninstall'
  $FixITFile = "$env:USERPROFILE\Downloads\Office365 ProPlus\o15-ctremove.diagcab"

  # OffScrub Files
  $Off = 'http://blogabout.cloud/download/162/'
  $OffDestinFile = "$env:USERPROFILE\Downloads\Office365 ProPlus\OffScrub VBS.zip"
  
  # Group Policy for Office 2016/2019/ProPlus
  $GP64 = 'https://download.microsoft.com/download/2/E/E/2EEEC938-C014-419D-BB4B-D184871450F1/admintemplates_x64_4768-1000_en-us.exe'
  $GP86 = 'https://download.microsoft.com/download/2/E/E/2EEEC938-C014-419D-BB4B-D184871450F1/admintemplates_x86_4768-1000_en-us.exe'
  $GPFilex64 = "$env:USERPROFILE\Downloads\Office365 ProPlus\admintemplates_x64_4768-1000_en-us.exe"
  $GPFilex86 = "$env:USERPROFILE\Downloads\Office365 ProPlus\admintemplates_x86_4768-1000_en-us.exe"
  
  # SQL Elements 
  $SQL64 = 'https://download.microsoft.com/download/E/A/E/EAE6F7FC-767A-4038-A954-49B8B05D04EB/Express%2064BIT/SQLEXPR_x64_ENU.exe'
  $SQLM64 = 'https://download.microsoft.com/download/E/A/E/EAE6F7FC-767A-4038-A954-49B8B05D04EB/MgmtStudio%2064BIT/SQLManagementStudio_x64_ENU.exe'
  $SQLFilex64 = "$env:USERPROFILE\Downloads\Office365 ProPlus\SQLEXPR_x64_ENU.exe"
  $SQLMFilex64 = "$env:USERPROFILE\Downloads\Office365 ProPlus\SQLManagementStudio_x64_ENU.exe"
  # Office Telemetry Dashboard 
  $OTDB = 'https://download.microsoft.com/download/6/0/C/60CC12AB-5140-4FB1-91C7-C0B4464E8288/TDadmSetUp.msi'
  $OTDBFile = "$env:USERPROFILE\Downloads\Office365 ProPlus\TDadmSetUp.msi"
  $TDadmSetup = '.\TDadmSetUp.msi'
  # Tags 
  $Tags = 'https://gallery.technet.microsoft.com/Create-Group-Policy-Object-7407976b/file/70476/1/Create-GroupPolicyObjectForTags.ps1'
  $TagsFile = "$env:USERPROFILE\Downloads\Office365 ProPlus\Create-GroupPolicyObjectForTags.ps1"
  $VBS = 'https://gallery.technet.microsoft.com/Add-Tags-Labels-for-Office-09c81574/file/70474/1/Add-Tags.vbs'
  $VBSFile = "$env:USERPROFILE\Downloads\Office365 ProPlus\Add-Tags.vbs"
  # DotNet 
  $Net35 = 'https://download.microsoft.com/download/7/0/3/703455ee-a747-4cc8-bd3e-98a615c3aedb/dotNetFx35setup.exe'
  $Net35File = "$env:USERPROFILE\Downloads\Office365 ProPlus\dotNetFx35setup.exe"
  # File names 
  $SQLFileName = "SQLEXPR_x64_ENU.exe"
  $SQLMFileName = "SQLManagementStudio_x64_ENU.exe"
  $OTDBFileName = "TDadmSetUp.msi"
  $TagsFileName = "Create-GroupPolicyObjectForTags.ps1"
  $VBSFileName = "Add-Tags.vbs"
  $Net35FileName = "dotNetFx35setup.exe"
  
  # Download Official Deployment Tool
  $DeploymentTool = "$env:USERPROFILE\Downloads\Office365 ProPlus\officedeploymenttool.exe"
  $SourceDeployment = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe'

  # Pre Configured Configuration.xml
  $SourceDownload = 'http://blogabout.cloud/download/165/16'
  $DownloadDir = "$env:USERPROFILE\Downloads\Office365 ProPlus"
  $DestinFile = "$env:USERPROFILE\Downloads\Office365 ProPlus\OfficeC2R.zip"
  
  # Office XMLs
  $Monthly64 = 'configuration-Office365-x64.xml'
  $Monthly86 = 'configuration-Office365-x86.xml'
  $Office2019x64 = 'configuration-Office2019Enterprise.xml'
  
  # Install Office MSI
  $officeVersions=@("
                     Microsoft Office Professional Plus 2010"
                     "Microsoft Office Professional Plus 2013",
                     "Microsoft Office Professional Plus 2016")                                
  $exists = Get-WMIObject win32_product | Where {$_.Name -like $officeVersions}

  # PowerShell Version
  $PSVersion = $PSVersionTable.PSVersion

  # OS Version
  $OSVersion = (Get-WmiObject Win32_OperatingSystem)

 #endregion Shortnames
 #region Menus
   [string] $FutureRoot = @'
 
  ┌─────────────────────────────────────────────────────────────┐
                      Office ProPlus Tool Kit               
           
                Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

  1)  Configure Current Channel                                -->
  2)  Configure Current Channel (Preview)                      -->
  3)  Configure Monthly Enterprise Channel                     -->
  4)  Configure Semi-Annual Enterprise Channel                 -->
  5)  Configure Semi-Annual Enterprise Channel (Preview)       -->
  
  7)  Download the Office Readiness Toolkit for Add-ins & VBS  -->
  8)  Download Microsoft FixIT Removal Tool                    -->
  9)  Download Offscrub Files (Office 03,07,10, O15 & O16)     -->
  10) Download Office 2016/2019/ProPlus Group Policy Templates -->
  11) Download Office Telemetry Requirements                   -->
  
  15) Download Office Deployment Tool (Official)               -->
  16) Download Pre-Loaded Office 365 Configuration Files       -->
  
  20) Install Office 365 ProPlus                               -->
  21) Install SQL Express                                      -->
  22) Install SQL Management Studio                            -->
  23) Install Office Telemetry Dashboard                       -->
     
      Please Note: We are working on fully automating the 
      installation of Office Telemetry Dashboard
  
  30) Build your own configuration.xml (config.office.com)     -->
  31) Install Office using your modified configuration.xml     --> 
     
     Please Note: All Downloads will be stored in 
                  Downloads\Office365 ProPlus Folder
  Q) Quit

  Select an option.. [1-99]?
'@
  [string] $Root = @'
 
  ┌─────────────────────────────────────────────────────────────┐
                      Office ProPlus Tool Kit               
           
                Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

  1)  Configure Current Channel                                -->
  2)  Configure Current Channel (Preview)                      -->
  3)  Configure Monthly Enterprise Channel                     -->
  4)  Configure Semi-Annual Enterprise Channel                 -->
  5)  Configure Semi-Annual Enterprise Channel (Preview)       -->
  6)  Check your Office 365 ProPlus Configuration              -->

  7)  Download the Office Readiness Toolkit for Add-ins & VBS  -->
  8)  Download Microsoft FixIT Removal Tool                    -->
  9)  Download Offscrub Files (Office 03,07,10, O15 & O16)     -->
  10) Download Office 2016/2019/ProPlus Group Policy Templates -->
  11) Download Office Telemetry Requirements                   -->
  
  21) Install SQL Express                                      -->
  22) Install SQL Management Studio                            -->
  23) Install Office Telemetry Dashboard                       -->
     
      Please Note: We are working on fully automating the 
      installation of Office Telemetry Dashboard
  
  30) Build your own configuration.xml (config.office.com)     -->
  31) Install Office using your modified configuration.xml     --> 
     
     Please Note: All Downloads will be stored in 
                  Downloads\Office365 ProPlus Folder
  Q) Quit

  Select an option.. [1-99]?
'@
  [string] $SubRoot16 = @'
 
  ┌─────────────────────────────────────────────────────────────┐
                      Office ProPlus Tool Kit                                 
  └─────────────────────────────────────────────────────────────┘
  This section of this script requires you to run option 1 first
  to download the Deployment Tool and pre-configured XML Files.

  1)  Download Office Deployment Tool with Modified XMLs          -->
  2)  Download Office 365 ProPlus Monthly Channel (x64)           -->
  3)  Download Office 365 ProPlus Monthly Channel (x86)           -->
  4)  Download Office 2019 Enterprise Editons     (x64)           -->

  98) Return to Main Menu
  Q) Exit Script

     
     Please Note: All Downloads will be stored in 
                  Downloads\Office365 ProPlus Folder

  Select an option.. [1-99]?
'@
  [string] $SubRoot20 = @'
 
  ┌─────────────────────────────────────────────────────────────┐
                      Office ProPlus Tool Kit                                 
  └─────────────────────────────────────────────────────────────┘
  This section of this script uses the downloaded XML files from
  option 16. If you require additional configuration options please
  edit the relevent XML file.

  Please Note: If you have download the Office Delta prior to
  installation, each option will download delta before installing
  Office.

  1)  Install Office 365 ProPlus Monthly Channel (x64)           -->
  2)  Install Office 365 ProPlus Monthly Channel (x86)           -->
  3)  Install Office 2019 Enterprise Editons     (x64)           -->
  

  98) Return to Main Menu
  Q) Quit

     
     Please Note: All Downloads will be stored in 
                  Downloads\Office365 ProPlus Folder

  Select an option.. [1-99]?
'@
 #endregion
 #region Functions
  function Get-Root    { # Menu Prompt
  
    Do {
      $MenuOption = Read-Host -Prompt $Root

      switch ($MenuOption){
        1  { # Monthly Channel
             Get-Root1
          } 
        2  { # Semi Annual (Targeted) Channel
             Get-Root2
          }
        3  { # Semi Annual Channel
             Get-Root3 
          }
        4  { # Monthly (Targeted) Channel
             Get-Root4
          }   
        5  { # Insider Fast Channel
             Get-Root5
          }
        6  { # Check 
             Get-Root6
          }
        7  {  # Download Readiness Toolkit for Office Add-ins and VBA
             Get-Root7
          }
        8  {  # Download FixIT Removal Tool
             Get-Root8 
          }
        9  {  # Download Offscrub Files
             Get-Root9
          }
        10  {  # Download Group Policy
             Get-Root10
          }
        11 {  # Configure Workstation/Server for Office Telemetry Dashboard  
             Get-Root11
          }
        15 { # Download Office Deployment Tool
             Get-Root15
          }
        16 { # Download Pre-Configured Configuration.xml with Office Deltas
             Get-Root16
          }
        20 { # Install Office 2016
             Get-Root20
           }
        21 { # Install SQL Express
             Get-Root21
           }
        22 { # Install SQL Management Studio
             Get-Root22
           }
        23 { # Install Office Telemetry Dashboard
             Get-Root23
           }
        30 { # Launch config.office.com
             Get-Root30
           }  
        31 { # Launch Install Office with your configuration file
             Get-Root31
           }
        32 { # Launch configuration information from reg
             Get-Root32
           }  
        $Quit {exit} 
      }
    
    }  until ($Root -eq {$Quit})
  }
  function Get-Root1   { # Current Channel

    # Check if updatebranch REG_SZ Exists
    
    if (Get-ItemProperty -Path $path -ErrorAction $SilentlyContinue)
    {
 
     # Create REG_SZ Key for Monthly Channel
     Write-Host 'INFO:Checking if updatebranch REG_SZ already exists' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-host 'INFO:REG_SZ exists, Setting key for Current Channel' -ForegroundColor $White -BackgroundColor $DarkGreen
     Set-ItemProperty -Path $path -Name $updatebranch -Value $Current
    } 
    else 
    {
     Write-Host 'ERROR: Checking if updatebranch REG_SZ already exists' -ForegroundColor $White -BackgroundColor $Red
     Write-host 'ERROR: REG_SZ doesnt exist, creating Keys and setting key for Current Channel' -ForegroundColor $White -BackgroundColor $Red
     New-Item -Path $path1 -Name $Key1 -ErrorAction $SilentlyContinue 
     New-Item -Path $path2 -Name $Key2 -ErrorAction $SilentlyContinue 
     New-Item -Path $path3 -Name $Key3 -ErrorAction $SilentlyContinue 
     New-Item -Path $path4 -Name $Key4 -ErrorAction $SilentlyContinue 
     New-ItemProperty -Path $path -Name $updatebranch -PropertyType String -Value $Current -Force -ErrorAction $SilentlyContinue
     Write-host 'INFO: REG_SZ Keys, successful created' -ForegroundColor $White -BackgroundColor $DarkGreen
    }
 
    # Update Office
    Start-Process -FilePath $C2RPath -ArgumentList $DisplaySetting
 
    # Return to Menu
    Get-Root
  }
  function Get-Root2   { # Current Channel (Preview)

    # Check if updatebranch REG_SZ Exists
    if (Get-ItemProperty -Path $path -ErrorAction $SilentlyContinue)
    {
 
     # Create REG_SZ Key for Semi Annual (Targeted) Channel
     Write-Host 'INFO: Checking if updatebranch REG_SZ already exists' -ForegroundColor $White -BackgroundColor $Green
     Write-host 'INFO: REG_SZ exists, Setting key for Current Channel (Preview)' -ForegroundColor $White -BackgroundColor $Green
     Set-ItemProperty -Path $path -Name $updatebranch -Value $CurrentPreview
    } 
    else 
    {
     Write-Host 'ERROR: Checking if updatebranch REG_SZ already exists' -ForegroundColor $White -BackgroundColor $Red
     Write-host 'EEROR: REG_SZ doesnt exist, creating Keys and setting key for Current Channel (Preview)' -ForegroundColor $White -BackgroundColor $Red
     New-Item -Path $path1 -Name $Key1 -ErrorAction $SilentlyContinue
     New-Item -Path $path2 -Name $Key2 -ErrorAction $SilentlyContinue
     New-Item -Path $path3 -Name $Key3 -ErrorAction $SilentlyContinue
     New-Item -Path $path4 -Name $Key4 -ErrorAction $SilentlyContinue
     New-ItemProperty -Path $path -Name $updatebranch -PropertyType String -Value $CurrentPreview -Force
    }
 
    # Update Office
    Start-Process -FilePath $C2RPath -ArgumentList $DisplaySetting
 
    # Return to Menu
    Get-Root
  }
  function Get-Root3   { # Monthly Enterprise Channel
    # Check if updatebranch REG_SZ Exists
    if (Get-ItemProperty -Path $path -ErrorAction $SilentlyContinue)
    {
     # Create REG_SZ Key for Semi Annual Channel
     Write-Host 'INFO: Checking if updatebranch REG_SZ already exists' -ForegroundColor $White -BackgroundColor $Green
     Write-host 'INFO: REG_SZ exists, Setting key for Monthly Enterprise Channel' -ForegroundColor $White -BackgroundColor $Green
     Set-ItemProperty -Path $path -Name $updatebranch -Value $MonthlyEnterprise
    } 
    else 
    {
     Write-Host 'ERROR: Checking if updatebranch REG_SZ already exists' -ForegroundColor $White -BackgroundColor $Red
     Write-host 'ERROR: REG_SZ doesnt exist, creating Keys and setting key for Monthly Enterprise Channel' -ForegroundColor $White -BackgroundColor $Red
     New-Item -Path $path1 -Name $Key1 -ErrorAction $SilentlyContinue
     New-Item -Path $path2 -Name $Key2 -ErrorAction $SilentlyContinue
     New-Item -Path $path3 -Name $Key3 -ErrorAction $SilentlyContinue
     New-Item -Path $path4 -Name $Key4 -ErrorAction $SilentlyContinue
     New-ItemProperty -Path $path -Name $updatebranch -PropertyType String -Value $MonthlyEnterprise -Force
    }
 
    # Update Office
    Start-Process -FilePath $C2RPath -ArgumentList $DisplaySetting
 
    # Return to Menu
    Get-Root
  }
  function Get-Root4   { # Semi-Annual Enterprise Channel
    if (Get-ItemProperty -Path $path -ErrorAction $SilentlyContinue)
    {
     # Create REG_SZ Key for Semi Annual Channel
     Write-Host 'INFO: Checking if updatebranch REG_SZ already exists' -ForegroundColor $White -BackgroundColor $Green
     Write-host 'INFO: REG_SZ exists, Setting key for Semi-Annual Enterprise Channel' -ForegroundColor $White -BackgroundColor $Green
     Set-ItemProperty -Path $path -Name $updatebranch -Value $SemiAnnual
    } 
    else 
    {
     Write-Host 'ERROR: Checking if updatebranch REG_SZ already exists' -ForegroundColor $White -BackgroundColor $Red
     Write-host 'ERROR: REG_SZ doesnt exist, creating Keys and setting key for Semi-Annual Enterprise Channel' -ForegroundColor $White -BackgroundColor $Red
     New-Item -Path $path1 -Name $Key1 -ErrorAction $SilentlyContinue
     New-Item -Path $path2 -Name $Key2 -ErrorAction $SilentlyContinue
     New-Item -Path $path3 -Name $Key3 -ErrorAction $SilentlyContinue
     New-Item -Path $path4 -Name $Key4 -ErrorAction $SilentlyContinue
     New-ItemProperty -Path $path -Name $updatebranch -PropertyType String -Value $SemiAnnual -Force
    }
    # Update Office
    Start-Process -FilePath $C2RPath -ArgumentList $DisplaySetting
    # Return to Menu
    Get-Root
  }
  function Get-Root5   {  # Semi-Annual Enterprise Channel (Preview)
     if (Get-ItemProperty -Path $path -ErrorAction $SilentlyContinue)
    {
      # Create REG_SZ Key for Semi Annual Channel
      Write-Host 'INFO: Checking if updatebranch REG_SZ already exists' -ForegroundColor $White -BackgroundColor $Green
      Write-host 'INFO: REG_SZ exists, Setting key for Semi-Annual Enterprise Channel (Preview)' -ForegroundColor $White -BackgroundColor $Green
      Set-ItemProperty -Path $path -Name $updatebranch -Value $SemiAnnualPreview
    } 
    else 
    {
     Write-Host 'ERROR: Checking if updatebranch REG_SZ already exists' -ForegroundColor $White -BackgroundColor $Red
     Write-host 'ERROR: REG_SZ doesnt exist, creating Keys and setting key for Semi-Annual Enterprise Channel (Preview)' -ForegroundColor $White -BackgroundColor $Red
     New-Item -Path $path1 -Name $Key1 -ErrorAction $SilentlyContinue
     New-Item -Path $path2 -Name $Key2 -ErrorAction $SilentlyContinue
     New-Item -Path $path3 -Name $Key3 -ErrorAction $SilentlyContinue
     New-Item -Path $path4 -Name $Key4 -ErrorAction $SilentlyContinue
     New-ItemProperty -Path $path -Name $updatebranch -PropertyType String -Value $SemiAnnualPreview -Force
    }
    # Update Office
    Start-Process -FilePath $C2RPath -ArgumentList $DisplaySetting
    # Return to Menu
    Get-Root
  }
  function Get-Root6   {
  if (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -ErrorAction SilentlyContinue)
  {
  Write-Host 'INFO: Obtaining configuration information registry' $SetLocation -ForegroundColor $White -BackgroundColor $DarkBlue
  }
  else
  {
  Write-Host 'ERROR: Office 365 ProPlus installation not detected' -BackgroundColor $DarkRed
  }
  Get-Root  
  }
  function Get-Root7   { # Readiness Toolkit for Office Add-ins and VBA
    if (Get-ItemProperty -Path $DownloadDir -ErrorAction $SilentlyContinue)
    {
     Write-Host 'INFO: Correct download path detected' -ForegroundColor $White -BackgroundColor $DarkGreen
     # Create REG_SZ Key for Semi Annual Channel
     Import-module BitsTransfer 
     Start-BitsTransfer $Readiness $ReadinessFile
     Write-Host 'INFO: Download completed, returning to Main Menu' -ForegroundColor $White -BackgroundColor $DarkGreen
    } 
    else 
    {
     Write-Host 'ERROR: Correct download path not detected' -ForegroundColor $White -BackgroundColor $DarkRed
     Write-Host 'ERROR: Creating required folder directory $env:USERPROFILE\Downloads\Office365 ProPlus' -ForegroundColor $White -BackgroundColor $DarkRed
     New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\Downloads\Office365 ProPlus"
     Import-module BitsTransfer 
     Start-BitsTransfer $Readiness $ReadinessFile
     Write-Host 'INFO: Download completed, returning to Main Menu' -ForegroundColor $White -BackgroundColor $DarkGreen
    }
    Get-Root
    }
  function Get-Root8   { # Download Fix IT
    if (Get-ItemProperty -Path $DownloadDir -ErrorAction $SilentlyContinue)
    {
     Write-Host 'INFO: Correct download path detected' -ForegroundColor $White -BackgroundColor $DarkGreen
     
     # Create REG_SZ Key for Semi Annual Channel
     Import-module BitsTransfer 
     $client = new-object -TypeName System.Net.WebClient
     $client.DownloadFile($FixIT,$FixITFile)
     
     Write-Host 'INFO: Download completed, returning to Main Menu' -ForegroundColor $White -BackgroundColor $DarkGreen
    } 
    else 
    {
     Write-Host 'ERROR: Correct download path not detected' -ForegroundColor $White -BackgroundColor $DarkRed
     Write-Host 'ERROR: Creating required folder directory $env:USERPROFILE\Downloads\Office365 ProPlus' -ForegroundColor $White -BackgroundColor $DarkGreen
     New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\Downloads\Office365 ProPlus"
     I
     mport-module BitsTransfer 
     Start-BitsTransfer $FixIT $FixITFile
     Write-Host 'INFO: Download completed, returning to Main Menu' -ForegroundColor $White -BackgroundColor $DarkGreen
    }
    Get-Root
    }
  function Get-Root9   { # Download Offscrub
    if ($psversion.Major -gt 4)
    {
      Write-Host 'INFO: PowerShell Version 5 detected' -ForegroundColor $White -BackgroundColor $DarkBlue
      New-Item -ItemType Directory -Force -Path $DownloadDir
    
    # Download from OneDrive
      Write-Output -InputObject 'INFO: Downloading zip file'
      New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\Downloads\Office365 ProPlus" 
      $client = new-object -TypeName System.Net.WebClient
      $client.DownloadFile($Off,$OffDestinFile)
    # Checking
    if (Get-ItemProperty -Path $OffDestinFile)
    {
      # Check if Download was successful
      Write-Host 'INFO: Download completed, continuing to next step' -ForegroundColor $White -BackgroundColor $DarkGreen
      # Extract ZIP File
      Write-Host 'Please Note: This next step requires Powershell Version 5. Unzipping OffScrub VBS.zip' -ForegroundColor $White -BackgroundColor $DarkBlue
      Expand-Archive -Path "$env:USERPROFILE\Downloads\Office365 ProPlus\OffScrub VBS.zip" -DestinationPath "$env:USERPROFILE\Downloads\Office365 ProPlus" -force -ErrorAction $SilentlyContinue
    } 
    else 
    {
      Write-Host 'ERROR: Download incompleted, returning to Main Menu' -ForegroundColor $White -BackgroundColor $Red 
    }
    }
    else
    {
      Write-Host 'ERROR: PowerShell Version 5 not detected' -ForegroundColor $White -BackgroundColor $Darkred
      New-Item -ItemType Directory -Force -Path $DownloadDir
    # Download from OneDrive
      Write-Output -InputObject 'INFO: Downloading zip file'
      New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\Downloads\Office365 ProPlus" 
      $client = new-object -TypeName System.Net.WebClient
      $client.DownloadFile($Off,$OffDestinFile)
    # Checking
      if (Get-ItemProperty -Path $OffDestinFile)
    {
      # Check if Download was successful
      Write-Host 'INFO: Download completed, manual extraction of zip file required' -ForegroundColor $White -BackgroundColor $DarkGreen
      Get-FileName -initialdirectory "$env:USERPROFILE\downloads\Office365 ProPlus"
      
    } 
    else 
    {
      Write-Host 'ERROR: Download incompleted, returning to Main Menu' -ForegroundColor $White -BackgroundColor $Red 
    }
    }
      Get-Root
    }
  function Get-Root10  { # Download Group Policy Template
    if (Get-ItemProperty -Path $DownloadDir -ErrorAction $SilentlyContinue)
    {
     Write-Host 'INFO: Download path detected' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Downloading x64 Group Policy Templates' -ForegroundColor $White -BackgroundColor $DarkBlue
     # Download Group Policy
     $client = new-object -TypeName System.Net.WebClient
     $client.DownloadFile($GP64,$GPFilex64)
     Write-Host 'INFO: Download Completed X64 Group Policy Template' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Downloading x86 Group Policy Templates' -ForegroundColor $White -BackgroundColor $DarkBlue
     $client.DownloadFile($GP86,$GPFilex86)
     Write-Host 'INFO: Download Completed X86 Group Policy Template' -ForegroundColor $White -BackgroundColor $DarkGreen
    } 
    else 
    {
     Write-Host 'ERROR: Download Path not detected' -ForegroundColor $White -BackgroundColor $Red
     Write-Host 'ERROR: Creating required folder directory $env:USERPROFILE\Downloads\Office365 ProPlus' $White -BackgroundColor $Red
     New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\Downloads\Office365 ProPlus"
     $client = new-object -TypeName System.Net.WebClient
     $client.DownloadFile($FixIT,$FixITFile)
     Write-Host 'INFO: Download completed, returning to Main Menu' -ForegroundColor $White -BackgroundColor $DarkGreen
    }
    Get-Root
    }    
  function Get-Root11  { # Download Office Telemetry Requirements
    if (Get-ItemProperty -Path $DownloadDir -ErrorAction SilentlyContinue)
    {
     # Path Exists
     Write-Host 'INFO: Download path dectected' -ForegroundColor $White -BackgroundColor $DarkGreen
     
if (Get-Item -Path $SQLFilex64 -ErrorAction $SilentlyContinue)
     {
     # Detected SQL Server Express
     Write-Host "INFO: $SQLFileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not Detected SQL Server Express
     Write-Host "INFO: $SQLFileName has not been detected, starting download" -ForegroundColor $White -BackgroundColor $DarkGray
     Start-BitsTransfer -Source $SQL64 -Destination $SQLFilex64
     #$client.DownloadFile($SQL64,$SQLFilex64)
     Write-Host 'INFO: Download completed for SQL Express Installation' -ForegroundColor $White -BackgroundColor $darkgreen
     Write-Host 'INFO: Downloading SQL Management Installation ' -ForegroundColor $White -BackgroundColor $DarkBlue
     }

if (Get-Item -Path $SQLMFilex64 -ErrorAction $SilentlyContinue)
     {
     # Detected SQL Server Management
     Write-Host "INFO: $SQLMFileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not Detected SQL Server Management
     Write-Host "INFO: $SQLMFileName has not been detected, starting download" -ForegroundColor $White -BackgroundColor $DarkGray
     Start-BitsTransfer -Source $SQLM64 -Destination $SQLMFilex64
     #$client.DownloadFile($SQLM64,$SQLMFilex64)
     Write-Host 'INFO: Download completed for SQL Management' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Downloading SQL Express Installation ' -ForegroundColor $White -BackgroundColor $DarkBlue
     }

if (Get-Item -Path $OTDBFile -ErrorAction $SilentlyContinue)
     {
     # Office Telemetry Dashboard
     Write-Host "INFO: $OTDBFileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not Office Telemetry Dashboard
     Write-Host "INFO: $OTDBFileName has not been detected, starting download"
     Start-BitsTransfer -Source $OTDB -Destination $OTDBFile
     #$client.DownloadFile($OTDB,$OTDBFile)
     Write-Host 'INFO: Download completed for SQL Express Tool' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Downloading Group Policy Telemetry Tags - PowerShell script' -ForegroundColor $White -BackgroundColor $DarkBlue
     }

if (Get-Item -Path $TagsFile -ErrorAction $SilentlyContinue)
     {
     # Tags
     Write-Host "INFO: $TagsFileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not Tags
     Write-Host "INFO: $TagsFileName has not been detected, starting download"
     Start-BitsTransfer -Source $Tags -Destination $TagsFile
     #$client.DownloadFile($Tags,$TagsFile)
     Write-Host 'INFO: Download completed for Group Policy Telemetry Tags - PowerShell script' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Downloading Group Policy Telemetry Tags - VBS File' -ForegroundColor $White -BackgroundColor $DarkBlue
     }

if (Get-Item -Path $VBSFile -ErrorAction $SilentlyContinue)
     {
     # VBS
     Write-Host "INFO: $VBSFileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not VBS
     Write-Host "INFO: $VBSFileName has not been detected, starting download"
     Start-BitsTransfer -Source $VBS -Destination $VBSFile
     #$client.DownloadFile($VBS,$VBSFile)
     Write-Host 'INFO: Download completed for Group Policy Telemetry Tags - VBS file' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Opening Internet Explorer for more information about Telemetry Dashboard' -ForegroundColor $White -BackgroundColor $DarkBlue
     }

if (Get-Item -Path $Net35File -ErrorAction $SilentlyContinue)
     {
     # DotNet
     Write-Host "INFO: $Net35FileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not DotNet
     Write-Host "INFO: $Net35FileName has not been detected, starting download"
     Start-BitsTransfer -Source $Net35 -Destination $Net35File
     Write-Host 'INFO: Download completed for SQL Express Installation' -ForegroundColor $White -BackgroundColor $darkgreen
     Write-Host 'INFO: Downloading SQL Management Installation ' -ForegroundColor $White -BackgroundColor $DarkBlue
     }
     
     $ie = New-Object -ComObject InternetExplorer.Application
     $ie.Navigate('https://technet.microsoft.com/en-us/library/jj219431.aspx')
     $ie.Visible = $true
    } 
    else 
    {
     # Path Creation
     Write-Host 'ERROR: Download path not detected' -ForegroundColor $White -BackgroundColor $DarkRed
     Write-host 'ERROR: Creating required folder directory $env:USERPROFILE\Downloads\Office365 ProPlus' -ForegroundColor $White -BackgroundColor $DarkRed
     New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\Downloads\Office365 ProPlus"
     Write-Host 'INFO: New folder directory created $env:USERPROFILE\Downloads\Office365 ProPlus' -ForegroundColor $White -BackgroundColor $DarkGreen
     
     # Download Office Telemetry Components
if (Get-Item -Path $SQLFilex64 -ErrorAction $SilentlyContinue)
     {
     # Detected SQL Server Express
     Write-Host "INFO: $SQLFileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not Detected SQL Server Express
     Write-Host "INFO: $SQLFileName has not been detected, starting download"
     Start-BitsTransfer -Source $SQL64 -Destination $SQLFilex64
     #$client.DownloadFile($SQL64,$SQLFilex64)
     Write-Host 'INFO: Download completed for SQL Express Installation' -ForegroundColor $White -BackgroundColor $darkgreen
     Write-Host 'INFO: Downloading SQL Management Installation ' -ForegroundColor $White -BackgroundColor $DarkBlue
     }

if (Get-Item -Path $SQLMFilex64 -ErrorAction $SilentlyContinue)
     {
     # Detected SQL Server Management
     Write-Host "INFO: $SQLMFileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not Detected SQL Server Management
     Write-Host "INFO: $SQLMFileName has not been detected, starting download"
     Start-BitsTransfer -Source $SQLM64 -Destination $SQLMFilex64
     #$client.DownloadFile($SQLM64,$SQLMFilex64)
     Write-Host 'INFO: Download completed for SQL Management' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Downloading SQL Express Installation ' -ForegroundColor $White -BackgroundColor $DarkBlue
     }

if (Get-Item -Path $OTDBFile -ErrorAction $SilentlyContinue)
     {
     # Office Telemetry Dashboard
     Write-Host "INFO: $OTDBFileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not Office Telemetry Dashboard
     Write-Host "INFO: $OTDBFileName has not been detected, starting download"
     Start-BitsTransfer -Source $OTDB -Destination $OTDBFile
     #$client.DownloadFile($OTDB,$OTDBFile)
     Write-Host 'INFO: Download completed for SQL Express Tool' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Downloading Group Policy Telemetry Tags - PowerShell script' -ForegroundColor $White -BackgroundColor $DarkBlue
     }

if (Get-Item -Path $TagsFile -ErrorAction $SilentlyContinue)
     {
     # Tags
     Write-Host "INFO: $TagsFileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not Tags
     Write-Host "INFO: $TagsFileName has not been detected, starting download"
     Start-BitsTransfer -Source $Tags -Destination $TagsFile
     #$client.DownloadFile($Tags,$TagsFile)
     Write-Host 'INFO: Download completed for Group Policy Telemetry Tags - PowerShell script' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Downloading Group Policy Telemetry Tags - VBS File' -ForegroundColor $White -BackgroundColor $DarkBlue
     }

if (Get-Item -Path $VBSFile -ErrorAction $SilentlyContinue)
     {
     # VBS
     Write-Host "INFO: $VBSFileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not VBS
     Write-Host "INFO: $VBSFileName has not been detected, starting download"
     Start-BitsTransfer -Source $VBS -Destination $VBSFile
     #$client.DownloadFile($VBS,$VBSFile)
     Write-Host 'INFO: Download completed for Group Policy Telemetry Tags - VBS file' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Opening Internet Explorer for more information about Telemetry Dashboard' -ForegroundColor $White -BackgroundColor $DarkBlue
     }

if (Get-Item -Path $Net35File -ErrorAction $SilentlyContinue)
     {
     # DotNet
     Write-Host "INFO: $Net35FileName detected, skipping download" -ForegroundColor $White -BackgroundColor $DarkCyan
     }
     else
     {
     # Not DotNet
     Write-Host "INFO: $Net35FileName has not been detected, starting download"
     Start-BitsTransfer -Source $Net35 -Destination $Net35File
     Write-Host 'INFO: Download completed for SQL Express Installation' -ForegroundColor $White -BackgroundColor $darkgreen
     Write-Host 'INFO: Downloading SQL Management Installation ' -ForegroundColor $White -BackgroundColor $DarkBlue
     }

     $ie = New-Object -ComObject InternetExplorer.Application
     $ie.Navigate('https://technet.microsoft.com/en-us/library/jj219431.aspx')
     $ie.Visible = $true
    }
    Get-Root 
    }
  function Get-Root15  { # Download Office Deployment Tool
    $client = new-object -TypeName System.Net.WebClient
    $client.DownloadFile($SourceDeployment,$DeploymentTool)
    # Checking
    
    if (Get-ItemProperty -Path $DeploymentTool)
    {
     # Check if Download was successful
     Write-Host 'INFO: Download completed, returning to Main Menu' -ForegroundColor $White -BackgroundColor $DarkGreen
     Set-Location $DownloadDir
     Start-Process -FilePath $DeploymentTool -Wait 
    } 
    else 
    {
     Write-Host 'ERROR: Download incompleted, returning to Main Menu' -ForegroundColor $White -BackgroundColor $DarkRed
    }
    Get-Root
    }
  function Get-Root16  { # Download Office Deployment with configured XMLs
    Do {
     $MenuOption = Read-Host -Prompt $SubRoot16

     switch ($MenuOption){
        
        1 { # Download Deployment Tool and XMLs
            Get-Root161
          }
        2 { # Office 365 ProPlus x64
            Get-Root162
          } 
        3 { # Office 365 ProPlus x86
            Get-Root163
          }
        4 { # Office 2019 x64
            Get-Root164
          }
        
        98{ # Return to Main Menu
            Get-Root16 
        }
          
        $Quit {exit} 
      }
    
    }  until ($Root -eq {$Quit})
  }
  function Get-Root20  { # Install Office 2016
    Do {
     $MenuOption = Read-Host -Prompt $SubRoot20

     switch ($MenuOption){
        
        1 { # Office 365 ProPlus x64
            Get-Root201
          }
        2 { # Office 365 ProPlus x84
            Get-Root202
          }
        3 { # Office 2019 x64
            Get-Root203
          }
        98{ # Return to Main Menu
            Get-Root 
        }
          
        $Quit {exit} 
      }
    
    }  until ($Root -eq {$Quit})
  }
  function Get-Root21  { # Install SQL Express2

    set-location -Path "$env:USERPROFILE\Downloads\Office365 ProPlus"
    start-process .\SQLEXPR_x64_ENU.exe -Wait
    Write-Host 'Installation now completed' -ForegroundColor $White -BackgroundColor $DarkBlue
    Get-Root 
    }
  function Get-Root22  { # Install SQL Management Studio
    set-location -Path "$env:USERPROFILE\Downloads\Office365 ProPlus"
    start-process .\SQLManagementStudio_x64_ENU.exe -Wait
    Write-Host 'Installation now completed' -ForegroundColor $White -BackgroundColor $DarkBlue
    Get-Root 
    }
  function Get-Root23  { # Install Office Telemetry Dashboard
    if (Get-ItemProperty -Path $Path5 -ErrorAction $SilentlyContinue)
    {
     # Path Exists
     Write-Host '.NET 3.5 Prerequiste met, continuing installation' -ForegroundColor $White -BackgroundColor $DarkGreen
     set-location -Path "$env:USERPROFILE\Downloads\Office365 ProPlus"
     Start-Process -FilePath $TDadmSetup -Wait
     Write-Host 'Installation now completed' -ForegroundColor $White -BackgroundColor $DarkBlue
    } 
    else 
    {
     # Install .NET 3.5
     Write-Host '.NET 3.5 Prerequiste not met, downloading .NET 3.5' -ForegroundColor $White -BackgroundColor $DarkRed
     $client = new-object -TypeName System.Net.WebClient
     $client.DownloadFile($Net35,$Net35File)
     Write-Host '.NET 3.5 now downloaded, installation process now starting' -ForegroundColor $White -BackgroundColor $DarkGreen
     set-location -Path "$env:USERPROFILE\Downloads\Office365 ProPlus"
     Start-Process -FilePath .\dotNetFx35setup.exe -Wait 
     Write-Host '.NET 3.5 installation now completed' -ForegroundColor $White -BackgroundColor $DarkBlue
     # Install Office Telemetry Dashboard
     Write-Host 'Installing Office Telemetry Dashboard' -ForegroundColor $White -BackgroundColor $DarkGreen
     start-process -FilePath $TDadmSetup -Wait
     Write-Host 'Office Telemetry Dashboard installation now completed' -ForegroundColor $White -BackgroundColor $DarkBlue
        }
     Get-Root 
  }
  function Get-Root30  {
     Write-Host "Launching Internet Explorer and visiting http://config.office.com" -ForegroundColor $White -BackgroundColor $DarkGreen
     $ie = New-Object -ComObject InternetExplorer.Application
     $ie.Navigate('https://config.office.com')
     $ie.Visible = $true
     Get-Root 
  }
  function Get-Root31  {
     $SetLocation = Read-Host -Prompt 'Specify Location your modified configuration.xml and setup.exe'
     Set-Location $SetLocation
     Write-Host 'Location Set' $SetLocation -ForegroundColor $White -BackgroundColor $DarkBlue
     Write-Host 'Downloading Office Delta to' $SetLocation -ForegroundColor $White -BackgroundColor $DarkGreen
     .\setup.exe /download configuration.xml
     Write-Host 'Download completed, starting installation Office ProPlus' -ForegroundColor $White -BackgroundColor $DarkGreen
     .\setup.exe /configure configuration.xml
     Get-Root 
  }
  function Get-Root32  {
     Write-Host 'Obtaining configuration information registry' $SetLocation -ForegroundColor $White -BackgroundColor $DarkBlue
     Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration
     Get-Root 
  }
  Function Get-OfficeRemoval {
  # Remove Office Using Troubleshooting Tool
  Set-Location
  if (Get-ItemProperty -Path $FixITFile)
    {
      # FixIT Files Exists
      Write-Host 'INFO: Launching Microsoft FixIT' -ForegroundColor $White -BackgroundColor $DarkGreen
      Start-Process -FilePath $FixITFile -Wait
    } 
    else 
    {
     Write-Host 'ERROR: Microsoft FixIT not detected, starting download' -ForegroundColor $White -BackgroundColor $Red
     $client = new-object -TypeName System.Net.WebClient
     $client.DownloadFile($FixIT,$FixITFile)
     Write-Host 'INFO: Download completed, returning to Main Menu' -ForegroundColor $White -BackgroundColor $DarkGreen
     Start-Process -FilePath $FixITFile -Wait
    }
  }
  Function Get-FileName {  
  [CmdletBinding()]
  param
  (
    [Object]$initialdirectory
  )
  Add-Type -AssemblyName System.windows.forms | Out-Null
    
    $OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialdirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}
 #endregion Functions
 #region Sub Functions
 #region Option 16
  function Get-Root161   { # Download Offscrub
    if ($psversion.Major -gt 4)
    {
    # PowerShell Version Check
    Write-Host 'INFO: PowerShell Version 5 detected' -ForegroundColor $White -BackgroundColor $DarkBlue
    New-Item -ItemType Directory -Force -Path $DownloadDir
    
    # Download from OneDrive
    Write-Output -InputObject 'INFO: Downloading zip file'
    New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\Downloads\Office365 ProPlus"
    #Start-BitsTransfer -Destination $SourceDownload -Source $DestinFile
    $client = new-object -TypeName System.Net.WebClient
    $client.DownloadFile($SourceDownload,$DestinFile)
    
    # Checking
    if (Get-ItemProperty -Path $DestinFile)
    {
      # Check if Download was successful
      Write-Host 'Download Completed' -ForegroundColor $White -BackgroundColor $DarkGreen
      
      # Extract ZIP File
      Write-Host 'Please Note: This next step requires Powershell Version 5. Unzipping package' -ForegroundColor $White -BackgroundColor $DarkBlue
      Expand-Archive -Path "$env:USERPROFILE\Downloads\Office365 ProPlus\OfficeC2R.zip" -DestinationPath "$env:USERPROFILE\Downloads\Office365 ProPlus" -ErrorAction $SilentlyContinue
    } 
    else 
    {
      Write-Host 'ERROR: Download incompleted, returning to Main Menu' -ForegroundColor $White -BackgroundColor $Red
      get-root16  
    }
    }
    else
    {
      # PowerShell Version Check
      Write-Host 'ERROR: PowerShell Version 5 not detected' -ForegroundColor $White -BackgroundColor $Darkred
      New-Item -ItemType Directory -Force -Path $DownloadDir

    # Download from OneDrive
    Write-Output -InputObject 'INFO: Downloading zip file'
    New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\Downloads\Office365 ProPlus"
    #Start-BitsTransfer -Destination $SourceDownload -source $DestinFile
    $client = new-object -TypeName System.Net.WebClient
    $client.DownloadFile($SourceDownload,$DestinFile)
      
    # Checking
      if (Get-ItemProperty -Path $DestinFile)
    {
      # Check if Download was successful
      Write-Host 'INFO: Download completed, manual extraction of zip file required' -ForegroundColor $White -BackgroundColor $DarkGreen
      Get-FileName -initialdirectory "$env:USERPROFILE\downloads\Office365 ProPlus"
      
    } 
    else 
    {
      Write-Host 'ERROR: Download incompleted, returning to Main Menu' -ForegroundColor $White -BackgroundColor $Red 
     Get-root16
    }
    }
      Get-Root
    }  
  function Get-Root162 { # Download configuration-Office365-x64
    Set-Location -Path $DownloadC2R
     Write-Host 'Downloading x64 Deltas for Monthly Channel.. Please wait, while we download the required Deltas' -ForegroundColor $White -BackgroundColor $DarkGreen
    .\setup.exe $download $Monthly64
    Get-Root16
    }
  function Get-Root163 { # Download configuration-Office365-x86
    Set-Location -Path $DownloadC2R
    Write-host 'INFO: Downloading x86 Deltas for Monthly Channel.. Please wait, while we download the required Deltas' -ForegroundColor $White -BackgroundColor $DarkGreen
    .\setup.exe $download $Monthly86
    Get-Root16
    }
  function Get-Root164 { # Download configuration-Office2019Enterprise
    Set-Location -Path $DownloadC2R
    Write-Host 'INFO: Downloading base Office 2019 x64 deltas .. Please wait, while we download the required Deltas' -ForegroundColor $White -BackgroundColor $DarkBlue
    Write-Host 'INFO: The following sample allows you to download and install Office 2019 Professional Plus,Visio 2019 Professional, and Project 2019 Professional directly from the Office CDN.' -ForegroundColor $white -BackgroundColor $DarkCyan
    .\setup.exe $download $Office2019x64
    Get-Root16
    }
  #endregion Option 16
 #region Option 20
  function Get-Root201 { # Install 2016 Monthly Channel (x64)
    if (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\RegisteredApplications' -ErrorAction $SilentlyContinue)
    {
     # Check if Download was successful
     Write-Host 'ERROR: An Office 2016 installation already exists, returning to main menu' -ForegroundColor $White -BackgroundColor $DarkRed
    } 
    else 
    {
     Write-Host 'INFO: No Office Applications detected, installing Office 2016'
     Set-location -Path $DownloadC2R
     Write-Host 'INFO: Setting location of setup.exe and XML file' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Configuring your Office 2016 installation now' -ForegroundColor $White -BackgroundColor $DarkGreen
     .\setup.exe /configure $Monthly64 
    }
    Get-Root20
    }
  function Get-Root202 { # Install 2016 Monthly Channel (x86)
    if (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\RegisteredApplications' -ErrorAction $SilentlyContinue)
    {
     # Check if Download was successful
     Write-Host 'ERROR: An Office 2016 installation already exists, returning to main menu' -ForegroundColor $White -BackgroundColor $DarkRed
    } 
    else 
    {
     Write-Host 'INFO: No Office Applications detected, installing Office 2016'
     Set-location -Path $DownloadC2R
     Write-Host 'INFO: Setting location of setup.exe and XML file' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Configuring your Office 2016 installation now' -ForegroundColor $White -BackgroundColor $DarkGreen
     .\setup.exe /configure $Monthly86 
    }
    Get-Root20
    }
  function Get-Root203 { # Install 2019
   
    if ($exists) 
    {
    Write-Host 'INFO: An installation of Office has not been detected' -ForegroundColor $White -BackgroundColor $DarkRed 
    Write-Host 'INFO: Installing Office 2019 with removal of Office MSI version' -ForegroundColor $White -BackgroundColor $DarkCyan
    }
    else 
    {
     Write-Host 'ERROR: An installation of Office has been detected' -ForegroundColor $White -BackgroundColor $DarkRed 
     Write-Host 'INFO: Installing Office 2019 with removal of Office MSI version' -ForegroundColor $White -BackgroundColor $DarkCyan
     Set-location -Path $DownloadC2R
     Write-Host 'INFO: Setting location of setup.exe and XML file' -ForegroundColor $White -BackgroundColor $DarkGreen
     Write-Host 'INFO: Configuring your Office 2019 installation now' -ForegroundColor $White -BackgroundColor $DarkGreen
     .\setup.exe /configure $Office2019x64
    }
    Get-Root20
    }
  
  #endregion Option 20
 #endregion Sub Functions
 #region Script Block
   clear-host
   Write-host 'Version information - You are running script version 1.1' -ForegroundColor $White -BackgroundColor $DarkGray
   Test-IsAdmin
   Root
 #endregion Script Block

