 <#Information
     Skype for Business - AlwaysOn Database Configurator

     .DESCRIPTION
     Tool to assist with the deployment of Skype for Business Server 2015/2019

     .NOTES
     Version				: 1.0
     Wish list			    : Better error trapping
                            : 
     Rights Required	    : Local administrator on Workstation
                            : 
                            : Powershell in Administrator Mode
     Sched Task Required	: No
     Windows Server Version	: 
     Author/Copyright	    : © Andrew J. Price - All Rights Reserved
     Email/Blog/Twitter	    : 
     Dedicated Post	        : 
 
     Disclaimer             : You running this script means you won't blame me if this breaks your stuff. This script is provided AS IS without warranty of any kind. 
                              I disclaim all implied warranties including,
                              without limitation, any implied warranties of merchantability or of fitness
                              .SYNOPSIS for a particular
                              purpose. The entire risk arising out of the use or performance of the sample scripts and
                              documentation remains with you. In no event shall I be liable for any damages whatsoever
                              (including, without limitation, damages for loss of business profits, business interruption,
                              loss of business information, or other pecuniary loss) arising out of the use of or inability
                              to use the script or documentation.

     Acknowledgements 	        : Andrew Offor
                                : 
                                :
     Assumptions		        : ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
                                :  
     Limitations			    : 
     Known issues				:   
     Version Changes            : 0.1 Initial Script Build
                                : 0.9 Test Build
                                : 1.0 Initial Build

     .LINK
     

     .EXAMPLE
     .\

     Description
     -----------
     Runs script with default values.


     .INPUTS
     None. You cannot pipe objects to this script.
 #>
#region Shortnames
  
  # Colours
  $White = 'White'
  $Red = 'Red'
  $Green = 'Green'
  $Cyan = 'Cyan'
  $DarkGreen = 'DarkGreen'
  $DarkBlue = 'DarkBlue'
  $DarkGray = 'DarkGray'
  $DarkRed = 'DarkRed' 
  
  # Generic 
  $SilentlyContinue = 'SilentlyContinue'
  $Space = '' 

 #endregion Shortnames
#region Menus
[string] $Root = @'
 
  ┌─────────────────────────────────────────────────────────────┐
   Skype for Business 2015/2019 - AlwaysOn Database Configurator              
           
                Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘
  This script is designed to run on your Primary SQL Server to 
  assist with Skype for Business installations  with SQL 
  Enterprise 2014 or 2016 with AlwaysOn Functionality.

  Steps:
  Check if an installation of SQL Server is present
  Install, Validate and Configure Failover Cluster
  Identify and Configure your Skype for Business Databases
  Create folder directories on Secondary Node

'@
[string] $ListofReasons = @'
 
  Potential reasons for failure;
  - An installation of SQL Server 2014 or 2016 is missing
  - Windows Firewall preventing check
  - SQL Server maybe shutdown
  - FFS (Fat Finger Syndrome) on computer name

'@
[string] $FailedDatabases = @'
 
  Potential reasons for failure;
  - Skype for Business Database are not deployed to SQL Server
  - Access Denied to SQL Instance


'@
[string] $ListofReasons = @'
 
  Potential reasons for failure;
  - Just didnt work - Install manually

'@
[string] $Step2 = @'
 
  List of Actions;
  - Check, Modify, Backup all Skype for Business Databases found.
  - Report if any Databases are not found
 

'@
[string] $Step3 = @'
 
  List of Actions;
  - Specify the folder locations for your Database, Logs and Backup on Primary SQL Server
  - Specify the folder locations for your Database, Logs and Backup on Secondary SQL Server
  - Report if any Databases are not found

'@
[string] $FinishLine = @'
 
  List of Actions;
  - List all things that need to be completed after this script
  - Configure AlwaysOn High Availability on all SQL Server (InstanceName) Service.
  - Restart SQL Server (Instance Name)
  - Configure SQL Manangement Studio with Always On Availbility Groups and Witness.

'@
[string] $GetAlwaysOn = @'
 
  List of Actions;
  - Launching SQL Server Configuration Mananger
  - Configure SQL Server (Instance Name) under SQL Server Service for AlwaysOn High Availability
  - Restart Service
  - Manually repeat above steps for your Secondary SQL Server Node

'@

 #endregion
#region Functions
# Test for Administrator Privilges 
function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')
    
    if (!(Test-IsAdmin)){
    throw 'Please Note: You are trying to run this script without evalated Administator Priviliges. In order to run this script you will required Powershell running in Administrator Mode'
    }
    else {
    Write-Verbose -Message 'Are you running as evalated Administator' -verbose
    }
    }
# Step 1 - Detect if Failover Clustering Installed or Configure
Function Get-SQLService {
   
   # Call for Servers
   $SQLServers = $PrimarySQLNode, $SecondarySQLNode

   # Check if Failover Clustering is installed
   foreach ($SQLService in $SQLServers)  
   {
   # Get SQL Service
   $serviceName = 'SQLBrowser'
   If (Get-Service -ComputerName $SQLService -Name $serviceName -ErrorAction $SilentlyContinue) {
   If ((Get-Service -Name $serviceName).Status -eq 'Running') {
   Write-Host ('Detected that {0} exists on {1}.{2} SQL Server Node' -f $serviceName, $SQLService, $FQDN) -ForegroundColor $Green
   } Else {    
   Write-Host ('Deteched that {0} exist on {1}.{2}, howvever it is not running.' -f $serviceName, $SQLService, $FQDN) -ForegroundColor $Green
   }
   } Else {
   Write-Host ('Unable to detected an installation of SQL Server on {0}.{1} by checking for {2} service, please investigate before continuing' -f $SQLService, $FQDN, $serviceName) -ForegroundColor $Red
   $ListofReasons  
   Write-Host 'Exiting now script'
   exit
   }
   }
   }
Function Get-Step1 {
   # SQL Servers
   $SQLServers = $PrimarySQLNode, $SecondarySQLNode

   ForEach ($computername in $SQLServers) {
   If (
   Get-WindowsFeature -Name 'Failover-Clustering' -ComputerName $computername | Where-Object {$_.Installed -eq $True}  )
   {
      Write-Host ('Detected that Failover Clustering exists on {0}.{1} SQL Server Node' -f $computername, $FQDN) -ForegroundColor $Green
   } Else {
   # Install Failover Clustering and Configure Role
   Write-Host ('Unable to detect an installation of Failover Clustering on {0}.{1}, installing feature' -f $computername, $fqdn) -ForegroundColor $Red
   Install-WindowsFeature -Name Failover-Clustering, RSAT-Clustering -IncludeManagementTools -ComputerName $computername
   
   }
   }
    Repeat-FailoverCluster
   }       
# Step 2 - Modify Databases for SQL AlwaysOn
Function Get-Step2 {
   # Configure Skype for Business Databases
   $Step2
   
   Import-Module -Name sqlps -DisableNameChecking -ErrorAction SilentlyContinue
   Write-Host -Foreground $Green ======== ('Please wait! Detecting Skype for Business Databases located on {0}.{1}' -f $PrimarySQLNode, $FQDN) ======== 

   # Back End databases
   $BEDB = 'cpsdyn','rgsconfig','rgsdyn','rtcab','rtcshared','rtcxds'
   $SvrInst = ('{0}\{1}' -f $PrimarySQLNode, $instance)
   Foreach ($DB1 in $BEDB) {
   Try {
   Invoke-Sqlcmd -Query ('ALTER DATABASE {0} SET RECOVERY FULL WITH NO_WAIT;' -f $DB1) -ServerInstance $SvrInst -ErrorAction Stop
   ; Backup-SqlDatabase -ServerInstance $SvrInst -Database $DB1 -ErrorAction $SilentlyContinue
   Write-Host ("Detected {0} - This database has now been changed to 'Full Recovery Mode and backed up'" -f $DB1) -ForegroundColor $Green
   }
   Catch {
   Write-Host ('Unabled to detect {0} - Skipping' -f $DB1) -ForegroundColor $Red
   }
   } 
   
   # CMS Databases
   $CMSDB = 'xds','lis'
   $SvrInst = ('{0}\{1}' -f $PrimarySQLNode, $instance)
   Foreach ($DB2 in $CMSDB) {
   Try {
   Invoke-Sqlcmd -Query ('ALTER DATABASE {0} SET RECOVERY FULL WITH NO_WAIT;' -f $DB2) -ServerInstance $SvrInst -ErrorAction Stop
   ; Backup-SqlDatabase -ServerInstance $SvrInst -Database $DB2 -ErrorAction $SilentlyContinue
   Write-Host ("Detected {0} - This database has now been changed to 'Full Recovery Mode and backed up'" -f $DB2) -ForegroundColor $Green
   }
   Catch {
   Write-Host ('Unabled to detect {0} - Skipping' -f $DB2) -ForegroundColor $Red
   }
   } 
   
   # Monitoring Databases
   $MDDB = 'LcsCDR','QoEMetrics'
   $SvrInst = ('{0}\{1}' -f $PrimarySQLNode, $instance)
   Foreach ($DB3 in $MDDB) {
   Try {
   Invoke-Sqlcmd -Query ('ALTER DATABASE {0} SET RECOVERY FULL WITH NO_WAIT;' -f $DB3) -ServerInstance $SvrInst -ErrorAction Stop
   ; Backup-SqlDatabase -ServerInstance $SvrInst -Database $DB3 -ErrorAction $SilentlyContinue
   Write-Host ("Detected {0} - This database has now been changed to 'Full Recovery Mode and backed up'" -f $DB3) -ForegroundColor $Green
   }
   Catch {
   Write-Host ('Unabled to detect {0} - Skipping' -f $DB3) -ForegroundColor $Red
   }
   } 
   
   # Archiving Database
   $ADDB = 'LcsLog'
   $SvrInst = ('{0}\{1}' -f $PrimarySQLNode, $instance)
   Foreach ($DB4 in $ADDB) {
   Try {
   Invoke-Sqlcmd -Query ('ALTER DATABASE {0} SET RECOVERY FULL WITH NO_WAIT;' -f $DB4) -ServerInstance $SvrInst -ErrorAction Stop
   ; Backup-SqlDatabase -ServerInstance $SvrInst -Database $DB4 -ErrorAction $SilentlyContinue
   Write-Host ("Detected {0} - This database has now been changed to 'Full Recovery Mode and backed up'" -f $DB4) -ForegroundColor $Green
   }
   Catch {
   Write-Host ('Unabled to detect {0} - Skipping' -f $DB4) -ForegroundColor $Red
   }
   }
   Get-Databases 
   }
Function Get-Databases {
   $title = 'Skype for Business Database'
   $Message = 'Did we fail detecting all your deployed Skype for Business Databases?'
   $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes',`
   'Launching Failover Clustering MMC .'
   $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No',`
   'Skipping..'
   $options = [management.automation.host.choicedescription[]]($yes, $no)
   $result = $host.ui.PromptForChoice($title,$Message,$options,1)
   Switch ($result)
   {
    0 {Write-Host $FailedDatabases 
       Exit}
    1 {Step3}
   }
   }
Function Get-FailoverCluster {
   # Launch Failover Clustering
   Write-Host 'Launching Failover Cluster Manager' -ForegroundColor $Green
   & "$env:windir\system32\mmc.exe" Cluadmin.msc
   Repeat-FailoverCluster-Completed
   }
Function Repeat-FailoverCluster {
   $title = 'SQL Failover Clustering'
   $Message = 'Do you need to configure your Failover Clustering Role?'
   $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes',`
   'Launching Failover Clustering MMC .'
   $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No',`
   'Skipping..'
   $options = [management.automation.host.choicedescription[]]($yes, $no)
   $result = $host.ui.PromptForChoice($title,$Message,$options,1)
   Switch ($result)
   {
    0 {Get-FailoverCluster}
    1 {Step2}
   }
   }
Function Repeat-FailoverCluster-Completed {
   $title = 'SQL Failover Clustering'
   $Message = 'Did you complete the configuration your Failover Clustering Role?'
   $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes',`
   'Launching Failover Clustering MMC .'
   $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No',`
   'Skipping..'
   $options = [management.automation.host.choicedescription[]]($yes, $no)
   $result = $host.ui.PromptForChoice($title,$Message,$options,1)
   Switch ($result)
   {
    0 {Step2}
    1 {Get-FailoverCluster}
   }
   }
# Step 3
Function Get-Step3 {
     
   $Step3
   # Specify Source SQL Database 
   $DatabasePath1 = Read-Host -Prompt 'Specify your SQL Database Drive Location on your Primary SQL Server for example (e$ or e$\csdata)'
   Write-Host -Foreground $Cyan ('You have specified {0}\{1} as your database path' -f $PrimarySQLNode, $DatabasePath1)
   $LogPath1 = Read-Host -Prompt 'Specify your SQL Log Drive Location Primary SQL Server for example (e$ or e$\csdata)'
   Write-Host -Foreground $Cyan ('You have specified {0}\{1} as your log path' -f $PrimarySQLNode, $LogPath1)
   
   # Target SQL Database
   $DatabasePath2 = Read-Host -Prompt 'Specify your SQL Database Drive Location on your Secondary SQL Server for example (e$ or e$\csdata)'
   Write-Host -Foreground $Cyan ('You have specified {0}\{1} as your database path' -f $SecondarySQLNode, $DatabasePath2)
   $LogPath2 = Read-Host -Prompt 'Specify your SQL Log Drive Location Secondary SQL Server for example (e$ or e$\csdata)'
   Write-Host -Foreground $Cyan ('You have specified {0}\{1} as your log path' -f $SecondarySQLNode, $LogPath2)

   # Defined Skype for Business Folder Names
   $DBFolder = 'ABSStore','ApplicationStore','ArchivingStore','BackendStore','CentralMgmtStore','MonitoringStore'
   $LogFolder = 'ABSStore','ApplicationStore','ArchivingStore','BackendStore','CentralMgmtStore','MonitoringStore'

   Write-Host "Depending on your Deployment you might see 'Folder Not Found' this is common and error output for informational purposes " -Foreground $Cyan
   Write-Host -Foreground $Cyan 'Copying from Primary SQL Server Node'

   # Copy Database
   Foreach ($RC1 in $DBFolder) {
     if (Get-ItemProperty -Path  ('\\{0}\{1}\{2}' -f $PrimarySQLNode, $DatabasePath1, $RC1) -ErrorAction $SilentlyContinue) 
       {
       & "$env:windir\system32\robocopy.exe" ('\\{0}\{1}\{2}' -f $PrimarySQLNode, $DatabasePath1, $RC1) ('\\{0}\{1}\{2}' -f $SecondarySQLNode, $DatabasePath2, $RC1) /MIR /njh /njs /tee /ZB /W:1 /NOCOPY
       Write-Host ('Detected {0} - Copying' -f $RC1) -ForegroundColor $Green
       }
       Else {
       Write-Host ('Unable to detect {0} Folder - Skipping' -f $RC1) -ForegroundColor $Red
       }
       }
   
   Write-Host $Space 
 
      
  # Copy Log
  Foreach ($RC2 in $LogFolder){
    if (Get-ItemProperty -Path  ('\\{0}\{1}\{2}' -f $PrimarySQLNode, $LogPath1, $RC2) -ErrorAction $SilentlyContinue)
      {
      & "$env:windir\system32\robocopy.exe" ('\\{0}\{1}\{2}' -f $PrimarySQLNode, $LogPath1, $RC2) ('\\{0}\{1}\{2}' -f $SecondarySQLNode, $LogPath2, $RC2) /MIR /njh /njs /tee /ZB /W:1 /NOCOPY
      Write-Host ('Detected {0} - Copying' -f $RC2) -ForegroundColor $Green
      }
      Else {
      Write-Host ('Unable to detect {0} Folder - Skipping' -f $RC2) -ForegroundColor $Red
      }
      }
    Repeat-FolderCopy
   }
Function Get-Step31 {
     
   $Step31
   # Specify Additional SQL Server
   $AdditionalSQLNode = Read-Host -Prompt 'Please Enter your Additional SQL Server Name in none FQDN format i.e DB03'
   Write-Host -Foreground DarkCyan ('{0}.{1} has been specified your SQL Server Name' -f $AdditionalSQLNode, $FQDN) 
   # Specify Source SQL Database 
   $DatabasePath1 = Read-Host -Prompt 'Specify your SQL Database Drive Location on your Primary SQL Server for example (e$ or e$\csdata)'
   Write-Host -Foreground $Cyan ('You have specified {0}\{1} as your database path' -f $PrimarySQLNode, $DatabasePath1)
   $LogPath1 = Read-Host -Prompt 'Specify your SQL Log Drive Location Primary SQL Server for example (e$ or e$\csdata)'
   Write-Host -Foreground $Cyan ('You have specified {0}\{1} as your log path' -f $PrimarySQLNode, $LogPath1)
   
   # Target SQL Database
   $DatabasePath2 = Read-Host -Prompt 'Specify your SQL Database Drive Location on your Secondary SQL Server for example (e$ or e$\csdata)'
   Write-Host -Foreground $Cyan ('You have specified {0}\{1} as your database path' -f $AdditionalSQLNode, $DatabasePath2)
   $LogPath2 = Read-Host -Prompt 'Specify your SQL Log Drive Location Secondary SQL Server for example (e$ or e$\csdata)'
   Write-Host -Foreground $Cyan ('You have specified {0}\{1} as your log path' -f $AdditionalSQLNode, $LogPath2)

   # Defined Skype for Business Folder Names
   $DBFolder = 'ABSStore','ApplicationStore','ArchivingStore','BackendStore','CentralMgmtStore','MonitoringStore'
   $LogFolder = 'ABSStore','ApplicationStore','ArchivingStore','BackendStore','CentralMgmtStore','MonitoringStore'

   Write-Host -Foreground $Green 'Starting Database Location Folders'
    # Copy Database
   Foreach ($RC1 in $DBFolder) {
   Try {
       & "$env:windir\system32\robocopy.exe" ('\\{0}\{1}\{2}' -f $PrimarySQLNode, $DatabasePath1, $RC1) ('\\{0}\{1}\{2}' -f $AdditionalSQLNode, $DatabasePath2, $RC1) /MIR /njh /njs /ndl /tee /ZB /W:1 /NOCOPY
       Write-Host ('{0} Recovery Set' -f $RC1) -ForegroundColor White -BackgroundColor Green
       }
       Catch {
       Write-Host ('{0} Folder Not Found - continuing' -f $RC1) -ForegroundColor White -BackgroundColor Red
       }
       }
   
   Write-Host $Space 
   Write-Host -Foreground $Green 'Starting Log Location Folders'  
  # Copy Log
  Foreach ($RC1 in $LogFolder){
  Try {
      & "$env:windir\system32\robocopy.exe" ('\\{0}\{1}\{2} \\{3}\{4}\{5}' -f $PrimarySQLNode, $LogPath1, $RC1, $AdditionalSQLNode, $LogPath2) /MIR /ZB /njh /njs /ndl /tee /W:1 /NOCOPY
      Write-Host ('{0} Recovery Set' -f $RC1) -ForegroundColor White -BackgroundColor Green
      }
      Catch {
      Write-Host ('{0} Folder Not Found - continuing' -f $RC1) -ForegroundColor White -BackgroundColor Red
      }
      }
    Repeat-FolderCopy
   }
# Step 4 - Detect if SQL Server been configured for AlwaysOn and if AlwaysOn Availabilty Group present
Function Repeat-AlwaysOn {
   $title = 'SQL AlwaysOn Availability Group'
   $Message = 'Do you need to configure your SQL Server for AlwaysOn Availability Group?'
   $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes',`
   'Launching SQL AlwaysOn .'
   $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No',`
   'Skipping..'
   $options = [management.automation.host.choicedescription[]]($yes, $no)
   $result = $host.ui.PromptForChoice($title,$Message,$options,1)
   Switch ($result)
   {
    0 {Get-AlwaysOn}
    1 {Get-FinishLine}
   }
   }
Function Get-AlwaysOn {
$GetAlwaysOn
if (Test-Path C:\Windows\SysWOW64\SQLServerManager12.msc) {
   Write-Host -ForegroundColor $Green 'Info: SQL Server 2014 Management Studio detected'
   & "$env:windir\system32\mmc.exe" /32 $env:windir\SysWOW64\SQLServerManager12.msc
   }
   else {
   Write-Host -ForegroundColor $Green 'Info: SQL Server 2016 Management Studio detected'
   & "$env:windir\system32\mmc.exe" /32 $env:windir\SysWOW64\SQLServerManager13.msc
   }
   Get-FinishLine
   }
Function Repeat-FolderCopy {
   $title = 'SQL'
   $Message = 'Do you need to update additional SQL replicas?'
   $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes',`
   'Updates additional SQL replicas.'
   $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No',`
   'No additonal SQL replicas to update'
   $options = [management.automation.host.choicedescription[]]($yes, $no)
   $result = $host.ui.PromptForChoice($title,$Message,$options,1)
   Switch ($result)
   {
    0 {Step31}
    1 {Repeat-AlwaysOn}
   }
   }
Function Get-FinishLine {
   $FinishLine
   
Exit
}    
#endregion Functions
#region Script Block
#Test-IsAdmin
Write-host 'Version information - You are running script version 1.0' -ForegroundColor $White -BackgroundColor $DarkGray
$Root

   # Obtain FQDN
   $FQDN=(Get-WmiObject -Class win32_computersystem).Domain
   # Prompt for SQL Server Names
   $PrimarySQLNode = Read-Host -Prompt 'Please Enter your Primary SQL Server Name in none FQDN format i.e DB01'
   Write-Host -Foreground $Green ('{0}.{1} has been specified as your SQL Server Name' -f $PrimarySQLNode, $FQDN) 
   $SecondarySQLNode = Read-Host -Prompt 'Please Enter your Primary SQL Server Name in none FQDN format i.e DB02'
   Write-Host -Foreground $Green ('{0}.{1} has been specified as your SQL Server Name' -f $SecondarySQLNode, $FQDN)
   $Instance = Read-Host -Prompt 'Enter your Primary SQL Server Instance Name'
   Write-Host -Foreground $Green ('{0} has been specified for {1}.{2}' -f $Instance, $PrimarySQLNode, $FQDN)
   Write-Host ''

Get-SQLService
Get-Step1
#endregion Script Block

#region Future Updates
#Function Get-CheckDBs {
#$Output = @()
# (Invoke-SQLCMD -Query 'SELECT * FROM sysdatabases WHERE dbid > 4') | ForEach-Object { # skipping first 4 databases: master, tempdb, model, msdb
#    $DBProps = Invoke-Sqlcmd -Query ("SELECT * FROM sys.databases WHERE name = '{0}'" -f $_.name) 
#    $Output += New-Object -TypeName PSObject -Property ([Ordered]@{
#        ServerName    = $PrimarySQLNode
#        DatabaseName  = $_.name
#        RecoveryModel = $DBProps.recovery_model_desc
#    })
#}
#$Output | Sort-Object -Property RecoveryModel | Select-Object -Property *
#}
#endregion Future Updates

