Clear-Host
 <#Information
     Gather information about Exchange On-Premises Infrastructure.

     .NOTES
     Version				      : 1.2
     Wish list			      : Better error trapping
     Rights Required	    : Local administrator on Workstation
                          : Powershell in Administrator Mode
     Sched Task Required	: No
     Author/Copyright	    : © The Watcher - All Rights Reserved
     Email/Blog/Twitter	  : thewatcher@blogabout.cloud @thewatchernode
     Acknowledgements     : 
     Disclaimer           : You running this script means you won't blame me if this breaks your stuff. This script is provided AS IS without warranty of any kind. 
                            I disclaim all implied warranties including,
                            without limitation, any implied warranties of merchantability or of fitness
                            .SYNOPSIS for a particular
                            purpose. The entire risk arising out of the use or performance of the sample scripts and
                            documentation remains with you. In no event shall I be liable for any damages whatsoever
                            (including, without limitation, damages for loss of business profits, business interruption,
                            loss of business information, or other pecuniary loss) arising out of the use of or inability
                            to use the script or documentation.

     Acknowledgements 	  : 
     Assumptions		      : ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended) 
     Limitations			    : 
     Known issues			    :
     Version Changes      : 0.1 Initial Script Build
                          : 1.0 Public Release
                          : 1.1 Minor scripting updates
                          : 1.2 Minor scripting updates
     
     .EXAMPLE
     .\get-exchangediscovery.ps1

     Description
     -----------
     Runs script with default values.


     .INPUTS
     None. You cannot pipe objects to this script.
 #>
#region Variables

# Colours
$DarkCyan = 'DarkCyan'
$DarkGray = 'DarkGray'
$Red = 'Red'
$Yellow = 'Yellow'
$White = 'White'

# Defined Variables
$Unlimited = 'unlimited'
$NoteProperty = 'NoteProperty'
$Quit = 'Q'
$ExPath = "$env:HOMEDRIVE\ExchangeDiscovery"
$PSV = $PSVersionTable.PSVersion
#endregion
#region Banner
[string] $Root = @'
 
  ┌─────────────────────────────────────────────────────────────┐
        Exchange Discovery Script for Office 365 Migrations             
           
              Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

  1)  Import Exchange Server 2007 PowerShell Module             -->
  2)  Import Exchange Server 2010 PowerShell Module             -->
  3)  Import Exchange Server 2013 -- > PowerShell Module        -->

  5)  Get Exchange on-premise Mailbox Data                      -->
  6)  Get Exchange Server Configuration                         -->
  7)  Get Exchange Server Hybrid Configuration                  -->
     
  Q) Quit

  Select an option.. [1-99]?
'@
[string] $Count = @'
 
  ┌─────────────────────────────────────────────────────────────┐
        Exchange Discovery Script for Office 365 Migrations               
           
              Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

   This section will gather the following information

   - Count of User Mailboxes
   - Count of Shared Mailboxes
   - Gather report of Permissions
   - Count of Room Mailboxes
   - Gather report of Permissions
   - Count of Equipments Mailboxes
   - Gather report of Permissions
   - Gather report of Mailboxes with Forwarding enabled
   - Gather report of Mailboxes with multiple alias
   - Gather report of Disabled Users with Mailboxes
   - Gather report of Distributions Lists
'@
[string] $ServerInfo = @'
 
  ┌─────────────────────────────────────────────────────────────┐
                       
           
              Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

   This section will gather the following information

   - Check and Get Database Availabilty Group Info if any or all EDBs are DAG Protected
   - Gather Mailbox and Public Folder Database information
   - Gather Send and Receiver Connector information
   - Gather Email Address Policy information
   - Gather Virtual Directories information
   - Gather Client Access Server information
   - Gather Outlook Anywhere information
'@
[string] $HybridInfo = @'
 
  ┌─────────────────────────────────────────────────────────────┐
              Hybrid Configuration Configuration               
           
              Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

   This section will gather the following information

   - Gather Exchange Hybrid Configuration
   - Gather Exchange Intra OrganizationConnector
   - Connect to Azure Cloud Shell
        - Gather Exchange Mailflow configuration

'@
[string] $HybridInfo = @'
 
  ┌─────────────────────────────────────────────────────────────┐
                        WARNING !!!!                    
  └─────────────────────────────────────────────────────────────┘

   In order to gather information about Exchange Hybrid Flow we
   will need to connect to Exchange Online using the Azure AZ
   module. If you have AzureRM installed and running PowerShell
   Version 5.1, we will need to install PowerShell Core Version 6

   If you would like to skip this option, press Yes.

'@
#endregion Banner
#region Functions
  function Get-Root    {
  # Menu Prompt
  
    Do {
      $MenuOption = Read-Host -Prompt $Root
      #Clear-Host
      switch ($MenuOption){
        1 { # Option 1
             Get-EXModule2007
          } 
        2 { # Option 2
             Get-EXModule2010
          }
        3 { # Not in use
             Get-EXModuleAny
          }
        4 { # Not in use
             Get-PSVersion
          }
        5 { # Option 5
             Get-Counts
          }
        6 { # Option 6
            Get-ServerInfo
          }
        7 { # Option 7
            Get-HybridHealth
          }  
        
        $Quit {exit} 
      }
    
    }  until ($Root -eq {$Quit})
  }
  Function Get-ExFolder {
    
    if (Test-Path -Path $ExPath)
    {
      Write-Host 'Info: Detected that',$ExPath,'directory exist' -BackgroundColor DarkCyan
    }
    else
    {
      Write-Host 'Warning: Detected that,'$ExPath,"directory doesn't exists, creating folder location" -ForegroundColor Yellow
      New-Item -ItemType Directory -Force -Path $ExPath
    }
  }
  Function Get-WaitISE {
  	
  $Shell = New-Object -ComObject 'WScript.Shell'
  $Button = $Shell.Popup('Click OK to continue.', 0, 'Waiting...', 0)
  }
  Function Get-PSVersion {
  
  if ($PSV.Major -gt 5){
  Write-Host "Info: Detected your PowerShell Version is",$PSV -BackgroundColor $DarkCyan
  }
  else
  {
  Write-Host "Warning: Detected your PowerShell Version is less than the required version, your current installed version is",$PSV -BackgroundColor $Red
  Write-Host "Returning to Main Menu" -ForegroundColor $Yellow
  Get-Root
  }
  }
  Function Get-PSCore6 {
  
  iex "& { $(irm https://aka.ms/install-powershell.ps1) } -UseMSI"
  Get-root
  }
  Function Connect-AzureExchange {}
# Option 1-3
  Function Get-EXModule2007 {
    Write-host 'Info: Importing Exchange PowerShell Module for 2007' -BackgroundColor $DarkCyan
    Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
 
    Get-Root
  }
  Function Get-EXModule2010 {
    Write-host 'Info: Importing Exchange PowerShell Module for 2010' -BackgroundColor $DarkCyan
    Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010

    Get-Root
  }
  Function Get-EXModuleAny {
    Write-host 'Info: Importing Exchange PowerShell Module for 2013 or Greater' -BackgroundColor $DarkCyan
    Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn
 
    Get-Root
  }
# Option 5
  Function Get-MailboxPermissionReport {
 
}
  Function Get-SharedMailboxReport {
  Write-Host ('Info: Downloading Shared Mailboxes Permissions Report to {0}' -f $ExPath) -BackgroundColor $DarkCyan
  Get-Mailbox -RecipientTypeDetails SharedMailbox | Get-MailboxPermission | Select-Object -Property Identity,User,@{Name='Access Rights'
  Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -Path $ExPath\SharedMailbox_Permissions.csv -NoTypeInformation
}
  Function Get-RoomMailboxReport {
  Write-Host ('Info: Downloading Room Mailboxes Permissions Report to {0}' -f $ExPath) -BackgroundColor $DarkCyan
  Get-Mailbox -RecipientTypeDetails RoomMailbox | Get-MailboxPermission | Select-Object -Property Identity,User,@{Name='Access Rights'
  Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -Path $ExPath\RoomMailbox_Permissions.csv -NoTypeInformation
}
  Function Get-EquipmentMailboxReport {
  Write-Host ('Info: Downloading Equipment Mailboxes Permissions Report to {0}' -f $ExPath) -BackgroundColor $DarkCyan
  Get-Mailbox -RecipientTypeDetails EquipmentMailbox | Get-MailboxPermission | Select-Object -Property Identity,User,@{Name='Access Rights'
  Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -Path $ExPath\EquipmentMailbox_Permissions.csv -NoTypeInformation
}
  Function Get-MailboxForwarding {
  Write-Host ('Info: Downloading Mailboxes Forwarding Permissions Report to {0}' -f $ExPath) -BackgroundColor $DarkCyan
  Get-Mailbox -ResultSize Unlimited| Where-Object {$_.ForwardingAddress -ne $null} | Select-Object -Property Name, ForwardingAddress, DeliverToMailboxAndForward | Export-CSV -Path $ExPath\MailboxForwardingReport.csv
}
  Function Get-MailboxAlias {
  Write-Host ('Info: Downloading Mailboxes Alias Report to {0}' -f $ExPath) -BackgroundColor $DarkCyan
  $Mailboxes = Get-Recipient -ResultSize Unlimited -RecipientType UserMailbox |
  Sort-Object -Property @{ Expression = { $_.EmailAddresses.Count } } -Descending

  $Results = foreach( $Mailbox in $Mailboxes ){

    $Stats = $Mailbox | Get-MailboxStatistics
    $Properties = @{

        FirstName          = $Mailbox.FirstName
        LastName           = $Mailbox.LastName
        DisplayName        = $Mailbox.DisplayName
        TotalItemSize      = $Stats.TotalItemSize
        PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
        }
    $AltAddresses = $Mailbox.EmailAddresses |
    Where-Object { $_ -match '^smtp:' -and $_ -ne $Mailbox.PrimarySmtpAddress }
    $i = 1
    foreach( $Address in $AltAddresses ){
        $Properties.Add( ('AltAddress{0}' -f $i), $Address -replace '^smtp:' )
        $i++
        }
    New-Object -TypeName PSObject -Property $Properties
    }

  $Results | Export-Csv -Path $ExPath\MailboxAliasReport.csv -NoTypeInformation

}
  Function Get-DisabledUserwithMailbox {
  Write-Host ('Info: Downloading Disabled User Mailboxes Report to {0}' -f $ExPath)
  $Mailboxes = Get-Mailbox | Where-Object {$_.RecipientTypeDetails -eq 'UserMailbox'}
  $Disabled = @()

  Foreach ($Mailbox in $Mailboxes) {
    if((Get-msolUser -id $Mailbox.samaccountname).Enabled -eq $False){
        $Disabled += Get-MailboxStatistics $Mailbox.samaccountname | Select-Object -Property DisplayName,TotalItemSize
    }    
  }
  $Disabled | Export-Csv -Path $ExPath\DisabledUserwithMailbox.csv -NoTypeInformation
  }
  Function Get-DistributionGroup {
   
    $dist = @(Get-DistributionGroup -resultsize unlimited)
    $reportoutput = @()

    # Report on Distribution List
    foreach ($dl in $dist)
    {
      $count =@(Get-DistributionGroup $dl.samaccountname).count
      $report = New-Object -TypeName PSObject
      $report | Add-Member -MemberType $NoteProperty -Name 'Group Name' -Value $dl.Name
      $report | Add-Member -MemberType $NoteProperty -Name 'samAccountname' -Value $dl.samaacountname
      $report | Add-Member -MemberType $NoteProperty -Name 'Group Type' -Value $dl.grouptype
      $report | Add-Member -MemberType $NoteProperty -Name 'DN' -Value $dl.distinguishedName
      $report | Add-Member -MemberType $NoteProperty -Name 'Manager' -Value $dl.managedby
      $report | Add-Member -MemberType $NoteProperty -Name 'Member Depart Restriction' -Value $dl.memberdepartrestriction
      $report | Add-Member -MemberType $NoteProperty -Name 'Member Join Restriction' -Value $dl.memberjoinrestriction
      $report | Add-Member -MemberType $NoteProperty -Name 'PrimarySMTPAddress' -Value $dl.primartysmtpaddress
      $report | Add-Member -MemberType $NoteProperty -Name 'Number of Members' -Value $count
      Write-Host ('INFO: {0} has {1} members' -f $dl.name, ($count))
      $reportoutput += $report
    }
    #Write-Host 'Info: Downloading Distribution Group Report to {0}' -f $AcoraPath
    $reportoutput | Export-Csv -Path $ExPath\DistributionListReport.csv -NoTypeInformation -Encoding UTF8 -append
  }
  Function Get-DatabaseCount {
  Write-Host ('Info: Downloading Mailbox Count per Database Report to {0}' -f $ExPath) -BackgroundColor $DarkCyan
  Get-Mailbox -resultsize $unlimited | Group-Object -Property:Database | Select-Object -Property Name,Count | Sort-Object -Property Name | Select-Object -Property * | Out-File -FilePath $ExPath\MailboxCountPerDatabase.txt
}    
  Function Get-Counts {
    
    $count
    Get-WaitISE
    
    # Folder
    Get-ExFolder
    
    # Transcript
    Start-Transcript -Path $ExPath\Log_MailboxDataReport.txt
    
    # Mailbox Counts
    $mailbox = get-mailbox -resultsize $Unlimited
    $user = get-mailbox -resultsize $Unlimited -RecipientType UserMailbox 
    $Shared =  get-mailbox -resultsize $Unlimited -RecipientType SharedMailbox
    $Room = get-mailbox -resultsize $Unlimited -RecipientType RoomMailbox
    $Equipment = get-mailbox -resultsize $Unlimited -RecipientType EquipmentMailbox
    
    # Mailbox Reporting
    Write-Host ('Info: Identified {0} UserMailboxes' -f $user.count) -ForegroundColor $Yellow
    Get-MailboxPermissionReport
    Write-Host ('Info: Identified {0} SharedMailboxes' -f $shared.count) -ForegroundColor $Yellow
    Get-SharedMailboxReport
    Write-Host ('Info: Identified {0} RoomMailboxes' -f $Room.count) -ForegroundColor $Yellow
    Get-RoomMailboxReport
    Write-Host ('Info: Identified {0} EquipmentMailboxes' -f $Equipment.count) -ForegroundColor $Yellow
    Get-EquipmentMailboxReport
    Write-Host 'Info: Gathering Mailbox Alias' -ForegroundColor $Yellow
    Get-MailboxAlias
    Write-Host 'Info: Identified Mailbox count per Exchange database' -ForegroundColor $Yellow
    Get-DatabaseCount
    Write-Host 'Info: Gathering Mailboxes with Forwarding Rules' -ForegroundColor $Yellow
    Get-MailboxForwarding
    Write-Host 'Info: Gathering Distribution List Report' -ForegroundColor $Yellow
    Get-DistributionGroup
    
    Stop-Transcript
  }
# Option 6
  Function Get-ExchangeVersion {
    
    $ExchangeServers = Get-ExchangeServer  | Sort-Object Name

    ForEach  ($Server in $ExchangeServers)
    {
    Get-ExchangeServer -name $Server.Name | Format-List Name, Edition, AdminDisplayVersion
    }

  }
  Function Get-DAG {
  If (Get-MailboxDatabase | Where-Object {$_.MasterType -eq 'DatabaseAvailabilityGroup'})
  {
    Write-Host 'Info: Gathering Database Availabilty Group information' -BackgroundColor $DarkCyan
    Get-DatabaseAvailabilityGroup -Status | Format-List | Out-File -FilePath $ExPath\DAGConfig.txt
    Get-DatabaseAvailabilityGroupNetwork | Format-List | Out-File -FilePath $ExPath\DAGNetworkConfig.txt
  }
  Else {Write-Host 'Warning: Unable to detected an active DAG' -ForegroundColor $Red
  }
}
  Function Get-Databases {
  Write-Host 'Info: Gathering Mailbox Databases information' -BackgroundColor $DarkCyan
  Get-MailboxDatabase -Status | Select-Object Name,Server,Recovery,EDBFilePath,LogFolderPath,IndexEnabled,*Quota*,LastFullBackup | Out-File -FilePath $ExPath\ExchangeMailboxDatabases.txt
  Write-Host 'Info: Gathering Public Folder Database information' -BackgroundColor $DarkCyan
  Get-PublicFolderDatabase -Status | Select-Object Name,Server,EDBFilePath,LogFolderPath,IndexEnabled,*Quota*,*Backup* | Out-File -FilePath $ExPath\ExchangePublicFolderDatabases.txt
 }
  Function Get-Mailflow {
  Write-Host 'Info: Gathering Send Connector information' -BackgroundColor $DarkCyan
  Get-SendConnector | Format-List | Out-File -FilePath $ExPath\SendConnectorConfiguration.txt -Append
  Write-Host 'Info: Gathering Receive Connector information' -BackgroundColor $DarkCyan
  Get-ReceiveConnector | Format-List | Out-File -FilePath $ExPath\ReceiveConnectorConfiguration.txt -Append
  }
  Function Get-AddressPolicy {
  Write-Host 'Info: Gathering Email Address Policy information' -BackgroundColor $DarkCyan
  Get-EmailAddressPolicy | Select-Object Name,RecipientFilter,LdapRecipientFilter,RecipientFilterApplied,IncludedRecipients,Precanned,Lowest,EnabledPrimarySMTPAddressTemplate,EnabledEmailAddressTemplates,DisabledEmailAddressTemplates,Enabled,HasEmailAddressSetting,HasMailboxManagerSetting | Export-Csv -Path $ExPath\EmailAddressPolicies.csv -NoTypeInformation -Append
  }
  Function Get-VirtualDirectory {
  
    # Get OWA
    Write-Host 'Info: Gathering OWA Virtual Directory information' -BackgroundColor $DarkCyan
    Get-OwaVirtualDirectory | Format-List | Out-File -FilePath $ExPath\NamespaceConfiguration.txt

    # Get ECP
    Write-Host 'Info: Gathering ECP Virtual Directory information' -BackgroundColor $DarkCyan
    Get-EcpVirtualDirectory | Format-List | Out-File -FilePath $ExPath\NamespaceConfiguration.txt -Append

    # Get OAB
    Write-Host 'Info: Gathering OAB Virtual Directory information' -BackgroundColor $DarkCyan    
    Get-OabVirtualDirectory | Format-List | Out-File -FilePath $ExPath\NamespaceConfiguration.txt -Append
    
    # Get Active Sync
    Write-Host 'Info: Gathering Active Sync Virtual Directory information' -BackgroundColor $DarkCyan
    Get-ActiveSyncVirtualDirectory | Format-List | Out-File -FilePath $ExPath\NamespaceConfiguration.txt -Append

    # Get EWS
    Write-Host 'Info: Gathering EWS Virtual Directory information' -BackgroundColor $DarkCyan
    Get-WebServicesVirtualDirectory | Format-List | Out-File -FilePath $ExPath\NamespaceConfiguration.txt -Append

    # Get Mapi Virtual Directory
    Write-Host 'Info: Gathering Mapi Virtual Directory information' -BackgroundColor $DarkCyan
    Get-MapiVirtualDirectory | Format-List | Out-File -FilePath $ExPath\NamespaceConfiguration.txt -Append

    # Get Outlook Anywhere
    Write-Host 'Info: Gathering Outlook Anywhere Virtual Directory information' -BackgroundColor $DarkCyan
    Get-OutlookAnywhere | Format-List | Out-File -FilePath $ExPath\NamespaceConfiguration.txt -Append
  
  }
  Function Get-CASDetails {
  Write-Host 'Info: Gathering Client Access Server information' -BackgroundColor $DarkCyan
  Get-ClientAccessServer | Select-Object Server,AutoDiscoverServiceInternalUri | Format-List | Out-File -FilePath $ExPath\NamespaceConfiguration.txt -Append
  }
  Function Get-OuAnywhere {
  Write-Host 'Info: Gathering Client Access Server information' -BackgroundColor $DarkCyan
  Get-OutlookAnywhere | Format-List | Out-File -FilePath $ExPath\OutlookAnywhereConfiguration.txt -Append
  }
  Function Get-ServerInfo {
    
    $ServerInfo
    Get-WaitISE
    
    # Folder
    Get-ExFolder
    
    # Transcript
    Start-Transcript -Path $ExPath\Log_ServerConfigReport.txt
    
    # Server Reporting

    Write-Host 'Info: Gathering Exchange Server Versioning' -ForegroundColor $Yellow
    Get-ExchangeVersion
    
    Write-Host 'Info: Gathering Database Availability Group Detail' -ForegroundColor $Yellow
    Get-DAG
    
    Write-Host 'Info: Gathering Mail and Public Folder Database Detail' -ForegroundColor $Yellow
    Get-Databases
    
    Write-Host 'Info: Gathering Send and Receive Connector Detail' -ForegroundColor $Yellow
    Get-Mailflow
    
    Write-Host 'Info: Gathering Email Address Policy Detail' -ForegroundColor $Yellow
    Get-AddressPolicy
    
    Write-Host 'Info: Gathering Virtual Directory Detail' -ForegroundColor $Yellow
    Get-VirtualDirectory
    
    Write-Host 'Info: Gathering Client Access Server Detail' -ForegroundColor $Yellow
    Get-CASDetails
    
    Write-Host 'Info: Gathering Outlook Anywhere Detail' -ForegroundColor $Yellow
    Get-OuAnywhere
    
    Stop-Transcript
  }
# Option 7
  Function Get-ConfigHybrid {
  Write-Host 'Info: Gathering Exchange Hybrid Configuration information' -BackgroundColor $DarkCyan
  Get-HybridConfiguration | Out-File -FilePath $ExPath\HybridConfiguration.txt -Append
  }
  Function Get-IntraOrgConnect {
  Write-Host 'Info: Gathering Exchange IntraOrganization Configuration information' -BackgroundColor $DarkCyan
  Get-IntraOrganizationConfiguration | Out-File -FilePath $ExPath\HybridMailFlowConfiguration.txt -Append
  }
  Function Get-AZModule {
  
  $rmmodule = Get-InstalledModule -Name AzureRM -ErrorAction SilentlyContinue
  $azmodule = Get-InstalledModule -Name Az -ErrorAction SilentlyContinue

  if ($rmmodule)
  {
  Write-Host "Detected: Installation of Azure RM Module, removing installation" -BackgroundColor DarkCyan
  Uninstall-AzureRM -Force
  if ($azmodule)
   {
   Write-Host "Detected: Installation of Azure AZ Module, checking for an update" -BackgroundColor DarkCyan
   #Get-AzureAZModule
   }
   else
   {
   Write-Host "Info: No Installation of Azure AZ Module" -ForegroundColor Yellow
   Install-Module -Name Az -AllowClobber
   }
  }
  else
  {
  Write-Host "Info: No Installation of Azure RM Module" -ForegroundColor Yellow
  if ($azmodule)
   {
   Write-Host "Detected: Installation of Azure AZ Module, checking for an update" -BackgroundColor DarkCyan
   #Get-AzureAZModule
   }
   else
   {
   Write-Host "Info: No Installation of Azure AZ Module" -ForegroundColor Yellow
   Install-Module -Name Az -AllowClobber
   }
  }  
  }
  Function Get-MailflowHybrid {
  
  Get-PSVersion

  # Cloud Only
  Write-Host 'Info: Gathering Exchange Hybrid Mail Flow information' -BackgroundColor $DarkCyan
  Get-HybridMailFlow | Out-File -FilePath $ExPath\HybridMailFlowConfiguration.txt -Append

  }
  Function Get-HybridHealth {
    
    $HybridInfo
    Get-WaitISE
    
    # Acora Folder
    Get-AcoraFolder
    
    # Transcript
    Start-Transcript -Path $ExPath\Log_HybridHealthReport.txt
    
    # Hybrid Health
    
    Write-Host 'Info: Gathering Exchange Hybrid Configuration' -ForegroundColor $Yellow
    Get-ConfigHybrid

    Write-Host 'Info: Gathering Exchange Intra Organization Connector Configuration' -ForegroundColor $Yellow
    Get-IntraOrgConnect

    $HybridInfoWarning
    Write-Host 'WARNING: In order to gather information from Exchange Online, you need to meet at '

    #Write-Host 'Info: Gathering Exchange Hybrid Mailflow Configuration' -ForegroundColor $Yellow
    #Get-MailflowHybrid
    
    Stop-Transcript
  }

  
#endregion
#region Code Launch
clear-host
Write-host 'Version information - You are running script version 1.2' -ForegroundColor $White -BackgroundColor $DarkGray
Get-Root
#endregion Code Launch
