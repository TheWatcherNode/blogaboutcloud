 <#Information
     
     .DESCRIPTION
     Configure Outbound Connector for Mimecast SmartHost

     .NOTES
     Version				: 1.1
     Wish list			    : Better error trapping, Report on SQL Database state after conversion.
     Rights Required	    : Local administrator on Workstation
                            : Powershell in Administrator Mode
     Sched Task Required	: No
     Windows Server Version	: 
     Author/Copyright	    : © The WatcherNode - All Rights Reserved
     Email/Blog/Twitter	    : 
     Dedicated Post	        : 
 
     Disclaimer             : You running this script means you won't blame me if this breaks your stuff. This script is provided AS IS without warranty of any kind. 
                              I disclaim all implied warranties including, without limitation, any implied warranties of merchantability or of fitness
                              .SYNOPSIS for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and
                              documentation remains with you. In no event shall I be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption,
                              loss of business information, or other pecuniary loss) arising out of the use of or inability
                              to use the script or documentation.

     Acknowledgements 	        : Andrew Offor, Rachel Moran
     Assumptions		        : ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
     Limitations			    : 
     Known issues				:   
     Version Changes            : 0.1 Initial Script Build
                                : 1.0 Initial Build
                                : 1.1 Additional Options (Journal Connector)
     .LINK
     

     .EXAMPLE
     .\set-o365-mimecastconnector.ps1

     Description
     -----------
     Runs script with default values.

     .INPUTS
     None. You cannot pipe objects to this script.
 #>
#region Variables
# Smart Host
$America1 = 'us-smtp-o365-outbound-1.mimecast.com'
$America2 = 'us-smtp-o365-outbound-2.mimecast.com'
$Australia1 = 'au-smtp-o365-outbound-1.mimecast.com'
$Australia2 = 'au-smtp-o365-outbound-2.mimecast.com'
$Europe1  = 'eu-smtp-o365-outbound-1.mimecast.com'
$Europe2  = 'eu-smtp-o365-outbound-2.mimecast.com'
$Germany1 = 'de-smtp-o365-outbound-1.mimecast.com'
$Germany2 = 'de-smtp-o365-outbound-2.mimecast.com'
$Offshore1 = 'je-smtp-o365-outbound-1.mimecast-offshore.com'
$Offshore2 ='je-smtp-o365-outbound-2.mimecast-offshore.com'
$SouthAfrica1 = 'za-smtp-o365-outbound-1.mimecast.co.za'
$SouthAfrica2 = 'za-smtp-o365-outbound-2.mimecast.co.za'

# Journaling
$Americaj1 = 'us-smtp-journal-1.mimecast.com'
$Americaj2 = 'us-smtp-journal-2.mimecast.com'
$Australiaj1 = 'au-smtp-journal-1.mimecast.com'
$Australiaj2 = 'au-smtp-journal-2.mimecast.com'
$Europej1  = 'eu-smtp-journal-1.mimecast.com'
$Europej2  = 'eu-smtp-journal-2.mimecast.com'
$Germanyj1 = 'de-smtp-journal-1.mimecast.com'
$Germanyj2 = 'de-smtp-journal-2.mimecast.com'
$Offshorej1 = 'je-smtp-journal-1.mimecast.com'
$Offshorej2 ='je-smtp-journal-2.mimecast.com'
$SouthAfricaj1 = 'za-smtp-journal-1.mimecast.za'
$SouthAfricaj2 = 'za-smtp-journal-2.mimecast.za'

$tlssettings = 'certificatevalidation'
$onprem = 'Partner' 

$White = 'White'
$Red = 'Red'
$Green = 'Green'
$DarkGray = 'DarkGray'
$Cyan = 'Cyan'
$Quit = 'Q'
#endregion
#region Banner
[string] $test = @'
 
  ┌─────────────────────────────────────────────────────────────┐
             Office 365 Mimecast Outbound Connector             
           
              Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

  
  1) Configure Outbound Routing
  2) Configure Inbound Routing
  3) Configure Journaling to Mimecast
     
  Q) Quit

  Select an option.. [1-99]?
'@
[string] $test1 = @'
 
  ┌─────────────────────────────────────────────────────────────┐
             Office 365 Mimecast Outbound Connector             
           
              Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

  Outbound Routing
  1)  Connector to America                 -->
  2)  Connector to Australia               -->
  3)  Connector to Europe (except Germany) -->
  4)  Connector to Germany                 -->
  5)  Connector to OffShore                -->
  6)  Connector to South Africa            -->
     
  Q) Quit

  Select an option.. [1-99]?
'@
[string] $test2 = @'
 
  ┌─────────────────────────────────────────────────────────────┐
             Office 365 Mimecast Outbound Connector             
           
              Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

  Inbound Routing - Spam Filter IP Ranges
  1)  Configure Spam Filter for America                 -->
  2)  Configure Spam Filter for Australia               -->
  3)  Configure Spam Filter for Europe (except Germany) -->
  4)  Configure Spam Filter for Germany                 -->
  5)  Configure Spam Filter for OffShore                -->
  6)  Configure Spam Filter for South Africa            -->
     
  Q) Quit

  Select an option.. [1-99]?
'@
[string] $test3 = @'
 
  ┌─────────────────────────────────────────────────────────────┐
             Office 365 Mimecast Outbound Connector             
           
              Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

  Configure Journaling to Mimecast
  0)  Configure Mail Contact for Connector -->
  1)  Connector to America                 -->
  2)  Connector to Australia               -->
  3)  Connector to Europe (except Germany) -->
  4)  Connector to Germany                 -->
  5)  Connector to OffShore                -->
  6)  Connector to South Africa            -->
     
  Q) Quit

  Select an option.. [1-99]?
'@
[string] $Root = @'
 
  ┌─────────────────────────────────────────────────────────────┐
             Office 365 Mimecast Outbound Connector             
           
              Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

  Outbound Routing
  1)  Connector to America                 -->
  2)  Connector to Australia               -->
  3)  Connector to Europe (except Germany) -->
  4)  Connector to Germany                 -->
  5)  Connector to OffShore                -->
  6)  Connector to South Africa            -->

  Configure Journaling to Mimecast
  10)  Configure Mail Contact for Connector -->
  11)  Connector to America                 -->
  12)  Connector to Australia               -->
  13)  Connector to Europe (except Germany) -->
  14)  Connector to Germany                 -->
  15)  Connector to OffShore                -->
  16)  Connector to South Africa            -->
     
  Q) Quit

  Select an option.. [1-99]?
'@
[string] $Warning = @'
 
  ┌─────────────────────────────────────────────────────────────┐
   Warning! If you have not configured the a Mail Contact this
   process will not be successful                 
  └─────────────────────────────────────────────────────────────┘
'@
#endregion Banner
function Get-Test    {
  # Menu Prompt
  
    Do {
      $MenuOption = Read-Host -Prompt $test
      Clear-Host
      switch ($MenuOption){
        1 { # Option 1
             Get-test1
          } 
        2 { # Option 2
             Get-test2
          }
        3 { # Option 2
             Get-test3
          }   
        $Quit {exit} 
      }
    
    }  until ($Root -eq {$Quit})
  }
function Get-Root    {
  # Menu Prompt
  
    Do {
      $MenuOption = Read-Host -Prompt $Root
      Clear-Host
      switch ($MenuOption){
        1 { # Option 1
             Get-Root1
          } 
        2 { # Option 2
             Get-Root2
          }
        3 { # Option 2
             Get-Root3
          }
        4 { # Option 2
             Get-Root4
          }
        5 { # Option 5
             Get-Root5
          }
        6 { # Option 6
             Get-Root6
          }
        10{ # Option 10
             Get-Root10
          }
        11{ # Option 11
             Get-Root11
          }
        12{ # Option 12
             Get-Root12
          }
        13{ # Option 13
             Get-Root13
          }
        14{ # Option 14
             Get-Root14
          }
        15{ # Option 15
             Get-Root15
          }
        16{ # Option 16
             Get-Root16
          }     
        
        $Quit {exit} 
      }
    
    }  until ($Root -eq {$Quit})
  }
# Outbound Connector
function Get-Root1   { # Connector for Americas

  $ConnectorName = Read-Host -Prompt 'Specify the name for your Outbound Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  new-outboundconnector -name $ConnectorName -smarthosts $America1,$America2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Outbound Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Outbound Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    }
}
function Get-Root2   { # Connector for Australia

  $ConnectorName = Read-Host -Prompt 'Specify the name for your Outbound Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  new-outboundconnector -name $ConnectorName -smarthosts $Australia1,$Australia2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Outbound Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Outbound Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    }
}
function Get-Root3   { # Connector for Europe

  $ConnectorName = Read-Host -Prompt 'Specify the name for your Outbound Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  new-outboundconnector -name $ConnectorName -smarthosts $Europe1,$Europe2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Outbound Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Outbound Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    } 
}
function Get-Root4   { # Connector for Germany

  $ConnectorName = Read-Host -Prompt 'Specify the name for your Outbound Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  new-outboundconnector -name $ConnectorName -smarthosts $Germany1,$Germany2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Outbound Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Outbound Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    }
}
function Get-Root5   { # Connector for Offshore

  $ConnectorName = Read-Host -Prompt 'Specify the name for your Outbound Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  new-outboundconnector -name $ConnectorName -smarthosts $Offshore1,$Offshore2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Outbound Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Outbound Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    }
}
function Get-Root6   { # Connector for South Africa

  $ConnectorName = Read-Host -Prompt 'Specify the name for your Outbound Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  new-outboundconnector -name $ConnectorName -smarthosts $SouthAfrica1,$SouthAfrica2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Outbound Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Outbound Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    }
  }
# Journal Connector
Function Get-MailboxContact {
   $title = 'Mail Contact'
   $Message = 'Have you created a Mail Contact for the Journal Connect i.e journal.domain.com?'
   $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes',`
   'Launching Failover Clustering MMC .'
   $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No',`
   'Skipping..'
   $options = [management.automation.host.choicedescription[]]($yes, $no)
   $result = $host.ui.PromptForChoice($title,$Message,$options,1)
   Switch ($result)
   {
    0 {Get-JournalA}
    1 {Get-Root10b}
   }
   }
function Get-Root10 {
   $Warning
   $Name = Read-Host -Prompt 'Please specify a name of the Mail Contact'
   Write-Host -Foreground $Cyan ('You have specified {0} as the name of the Mail Contact' -f $Name)
   $ExternalEmailAddress = Read-Host -Prompt 'Please specify the External Email Address (i.e journal@domain.com)'
   Write-Host -Foreground $Cyan ('You have specified {0} as your External Email Address' -f $ExternalEmailAddress)

   New-MailContact -Name $Name -ExternalEmailAddress $ExternalEmailAddress
   $newmailcontact = Get-MailContact -Name $Name

}
function Get-Root11 { # Connector for Americas

  $Warning
  $ConnectorName = Read-Host -Prompt 'Specify the name for your Journal Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  $Journal = Read-Host -Prompt 'Specify the name for your Mail Contact (journal@domain.com)'
  Write-Host -Foreground $Cyan ('You have specified {0} as your mail contact for recipientdomain' -f $Journal)

  new-outboundconnector -name $ConnectorName -recipientdomain $journal -smarthosts $Americaj1,$Americaj2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Journal Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Journal Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    }
}
function Get-Root12 { # Connector for Australia

  $ConnectorName = Read-Host -Prompt 'Specify the name for your Journal Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  $Journal = Read-Host -Prompt 'Specify the name for your Mail Contact (journal@domain.com)'
  Write-Host -Foreground $Cyan ('You have specified {0} as your mail contact for recipientdomain' -f $Journal)

  new-outboundconnector -namel $ConnectorName -recipientdomain $journal -smarthosts $Australiaj1,$Australiaj2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Journal Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Journal Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    }
}
function Get-Root13 { # Connector for Europe

  $ConnectorName = Read-Host -Prompt 'Specify the name for your Journal Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  $Journal = Read-Host -Prompt 'Specify the name for your Mail Contact (journal@domain.com)'
  Write-Host -Foreground $Cyan ('You have specified {0} as your mail contact for recipientdomain' -f $Journal)

  new-outboundconnector -name $ConnectorName -recipientdomain $journal -smarthosts $Europej1,$Europej2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Journal Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Journal Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    } 
}
function Get-Root14 { # Connector for Germany

  $ConnectorName = Read-Host -Prompt 'Specify the name for your Journal Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  $Journal = Read-Host -Prompt 'Specify the name for your Mail Contact (journal@domain.com)'
  Write-Host -Foreground $Cyan ('You have specified {0} as your mail contact for recipientdomain' -f $Journal)

  new-outboundconnector -name $ConnectorName -recipientdomain $journal -smarthosts $Germanyj1,$Germanyj2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Journal Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Journal Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    }
}
function Get-Root15 { # Connector for Offshore

  $ConnectorName = Read-Host -Prompt 'Specify the name for your Journal Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  $Journal = Read-Host -Prompt 'Specify the name for your Mail Contact (journal@domain.com)'
  Write-Host -Foreground $Cyan ('You have specified {0} as your mail contact for recipientdomain' -f $Journal)

  new-outboundconnector -name $ConnectorName -recipientdomain $journal -smarthosts $Offshorej1,$Offshorej2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Journal Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Journal Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    }
}
function Get-Root16 { # Connector for South Africa

  $ConnectorName = Read-Host -Prompt 'Specify the name for your Journal Connector'
  Write-Host -Foreground $Cyan ('You have specified {0} as your connector name' -f $ConnectorName)
  $Journal = Read-Host -Prompt 'Specify the name for your Mail Contact (journal@domain.com)'
  Write-Host -Foreground $Cyan ('You have specified {0} as your mail contact for recipientdomain' -f $Journal)

  new-outboundconnector -name $ConnectorName -recipientdomain $journal -smarthosts $SouthAfrica1,$SouthAfrica2 -tlssettings $tlssettings -recipientdomains * -routeallmessagesviaonpremises $false -connectortype $onprem -usemxrecord $false
  $ConnectCheck = Get-OutboundConnector $ConnectorName
  try {
    $ConnectCheck
      Write-Host ('Success: Your Journal Connector {0} has been created' -f $ConnectorName) -ForegroundColor $Green

    }
    catch
    {
      Write-Host ('Error: Your Journal Connector {0} has not been created' -f $ConnectorName) -ForegroundColor $Red
    }
  }

#region Code Launch
clear-host
Write-host 'Version information - You are running script version 1.1' -ForegroundColor $White -BackgroundColor $DarkGray
Get-Root
#endregion Code Launch