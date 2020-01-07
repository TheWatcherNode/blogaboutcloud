
# Variables
$Quit = 'Q'

#region Banner
[string] $Root = @'
 
  ┌─────────────────────────────────────────────────────────────┐
                  Get-DisabledUserwithMailbox              
           
              Follow me @thewatchernode on Twitter                   
  └─────────────────────────────────────────────────────────────┘

  1)  Get-ADUser      (On-prem)      -->
  2)  Get-AzureADUser (Cloud)        -->
     
  Q) Quit

  Select an option.. [1-99]?
'@
#endregion Banner
#region Functions
Function Get-ADUser {
  
  $Mailboxes = Get-Mailbox | Where-Object {$_.RecipientTypeDetails -eq 'UserMailbox'}
  $Disabled = @()

  Foreach ($Mailbox in $Mailboxes) {
    if((Get-msolUser -id $Mailbox.samaccountname).Enabled -eq $False){
        $Disabled += Get-MailboxStatistics $Mailbox.samaccountname | Select-Object -Property DisplayName,TotalItemSize
    }    
  }
  $Disabled | Export-Csv -Path $env:userprofile\desktop\DisabledADUserwithMailbox.csv -NoTypeInformation
  }
Function Get-AzureADUser {

  Connect-MsolService
  
  $Mailboxes = Get-Mailbox | Where-Object {$_.RecipientTypeDetails -eq 'UserMailbox'}
  $Disabled = @()

  Foreach ($Mailbox in $Mailboxes) {
    if((Get-msolUser -userprincipalname $Mailbox.userprincipalname).Enabled -eq $False){
        $Disabled += Get-MailboxStatistics $Mailbox.userprincipalname | Select-Object -Property DisplayName,TotalItemSize
    }    
  }
  $Disabled | Export-Csv -Path $env:userprofile\desktop\DisabledAzureADUserwithMailbox.csv -NoTypeInformation
}
function Get-Root    {
  # Menu Prompt
  
    Do {
      $MenuOption = Read-Host -Prompt $Root
      Clear-Host
      switch ($MenuOption){
        1 { # Option 1
             Get-ADUser
             Get-Root
          } 
        2 { # Option 2
             Get-AzureADUser
             Get-Root
          }
        Q { # Not in use
            Exit
          }
        4 { # Not in use
          }
        5 { # Option 5
            Get-Root5
          }
        6 { # Option 6
            Get-Root6
          }  
        
  
        $Quit {return} 
      }
    
    }  until ($Root -eq {$Quit})
  }
  #endregion Menu Prompt
  
  
#region Code Launch
clear-host
Get-Root
#endregion Code Launch

