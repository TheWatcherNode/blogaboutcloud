<#Information
 
    Author: thewatchernode
    Contact: author@blogabout.cloud
    Published: 20th November 2020

    .DESCRIPTION
    Change many privacy settings to be off by default. Remove Built-in advertising, Cortana, OneDrive, Cortana stuff 
    (all optional). 

    Version Changes            
    : 0.1 Initial Script Build
    : 1.0 Initial Build Release
    : 1.1 Tested on Windows 1909
    : 1.2 Tested on Windows 2020

    .LINK
     

    .EXAMPLE
    .\get-win10bloatware

    Description
    -----------
    Runs script with default values.


    .INPUTS
    No switches Disables unnecessary services and scheduled tasks. Removes all UWP apps except for some useful ones.Disables Cortana, OneDrive, restricts default privacy settings and cleans up the default start menu.
    -AllApps       Removes ALL apps including the store. Make sure this is what you want before you do it. It can be tough to get the store back.
    -LeaveTasks    Leaves scheduled tasks alone.
    -LeaveServices Leaves services alone.
    -AppAccess		Sets privacy permissions in Settings -> Privacy to restricted by default. Some apps may require these  permissions to function, you can just reenable them from the settings menu.
    -ClearStart    Empties the start menu completely leaving you with just the apps list.
    -OneDrive		Leaves all OneDrive content fully functional.
    -Tablet		Use this for tablets or 2-in-1s to leave location and sensors enabled.
    -Cortana		Leave Cortana and web enabled search intact.
    -Xbox			Leave xBox apps and related items.
    -AppsOnly      Only removes apps, doesn't touch privacy settings,services, and scheduled tasks. Cannot be used with -SettingsOnly switch. Can be used with all the others.
    -SettingsOnly  Only adjusts privacy settings, services, and scheduled tasks. Leaves apps. Cannot be used with -AppsOnly switch.  Can be used with all others (-AllApps won't do anything in that case, obviously).
#>


[cmdletbinding(DefaultParameterSetName='Win10Debloat')]
param (
  [switch]$AllApps, 
  [switch]$LeaveTasks,
  [switch]$LeaveServices,
  [switch]$AppAccess,
  [switch]$OneDrive,
  [switch]$Xbox,
  [switch]$Tablet,
  [switch]$Cortana,
  [switch]$ClearStart,
  [Parameter(ParameterSetName='AppsOnly')]
  [switch]$AppsOnly,
  [Parameter(ParameterSetName='SettingsOnly')]
  [switch]$SettingsOnly
)

<# Customize Script to your requirements 


    # Keeping a number of Apps installed 
    Sometime there maybe a requiredment to keep cerain apps in order to do this you will need to modify line 77  to make it active and specifiy the Apps you want to keep.

    Make sure not begin or end with a | (vertical line) ex: "app|app2" - good. "|app|app2|" - bad.


    Start Menu XML. 
    If you run the script without -ClearStart, the XML below will be 
    used for a custom start layout. By default it just leaves File 
    Explorer, classic Control Panel, and Snipping Tool tiles.
    Place your XML like so:
    $StartLayourStr = @"
    <**YOUR START LAYOUT XML**>
    "@
#>

# Remove hash to utilizse # $GoodApps =	"store|calculator|sticky|windows.photos|soundrecorder|mspaint|screensketch"
$StartLayoutStr = @' 
<LayoutModificationTemplate 
    Version="1" 
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification">
  <LayoutOptions StartTileGroupCellWidth="6" />
  <DefaultLayoutOverride>
    <StartLayoutCollection>
      <defaultlayout:StartLayout 
            GroupCellWidth="6" 
            xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout">
        <start:Group 
            Name="" 
            xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout">
          <start:DesktopApplicationTile 
                Size="2x2" Column="0" Row="0" 
                DesktopApplicationLinkPath="%APPDATA%\Microsoft\Windows\Start Menu\Programs\System Tools\File Explorer.lnk" />
          <start:DesktopApplicationTile 
                Size="2x2" Column="2" Row="0" 
                DesktopApplicationLinkPath="%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Accessories\Snipping Tool.lnk" />
		  <start:DesktopApplicationTile 
                Size="2x2" Column="0" Row="2" 
                DesktopApplicationLinkPath="%APPDATA%\Microsoft\Windows\Start Menu\Programs\System Tools\Control Panel.lnk" />
        </start:Group>
      </defaultlayout:StartLayout>
    </StartLayoutCollection>
  </DefaultLayoutOverride>
</LayoutModificationTemplate>
'@


#region Static Variables 

$Disabled = 'disabled'
$REG_DWORD = 'REG_DWORD'
$SilentlyContinue = 'silentlycontinue'
$DarkCyan = 'DarkCyan'
$DarkRed = 'DarkRed'
$DarkGray = 'DarkGray'
$Green = 'Green'
$Red = 'Red'
$Yellow = 'Yellow'
$White = 'White'

#endregion

#region Functions

#Appx removal
#First determines based on -OneDrive and -Xbox switches whether to also leave OneDrive apps, xBox apps, both, or none.  Then removes all apps or some apps depending on -AllApps.
#Apps that shouldn't be removed, don't edit this variable!
$SafeApps = 'sechealth|secureas'
Function Get-RemoveApps {
	
  If ($Xbox) { 
    $GoodApps = ('{0}|Xbox' -f $GoodApps) 
  }
	 
  If ($AllApps) {  
    Write-Verbose -Message 'INFO: Removing all apps and provisioned appx packages for this device.'
    Get-AppxPackage -allusers | where-object {$_.name -notmatch $SafeApps} | Remove-AppxPackage -erroraction $SilentlyContinue
    Get-AppxPackage -allusers | where-object {$_.name -notmatch $SafeApps} | Remove-AppxPackage -erroraction $SilentlyContinue
    Get-AppxProvisionedPackage -online |where-object {$_.displayname -notmatch $SafeApps} |  Remove-AppxProvisionedPackage -online -erroraction $SilentlyContinue
  }
  Else {
    Write-Verbose -Message 'INFO: Removing apps and provisioned appx packages for this device'      
    Get-AppxPackage -allusers | where-object {($_.name -notmatch $GoodApps) -and ($_.name -notmatch $SafeApps)} | Remove-AppxPackage -erroraction $SilentlyContinue
    Get-AppxPackage -allusers | where-object {($_.name -notmatch $GoodApps) -and ($_.name -notmatch $SafeApps)} | Remove-AppxPackage -erroraction $SilentlyContinue
    Get-AppxProvisionedPackage -online | where-object {($_.displayname -notmatch $GoodApps) -and ($_.displayname -notmatch $SafeApps)} | Remove-AppxProvisionedPackage -online -erroraction $SilentlyContinue
  }        
}
Function Get-DisableTasks {
  If ($LeaveTasks) {
    Write-Verbose -Message 'WARNING: Leave Tasks switch was set - no changes will be made scheduled tasks ' 
  }
  Else {
    Write-Host 'INFO: Disabling unecessary scheduled tasks'
    Get-Scheduledtask -TaskName 'Microsoft Compatibility Appraiser','ProgramDataUpdater','Consolidator','KernelCeipTask','UsbCeip','Microsoft-Windows-DiskDiagnosticDataCollector', 'GatherNetworkInfo','QueueReporting' -ErrorAction silentlycontinue | Disable-scheduledtask 
  }
}
Function Get-DisableServices {
    
  If ($LeaveServices) {
    Write-Host 'WARNING: Leave Service switch was set - no changes will be made services'
  }
  Else {
    Write-Host 'INFO: Stopping and disabling unneccessary services'
    #Diagnostics tracking WMP Network Sharing
    Get-Service -Name Diagtrack,WMPNetworkSvc -ErrorAction $SilentlyContinue | Stop-Service -PassThru | Set-Service -startuptype $Disabled
		
    #XBox services
    If (!($Xbox)){
      #Disable xBox services - "xBox Game Monitoring Service" - XBGM - Can't be disabled (access denied)
      Get-Service -Name XblAuthManager,XblGameSave,XboxNetApiSvc -ErrorAction $SilentlyContinue | Stop-Service -PassThru | Set-Service -StartupType $Disabled
    }
  }
}
Function Get-LoadDefaultHive {
$matjazp72 = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' Default).Default
    reg load "$reglocation" $matjazp72\ntuser.dat
}
Function Get-UnloadDefaultHive {
  [gc]::collect()
  & "$env:windir\system32\reg.exe" unload "$reglocation"
}
Function Get-RegChange {
  Write-Host 'INFO: Configuring registry for current and default user profiles, and policies for the local machine.'
  $reglocation = 'HKCU'
  regsetuser
  $reglocation = 'HKLM\AllProfile'
  loaddefaulthive
  regsetuser
  unloaddefaulthive
  $reglocation = $null
  regsetmachine
  Write-Host 'INFO: The registry has been set current user and default user profiles, and policies set for the local machine!***'
}
Function Get-RegSetUser {
  #Start menu suggestions

  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'SystemPaneSuggestionsEnabled' /D 0 /F
  #Show suggested content in settings
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'SubscribedContent-338393Enabled' /D 0 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'SubscribedContent-353694Enabled' /D 0 /F
  #Show suggestions occasionally
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'SubscribedContent-338388Enabled' /D 0 /F
  #Multitasking - Show suggestions in timeline
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'SubscribedContent-353698Enabled' /D 0 /F
  #Lockscreen suggestions, rotating pictures
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'SoftLandingEnabled' /D 0 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'RotatingLockScreenEnabled' /D 0 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'RotatingLockScreenOverlayEnabled' /D 0 /F
  #Preinstalled apps, Minecraft Twitter etc all that - Enterprise only it seems
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'PreInstalledAppsEnabled' /D 0 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'PreInstalledAppsEverEnabled' /D 0 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'OEMPreInstalledAppsEnabled' /D 0 /F
  #MS shoehorning apps quietly into your profile
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'SilentInstalledAppsEnabled' /D 0 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'ContentDeliveryAllowed' /D 0 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'SubscribedContentEnabled' /D 0 /F
  #Ads in File Explorer
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /T $REG_DWORD /V 'ShowSyncProviderNotifications' /D 0 /F
  #Show me the Windows welcome experience after updates and occasionally
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'SubscribedContent-310093Enabled' /D 0 /F
  #Get tips, tricks, suggestions as you use Windows 
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'SubscribedContent-338389Enabled' /D 0 /F 
  #Ask for feedback
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Siuf\Rules" /T $REG_DWORD /V 'NumberOfSIUFInPeriod' /D 0 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Siuf\Rules" /T $REG_DWORD /V 'PeriodInNanoSeconds' /D 0 /F
	
  #Let apps use advertising ID
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo" /T $REG_DWORD /V 'Enabled' /D 0 /F
	
  #Tailored experiences - Diagnostics & Feedback settings
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Privacy" /T $REG_DWORD /V 'TailoredExperiencesWithDiagnosticDataEnabled' /D 0 /F
	
  #Let apps on other devices open messages and apps on this device - Shared Experiences settings
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CDP" /T $REG_DWORD /V 'RomeSdkChannelUserAuthzPolicy' /D 0 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CDP" /T $REG_DWORD /V 'CdpSessionUserAuthzPolicy' /D 0 /F
	
  #Speech Inking & Typing - comment out if you use the pen\stylus a lot
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\SettingSync\Groups\Language" /T $REG_DWORD /V 'Enabled' /D 0 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\InputPersonalization" /T $REG_DWORD /V 'RestrictImplicitTextCollection' /D 1 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\InputPersonalization" /T $REG_DWORD /V 'RestrictImplicitInkCollection' /D 1 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore" /T $REG_DWORD /V 'HarvestContacts' /D 0 /F
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Personalization\Settings" /T $REG_DWORD /V 'AcceptedPrivacyPolicy' /D 0 /F
	
  #Improve inking & typing recognition
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Input\TIPC" /T $REG_DWORD /V 'Enabled' /D 0 /F
	
  #Pen & Windows Ink - Show recommended app suggestions
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\PenWorkspace" /T $REG_DWORD /V 'PenWorkspaceAppSuggestionsEnabled' /D 0 /F
	
  #People
  #Show My People notifications
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People\ShoulderTap" /T $REG_DWORD /V 'ShoulderTap' /D 0 /F
	
  #Show My People app suggestions
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /T $REG_DWORD /V 'SubscribedContent-314563Enabled' /D 0 /F
	
  #People on Taskbar
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People" /T $REG_DWORD /V 'PeopleBand' /D 0 /F
	
  #Other Settings
  #Use Autoplay for all media and devices
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers" /T $REG_DWORD /V 'DisableAutoplay' /D 1 /F
	
  #Taskbar search, personal preference. 0 = no search, 1 = search icon, 2 = search bar
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /T $REG_DWORD /V 'SearchboxTaskbarMode' /D 0 /F
	
  #Allow search to use location if it's enabled
  & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /T $REG_DWORD /V 'AllowSearchToUseLocation' /D 0 /F
	
  #--Optional User Settings--
	
  #App permissions user settings, these are all available from the settings menu
  If (!($AppAccess)) { 	
    #Camera
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam" /T REG_SZ /V 'Value' /D Deny /F
    #Microphone
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone" /T REG_SZ /V 'Value' /D Deny /F
    #Account Info
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\userAccountInformation" /T REG_SZ /V 'Value' /D Deny /F
    #Contacts
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\contacts" /T REG_SZ /V 'Value' /D Deny /F	
    #Calendar
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\appointments" /T REG_SZ /V 'Value' /D Deny /F
    #Call history
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\phoneCallHistory" /T REG_SZ /V 'Value' /D Deny /F
    #Email
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\email" /T REG_SZ /V 'Value' /D Deny /F
    #Tasks
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\userDataTasks" /T REG_SZ /V 'Value' /D Deny /F
    #TXT/MMS
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\chat" /T REG_SZ /V 'Value' /D Deny /F
    #Cellular Data
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\cellularData" /T REG_SZ /V 'Value' /D Deny /F
    #Allow apps to run in background global setting - seems to reset during OOBE
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications" /T $REG_DWORD /V 'GlobalUserDisabled' /D 1 /F
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /T $REG_DWORD /V 'BackgroundAppGlobalToggle' /D 0 /F	
    #My Documents
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\documentsLibrary" /T REG_SZ /V 'Value' /D Deny /F
    #My Pictures
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\picturesLibrary" /T REG_SZ /V 'Value' /D Deny /F
    #My Videos
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\videosLibrary" /T REG_SZ /V 'Value' /D Deny /F
    #File System
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\broadFileSystemAccess" /T REG_SZ /V 'Value' /D Deny /F
  }	 
	
  #Disable Cortana - use -Cortana to leave it on
  If (!($Cortana)){
    #Disable Cortana and Bing search user settings
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /T $REG_DWORD /V 'CortanaEnabled' /D 0 /F
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /T $REG_DWORD /V 'CanCortanaBeEnabled' /D 0 /F
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /T $REG_DWORD /V 'BingSearchEnabled' /D 0 /F
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /T $REG_DWORD /V 'DeviceHistoryEnabled' /D 0 /F
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /T $REG_DWORD /V 'CortanaConsent' /D 0 /F
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /T $REG_DWORD /V 'CortanaInAmbientMode' /D 0 /F

    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Speech_OneCore\Preferences" /T $REG_DWORD /V 'VoiceActivationEnableAboveLockscreen' /D 0 /F
    #Disable Cortana search history
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /T $REG_DWORD /V 'HistoryViewEnabled' /D 0 /F
  }

  #Tablet Settings - use -Tablet switch to leave these on
  If (!($Tablet)) {
    #Deny access to location and sensors
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Permissions\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}" /T $REG_DWORD /V 'SensorPermissionState' /D 0 /F
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}" /T REG_SZ /V 'Value' /D Deny /F
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{E6AD100E-5F4E-44CD-BE0F-2265D88D14F5}" /T REG_SZ /V 'Value' /D Deny /F
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location" /T REG_SZ /V 'Value' /D Deny /F
  }
		
  #Game settings - use -Xbox to leave these on
  If (!$Xbox) {
    #Disable Game DVR
    & "$env:windir\system32\reg.exe" Add "$reglocation\System\GameConfigStore" /T $REG_DWORD /V 'GameDVR_Enabled' /D 0 /F
  }
	
  #OneDrive settings - use -OneDrive switch to leave these on
  If (!($OneDrive)) {
    #Disable OneDrive startup run user settings
    & "$env:windir\system32\reg.exe" Add "$reglocation\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /T REG_BINARY /V 'OneDrive' /D 0300000021B9DEB396D7D001 /F
  }	 

  #End user registry settings
}
Function Get-RegSetMachine {

  #/Application Compatibility
  #Turn off Application Telemetry			
  & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat' /T REG_DWORD /V 'AITEnable' /D 0 /F			
  #Turn off inventory collector			
  & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat' /T REG_DWORD /V 'DisableInventory' /D 1 /F

  #/Cloud Content			
  #Do not show Windows Tips			
  & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent' /T REG_DWORD /V 'DisableSoftLanding' /D 1 /F
  #Turn off Consumer Experiences			
  & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent' /T REG_DWORD /V 'DisableWindowsConsumerFeatures' /D 1 /F
  #3rd party suggestions
  & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent' /T REG_DWORD /V 'DisableThirdPartySuggestions' /D 1 /F
  	
  # /Data Collection and Preview Builds			
  #Set Telemetry to off (switches to 1:basic for W10Pro and lower)			
  & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection' /T REG_DWORD /V 'AllowTelemetry' /D 0 /F
  #Do not show feedback notifications			
  & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection' /T REG_DWORD /V 'DoNotShowFeedbackNotifications' /D 1 /F
	
  #/Microsoft Edge			
  #Always send do not track			
  & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\MicrosoftEdge\Main' /T REG_DWORD /V 'DoNotTrack' /D 1 /F

  #Windows Search
  #Set what info is shared in search to anonymous only (may be deprecated)
  & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search' /T $REG_DWORD /V 'ConnectedSearchPrivacy' /D 3 /F
	
  #Add "Run as different user" to context menu
  & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\Explorer' /T $REG_DWORD /V 'ShowRunasDifferentuserinStart' /D 1 /F
	
  #/Windows Update			
  #Turn off featured SOFTWARE notifications through WU (basically ads)			
  & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' /T $REG_DWORD /V 'EnableFeaturedSoftware' /D 0 /F

  #--Non Local GP Settings--		
  #Delivery Optimization settings - sets to 1 for LAN only, change to 0 for off
  & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config' /T $REG_DWORD /V 'DownloadMode' /D 1 /F
  & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config' /T $REG_DWORD /V 'DODownloadMode' /D 1 /F
  & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Settings' /T $REG_DWORD /V 'DownloadMode' /D 1 /F
	
  #Disabling advertising info and device metadata collection for this machine
  & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo' /T $REG_DWORD /V 'Enabled' /D 0 /F
  & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata' /V 'PreventDeviceMetadataFromNetwork' /T REG_DWORD /D 1 /F

  #Disable CEIP. GP setting at: Computer Config\Admin Templates\System\Internet Communication Managemen\Internet Communication settings
  & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Microsoft\SQMClient\Windows' /T $REG_DWORD /V 'CEIPEnable' /D 0 /F
	
    #Prevent using sign-in info to automatically finish setting up after an update
  & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' /T $REG_DWORD /V 'ARSOUserConsent' /D 2 /F
    
  #Enable diagnostic data viewer
  & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Diagnostics\DiagTrack\EventTranscriptKey' /T $REG_DWORD /V 'EnableEventTranscript' /D 1 /F
	
  #Disable Edge desktop shortcut
  & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer' /T $REG_DWORD /V 'DisableEdgeDesktopShortcutCreation' /D 1 /F

  #--Optional Machine Settings--
	
  #Disable Cortana - use -Cortana to leave it on
  If (!($Cortana)){
    #Cortana local GP - Computer Config\Admin Templates\Windows Components\Search			
    #Disallow Cortana			
    & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search' /T $REG_DWORD /V 'AllowCortana' /D 0 /F
  }	 

  #Tablet Settings - use -Tablet switch to leave these on
  If (!($Tablet)) {
    #Turn off location - global
    & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location' /T REG_SZ /V 'Value' /D Deny /F
  }	 
	
  #Game settings - use -Xbox to leave these on
  If (!($Xbox)) {
    #Disable Game Monitoring Service
    & "$env:windir\system32\reg.exe" Add 'HKLM\SYSTEM\CurrentControlSet\Services\xbgm' /T $REG_DWORD /V 'Start' /D 4 /F
    & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\GameDVR' /T $REG_DWORD /V 'AllowGameDVR' /D 0 /F
  }	 

  #OneDrive settings - use -OneDrive switch to leave these on
  If (!($OneDrive)) {
    #Prevent usage of OneDrive local GP - Computer Config\Admin Templates\Windows Components\OneDrive	
    & "$env:windir\system32\reg.exe" Add	'HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive' /T $REG_DWORD /V 'DisableFileSyncNGSC' /D 1 /F
    & "$env:windir\system32\reg.exe" Add 'HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive' /T $REG_DWORD /V 'DisableFileSync' /D 1 /F
    #Remove OneDrive from File Explorer
    & "$env:windir\system32\reg.exe" Add 'HKCR\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}' /T $REG_DWORD /V 'System.IsPinnedToNameSpaceTree' /D 0 /F
    & "$env:windir\system32\reg.exe" Add 'HKCR\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}' /T $REG_DWORD /V 'System.IsPinnedToNameSpaceTree' /D 0 /F
  }	 
	
  #End machine registry settings
}            
Function Get-ClearStartMenu {
  If ($ClearStart) {
    Write-Host 'INFO: Configuring empty start menu for new user profiles'
    #Don't edit this. Creates empty start menu if -ClearStart is used.
    $StartLayoutStr = @'
<LayoutModificationTemplate Version="1" xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification" xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout" xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout" xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout">
  <LayoutOptions StartTileGroupCellWidth="6" />
  <DefaultLayoutOverride>
    <StartLayoutCollection>
      <defaultlayout:StartLayout GroupCellWidth="6" xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout">
      </defaultlayout:StartLayout>
    </StartLayoutCollection>
  </DefaultLayoutOverride>
  </LayoutModificationTemplate>
'@
    Add-Content -Path $Env:TEMP\startlayout.xml -Value $StartLayoutStr
    Import-StartLayout -layoutpath $Env:TEMP\startlayout.xml -mountpath $Env:SYSTEMDRIVE\
    Remove-Item -Path $Env:TEMP\startlayout.xml
  }    Else {        
    Write-Host 'INFO: Configuring clean start menu for new profiles'
    #Custom start layout XML near the top of the script.

    Add-Content -Path $Env:TEMP\startlayout.xml -Value $StartLayoutStr
    Import-StartLayout -layoutpath $Env:TEMP\startlayout.xml -mountpath $Env:SYSTEMDRIVE\
    Remove-Item -Path $Env:TEMP\startlayout.xml
  }
}
Function Get-Goodbye {
  Write-Host 'INFO: This Windows 10 device has now been debloated '
}

#endregion


#region Script Block
Clear-Host
Write-host 'Version information - You are running script version 1.2' -ForegroundColor $White -BackgroundColor $DarkGray
@'
  ┌─────────────────────────────────────────────────────────────┐
                Removal of Windows 10 Bloatware

               Follow @thewatchernode on Twitter                               
  └─────────────────────────────────────────────────────────────┘
'@
Start-Transcript -Path $env:HOMEDRIVE\Windows\Temp\Win10BloatLog.txt
If ($AppsOnly) {
  Get-RemoveApps
    Get-ClearStartMenu
    Get-Goodbye
}Elseif ($SettingsOnly) {
    Get-DisableTasks
    Get-DisableServices
    Get-RegChange
    Get-ClearStartMenu
    Get-Goodbye
}Else {
    Get-RemoveApps
    Get-DisableTasks
    Get-DisableServices
    Get-RegChange
    Get-ClearStartMenu
    Get-Goodbye
}
Stop-Transcript
#endregion

