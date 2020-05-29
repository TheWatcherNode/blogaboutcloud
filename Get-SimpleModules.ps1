﻿     # Self-elevate the script if required
    if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
     if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
      $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
      Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
      Exit
     }
    }


#region - Important Section 
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
Install-PackageProvider -name NuGet -Force

#endregion

#region - Install Modules 
Install-module -Name MSOnline
Install-Module -name ImportExcel
Install-Module -Name CloudConnector