###########################################################################
# Powershell Script to download and install
# Microsoft Windows 10 Software Updates at reboot and shutdown.
#
# Name: Download_and_Install_Win_Update.ps1
#
# Author: marcus s.
# Contact: https://github.com/dagrachon
#
# -------------------------------------------------------------------------
#
# You have a royalty-free right to use, modify, reproduce, and
# distribute this script file in any way you find useful, provided that
# you agree that the creator, owner above has no warranty, obligations,
# or liability for such use.
#
# -------------------------------------------------------------------------
#
# Version History:
#
# 2017-06-12: initial release;
# 2017-06-23: complete Script-rework after having multiple issues 
#             with Windows Upgrades;
#             switched to Windows Update PowerShell Module;
# 2017-06-26: implemented Notification-Task
#
# -------------------------------------------------------------------------
#
# Requires Powershell Version 3.0 or higher
#
###########################################################################

# customize reports (at least recommended for testing purposes)
$FileReport = $true
$FileReportPath = 'C:\Admin\Windows Update Reports\'
$Path = $FileReportPath + "$env:ComputerName" + ".txt"

#is an internet connection up?
if (!(Test-Connection -ComputerName google.com -Count 1 -Quiet))
    {
        if ($FileReport -eq $true) 
		{
			(Get-Date -Format dd-MM-yyyy_HH-mm).ToString(),
            "No network connection up. WindowsUpdate could not be started.`r`n" | Out-File -Append -FilePath $path
        }
    exit
    }

#Testing if there are any pending reboots from earlier Windows Update sessions

$RegistryKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
$RegistryEntry = "InstallWindowsUpdates"

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
    {
        #Report to file if enabled
		if ($FileReport -eq $true) 
		{
			(Get-Date -Format dd-MM-yyyy_HH-mm).ToString(),
			"WindowsUpdate was run on $env:ComputerName, please run WindowsUpdate again when the client is rebooted.`r`n" | Out-File -Append -FilePath $path
		}
        exit        
    } 
 
# Search for available Updates

$Updates = @(Get-WUList -UpdateType Software)
if ($Updates.Length -eq 0)
    {
    # no new updates there, nothing to do.
        if ($FileReport -eq $true)
        {
            (Get-Date -Format dd-MM-yyyy_HH-mm).ToString(),
		    "WindowsUpdate was run on $env:ComputerName, no new updates were found. `r`n" | Out-File -Append -FilePath $path
        }
        exit
    }
else
    {
    # Any major Upgrades there?
    #
    # If so, a ScheduledTask will be registered/activated to notify the next user logging in to install the Upgrade.
    # You can manipulate the $principal variable to notify an specific user (e.g. your administrator) 
    #
    # more about creating ScheduledTasks:
    # https://technet.microsoft.com/de-de/library/jj649811(v=wps.630).aspx
    #
    # more about notifying users through the action-center:
    # https://docs.microsoft.com/en-us/uwp/api/windows.ui.notifications.toastnotificationmanager

    $taskName = 'UpgradeInstallation'

    $Upgrades = @(Get-WUList -RootCategories Upgrades)

        if (!($Upgrades.Length -eq 0))
        {
            if (Get-ScheduledTask | Where-Object {$_.TaskName -like $taskName })
            {
                Enable-ScheduledTask -TaskName $taskName
                (Get-Date -Format dd-MM-yyyy_HH-mm).ToString(),
                "Upgrade installation after reboot required! Next User logging in will be notified.`r`n" | Out-File -Append -FilePath $path
            }
            else
            {
                $time = (Get-Date).AddMinutes(2).ToLongTimeString()
                $trigger = New-ScheduledTaskTrigger -AtLogOn
                $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew
                $principal = New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545" -RunLevel Highest
                $action = New-ScheduledTaskAction -Execute powershell.exe -Argument '-NoProfile -NoLogo -NonInteractive -File  C:\Admin\WindowsUpgradeNotification.ps1'

                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal
                # the notificationscript is found here: https://gist.github.com/altrive/72594b8427b2fff16431 
                (Get-Date -Format dd-MM-yyyy_HH-mm).ToString(),
                "Upgrade installation after reboot required! Next User logging in will be notified.`r`n" | Out-File -Append -FilePath $path
            }
        }
        else
        {
            if (Get-ScheduledTask | Where-Object {$_.TaskName -like $taskName })
            {
                Disable-ScheduledTask -TaskName $taskName
            }
        }

    #Install the updates:
    if ($FileReport -eq $true)
        {
		(Get-Date -Format dd-MM-yyyy_HH-mm).ToString(),
        (Get-WUInstall -UpdateType Software -NotCategory Upgrades -AcceptAll -AutoReboot) | Out-File -Append -FilePath $path
        }
    else
        {
        Get-WUInstall -UpdateType Software -NotCategory Upgrades -AcceptAll -AutoReboot
        }
    }