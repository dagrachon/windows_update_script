Readme:

The following Powershell-Script uses the Windows Update PowerShell Module to Download and Install Windows Updates at Shutdown / Restart.

1. check internet-connection 
2. look for pending reboots
3. search for available updates
- if Updates are available, we look for any Upgrades.
- if Upgrades are available, a scheduled task, to notify the next logged on user, will be registered.
- if Updates are available, the installation process starts.


Usage:

Download and Save the Powershell-Script in to your Shutdown-Script directory:
C:\Windows\System32\GroupPolicy\Machine\Scripts\Shutdown

Activate the GroupPolicy Element to run lokal Powershell-Scripts:
Computer Configuration -> Administrative Templates -> Windows Components -> Windows PowerShell and configure the Turn On Script Execution setting

Configure the GroupPolicy Element
Computerconfiguration -> Windows -> Scripts -> Shutdown:
Use the tab Powershell-Scripts, select the saved script and complete the configuration.

Download and Save the Notification-Script
https://github.com/dagrachon/WindowsUpgradeNotification
Run the notification-script once. 
Afterwards, go to Settings -> System -> Notifications & Actions -> Scroll down to 'Automatic WindowsUpdate' (or however you named the $notifier CreateToastNotifier("xxx")) and activate the setting 'show in action-center'

Install the WindowsUpdate PowerShell Module
https://www.powershellgallery.com/packages/PSWindowsUpdate/2.1.1.2

Create the folder for the Error-Log C:\Admin\Windows Update Logs\ (or wherever you want this logs)
