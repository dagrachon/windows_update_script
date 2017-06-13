Readme:

The following Powershell-Script uses the Windows Update Agent API to Download and Install Windows Updates at Shutdown / Restart.

It's originally from Jan Egil Ring (jan.egil.ring@powershell.no) first published 2010 
(http://blog.powershell.no/2010/06/25/manage-windows-update-installations-using-windows-powershell)
based on a blog-post from James ONeill from 2009 (https://blogs.technet.microsoft.com/jamesone/2009/01/27/managing-windows-update-with-powershell/)

I made a few changes based on user comments on the originally blog-post and my own needs in my current working environment.

Full documentation of the Windows Update Agent API is found here:
https://msdn.microsoft.com/de-de/library/windows/desktop/aa387099(v=vs.85).aspx

Usage:

Download and Save the Powershell-Script in to your Shutdown-Script directory:
C:\Windows\System32\GroupPolicy\Machine\Scripts\Shutdown

Configure the GroupPolicy Element
Computerconfiguration -> Windows -> Scripts -> Shutdown:
Use the tab Powershell-Scripts, select the saved script and complete the configuration.

Activate the GroupPolicy Element to run lokal Powershell-Scripts:
Computer Configuration -> Administrative Templates -> Windows Components -> Windows PowerShell and configure the Turn On Script Execution setting

Create the folder for the Error-Log C:\Admin\Windows Update Logs\ (or wherever you want this logs)