###########################################################################"
#
# NAME: Download_and_Install_Win_Update.ps1
#
# AUTHOR: Jan Egil Ring
# EMAIL: jan.egil.ring@powershell.no
#
# COMMENT: Script to download and install updates from Windows Update/WSUS. Reporting and rebooting may be customized.
#          For more details, see the following blog-post:
#          http://blog.powershell.no/2010/06/25/manage-windows-update-installations-using-windows-powershell
#
# You have a royalty-free right to use, modify, reproduce, and
# distribute this script file in any way you find useful, provided that
# you agree that the creator, owner above has no warranty, obligations,
# or liability for such use.
#
# VERSION HISTORY:
# 1.0 25.06.2010 - Initial release
# 1.1 06.12.2013 - Modifications by Brian Clark
# 2.0 02.06.2017 - Modifications by Markus S.
#			Modifications depending on the purpose to use this script at windows 10 client shutdown
#
#	Requires Powershell Version 2.0
#
###########################################################################
#
# using this script at shutdown, the shutdown commands in the script are not needed.
# as is the mail report function not needed in my actual environment.
#
###########################################################################

$FileReport = $true
$FileReportPath = "C:\Admin\Windows Update Reports\"
#$AutoShutdown = $true
#$AutoShutdownIfPending = $true
 
$Path = $FileReportPath + "$env:ComputerName" + ".txt"

#Testing if a internet connection is up
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
if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
    {
		#Report to file if enabled
		if ($FileReport -eq $true) 
		{
			(Get-Date -Format dd-MM-yyyy_HH-mm).ToString(),
			"WindowsUpdate was run on $env:ComputerName, please run WindowsUpdate again when the client is rebooted.`r`n" | Out-File -Append -FilePath $path
		}
 
#		#Shutdown if autoshutdown for pending updates is enabled
#		if ($AutoShutdownIfPending) 
#		{
#			shutdown.exe /t 0 
#		}
#    exit
	}	
#else
#    {
#		#Report to file if enabled
#		if ($FileReport -eq $true) 
#			{
#				(Get-Date -Format dd-MM-yyyy_HH-mm).ToString(),
#				"WindowsUpdate - check for reboot was run on $env:ComputerName, no client reboot required.`r`n" | Out-File -Append -FilePath $path
#			}
#	}


$criteria="IsInstalled=0 and IsHidden=0 and Type='Software'" 
$updateSession = new-object -com "Microsoft.Update.Session"
write-progress -Activity "Updating" -Status "Checking available updates"  
$updates=$updateSession.CreateupdateSearcher().Search($criteria).Updates
$downloader = $updateSession.CreateUpdateDownloader() 
$downloader.Updates = $Updates 
        
if ($downloader.Updates.Count -eq "0") 
	{
		#Report to file if enabled
		if ($FileReport -eq $true) 
		{
			(Get-Date -Format dd-MM-yyyy_HH-mm).ToString(),
			"WindowsUpdate was run on $env:ComputerName, but no new updates were found.`r`n" | Out-File -Append -FilePath $Path
		}
    exit
	}
else 
	{
		#If updates are available, check to accept Eulas
        For($i=0; $i -lt $downloader.Updates.Count; $i++)
        {
            $update=$downloader.Updates.Item($i)
            $update.EulaAccepted
            if(-not $update.EulaAccepted)
            {
                write-progress -Activity 'Preparing' -Status "Accepting EULA for $update"
                $update.AcceptEula()
            }
        }

        #If updates are available, download and install
		write-progress -Activity 'Updating' -Status "Downloading $($downloader.Updates.count) update(s)"
		
		$resultcode= @{0="Not Started"; 1="In Progress"; 2="Succeeded"; 3="Succeeded With Errors"; 4="Failed" ; 5="Aborted" }
		$Result= $downloader.Download()
 
		if (($Result.Hresult -eq 0) -and (($result.resultCode -eq 2) -or ($result.resultCode -eq 3)) ) 
		{
			$updatesToInstall = New-object -com "Microsoft.Update.UpdateColl"
            
			$Updates | where {$_.isdownloaded} | foreach-Object {$updatesToInstall.Add($_) | out-null}
 
			$installer = $updateSession.CreateUpdateInstaller()      
			$installer.Updates = $updatesToInstall
 
			write-progress -Activity 'Updating' -Status "Installing $($Installer.Updates.count) update(s)"        
 
			$installationResult = $installer.Install()        
			$Global:counter=0      
 
			$Report = $installer.updates | 
				Select-Object -property @{Name='Datum';expression={(Get-Date -Format dd-MM-yyyy_HH-mm).ToString()}},Title,EulaAccepted,@{Name='Result';expression={$ResultCode[$installationResult.GetUpdateResult($Global:Counter).resultCode ] }},@{Name='Reboot required';expression={$installationResult.GetUpdateResult($Global:Counter++).RebootRequired }} 
		}
		#Report to file if enabled
		if ($FileReport -eq $true) 
		{
			$Report | Out-File -Append -FilePath $path
		}
#        #Shutdown if AutoShutdown is enabled and one or more updates are requiring a reboot
#        if ($AutoShutdown -and $installationResult.rebootRequired) 
#        {
#            shutdown.exe /t 0 
#        }
	}