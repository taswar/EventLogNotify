<#
    ** THIS SCRIPT IS PROVIDED WITHOUT WARRANTY, USE AT YOUR OWN RISK **     
     
    .SYNOPSIS 
        Get latest error event based on source and show as a notify ballon
 
    .DESCRIPTION 
        This script monitors system event log for latest error and display them as a NotifyIcon Button when the error happens.                 
 
    .REQUIREMENTS 
        1.    Set-Executionpolicy remotesigned 
 
    .NOTES 
        Tested with Windows 7
		To remove the Event one needs to find out event id by 
		Get-EventSubscriber 
		
		To remove use
		
		Unregister-Event [EventId]
 
    .AUTHOR 
        Taswar Bhatti | http://taswar.zeytinsoft.com    
		
	.INSTALL
		Place banana.ico in C:\Temp or change line 50 to location of icon
		Modify Application Error to the event log source name to monitor
		example '.Net Runtime', 'ASP.NET 4.0.30319.0', 'WMI', etc	
#> 

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#change the name of Application Error to whichever source name you wish to monitor
#query name of the source in eventlog
$query = "Select * from __InstanceCreationEvent 
Where TargetInstance ISA 'Win32_NTLogEvent' 
AND TargetInstance.logfile='Application' 
AND TargetInstance.EventType = 1 
AND TargetInstance.SourceName = 'Application Error'"

Register-WmiEvent -Query $query  -Action { 	
	$global:myevent = $event	
    
	$message = $event.SourceEventArgs.NewEvent.TargetInstance.Message
		
	$now = Get-Date
	
	$objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon 

	$objNotifyIcon.Icon = "C:\temp\banana.ico"
	
	$objNotifyIcon.BalloonTipIcon = "Error" 
			
	$objNotifyIcon.BalloonTipText = "Error : $message"
	
	$objNotifyIcon.BalloonTipTitle = "Application Error $now"
	 
	$objNotifyIcon.Visible = $True 
	
	$objNotifyIcon.ShowBalloonTip(10000)
	
	Start-Sleep -Milliseconds 10000
	
	$objNotifyIcon.Dispose()
}