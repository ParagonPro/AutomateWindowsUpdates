 #Print out update information
 Write-Host 'Obtaining Pending Update Information, please wait.'  
   
 #Begin Transcript  
 $ErrorActionPreference="SilentlyContinue"  
 Stop-Transcript | out-null  
 $ErrorActionPreference="Continue"  
 #Start-Transcript -path 'PendingUpdates.txt' -append  
   
 #Get All Assigned updates in $SearchResult  
 $UpdateSession = New-Object -ComObject Microsoft.Update.Session  
 $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()  
 $SearchResult = $UpdateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0")  
   
 #Matrix Results for type of updates that are needed  
 $Critical = $SearchResult.updates | where { $_.MsrcSeverity -eq "Critical" }  
 $important = $SearchResult.updates | where { $_.MsrcSeverity -eq "Important" }  
 $other = $SearchResult.updates | where { $_.MsrcSeverity -eq $null }  
   
 #Pending Updates Tally  
 Write-Host "=====Pending Updates Tally====="  
 Write-Host "Total Updates: $($SearchResult.updates.count)"  
 Write-Host "Critical Updates: $($Critical.count)" -foregroundcolor "red"  
 Write-Host "Important Updates: $($Important.count)" -foregroundcolor "yellow"  
 Write-Host "Other Updates: $($other.count)"  
   
 Write-Host "===== Pending Updates ====="  
 For($i=0;$i -lt $SearchResult.Updates.Count; $i++)  
 {  
   Write-Host "Item No: $($i + 1)"  
   Write-Host "Item Name: $($SearchResult.Updates.Item($i).Title)"  
   Write-Host "Description:"  
   Write-Host "$($SearchResult.Updates.Item($i).Description)"  
   Write-Host "`n" 
   $UpdateList = $UpdateList + "Item No: $($i + 1) <br/>Item Name: $($SearchResult.Updates.Item($i).Title) <br/>Description: $($SearchResult.Updates.Item($i).Description) <br/><br/>******************<br/><br/>" 
 }

#Send Email with updates

#Set variables
$thisComputer = $env:COMPUTERNAME
$smtpServer = "SMTP SERVER IP" #SMTP SERVER ADDRESS
$smtpFrom = "EMAIL ADDRESS HERE" #THIS IS WHERE THE EMAILS WILL COME FROM
$ReportRebootTo = "EMAIL ADDRESS HERE" #THIS IS THE ADDRESS REBOOTS WILL GO TO

if($SearchResult.Updates.Count -gt 0){

    #Email Variables

    $smtpTo = $ReportRebootTo;
    $messageSubject = $thisComputer + " Windows Update Report"
    $Message = New-Object System.Net.Mail.mailmessage $smtpFrom, $smtpTo
    $Message.IsBodyHtml=$true
    $Message.Subject = $messageSubject

    Write-Host "Sending Update Information Email" 
    $timeStamp = get-date -Format hh:mm
    $todaysDate = get-date -format D
    $RebootResult = "The server <b>" + $thisComputer + "</b> has installed the following updates: <br/><br/>=====Pending Updates Tally=====<br/><br/>Total Updates: $($SearchResult.updates.count)<br/>Critical Updates: $($Critical.count)<br/>Important Updates: $($Important.count)<br/>Other Updates: $($other.count)<br/><br/>===== Pending Updates =====<br/><br/>" + $UpdateList
    $Message.Body = $RebootResult
    $smtp = new-Object Net.Mail.SmtpClient($smtpServer)
    $smtp.Send($message)
    $UpdateList = ""

    #Download and Install Updates
    Write-Host "Downloading and Installing Updates"

    #Define update criteria.

    $Criteria = "IsAssigned=1 and IsHidden=0 and IsInstalled=0"

    #Search for relevant updates.

    $Searcher = New-Object -ComObject Microsoft.Update.Searcher

    $SearchResult = $Searcher.Search($Criteria).Updates

    #Download updates.

    $Session = New-Object -ComObject Microsoft.Update.Session

    $Downloader = $Session.CreateUpdateDownloader()

    $Downloader.Updates = $SearchResult

    $Downloader.Download()

    #Install updates.

    $Installer = New-Object -ComObject Microsoft.Update.Installer

    $Installer.Updates = $SearchResult

    $Result = $Installer.Install()

    #Check if Pending Reboot
    if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA Ignore) { $PendingReboot = 'True' }
    if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) { $PendingReboot = 'True' }
    if (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -EA Ignore) { $PendingReboot = 'True' }
    
    Write-Host "Reboot Pending"

    #If Reboot Is Required
    If ($PendingReboot -eq 'True')
	{ 
        #Email Variables
        $smtpTo = $smtpFrom;
        $messageSubject = $thisComputer + " - Reboot Required"
        $Message = New-Object System.Net.Mail.mailmessage $smtpFrom, $smtpTo
        $Message.IsBodyHtml=$true
        $Message.Subject = $messageSubject

        Write-Host "Done `n Letting IT Know About Pending Reboot"
		$timeStamp = get-date -Format hh:mm
		$todaysDate = get-date -format D
		$RebootResult = "The server <b>" + $thisComputer + "</b> has installed updates and requires a reboot."
		$Message.Body = $RebootResult
        
		$smtp = new-Object Net.Mail.SmtpClient($smtpServer)
		$smtp.Send($message)
		#shutdown.exe /r /t 60 /c "We will be rebooting in 60 seconds to install new updates. Open Command Promt and type 'shutdown -a' to abort"
        Start-Sleep -s 10
        Write-Host "Bye - Bye"
	}
}else{
    exit
}

