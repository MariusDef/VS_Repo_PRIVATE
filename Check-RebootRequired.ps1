$LogFile = "C:\Windows\Logs\Check-RebootRequired.log"

$RebootRequired = $false

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) ########################################################## " | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Start Check-Reboot Required Script" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) ########################################################## " | out-file $LogFile -Append

#Check Windows Updates for required reboot #1
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Check Windows Udpates #1..." | out-file $LogFile -Append
if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) REBOOT - Windows Udpate are installed and requires a reboot!" | out-file $LogFile -Append
    $RebootRequired = $true
}
else
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) OK - Windows Udpate doesn't require a reboot" | out-file $LogFile -Append
}

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) ------------------------------------------------" | out-file $LogFile -Append

#Check Windows Updates for required reboot #^2
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Check Windows Udpates #2..." | out-file $LogFile -Append
if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) REBOOT - Windows Udpate are installed and requires a reboot!" | out-file $LogFile -Append
    $RebootRequired = $true
}
else
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) OK - Windows Udpate doesn't require a reboot" | out-file $LogFile -Append
}


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) ------------------------------------------------" | out-file $LogFile -Append

#Check PendingFileRenameOperations
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Check pending file rename operations..." | out-file $LogFile -Append
$prop = Get-ItemProperty "HKLM:SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction Ignore
if($prop -ne $null) 
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) REBOOT - Pending file rename operations requires a reboot!" | out-file $LogFile -Append
    $RebootRequired = $true
}
else
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) OK - Pending file rename operations doesn't require a reboot." | out-file $LogFile -Append
}


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) ------------------------------------------------" | out-file $LogFile -Append

#Final Script Results
if ($RebootRequired -eq $true)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) A reboot is required. Exit Script with Reboot Code: 9 " | out-file $LogFile -Append
    Exit 9
}
else
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) No reboot is required. Exit Script with Exit Code: 0 " | out-file $LogFile -Append
    Exit 0
}