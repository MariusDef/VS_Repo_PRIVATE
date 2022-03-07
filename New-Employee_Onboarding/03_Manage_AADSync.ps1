#################################################################################
#   Author: Marius Deffner, cellent GmbH, 13.11.2018
#
#   Possible Exit Codes:
#      0:  Successfull
#      7:  Any job in script failed
#
#
#################################################################################



$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0
$AADSyncSrv = "DE-AAD-210"



"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append









##############################################################################################
#   Connect to AAD Sync Server
##############################################################################################

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Connect to AAD Sync Server for Start Sync (Server: $($AADSyncSrv))..." | out-file $LogFile -Append
try
{
    $SessionOnline = New-PSSession -ComputerName $AADSyncSrv -Authentication Kerberos
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Import Commands..." | out-file $LogFile -Append
    Import-PSSession $SessionOnline -CommandName Start-ADSyncSyncCycle, Get-ADSyncScheduler -AllowClobber
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             successfully connected and imported" | out-file $LogFile -Append
}
catch
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
    $Result = 7
}


##############################################################################################
#   Set Funktion for wait process
##############################################################################################
Function Wait-ForADSync
{
    do
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Wait for Sync. Is currently in sync: '$((Get-ADSyncScheduler).SyncCycleInProgress)'..." | out-file $LogFile -Append   
        Start-Sleep 10
    }while ((Get-ADSyncScheduler).SyncCycleInProgress -eq $True)
}


##############################################################################################
#   Manage Sync
##############################################################################################
if ($Result -ne 7)
{  
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage Sync progress..." | out-file $LogFile -Append
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Current Sync status (Is currently in sync): '$((Get-ADSyncScheduler).SyncCycleInProgress)'" | out-file $LogFile -Append
    if ((Get-ADSyncScheduler).SyncCycleInProgress -eq $false)
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       start Delta Sync..." | out-file $LogFile -Append
        Start-ADSyncSyncCycle -PolicyType Delta
        Wait-ForADSync
    }
    else
    {
        Wait-ForADSync
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       start Delta Sync..." | out-file $LogFile -Append
        Start-ADSyncSyncCycle -PolicyType Delta
        Wait-ForADSync
    }
}









"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result