#################################################################################
#   Author: Marius Deffner, cellent GmbH, 13.11.2018
#
#   Possible Exit Codes:
#      0:  Successfull
#      7:  Any job in script failed
#
#
#################################################################################


Param
(
    [Parameter(Mandatory = $true)]
    [string] $UserDisplayname,
    [Parameter(Mandatory = $true)]
    [DateTime] $ChangeValue
    
)

$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0
$DomainController = "DE-DCO-202"


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append





##############################################################################################
#   Connect to AAD Sync Server
##############################################################################################





if ($UserDisplayname.IndexOf("(") -gt 0)
{
    $UserID = ($UserDisplayname.Substring($UserDisplayname.IndexOf("(")+1)).trimend(")")
    #$UserID = $CorrectUserID
}

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Parameter: UserDisplayName: '$($UserDisplayname)'; ChangeValue: '$($ChangeValue)" | out-file $LogFile -Append


$OldExpireDate = (Get-ADUser -Identity $UserID -Properties accountExpires -Server $DomainController).accountExpires
$oldExpireDate = [datetime]::FromFileTime($OldExpireDate)

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Actual accountExpires from AD: '$($OldExpireDate)'" | out-file $LogFile -Append

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Set accountExpires: '$($ChangeValue)'" | out-file $LogFile -Append

$ChangeValue = $ChangeValue.AddDays(1)

Try
{
    Set-ADAccountExpiration -Identity $UserID -DateTime $ChangeValue -Server $DomainController
}
Catch
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
}  



$CheckValue = (Get-ADUser -Identity $UserID -Properties accountExpires -Server $DomainController).accountExpires
$Checkvalue = [datetime]::FromFileTime($CheckValue)

if ($CheckValue -eq $ChangeValue)
{
    $Result = 0
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       accountExpires successfully changed" | out-file $LogFile -Append
}
else
{
    $Result = 7
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ERROR - Change accountExpires failed!" | out-file $LogFile -Append
}


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())  Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result