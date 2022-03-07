#################################################################################
#   Author: Marius Deffner, cellent GmbH, 17.10.2018
#
#   Possible Exit Codes:
#      7: Any job in script failed
#      8: Mailbox already exists. Nothing to do
#
#
#################################################################################

Param
(
    [Parameter(Mandatory = $true)]
    [string] $UserDisplayname
)


$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0
$DisabledOU = "OU=Disabled,OU=Internal User,OU=Usermanagement,DC=cellent,DC=int"

#### Manage Variables
if ($UserDisplayname.IndexOf("(") -gt 0)
{
    $Logonname = ($UserDisplayname.Substring($UserDisplayname.IndexOf("(")+1)).trimend(")")
    #$UserID = $CorrectUserID
}


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Parameter:" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       UserDisplayname: '$($UserDisplayname)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Logonname: '$($Logonname)'" | out-file $LogFile -Append

$DomainController = "de-dco-202.cellent.int" #(get-ADDomainController -Discover).hostname
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Use Domaincontroller: '$($DomainController)'" | out-file $LogFile -Append





#### Move User to disabled OU
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Start moveing User to disabled OU..." | out-file $LogFile -Append

    $DNCurrent = (Get-ADUser -Identity $Logonname -Properties distinguishedName -Server $DomainController).distinguishedName

    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       disable User Account..." | out-file $LogFile -Append
    Set-ADUser -Identity $DNCurrent -Enabled $False
    
    
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       current user location: '$($DNCurrent)'" | out-file $LogFile -Append

    

    Move-ADObject -Identity $DNCurrent -TargetPath $DisabledOU -Server $DomainController

    if ((Get-ADUser -Identity $Logonname -Properties distinguishedName -Server $DomainController).distinguishedName -like "*$($DisabledOU)")
    {
    
    }
    else
    {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ERROR: user move failed!" | out-file $LogFile -Append    
            $Result = 7
    }

}





"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result