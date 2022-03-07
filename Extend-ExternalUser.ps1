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

if ($UserDisplayname.IndexOf("(") -gt 0)
{
    $UserID = ($UserDisplayname.Substring($UserDisplayname.IndexOf("(")+1)).trimend(")")
    #$UserID = $CorrectUserID
}

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Parameter: UserDisplayName: '$($UserDisplayname)'; ChangeValue: '$($ChangeValue)" | out-file $LogFile -Append



$OldExpireDate = (Get-ADUser -Identity $UserID -Properties accountExpires -Server $DomainController).accountExpires
$oldExpireDate = [datetime]::FromFileTime($OldExpireDate)

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Actual accountExpires from AD: '$($OldExpireDate)'" | out-file $LogFile -Append





Try
{
    Set-ADAccountExpiration -Identity $UserID -DateTime $ChangeValue -Server $DomainController
}
Catch
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
}  



$CheckValue = (Get-ADUser -Identity $UserID -Properties accountExpires -Server $DomainController).accountExpires
$Checkvalue = [datetime]::FromFileTime($CheckValue)

if ($CheckValue -eq $ChangeValue)
{
    $Result = 0
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - accountExpires successfully changed" | out-file $LogFile -Append
}
else
{
    $Result = 7
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR - Change accountExpires failed!" | out-file $LogFile -Append
}


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result