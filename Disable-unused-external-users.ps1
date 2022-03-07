cls
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$DomainController = "de-dco-202"
$path = "OU=External User,OU=Usermanagement,DC=cellent,DC=int"
$comment = ""
#get-aduser -searchbase $path -filter * -properties name, AccountExpires, lastlogontimestamp, Enabled | select Name, @{N='AccountExpires'; E={[DateTime]::FromFileTime($_.AccountExpires)}}, @{N='lastlogontimestamp'; E={[DateTime]::FromFileTime($_.lastlogontimestamp)}}, Enabled | Export-Csv 'c:\temp\externalUsers.csv' -notype

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) OU to check for accounts enabled, but not loged on last 14 days: $($path)" | out-file $LogFile -Append
    
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Get list of users..." | out-file $LogFile -Append
$users = get-aduser -searchbase $path -filter * -properties name, AccountExpires, lastlogontimestamp, Enabled, info, sAMAccountName  -Server $DomainController  | select Name, @{N='AccountExpires'; E={[DateTime]::FromFileTime($_.AccountExpires)}}, @{N='lastlogontimestamp'; E={[DateTime]::FromFileTime($_.lastlogontimestamp)}}, Enabled, info, sAMAccountName


foreach($user in $users)
{
    if(($user.lastlogontimestamp -lt (Get-Date).AddDays(-14)) -and (-not($user.AccountExpires -lt (Get-Date))) -and ($user.Enabled))
    {
        $Text = "!! ERROR - Last logon older than 14 days, Account expires not set and user is enabled: $($user.Name) - lastlogontimestamp: $($user.lastlogontimestamp) - AccountExpires: $($user.AccountExpires)"
        $Text
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       $($Text)" | out-file $LogFile -Append
        $comment = $($user.info) #get-aduser -Identity $user -Properties info | select Info
        $comment = "($(get-date)) by Script: disabled account while last logon older than 14 days`r`n" + $comment
        #$comment
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             add AD-User infotext and disable the account..." | out-file $LogFile -Append
        Set-AdUser $user.sAMAccountName -Replace @{Info=$comment} -Server $DomainController
        Disable-ADAccount $user.sAMAccountName -Server $DomainController
        #get-aduser $user.sAMAccountName -Properties Enabled | select Enabled
        if ((get-aduser $user.sAMAccountName -Properties Enabled -Server $DomainController | select Enabled).enabled -eq $true)
        {
            $Text = "                  ...error - still enabled"
            $Text
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($Text)" | out-file $LogFile -Append
        }
        else
        {
            $Text = "                  ...now disabled"
            $Text
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($Text)" | out-file $LogFile -Append
        }
        exit
    }
    if (($user.AccountExpires -lt (Get-Date)))
    {
        $Text = "Info - AccountExpires is in the past. So, it is disabled: $($user.Name) - AccountExpires: $($user.AccountExpires)"
        $Text
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       $($Text))" | out-file $LogFile -Append
    }
    if (($user.Enabled -eq $false))
    {
        $Text = "Info - User is disabled: $($user.Name)"
        $Text
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       $($Text))" | out-file $LogFile -Append
    }
    if ($user.lastlogontimestamp -gt (Get-Date).AddDays(-14))
    {
        $Text = "Info - Last logon newer than 14 days: $($user.Name) - lastlogontimestamp: $($user.lastlogontimestamp)"
        $Text
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       $($Text))" | out-file $LogFile -Append
    }
}

#(get-aduser 'csiebert' -Properties Enabled).Enabled