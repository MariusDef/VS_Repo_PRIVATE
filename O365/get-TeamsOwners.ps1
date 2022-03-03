﻿#$credentials = Get-Credential

connect-microsoftteams #-Credential $credentials

$TeamColl = get-team

"" | Out-File "C:\Users\wideffnerm\OneDrive - WIRTGEN GROUP\VS_Repo\get-TeamsOwners.ps1.log"

foreach ($team in $TeamColl)
{
    $Owners = Get-TeamUser -GroupId $team.GroupId -Role Owner | Select-Object -Property User
    
    foreach($Owner in $Owners)
    {
        $Owner.User | Out-File "C:\Users\wideffnerm\OneDrive - WIRTGEN GROUP\VS_Repo\get-TeamsOwners.ps1.log" -Append
    }
}

get-command *connection* 
    
Get-CsCloudCallDataConnection