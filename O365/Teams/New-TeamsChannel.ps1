$TeamOwner = "Stefan.Stude@wirtgen.de" #Read-Host "Enter the MailID of the Owner"
$TeamName = "WI Controlling GmbH" #Read-Host "Enter the Name of the Team"
$TeamDescription = "WI Controlling GmbH" #Read-Host "Enter the Description of the Team"
$TicketNumber = "WGSP-117305" #Read-Host "Enter the Ticketnumber of the Teams Channel Request"

#$TeamOwner = "Marius.Deffner@wirtgen-group.com"
#$TeamName = "Marius PS Test Team"
$MailNickName = $TeamName -replace '\s',''
#$TeamDescription = "Marius: Team dient zum Test der automatischen Erstellung neuer Channels. Wird wieder gelöscht"
#$TicketNumber = "999999"




function CheckTeamMember ([string]$GroupID, [string]$UserID)
{
    $Owners = Get-TeamUser -GroupId $GroupID -Role Owner | Select-Object -Property User
    return $false
    foreach($Owner in $Owners)
    {
        if ($Owner.User = $UserID)
        {return $true}
    }
}


#Azure connection
try {
    $conTeams = Connect-MicrosoftTeams


    #Validate TeamOwner
    if (-not(get-azureaduser -ObjectId $TeamOwner)){
    write-host "Team Owner does not exist. Exit Script"
    exit
    }
    else{Write-Host "Team Owner verified. Continue Script"}

    #Create Team
    if (-not(Get-Team -DisplayName $TeamName))
    {
        $result = New-Team -DisplayName "$($TeamName)" -Description "$($TicketNumber): $($TeamDescription)" -Visibility Private -MailNickName $MailNickName
        Write-Host "Group ID: $($result.GroupId)"
        if (-not($result.GroupId)){
            
            Write-Host "Team creation failed. Quit Script"
        }
        else {
            try {
                Write-Host "Team successfully created. Add Owner to the Team..."
                #Add Owner        
                Add-TeamUser -GroupId $result.GroupID -User $TeamOwner -Role Owner
                
                if (CheckTeamMember -GroupID $result.GroupID -UserID $TeamOwner)
                {Write-Host "User '$($TeamOwner)' is now Owner"}
                else {
                    Write-Host "Error: Make User '$($TeamOwner)' an Owner failed"
                    exit
                }
                else {
                    #Remove Admin from users
                    Remove-TeamUser -GroupId $result.GroupId -User $conTeams.Account
                }
            }            
            catch {
                "Error occurred:"
                ""
                $_
            }
        }
    }     
    else {
        Write-Host "Team already exists. Exit script"
        exit
    }      
}
catch {
    "Error occurred:"
    ""
    $_
}
Finally {
    $conTeams = $null
}




