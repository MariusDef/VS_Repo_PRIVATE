#$TeamOwner = Read-Host "Enter the MailID of the Owner"
#$TeamName = Read-Host "Enter the Name of the Team"
#$TeamDescription = Read-Host "Enter the Description of the Team"
#$TicketNumber = Read-Host "Enter the Ticketnumber of the Teams Channel Request"

Sub MyConnect-Teams
{
    if ($null -eq $conTeams.Account)
    {
        $conTeams = Connect-MicrosoftTeams
        if ($null -eq $conTeams.Account)
        {
            Write-Host "Connection to Teams failed. Exit Script"   
            Exit
        }
        else{Write-Host "Connected successfully to Teams"}
    }
}


$PlanName = "Plan to Test PS Plan"
$TeamName = "Marius PS Test Team"

MyConnect-Teams

#Azure connection


try {
    $Team = Get-Team -DisplayName $TeamName
    
    $params = @{
        Owner = $Team.GroupID
        Title = $PlanName
    }
    #$conGraph = Connect-MgGraph
    
    $result = New-MgPlannerPlan -BodyParameter $params
    write-host $result

    Get-MgPlannerPlan -Search $params
}
catch {
    "Error occurred:"
    ""
    $_
}




