Import-Module AzureAD
#connect-azuread
$DomainName = "Morgana1.onmicrosoft.com"
$ExportCSV = "C:\Temp\AzureUsers_active.csv"

$Measure_ImportCSV = Measure-Command {
    $Users = Import-Csv $ExportCSV
    $users | Out-Null
}
"Import from CSV (total minutes): " + $Measure_ImportCSV.TotalMinutes

[int]$i = 0
$TotalUsers = $Users.count

foreach($User in $Users)
{
    $i++
    Write-Progress -Activity "Processing $($User.UserPrincipalName)" -Status "$($i) out of $($TotalUsers)..."
    try {
        $Filter = "$($($user.Mail).replace('@','_'))#EXT#@$($DomainName)"
        $UserAD = get-azureaduser -Filter "userprincipalname eq '$($filter)' and UserType eq 'Guest'"
        if ($UserAD) {
            Write-Host "User $($user.UserPrincipalName) does already exist. Skip it" -ForegroundColor Yellow
        }
        elseif($User.Mail) {
            $InviteResult = New-AzureADMSInvitation -InvitedUserDisplayName $User.DisplayName -InvitedUserEmailAddress $user.Mail -SendInvitationMessage $true `
                -InviteRedirectUrl "https://morgana1.sharepoint.com/sites/UkraineTest"
            
            Set-AzureADUser -ObjectId $InviteResult.ID -GivenName $user.GivenName -Surname $user.Surname -City $user.City -CompanyName $user.CompanyName `
                -Country $user.Country

            Write-Host "Invitation send to $($user.Mail) and updated him" -ForegroundColor Green
        }
        elseif(!$user.Mail){Write-Host "User $($user.UserPrincipalName) has no mail. Skip it" -ForegroundColor Yellow}

    }
    catch {
        Write-Host "Error occured for $($user.UserPrincipalName)" -ForegroundColor Yellow
        Write-Host $_ -ForegroundColor Red
    }
    Write-Progress -Activity "Processing $($User.UserPrincipalName)" -Status "$($i) out of $($TotalUsers) completed"
}

#$Filter = "$($($users.Mail).replace('@','_'))#EXT#@$($DomainName)"
#$filter
#get-azureaduser -Filter "userprincipalname eq '$($filter)' and UserType eq 'Guest'"
#get-azureaduser -Filter "userprincipalname eq 'Marius.Deffner_wirtgen-group.com#EXT#@Morgana1.onmicrosoft.com' and UserType eq 'Guest'"