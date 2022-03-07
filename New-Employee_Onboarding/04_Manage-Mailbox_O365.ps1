#################################################################################
#   Author: Marius Deffner, cellent GmbH, 17.10.2018
#
#   Possible Exit Codes:
#      0:  Successfull
#      7:  Any job in script failed
#      8:  Timeout. Migration is working and not finished in a timely matter. License mapping missing
#      9:  Credentials for O365 missing
#      6: Credentials for Exchange On Prem Missing
#      4: Cannot set UsageLocation
#      3: Cannot set License
#
#
#################################################################################

Param
(
    [Parameter(Mandatory = $true)]
    [string] $Surname,
    [Parameter(Mandatory = $true)]
    [string] $Prename,
    [Parameter(Mandatory = $true)]
    [int] $PersonalNumber
)


$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0
$ExchangeSrv = "DE-EXC-211"



"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Parameter:" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Surname: '$($Surname)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Prename: '$($Prename)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Personal Number: '$($PersonalNumber)'" | out-file $LogFile -Append

$DomainController = "de-dco-202.cellent.int" #(get-ADDomainController -Discover).hostname
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Use Domaincontroller: '$($DomainController)'" | out-file $LogFile -Append



##############################################################################################
#   Im AD prüfen auf Parametrisierten User und holen von Userobjekt
##############################################################################################
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Check, if user already exits..." | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Variables for existance check: PreName, Surname, PersonalNumber" | out-file $LogFile -Append
    
if ($CurrentUser = Get-ADUser -Server $DomainController -Filter {(givenname -eq $Prename) -and (sn -eq $Surname) -and (ExtensionAttribute3 -eq $PersonalNumber)})
{
    $UserFound = $CurrentUser.sAMAccountName
    $Logonname = $UserFound
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       User found with sAMAccountName: $($UserFound)" | out-file $LogFile -Append
}
else 
{
    $UserFound = $false
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       User not found" | out-file $LogFile -Append
    $Result = 7
}




if ($Result -ne 7)
{  
    
    ##############################################################################################
    #   Verwalten der notwendigen Credentials
    ##############################################################################################
    #Read-Host -Prompt "Enter PW for MSOnline User 'svc_itsc@cellent.de'" -AsSecureString | ConvertFrom-SecureString | Out-File ("\\de-frs-201\e$\Heat_Automation\New-Employee_Onboarding\Manage-Mailbox_online.ps1_O365_cred.txt")
    #Read-Host -Prompt "Enter PW for MSOnline User 'cellent\svc_de_heat'" -AsSecureString | ConvertFrom-SecureString | Out-File ("\\de-frs-201\e$\Heat_Automation\New-Employee_Onboarding\Manage-Mailbox_online.ps1_OnPrem_cred.txt")

    $UsernameO365 = "svc_itsc@cellent.de"
    $PasswordO365 = Get-Content ($MyInvocation.MyCommand.Definition + "_O365_cred.txt") | ConvertTo-SecureString
    #$PasswordO365 = Get-Content ("\\de-frs-201\e$\Heat_Automation\New-Employee_Onboarding\Manage-Mailbox_O365.ps1_O365_cred.txt") | ConvertTo-SecureString
    $MSOLCred = New-Object System.Management.Automation.PSCredential $UsernameO365, $PasswordO365
    #$MSOLCred = Get-Credential
    if ($MSOLCred -eq $null){$Result = 9}

    $UsernameOnPrem = "cellent\svc_de_heat"
    $PasswordOnPrem = Get-Content ($MyInvocation.MyCommand.Definition + "_OnPrem_cred.txt") | ConvertTo-SecureString
    #$PasswordOnPrem = Get-Content ("\\de-frs-201\e$\Heat_Automation\New-Employee_Onboarding\Manage-Mailbox_O365.ps1_OnPrem_cred.txt") | ConvertTo-SecureString
    $ExcOPremCred = New-Object System.Management.Automation.PSCredential $UsernameOnPrem, $PasswordOnPrem

    if ($ExcOPremCred -eq $null){$Result = 6}
    
    
    if (($Result -ne 9) -and ($Result -ne 6))
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Connect to Exchage Online for Mailbox management (Server: $($ExchangeSrv))..." | out-file $LogFile -Append
        try
        {
            $SessionOnline = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $MSOLCred -Authentication Basic -AllowRedirection
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Import Commands..." | out-file $LogFile -Append
            Import-PSSession $SessionOnline -CommandName Get-Mailbox, new-moveRequest, get-moverequest -AllowClobber
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             successfully connected and imported" | out-file $LogFile -Append
        }
        catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
            $Result = 7
        }
    }

    ##############################################################################################
    #   Wenn Session offen dann prüfen, ob Mailbox bereits migriert
    ##############################################################################################
    if (($Result -ne 7) -and ($Result -ne 9) -and ($Result -ne 6))
    {
        Try
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Check, if Mailbox already moved to O365..." | out-file $LogFile -Append
            if (Get-Mailbox -Identity "$($Prename).$($Surname)@cellent.de")
            {
                $MailboxExist = $true
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Mailbox is already migrated. Skip migration" | out-file $LogFile -Append
            }
            Else
            {
                $MailboxExist = $False
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Mailbox is not migrated. Start migration" | out-file $LogFile -Append
            }
        }
        catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
            $Result = 7
        }
        
        #Get-PSSession | Remove-PSSession

        ##############################################################################################
        #   Wenn Session noch offen und Mailbox nicht migriert -> migrieren
        ##############################################################################################
        if (($Result -ne 7)-and ($MailboxExist -eq $false))
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Start Mailbox migration and wait 10 Minutes for migration state 'completed'..." | out-file $LogFile -Append
            try
            {
                $bla = New-MoveRequest -Identity "$($Prename).$($Surname)@cellent.de" -Remote -RemoteHostName "owa.cellent.de" -TargetDeliveryDomain cellentAG.mail.onmicrosoft.com -RemoteCredential $ExcOPremCred -BadItemLimit 1000
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Result from Move-Request: $($bla)" | out-file $LogFile -Append
                [int]$i = 0
                do
                {
                    $i +=1
                    sleep -Seconds 10
                    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             i: $($i) - Status: $((Get-MoveRequest | Where-Object DisplayName -EQ "$($Surname), $($Prename)" | Select-Object Status).status)" | out-file $LogFile -Append
                    if ($i -ge 30) {$i=9999; break}
                } while ((Get-MoveRequest | Where-Object DisplayName -EQ "$($Surname), $($Prename)" | Select-Object Status).status -ne "Completed")
                if ($i -eq 9999)
                {
                    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Warning. Script timed out. Quit Script" | out-file $LogFile -Append   
                    $Result = 8
                }
            }
            catch
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
                $Result = 7
            }
        }
    }
}




##############################################################################################
#   Wenn Migration erfolgreich abgeschlossen und Session offen, UsageLocation zuweisen
##############################################################################################
if (($Result -ne 7) -and ($Result -ne 8))
{  
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Connect to O365 to assign Location to user..." | out-file $LogFile -Append
    Try
    {
        Import-Module msonline
        Connect-MsolService -Credential $MSOLCred   


        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Set UsageLanguage to 'DE', if not already set..." | out-file $LogFile -Append
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Current UsageLocation: '$((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property UsageLocation).UsageLocation)'..." | out-file $LogFile -Append
        if ((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property UsageLocation).UsageLocation -ne "DE")
        {
            Set-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" -UsageLocation "DE"    
        }
        if ((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property UsageLocation).UsageLocation -ne "DE")
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Error setting UsageLocation to DE. Quit Script" | out-file $LogFile -Append
            $Result = 4
        }   
        else
        {
             "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Location successfully assigned..." | out-file $LogFile -Append
        }
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
        $Result = 4
    }
}



##############################################################################################
#   Lizenz zuweisen
##############################################################################################
if (($Result -ne 7) -and ($Result -ne 8) -and ($Result -ne 6) -and ($Result -ne 4))
{  
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Connect to O365 to assign License to user..." | out-file $LogFile -Append
    Try
    {
        Import-Module msonline
        Connect-MsolService -Credential $MSOLCred   


        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Assign License..." | out-file $LogFile -Append
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Is currently licensed?: '$((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property isLicensed).isLicensed)'..." | out-file $LogFile -Append
        if ((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property isLicensed).isLicensed -ne "True")
        {
            Set-MsolUserLicense -UserPrincipalName "$($Logonname)@cellent.de" -AddLicenses "cellentAG:ENTERPRISEPACK"
        }
        if ((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property isLicensed).isLicensed -ne "True")
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Error assignung License to user. Quit Script" | out-file $LogFile -Append
            $Result = 3
        }   
        else
        {
             "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             License successfully assigned..." | out-file $LogFile -Append
        }
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
        $Result = 3
    }
}




Get-PSSession | Remove-PSSession



"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result