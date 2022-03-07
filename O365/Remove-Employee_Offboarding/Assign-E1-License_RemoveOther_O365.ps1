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
    [string] $Displayname,
    [Parameter(Mandatory = $true)]
    [string] $OnlyRemoveLicense
)


$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0
$DomainController = "de-dco-202.cellent.int"


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Parameter:" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Displayname: '$($Displayname)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       OnlyRemoveLicense: '$($OnlyRemoveLicense)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Use Domaincontroller: '$($DomainController)'" | out-file $LogFile -Append



if ($Displayname.IndexOf("(") -gt 0)
{
    $UserID = ($Displayname.Substring($Displayname.IndexOf("(")+1)).trimend(")")
}


##############################################################################################
#   Im AD prüfen auf Parametrisierten User und holen von Userobjekt
##############################################################################################
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Check, if user already exits..." | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Variables for existance check: LoginID ($($UserID))" | out-file $LogFile -Append
    
$CurrentUser = Get-ADUser -Server $DomainController -Filter {(sAMAccountName -eq $UserID)}

if ($CurrentUser)
{
    $UserFound = $CurrentUser.sAMAccountName
    $Logonname = $UserFound
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       User found with sAMAccountName: $($UserFound)" | out-file $LogFile -Append
}
else 
{
    #$UserFound = $false
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       User not found" | out-file $LogFile -Append
    $Result = 7
}




if ($Result -ne 7)
{  
    
    ##############################################################################################
    #   Verwalten der notwendigen Credentials
    ##############################################################################################
    #Read-Host -Prompt "Enter PW for MSOnline User 'svc_itsc@cellent.de'" -AsSecureString | ConvertFrom-SecureString | Out-File ("E:\Heat_Automation\Remove-Employee_Offboarding\O365_cred.txt")
    #Read-Host -Prompt "Enter PW for MSOnline User 'cellent\svc_de_heat'" -AsSecureString | ConvertFrom-SecureString | Out-File ("E:\Heat_Automation\Remove-Employee_Offboarding\O365_cred.txt")

    $UsernameO365 = "svc_itsc@cellent.de"
    $PasswordO365 = Get-Content ((split-path -parent $MyInvocation.MyCommand.Definition) + "\O365_cred.txt") | ConvertTo-SecureString
    #$PasswordO365 = Get-Content ("E:\Heat_Automation\Remove-Employee_Offboarding\O365_cred.txt") | ConvertTo-SecureString
    $MSOLCred = New-Object System.Management.Automation.PSCredential $UsernameO365, $PasswordO365
    #$MSOLCred = Get-Credential
    if ($MSOLCred -eq $null){$Result = 7}    
}



##############################################################################################
#   Assign UsageLocation zuweisen
##############################################################################################
if ($Result -ne 7)
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
            $Result = 7
        }   
        else
        {
             "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Location successfully assigned..." | out-file $LogFile -Append
        }
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
        $Result = 7
    }
}



##############################################################################################
#   Alle Lizenzen entfernen
##############################################################################################
if ($Result -ne 7)
{  
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Connect to O365 to remove Licenses from user..." | out-file $LogFile -Append
    Try
    {
        Import-Module msonline
        Connect-MsolService -Credential $MSOLCred   


        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Is currently licensed?: '$((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property isLicensed).isLicensed)'..." | out-file $LogFile -Append
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Assigned Licenses: '$($($(Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de").Licenses).AccountSkuID)'..." | out-file $LogFile -Append
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Remove Licenses..." | out-file $LogFile -Append
        
        
        
        ((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de").Licenses).AccountSkuID | Foreach{Set-MsolUserLicense -UserPrincipalName "$($Logonname)@cellent.de" -RemoveLicenses $_}
        
        if ((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property isLicensed).isLicensed -eq "True")
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Error removing Licenses from user. Quit Script" | out-file $LogFile -Append
            $Result = 7
        }   
        else
        {
             "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             License successfully removed..." | out-file $LogFile -Append
        }
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
        $Result = 7
    }
}



if ($OnlyRemoveLicense -eq "False")
{
    sleep -Seconds 8

    ##############################################################################################
    #   Lizenz zuweisen
    ##############################################################################################
    if ($Result -ne 7)
    {  
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Connect to O365 to assign E1 License to user..." | out-file $LogFile -Append
        Try
        {
            Import-Module msonline
            Connect-MsolService -Credential $MSOLCred   


            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Assign E1 License..." | out-file $LogFile -Append

            if ((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property isLicensed).isLicensed -ne "True")
            {
                Set-MsolUserLicense -UserPrincipalName "$($Logonname)@cellent.de" -AddLicenses "cellentAG:STANDARDPACK"
            }
            if ((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property isLicensed).isLicensed -ne "True")
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Error assignung E1 License to user. Quit Script" | out-file $LogFile -Append
                $Result = 7
            }   
            else
            {
                 "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             E1 License successfully assigned..." | out-file $LogFile -Append
            }
        }
        catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
            $Result = 3
        }
    }
}



Get-PSSession | Remove-PSSession



"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result