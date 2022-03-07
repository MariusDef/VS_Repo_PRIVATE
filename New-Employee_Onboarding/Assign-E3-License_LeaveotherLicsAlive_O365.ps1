#################################################################################
#   Author: Marius Deffner, cellent GmbH, 27.09.2019
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
    [string] $Displayname
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
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Use Domaincontroller: '$($DomainController)'" | out-file $LogFile -Append



if ($Displayname.IndexOf("(") -gt 0)
{
    $UserID = ($Displayname.Substring($Displayname.IndexOf("(")+1)).trimend(")")
}


##############################################################################################
#   Im AD prüfen auf Parametrisierten User und holen von Userobjekt
##############################################################################################
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Check, if user already exits..." | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Variables for existance check: LoginID '$($UserID)'" | out-file $LogFile -Append
    
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
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage O365 Credentials..." | out-file $LogFile -Append
    
    #Read-Host -Prompt "Enter PW for MSOnline User 'svc_itsc@cellent.de'" -AsSecureString | ConvertFrom-SecureString | Out-File ("E:\Heat_Automation\Remove-Employee_Offboarding\O365_cred.txt")
    #Read-Host -Prompt "Enter PW for MSOnline User 'cellent\svc_de_heat'" -AsSecureString | ConvertFrom-SecureString | Out-File ("E:\Heat_Automation\Remove-Employee_Offboarding\O365_cred.txt")

    $UsernameO365 = "svc_itsc@cellent.de"
    #write-host ((split-path -parent $MyInvocation.MyCommand.Definition) + "\O365_cred.txt")
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
#   Lizenz zuweisen
##############################################################################################
if ($Result -ne 7)
{  
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Connect to O365 to assign E3 License to user..." | out-file $LogFile -Append
    Try
    {
        Import-Module msonline
        Connect-MsolService -Credential $MSOLCred   


        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Assign E3 License..." | out-file $LogFile -Append

        #(Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property Licenses).Licenses
        #{
       Set-MsolUserLicense -UserPrincipalName "$($Logonname)@cellent.de" -AddLicenses "cellentAG:ENTERPRISEPACK"
       #Set-MsolUserLicense -UserPrincipalName "$($Logonname)@cellent.de" -AddLicenses "cellentAG:STANDARDPACK"
        #}
        
        $Success = $false
        (Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property Licenses).Licenses | ForEach-Object {if ($_.AccountSkuID -eq "cellentAG:ENTERPRISEPACK"){$Success = $true} }
        if ($success -eq $false)
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Error assignung E3 License to user. Quit Script" | out-file $LogFile -Append
            $Result = 7
        }   
        elseif ($Success -eq $true)
        {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             E3 License successfully assigned..." | out-file $LogFile -Append
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