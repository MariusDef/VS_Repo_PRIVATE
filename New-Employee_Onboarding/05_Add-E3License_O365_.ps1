#################################################################################
#   Author: Marius Deffner, cellent GmbH, 17.10.2018
#
#   Possible Exit Codes:
#      7: Any job in script failed
#      9: MS Online Credential Error
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
    [int] $PersonalNumber,
    [Parameter(Mandatory = $false)]
    [String] $IsExternalString
)


$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0

#### Manage Variables
if($IsExternalString -eq 'false'){$IsExternal = $false}
if($IsExternalString -eq ''){$IsExternal = $false}
if($IsExternalString -eq 'true'){$IsExternal = $true}


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Parameter:" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Surname: '$($Surname)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Prename: '$($Prename)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Personal Number: '$($PersonalNumber)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       IsExternal: '$($IsExternal)'" | out-file $LogFile -Append

$DomainController = "de-dco-202.cellent.int" #(get-ADDomainController -Discover).hostname
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Use Domaincontroller: '$($DomainController)'" | out-file $LogFile -Append



$Surname = $Surname.Trim()
$Prename = $Prename.Trim()

#### Get current userobject from AD
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
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())                  no. set usage location..." | out-file $LogFile -Append    
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

sleep -Seconds 14

##############################################################################################
#   Lizenz zuweisen
##############################################################################################
if (($Result -ne 7) -and ($Result -ne 8) -and ($Result -ne 6) -and ($Result -ne 4))
{  
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Connect to O365 to assign E3 License to user..." | out-file $LogFile -Append
    Try
    {
        Import-Module msonline
        Connect-MsolService -Credential $MSOLCred   


        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Assign E3 License..." | out-file $LogFile -Append
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Is currently licensed?: '$((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property isLicensed).isLicensed)'..." | out-file $LogFile -Append
        if ((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property isLicensed).isLicensed -ne "True")
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())                  no. set license..." | out-file $LogFile -Append    
            Set-MsolUserLicense -UserPrincipalName "$($Logonname)@cellent.de" -AddLicenses "cellentAG:ENTERPRISEPACK"
        }
        if ((Get-MsolUser -UserPrincipalName "$($Logonname)@cellent.de" | Select-Object -Property isLicensed).isLicensed -ne "True")
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Error assignung E3 License to user. Quit Script" | out-file $LogFile -Append
            $Result = 3
        }   
        else
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