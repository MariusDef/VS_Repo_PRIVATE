#################################################################################
#   Author: Marius Deffner, cellent GmbH, 17.10.2018
#
#   Possible Exit Codes:
#      7: Any job in script failed
#
#      Info: Passwort ist immer 'Secure@cellent99'
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
    [Parameter(Mandatory = $true)]
    [string] $Manager,
    [Parameter(Mandatory = $true)]
    [string] $ManagerMail,
    [Parameter(Mandatory = $true)]
    [string] $CostCenter,
    [Parameter(Mandatory = $false)]
    [string] $JobTitle,
    [Parameter(Mandatory = $false)]
    [string] $Department,
    [Parameter(Mandatory = $false)]
    [string] $Location,
    [Parameter(Mandatory = $true)]
    [String] $ServiceRequestID,
    [Parameter(Mandatory = $false)]
    [String] $IsExternalString,
    [Parameter(Mandatory = $false)]
    [String] $PracticeDepartment,
    [Parameter(Mandatory = $false)]
    [String] $PracticeManager

)


$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0
$ManagerForExternal = "Vertraege, Vertraege (Vertraege)"
$ManagerMailForExternal = "vertraege@cellent.de"
$HomeDriveRoot = "\\de-fls-201\d$\Userhome\"

#### Manage Variables
if($IsExternalString -eq 'false'){$IsExternal = $false}
if($IsExternalString -eq ''){$IsExternal = $false}
if($IsExternalString -eq 'true'){$IsExternal = $true}


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Parameter:" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Surname: '$($Surname)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Prename: '$($Prename)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Personal Number: '$($PersonalNumber)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Manager: '$($Manager)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       CostCenter: '$($CostCenter)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       JobTitle: '$($JobTitle)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Department: '$($Department)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Location: '$($Location)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Company: '$($Company)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       EntryDate: '$($EntryDate)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ServiceRequestID: '$($ServiceRequestID)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       IsExternal: '$($IsExternal)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       PracticeDepartment: '$($PracticeDepartment)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       PracticeManager: '$($IsExtePracticeManagerrnal)'" | out-file $LogFile -Append

$DomainController = "de-dco-202.cellent.int" #(get-ADDomainController -Discover).hostname
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Use Domaincontroller: '$($DomainController)'" | out-file $LogFile -Append



#### Check, if User already exists
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Check, if user already exits..." | out-file $LogFile -Append
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Variables for existance check: PreName, Surname, JobTitle (and Department)" | out-file $LogFile -Append
    
    if ($CurrentUser = Get-ADUser -Server $DomainController -Filter {(givenname -eq $Prename) -and (sn -eq $Surname)})#  -and (Department -eq $Department)
    {
        $UserFound = $CurrentUser.sAMAccountName
        $Logonname = $UserFound
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       User found with sAMAccountName: $($UserFound)" | out-file $LogFile -Append
    }
    else 
    {
        $UserFound = $false
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       User not found" | out-file $LogFile -Append
    }
}

#### Check, Manager if external
if (($Result -ne 7)-and ($IsExternal -eq $true))
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Check, manager for external..." | out-file $LogFile -Append
    
    if ($Manager -ne $ManagerForExternal)
    {
        $Manager = $ManagerForExternal
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Manager was wrong. Change to : $($ManagerForExternal)" | out-file $LogFile -Append
    }
    else 
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Manager was correct" | out-file $LogFile -Append
    }
}

#### Manage sAMAccountName
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Check free sAMAccountName..." | out-file $LogFile -Append
if ($UserFound -eq $false)
{
    [int]$n = 0
    do
    {
        $n = $n + 1
        [String]$Logonname = $Prename.Remove($n) + $Surname
        $Logonname = $Logonname.Replace('ß', 'ss')
        $Logonname = $Logonname.Replace('ä', 'ae')
        $Logonname = $Logonname.Replace('ü', 'ue')
        $Logonname = $Logonname.Replace('ö', 'oe')
        $Logonname = $Logonname.Replace('Ä', 'Ae')
        $Logonname = $Logonname.Replace('Ü', 'Ue')
        $Logonname = $Logonname.Replace('Ö', 'Oe')
    } while (Get-ADUser -Server $DomainController -Filter {samAccountName -eq $Logonname})
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       use Logonname: '$($Logonname)'" | out-file $LogFile -Append
}
else
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       not neccessary, account does already exist" | out-file $LogFile -Append
}



#### Define Company wide 
if ($Result -ne 7)
{
    $Company = "cellent GmbH"
    if($IsExternal -eq $true)
    {
        $Company = 'Extern'
    }
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage Company wide settings..." | out-file $LogFile -Append
    switch ($Company)
    {
        'cellent GmbH' 
        {
            [String]$HomePage = "www.cellent.de"
            [String]$OU="OU=cellent,OU=Internal User,OU=Usermanagement,DC=cellent,DC=int"
        }
        'cMB' 
        {
            [String]$HomePage = "www.cellent-mittelstandsberatung.de"
            [String]$OU="OU=cMB,OU=Internal User,OU=Usermanagement,DC=cellent,DC=int"
        }
        'Extern' 
        {
            [String]$OU="OU=cellent-with-Mailbox,OU=External User,OU=Usermanagement,DC=cellent,DC=int"
        }
        default
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error getting a Company. Quit Script" | out-file $LogFile -Append
        }
    }
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Homepage: $($HomePage)" | out-file $LogFile -Append
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       OU: $($OU)" | out-file $LogFile -Append
}



#### Define Location wide settings
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage Location wide settings..." | out-file $LogFile -Append
    if($IsExternal -eq $true)
    {
        $Location = 'Extern'
    }
    [String]$Phone = "+49 (711) 52030 0"
    switch ($Location)
    {
        'Aalen' 
        {
            [String]$StreetAddress = "Gartenstraße 97"
            [String]$PostalCode = "73430"
            [String]$Country = "DE"
            [String]$Fax ="+49 (7361) 974 444"
            $GroupsList = @("GG_Intranet_Internal_User", "GG_VPN_Group", "GG-MDM-STD-O365", "V AA Alle Mitarbeiter", "V CE Aalen")
        }
        'Fellbach' 
        {
            [String]$StreetAddress = "Ringstraße 70"
            [String]$PostalCode = "70736"
            [String]$Country = "DE"
            [String]$Fax ="+49 (711) 52030 40"
            $GroupsList = @("GG_Intranet_Internal_User", "GG_VPN_Group", "GG-MDM-STD-O365", "V FE Alle Mitarbeiter")
        }
        'Holzgerlingen' 
        {
            [String]$StreetAddress = "Max-Eyth-Straße 38"
            [String]$PostalCode = "71088"
            [String]$Country = "DE"
            [String]$Fax ="+49 (0)7031 62345 19"
            $GroupsList = @("GG_Intranet_Internal_User", "GG_VPN_Group", "GG-MDM-STD-O365", "V HzG Alle Mitarbeiter")
        }
        'Dresden' 
        {
            [String]$StreetAddress = "Freiberger Straße 35"
            [String]$PostalCode = "01067"
            [String]$Country = "DE"
            [String]$Fax ="+49 (0)711 52030 40"
            $GroupsList = @("GG_Intranet_Internal_User", "GG_VPN_Group", "GG-MDM-STD-O365", "V DD Alle Mitarbeiter")
        }
        'Hamburg' 
        {
            [String]$StreetAddress = "Gotenstraße 12"
            [String]$PostalCode = "20095"
            [String]$Country = "DE"
            [String]$Fax ="+49 (0)711 52030 40"
            $GroupsList = @("GG_Intranet_Internal_User", "GG_VPN_Group", "GG-MDM-STD-O365", "V HH Alle Mitarbeiter")
        }
        'Karlsruhe' 
        {
            [String]$StreetAddress = "Willy-Andreas-Allee 19"
            [String]$PostalCode = "76131"
            [String]$Country = "DE"
            [String]$Fax ="+49 (0)721 16113 13"
            $GroupsList = @("GG_Intranet_Internal_User", "GG_VPN_Group", "GG-MDM-STD-O365", "V KA Alle Mitarbeiter")
        }
        'München' 
        {
            [String]$StreetAddress = "Lehrer-Wirth-Str. 2"
            [String]$PostalCode = "81829"
            [String]$Country = "DE"
            [String]$Fax ="+49 (0)89 997273 111"
            $GroupsList = @("GG_Intranet_Internal_User", "GG_VPN_Group", "GG-MDM-STD-O365", "V MU Alle Mitarbeiter")
        }
        'Neu-Ulm' 
        {
            [String]$StreetAddress = "Albrecht-Berblinger-Staße 6"
            [String]$PostalCode = "89231"
            [String]$Country = "DE"
            [String]$Fax ="+49 (0)711 52030 40"
            $GroupsList = @("GG_Intranet_Internal_User", "GG_VPN_Group", "GG-MDM-STD-O365", "V NU Alle Mitarbeiter-1-945565003")
        }
        'Extern' 
        {
            [String]$Country = "DE"
            $GroupsList = @("GG_Intranet_External_User", "V CE Extern")
        }
        Default
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error getting a Location. Quit Script" | out-file $LogFile -Append
        }
    }
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       StreetAddress: $($StreetAddress)" | out-file $LogFile -Append
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       PostalCode: $($PostalCode)" | out-file $LogFile -Append
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Country: $($Country)" | out-file $LogFile -Append
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Fax: $($Fax)" | out-file $LogFile -Append
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       arrGroups: $($arrGroups)" | out-file $LogFile -Append
}



#### Manage Homedrive
if (($Result -ne 7) -and ($IsExternal -ne $true))
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage Homedrive..." | out-file $LogFile -Append   
    [String]$HomeDirectory = "\\cellent.int\fls\Userhome\$($Logonname)"
    [String]$HomeDriveLetter = "U:"

    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       HomeDirectory: $($HomeDirectory)" | out-file $LogFile -Append
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       HomeDriveLetter: $($HomeDriveLetter)" | out-file $LogFile -Append
}
elseif(($Result -ne 7) -and ($IsExternal -eq $true))
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Is is external. No HomeDirectory necessary" | out-file $LogFile -Append
}




#### Manage UserDisplayname
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage DisplayName..." | out-file $LogFile -Append 
    [String]$UserDisplayname = "$($Surname), $($Prename)"
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       UserDisplayname: $($UserDisplayname)" | out-file $LogFile -Append
}




#### Create AD User
if (($Result -ne 7))
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Create the AD User..." | out-file $LogFile -Append 
    if ($UserFound -eq $false)
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       manage jobs for manager object..." | out-file $LogFile -Append
        if ($Manager.IndexOf("(") -gt 0){$ManagersAMAccountName = ($Manager.Substring($Manager.IndexOf("(")+1)).trimend(")")}
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             ManagersAMAccountName: $($ManagersAMAccountName)" | out-file $LogFile -Append
        Try
        {
            #"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             DomainController: $($DomainController)" | out-file $LogFile -Append
            $ManagerObject = Get-ADUser -Identity $ManagersAMAccountName -Server $DomainController
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             found and stored" | out-file $LogFile -Append
        }
        catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             ERROR finding Manager Object: '$($_.Exception.Message)'" | out-file $LogFile -Append
            $Result = 7
        }
    
        if ($Result -ne 7)
        {    
            try
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Create User Object..." | out-file $LogFile -Append
                #New-ADUser $UserParams -AccountPassword (ConvertTo-SecureString -AsPlainText "Secure@cellent99" -Force) -Server $DomainController
                #New-ADUser -Name $UserDisplayname -SamAccountName $Logonname -GivenName $Prename -Surname $Surname -Path $OU -Server $DomainController
                New-ADUser -Name $UserDisplayname -DisplayName $UserDisplayname -Manager $ManagerObject -UserPrincipalName "$($Logonname)@cellent.de" -SamAccountName $Logonname -Fax $Fax -Description $JobTitle -Department $Department -Company $Company -Enabled $true -Path $OU -City $Location -PostalCode $PostalCode -Country $Country -HomeDrive $HomeDriveLetter -HomeDirectory $HomeDirectory -Title $JobTitle -Office $Location -GivenName $Prename -Surname $Surname -ChangePasswordAtLogon $true -OfficePhone $Phone -HomePage $HomePage -StreetAddress $StreetAddress -AccountPassword (ConvertTo-SecureString -AsPlainText "Secure@cellent99" -Force) -Server $DomainController
                #New-ADUser -Name $UserDisplayname -DisplayName $UserDisplayname -SamAccountName $Logonname -Fax $Fax -Department $Department -Company $Company -Enabled $true -Path $OU -City $Location -PostalCode $PostalCode -Country $Country -GivenName $Prename -Surname $Surname -ChangePasswordAtLogon $true -AccountPassword (ConvertTo-SecureString -AsPlainText "Secure@cellent99" -Force) -Server $DomainController
            }
            catch
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ERROR Creating the User: '$($_.Exception.Message)'" | out-file $LogFile -Append
                $Result = 7
            }

            if ($Result -ne 7)
            {
                # Check User
                if (Get-ADUser -Server $DomainController -Identity $Logonname)
                {
                    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       successfully created the User" | out-file $LogFile -Append   
                }
                else
                {
                    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ERROR - Created User not found!" | out-file $LogFile -Append
                    $Result = 7
                }
            }
        }
    }
    else
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       not neccessary, user does already exist" | out-file $LogFile -Append
    }
}


#### Manage expiration if extern
if (($Result -ne 7) -and ($IsExternal -eq $true))
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage expiration of account..." | out-file $LogFile -Append   

    
    Set-ADAccountExpiration -Identity $Logonname -DateTime (get-date).AddMonths(3) -Server $DomainController
    
    $OldExpireDate = (Get-ADUser -Identity $Logonname -Properties accountExpires -Server $DomainController).accountExpires
    $oldExpireDate = [datetime]::FromFileTime($OldExpireDate)

    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Expiration Date: $($oldExpireDate)" | out-file $LogFile -Append
   
}



#### Manage ExtensionAttributes (CostCenter, ManagerMail, PersonalNumber)
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage ExtensionAttributes..." | out-file $LogFile -Append 
    try
    {
        if ($IsExternal -eq $true)
        {
            $ManagerMail = $ManagerMailForExternal
            $PersonalNumber = "99999"
        }

        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       DEBUG: ExtensionAttribute1='$($CostCenter)'; ExtensionAttribute2='$($ManagerMail)'; ExtensionAttribute3='$($PersonalNumber)'" | out-file $LogFile -Append 
        Set-AdUser $Logonname -Add @{ExtensionAttribute1="$($CostCenter)"; ExtensionAttribute2="$($ManagerMail)"; ExtensionAttribute3="$($PersonalNumber)"} -Server $DomainController
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ExtensionAttributes successfully added" | out-file $LogFile -Append

        if ($PracticeDepartment)
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       DEBUG: ExtensionAttribute4='$($PracticeDepartment)'" | out-file $LogFile -Append 
            Set-AdUser $Logonname -Add @{ExtensionAttribute4="$($PracticeDepartment)"} -Server $DomainController
        }
        if ($PracticeManager)
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       DEBUG: ExtensionAttribute5='$($PracticeManager)'" | out-file $LogFile -Append 
            Set-AdUser $Logonname -Add @{ExtensionAttribute5="$($PracticeManager)"} -Server $DomainController
        }
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error setting ExtensionAttributes: '$($_.Exception.Message)'" | out-file $LogFile -Append   
        $Result = 7
    }
}



#### Manage InfoText
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage InfoText..." | out-file $LogFile -Append 
    
    if ((Get-ADUser $Logonname -Properties Info -Server $DomainController).Info -notlike "*Created by script*ServiceRequest*")
    {
        [String]$InfoText = "Created by Script at '$(Get-Date)' for ServiceRequest# $($ServiceRequestID)"
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       InfoText: $($InfoText)" | out-file $LogFile -Append
        try
        {
            Set-AdUser $Logonname -Replace @{Info=$InfoText} -Server $DomainController
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       InfoText successfully added" | out-file $LogFile -Append
        }
        catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error setting Infotext: '$($_.Exception.Message)'" | out-file $LogFile -Append   
            $Result = 7
        }
    }
    else
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       not neccessary, Text already added" | out-file $LogFile -Append
    }
}




#### Adding Groups to the User
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Add User to Groups..." | out-file $LogFile -Append 
    foreach($Group in $GroupsList)
    {
        try
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Use Group: $($Group)" | out-file $LogFile -Append
            Add-ADGroupMember -Server $DomainController -Identity $Group -Members $Logonname
        }
        catch 
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ERROR adding user to Group: '$($_.Exception.Message)'" | out-file $LogFile -Append   
            if ($_.Exception.Message -notlike '*user account is already a member of the specified group*')
            {
                $Result = 7
            }
        }
    }
}


#### Change Default Group, if external
#if (($Result -ne 7)-and ($IsExternal -eq $true))
#{
#    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Change Primary Group..." | out-file $LogFile -Append
#    
#    if ($CurrentUser = Get-ADUser -Server $DomainController -Filter {(givenname -eq $Prename) -and (sn -eq $Surname) -and (Title -eq $JobTitle) -and (Department -eq $Department)})
#    {
#        $UserFound = $CurrentUser.sAMAccountName
#        $Logonname = $UserFound
#        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       User found with sAMAccountName: $($UserFound)" | out-file $LogFile -Append
#
#        $PrimaryGroup = Get-ADUser $Logonname -Server $DomainController -Properties PrimaryGroupID
#        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Current Primary Group: $($PrimaryGroup)" | out-file $LogFile -Append
#
#        if ($PrimaryGroup -eq 514)
#        {
#            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Group is alreay 'Domain Guests'. Nothing to do" | out-file $LogFile -Append
#        }
#        else
#        {
#            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Group is not 'Domain Guests'. change it..." | out-file $LogFile -Append
#
#            Set-ADUser $Logonname -Server $DomainController -Replace @{PrimaryGroupID="514"}
#
#            $PrimaryGroup = (Get-ADUser $Logonname -Server $DomainController -Properties PrimaryGroupID).PrimaryGroupID
#            if ($PrimaryGroup -eq 514)
#            {
#                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())                   Group successfully changed" | out-file $LogFile -Append
#            }
#            else
#            {
#                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())                   Group change failed!" | out-file $LogFile -Append
#                $Result  = 7
#            }
#        }
#    }
#    else 
#    {
#        $UserFound = $false
#        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       User not found" | out-file $LogFile -Append
#    }
#}



#### Remove Domain Users, if external
#if (($Result -ne 7)-and ($IsExternal -eq $true))
#{
#    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Remove Domain Users Group..." | out-file $LogFile -Append
#    
#    if ($CurrentUser = Get-ADUser -Server $DomainController -Filter {(givenname -eq $Prename) -and (sn -eq $Surname) -and (Title -eq $JobTitle) -and (Department -eq $Department)})
#    {
#        $UserFound = $CurrentUser.sAMAccountName
#        $Logonname = $UserFound
#        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       User found with sAMAccountName: $($UserFound)" | out-file $LogFile -Append
#
#                
#       try
#        {
#            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Try to remove" | out-file $LogFile -Append
#            Remove-ADGroupMember -Server $DomainController -Identity "Domain Users" -Members $Logonname -Confirm:$false
#        }
#        catch 
#        {
#            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ERROR removing user from Group: '$($_.Exception.Message)'" | out-file $LogFile -Append   
#            #if ($_.Exception.Message -notlike '*user account is already a member of the specified group*')
#            #{
#                $Result = 7
#            #}
#        } 
#       
#    }
#    else 
#    {
#        $UserFound = $false
#        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       User not found" | out-file $LogFile -Append
#    }
#}


#### Create Homedrive for user
if (($Result -ne 7) -and ($IsExternal -ne $true))
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage Homedrive for user..." | out-file $LogFile -Append
    if (Test-Path "$($HomeDriveRoot)$($Logonname)")
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       folder already exists" | out-file $LogFile -Append    
    }
    else
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       folder does not exist. Create it..." | out-file $LogFile -Append    
        New-Item -ItemType Directory -Path "$($HomeDriveRoot)$($Logonname)"

        if (Test-Path "$($HomeDriveRoot)$($Logonname)")
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             folder successfully created" | out-file $LogFile -Append    
        }
        else
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Error: Folder creation failed!" | out-file $LogFile -Append    
            $Result = 7
        }
    }
}



#### Set Homedrive Rights for user
if (($Result -ne 7) -and ($IsExternal -ne $true))
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage Homedrive rights for user..." | out-file $LogFile -Append
    
    $Acl = Get-Acl "$($HomeDriveRoot)$($Logonname)"
    $Ar = New-Object  system.security.accesscontrol.filesystemaccessrule($Logonname,"FullControl","Allow")
    $Acl.SetAccessRule($Ar)
    Set-Acl "$($HomeDriveRoot)$($Logonname)" $Acl
    
   
}



"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result