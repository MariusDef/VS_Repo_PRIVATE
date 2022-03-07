Param
(
    [Parameter(Mandatory = $true)]
    [string] $UserDisplayname,
    [Parameter(Mandatory = $true)]
    [string] $ChangeValue,
    [Parameter(Mandatory = $true)]
    [string] $ChangeType
)


$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0
$DomainController = "DE-DCO-202"

if ($UserDisplayname.IndexOf("(") -gt 0)
{
    $UserID = ($UserDisplayname.Substring($UserDisplayname.IndexOf("(")+1)).trimend(")")
    #$UserID = $CorrectUserID
}

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Parameter: UserDisplayName: '$($UserDisplayname)'; ChangeValue: '$($ChangeValue)'; ChangeType: '$($ChangeType)'" | out-file $LogFile -Append


Switch ($ChangeType)
{
    "Title"
    {
        $OldTitle = (Get-ADUser -Identity $UserID -Properties Title -Server $DomainController).Title
         "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Actual Title from AD: '$($OldTitle)'" | out-file $LogFile -Append

        Try
        {
            Set-ADUser -Identity $UserID -Server $DomainController -Replace @{title="$($ChangeValue)";description="$($ChangeValue)"}
        }
        Catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
        }  



        $CheckValue = (Get-ADUser -Identity $UserID -Properties Title -Server $DomainController).Title
        if ($CheckValue -eq $ChangeValue)
        {
            $Result = 0
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Title successfully changed" | out-file $LogFile -Append
        }
        else
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR - Change Title failed!" | out-file $LogFile -Append
        }

        $CheckValue = (Get-ADUser -Identity $UserID -Properties Description -Server $DomainController).Description
        if ($CheckValue -eq $ChangeValue)
        {
            $Result = 0
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Description successfully changed" | out-file $LogFile -Append
        }
        else
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR - Change Description failed!" | out-file $LogFile -Append
        }
    }
    "Department"
    {
        $OldValue = (Get-ADUser -Identity $UserID -Properties Department -Server $DomainController).Department
         "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Actual Department from AD: '$($OldValue)'" | out-file $LogFile -Append

        Try
        {
            Set-ADUser -Identity $UserID -Server $DomainController -Replace @{department="$($ChangeValue)"}
        }
        Catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
        }  



        $CheckValue = (Get-ADUser -Identity $UserID -Properties Department -Server $DomainController).Department
        if ($CheckValue -eq $ChangeValue)
        {
            $Result = 0
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Department successfully changed" | out-file $LogFile -Append
        }
        else
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR - Change failed!" | out-file $LogFile -Append
        }
    }
    "CostCenter"
    {
        $OldValue = (Get-ADUser -Identity $UserID -Properties extensionAttribute1 -Server $DomainController).extensionAttribute1
         "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Actual CostCenter from AD: '$($OldValue)'" | out-file $LogFile -Append

        Try
        {
            Set-ADUser -Identity $UserID -Server $DomainController -Replace @{extensionAttribute1="$($ChangeValue)"}
        }
        Catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
        }  



        $CheckValue = (Get-ADUser -Identity $UserID -Properties extensionAttribute1 -Server $DomainController).extensionAttribute1
        if ($CheckValue -eq $ChangeValue)
        {
            $Result = 0
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - CostCenter successfully changed" | out-file $LogFile -Append
        }
        else
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR - Change failed!" | out-file $LogFile -Append
        }
    }
    "PracticeDepartment"
    {
        $OldValue = (Get-ADUser -Identity $UserID -Properties extensionAttribute4 -Server $DomainController).extensionAttribute4
         "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Actual PracticeDepartment from AD: '$($OldValue)'" | out-file $LogFile -Append

        Try
        {
            Set-ADUser -Identity $UserID -Server $DomainController -Replace @{extensionAttribute4="$($ChangeValue)"}
        }
        Catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
        }  



        $CheckValue = (Get-ADUser -Identity $UserID -Properties extensionAttribute4 -Server $DomainController).extensionAttribute4
        if ($CheckValue -eq $ChangeValue)
        {
            $Result = 0
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - PracticeDepartment successfully changed" | out-file $LogFile -Append
        }
        else
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR - Change failed!" | out-file $LogFile -Append
        }
    }
    "PracticeManager"
    {
        $OldValue = (Get-ADUser -Identity $UserID -Properties extensionAttribute5 -Server $DomainController).extensionAttribute5
         "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Actual PracticeManager from AD: '$($OldValue)'" | out-file $LogFile -Append

        Try
        {
            Set-ADUser -Identity $UserID -Server $DomainController -Replace @{extensionAttribute5="$($ChangeValue)"}
        }
        Catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
        }  



        $CheckValue = (Get-ADUser -Identity $UserID -Properties extensionAttribute5 -Server $DomainController).extensionAttribute5
        if ($CheckValue -eq $ChangeValue)
        {
            $Result = 0
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - PracticeManager successfully changed" | out-file $LogFile -Append
        }
        else
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR - Change failed!" | out-file $LogFile -Append
        }
    }
     "Location"
    {
        $OldLocation = (Get-ADUser -Identity $UserID -Properties l -Server $DomainController).l
         "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Actual Location from AD: '$($OldLocation)'" | out-file $LogFile -Append

        Try
        {
            Set-ADUser -Identity $UserID -Server $DomainController -Replace @{l="$($ChangeValue)";physicalDeliveryOfficeName="$($ChangeValue)"}
        }
        Catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
        }  



        $CheckValue = (Get-ADUser -Identity $UserID -Properties physicalDeliveryOfficeName -Server $DomainController).physicalDeliveryOfficeName
        if ($CheckValue -eq $ChangeValue)
        {
            $Result = 0
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - physicalDeliveryOfficeName successfully changed" | out-file $LogFile -Append
        }
        else
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR - Change physicalDeliveryOfficeName failed!" | out-file $LogFile -Append
        }

        $CheckValue = (Get-ADUser -Identity $UserID -Properties l -Server $DomainController).l
        if ($CheckValue -eq $ChangeValue)
        {
            $Result = 0
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - l successfully changed" | out-file $LogFile -Append
        }
        else
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR - Change l failed!" | out-file $LogFile -Append
        }
    }
    
    "Manager"
    {

        $OldValue1 = (Get-ADUser -Identity $UserID -Properties extensionAttribute2 -Server $DomainController).extensionAttribute2
        $OldValue2 = (Get-ADUser -Identity $UserID -Properties Manager -Server $DomainController).Manager
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Actual ExtensionAttribute2 from AD: '$($OldValue1)'" | out-file $LogFile -Append
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Actual Manager from AD: '$($OldValue2)'" | out-file $LogFile -Append

        if ($UserDisplayname.IndexOf("(") -gt 0){$UserIDManager = ($ChangeValue.Substring($ChangeValue.IndexOf("(")+1)).trimend(")")}

        $Manager = Get-ADUser -Identity $UserIDManager -Server $DomainController
        $ManagerMail = (Get-ADUser -Identity $UserIDManager -Properties Mail -Server $DomainController).Mail
        
        
        ########### Change Manager ################
        Try
        {
            Set-ADUser -Identity $UserID -Manager $Manager -Server $DomainController
        }
        Catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
        }
          
        $UserIDManager = (Get-ADUser -Identity $UserID -Properties Manager -Server $DomainController).Manager
        #$CheckValue = (Get-ADUser -Filter {distinguishedName -eq $UserIDManager} -Properties displayName, UserPrincipalName)
        #$CheckValuePrep = $CheckValue.DisplayName + " (" + $CheckValue.UserPrincipalName.TrimEnd("@cellent.de") + ")"
        
        
        #if ($ChangeValue -eq $CheckValuePrep)
        if ($UserIDManager -eq $Manager.DistinguishedName)
        {
            $Result = 0
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Manager successfully changed" | out-file $LogFile -Append
        }
        else
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR - Change failed!" | out-file $LogFile -Append
        }


        ########### Change Extensionattribute2 = ManagerMail ################
        Try
        {
            Set-ADUser -Identity $UserID -Server $DomainController -Replace @{extensionAttribute2="$($ManagerMail)"}
        }
        Catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
        }  

        
        $CheckValue = (Get-ADUser -Identity $UserID -Properties extensionAttribute2 -Server $DomainController).extensionAttribute2
        if ($CheckValue -eq $ManagerMail)
        {
            if ($Result -ne 7){$Result = 0}
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ExtensionAttribute2 successfully changed" | out-file $LogFile -Append
        }
        else
        {
            $Result = 7
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - ERROR - Change failed!" | out-file $LogFile -Append
        }
    }
}

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($ChangeType) - Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result