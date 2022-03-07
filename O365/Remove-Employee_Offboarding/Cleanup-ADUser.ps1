#################################################################################
#   Author: Marius Deffner, cellent GmbH, 17.10.2018
#
#   Possible Exit Codes:
#      7: Any job in script failed
#      8: Mailbox already exists. Nothing to do
#
#
#################################################################################

Param
(
    [Parameter(Mandatory = $true)]
    [string] $UserDisplayname
)


$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0
$HomeDriveRoot = "\\de-fls-201\d$\Userhome\"

#### Manage Variables
if ($UserDisplayname.IndexOf("(") -gt 0)
{
    $Logonname = ($UserDisplayname.Substring($UserDisplayname.IndexOf("(")+1)).trimend(")")
    #$UserID = $CorrectUserID
}


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Parameter:" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       UserDisplayname: '$($UserDisplayname)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Logonname: '$($Logonname)'" | out-file $LogFile -Append

$DomainController = "de-dco-202.cellent.int" #(get-ADDomainController -Discover).hostname
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Use Domaincontroller: '$($DomainController)'" | out-file $LogFile -Append




#### Function for Password generator
function Get-RandomCharacters($length, $characters) 
{
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}

Function PasswordGenerator 
{
    $password = Get-RandomCharacters -length 5 -characters 'abcdefghiklmnoprstuvwxyz'
    $password += Get-RandomCharacters -length 1 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
    $password += Get-RandomCharacters -length 1 -characters '1234567890'
    $password += Get-RandomCharacters -length 1 -characters '!"§$%&/()=?}][{@#*+'

    $characterArray = $password.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $password = -join $scrambledStringArray    
    
    return $password
}

#### Change Password
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Start changing password to random password..." | out-file $LogFile -Append
    try
    {
        $password = PasswordGenerator
        Set-ADAccountPassword -Identity $Logonname -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Password successfully set" | out-file $LogFile -Append
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error setting password: '$($_.Exception.Message)'" | out-file $LogFile -Append   
        $Result = 7
    }
    
}


#### Clear Phone, Fax, Mobilephone, Manager
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Start clearing user attributes..." | out-file $LogFile -Append
    try
    {
        Set-ADUser -Identity $Logonname -Clear telephoneNumber, facsimileTelephoneNumber, ipPhone, mobile, manager, homeDirectory, homeDrive -Server $DomainController
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Attributes successfully cleared" | out-file $LogFile -Append
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error clearing attributes: '$($_.Exception.Message)'" | out-file $LogFile -Append   
        $Result = 7
    }
}


#### Remove Group Memberships
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Remove Group Memberships..." | out-file $LogFile -Append
    try
    {
        $Groups = Get-ADUser -Identity $Logonname -Properties MemberOf 
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Current Groups: '$($Groups.MemberOf)'" | out-file $LogFile -Append
        
        ForEach-Object {$groups.MemberOf | Remove-ADGroupMember -Members $groups.DistinguishedName -Confirm:$false}

        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Groups memberships successfully removed" | out-file $LogFile -Append
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error clearing attributes: '$($_.Exception.Message)'" | out-file $LogFile -Append   
        $Result = 7
    }
}


#### Delete Homedrive
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Manage Homedrive for user..." | out-file $LogFile -Append
    if (Test-Path "$($HomeDriveRoot)$($Logonname)")
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       folder exists. Delete it" | out-file $LogFile -Append    
        Remove-Item "$($HomeDriveRoot)$($Logonname)" -Recurse -Force -Confirm:$false
        
        if (Test-Path "$($HomeDriveRoot)$($Logonname)")
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Error: Folder delete failed!" | out-file $LogFile -Append    
            $Result = 7
        }
        else
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Folder successfully deleted" | out-file $LogFile -Append    
        }
    }
    else
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       folder does not exist." | out-file $LogFile -Append        
    }
}


#### Add Note for script modification
if ($Result -ne 7)
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Add Note for script modification..." | out-file $LogFile -Append
    try
    {
        $Info = Get-ADUser -Identity $Logonname -Properties info | %{ $_.info}
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Current Info: '$($Info)'" | out-file $LogFile -Append
        
        Set-ADUser $Logonname -Replace @{info="$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) cleanup attributes, password & groups by script `r`n $($Info)"}

        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Groups memberships successfully removed" | out-file $LogFile -Append
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error clearing attributes: '$($_.Exception.Message)'" | out-file $LogFile -Append   
        $Result = 7
    }
}



"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result