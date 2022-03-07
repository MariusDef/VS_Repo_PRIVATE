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
$ExchangeSrv = "DE-EXC-211"

#### Manage Variables
if($IsExternalString -eq 'false'){$IsExternal = $false}
if($IsExternalString -eq ''){$IsExternal = $false}
if($IsExternalString -eq 'true'){$IsExternal = $true}
if($IsExternal){$PersonalNumber = "99999"}


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
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Connect to Exchage Server for Mailbox management (Server: $($ExchangeSrv))..." | out-file $LogFile -Append
    try
    {
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($ExchangeSrv)/powershell" -Authentication Kerberos
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Import Commands 'enable-mailbox' , 'get-mailbox'..." | out-file $LogFile -Append
        Import-PSSession $Session -CommandName Enable-Mailbox, Get-Mailbox, set-mailbox
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
        $Result = 7
    }

    if ($Session.State -eq "Opened")
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       check, if a mailbox currently exists..." | out-file $LogFile -Append
        if (get-mailbox $Logonname)
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Yes. Mailbox does exist. Do not create again. " | out-file $LogFile -Append

            #RemoteRecipientType (Migrated)
            #Database (de-o365-201)
            $Result = 0
        }
        else
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Mailbox not available. Create one..." | out-file $LogFile -Append
            if ($IsExternal -eq $true)
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())                   for external user" | out-file $LogFile -Append
                enable-mailbox $Logonname  
                #set-mailbox $Logonname -emailaddresses "$($Prename).$($Surname).ext@cellent.de"
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())                         set Primary SMTP Addresse to: $($Prename).$($Surname).ext@cellent.de" | out-file $LogFile -Append
                set-mailbox $Logonname -PrimarySmtpAddress "$($Prename).$($Surname).ext@cellent.de" -EmailAddressPolicyEnabled $false
            }
            else
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())                   for internal user" | out-file $LogFile -Append
                enable-mailbox $Logonname
            }
            

            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             Mailbox created. Check, if mailbox is available..." | out-file $LogFile -Append
            if (get-mailbox $Logonname)
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())                   Yes. Mailbox is available. Continue" | out-file $LogFile -Append
            }
            else
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())                   Error: No. Mailbox is not available" | out-file $LogFile -Append
                $Result = 7
            }
        }
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Remove Remoteing to Exchage Server for Mailbox management (Server: $($ExchangeSrv))..." | out-file $LogFile -Append
        exit-pssession
        Remove-PSSession -ComputerName $ExchangeSrv
    }
    else
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available. Quit Script" | out-file $LogFile -Append
        $Result = 7
    }
}




"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result