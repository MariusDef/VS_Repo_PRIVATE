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
    [string] $ManagerMail,
    [Parameter(Mandatory = $true)]
    [string] $EnableRedirect
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
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Manager Mail: '$($ManagerMail)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Enable Redirect: '$($EnableRedirect)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Use Domaincontroller: '$($DomainController)'" | out-file $LogFile -Append


if ($Displayname.IndexOf("(") -gt 0)
{
    $UserID = ($Displayname.Substring($Displayname.IndexOf("(")+1)).trimend(")")
}


##############################################################################################
#   Im AD prüfen auf Parametrisierten User und holen von Userobjekt
##############################################################################################
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Check, if user already exits..." | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Variables for existance check: Displayname" | out-file $LogFile -Append
    
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




##############################################################################################
#   Verwalten der notwendigen Credentials
##############################################################################################
if ($Result -ne 7)
{  
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
#   Verbinden zu Office 365 (Exchange Online)
##############################################################################################
if (($Result -ne 7))
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Connect to O365 for Mailbox management..." | out-file $LogFile -Append
        try
        {
            $SessionOnline = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri "https://ps.outlook.com/PowerShell-LiveID" -Credential $MSOLCred -Authentication Basic -AllowRedirection
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Import Commands..." | out-file $LogFile -Append
            Import-PSSession $SessionOnline -AllowClobber
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())             successfully connected and imported" | out-file $LogFile -Append
        }
        catch
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
            $Result = 7
        }
    }
    

##############################################################################################
#   Mails weiterleiten
##############################################################################################
if ($Result -ne 7)
{  
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Start Mail redirection configuration..." | out-file $LogFile -Append
    Try
    {
        
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Current Forwarding: '$((Get-Mailbox -Identity "$($Logonname)@cellent.de" | Select-Object -Property ForwardingSmtpAddress).ForwardingSmtpAddress)'..." | out-file $LogFile -Append
        
        if ($EnableRedirect -eq "True")
        {
            Set-Mailbox -Identity "$($Logonname)@cellent.de" -ForwardingSmtpAddress $ManagerMail -DeliverToMailboxAndForward $true

            if ($(Get-Mailbox -Identity "$($Logonname)@cellent.de" | Select-Object -Property ForwardingSmtpAddress).ForwardingSmtpAddress -eq "SMTP:$($ManagerMail)")
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Forwarding successfully set" | out-file $LogFile -Append
            }
            else
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Set forwarding failed!" | out-file $LogFile -Append
                $Result = 7
            }

        }


        if ($EnableRedirect -eq "False")
        {
            Set-Mailbox -Identity "$($Logonname)@cellent.de" -ForwardingSmtpAddress $null

            if ($(Get-Mailbox -Identity "$($Logonname)@cellent.de" | Select-Object -Property ForwardingSmtpAddress).ForwardingSmtpAddress -eq $null)
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Forwarding successfully disabled" | out-file $LogFile -Append
            }
            else
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: disable forwarding failed!" | out-file $LogFile -Append
                $Result = 7
            }
        }
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Session not available: '$($_.Exception.Message)'" | out-file $LogFile -Append
        $Result = 7
    }
}

Get-PSSession | Remove-PSSession



"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result