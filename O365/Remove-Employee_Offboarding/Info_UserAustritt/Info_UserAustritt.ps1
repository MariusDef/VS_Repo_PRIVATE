#################################################################################
#   Author: Marius Deffner, cellent GmbH, 30.09.2019
#
#   Possible Exit Codes:
#      0:  Successfull
#      7:  Any job in script failed
#
#
#################################################################################


Param
(
    [Parameter(Mandatory = $true)]
    [string] $Displayname,
    [Parameter(Mandatory = $true)]
    [string] $SRNumber,
    [Parameter(Mandatory = $true)]
    [string] $UserMail
)


$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

[String]$authUsername = 'sv43448t@wipro.com'
[String]$authUserpassword = 'S6h}\Xu$'
$EWSurl = 'https://webmail2.wipro.com/ews/exchange.asmx'
$dllPath = "E:\Heat_Automation\2.2\Microsoft.Exchange.WebServices.dll"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0
$DomainController = "de-dco-202.cellent.int"


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Parameter:" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Displayname: '$($Displayname)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       UserMail: '$($UserMail)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ServiceRequest: '$($SRNumber)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Use Domaincontroller: '$($DomainController)'" | out-file $LogFile -Append

if ($Displayname.IndexOf("(") -gt 0)
{
    $UserID = ($Displayname.Substring($Displayname.IndexOf("(")+1)).trimend(")")
}


###################################################
####  Functions ###################################
###################################################
#Read-Host -Prompt "Enter PW for SQL HeatService User 'heatservice'" -AsSecureString | ConvertFrom-SecureString | Out-File ("C:\Temp\Info_UserAustritt\SQLHeatService_cred.txt")
$Username = "heatservice"
$SecurePassword = Get-Content ((split-path -parent $MyInvocation.MyCommand.Definition) + "\SQLHeatService_cred.txt") | ConvertTo-SecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
$UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

$datasource = "de-sql-202\heatprd"
$database = "heatsm"
$connectionstring = "server=$($datasource);database=$($database);User ID=$($Username);Password=$($UnsecurePassword);"


Function Get-Assets([string]$UserName)
{
    

    try
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       open SQL Connection..." | out-file $LogFile -Append
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionstring
        $connection.Open()


    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: opening connection: '$($_.Exception.Message)'" | out-file $LogFile -Append
        $Result = 7
    }
    

    if($connection.State -eq 'Open')
    {
        $query = "select AssetID_ITSM, 
	        Case when CIType = 'AccessCard' then CIType when CIType = 'SIM' then CIType else Manufacturer end as Manufacturer, 
	        case when CIType = 'AccessCard' then Location_ITSM else Model end as Model, 
	        SerialNumber_ITSM, Equipment_ITSM
            from ci
            where owner = '$($UserID)'
            and Status = 'Active' and (not (CIType = 'Maintenance')) and (not (CIType = 'Service'))"
        #$query = “select AssetID_ITSM, Manufacturer, Model, SerialNumber_ITSM, Equipment_ITSM from ci where owner = 'mdeffner'and Status = 'Active' and (not (CIType = 'Maintenance'))”
   
        #"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Get all mapped License CIs for User: '$($UserDisplayname)' with filter on LicenseModel: '$($LicenseModel)'..." | out-file $LogFile -Append
        $conn = $connection.CreateCommand()
        $conn.CommandText = $query

        $Reader = $conn.ExecuteReader()
        $Table = New-Object System.Data.DataTable
        $Table.Load($Reader)
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       row count: $($Table.Rows.Count)..." | out-file $LogFile -Append
        if ($Table.Rows.Count -eq 0)
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       No CIs for user available" | out-file $LogFile -Append
            $Response = 1
        }
        else
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Yes CIs for user available. Build result table" | out-file $LogFile -Append
            $Response = 2

            $bla += '<table border=1 rules=all>'
            $bla += '<tr><th align="left">AssetID</th><th align="left">Manufacturer</th><th align="left">Model</th><th align="left">Serialnumber</th><th align="left">Equipment</th></tr>'
            for ($i=0;$i -lt $Table.Rows.Count;$i++)
            {
                #"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       AssetID: '$($Table.Rows[$i][2])' # Model: '$($Table.Rows[$i][1])'" | out-file $LogFile -Append
                if ($table.Rows[$i][2] -notlike '*CallManager*')
                {
                    $bla+= "<tr><td>" + $Table.Rows[$i][0] + "</td><td>" + $Table.Rows[$i][1] + "</td><td>" + $Table.Rows[$i][2] + "</td><td>" + $Table.Rows[$i][3] + "</td><td>" + $Table.Rows[$i][4] + "</td></tr>`n"
                }
            }
            $bla += "</table>"
        }
    
    }



    ##############################################################################################
    #   close SQL Connection
    ##############################################################################################
    try
    {
        if($connection.State -eq 'Open')
        {
            $connection.Close()
        }
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Connection cannot beeing closed: '$($_.Exception.Message)'" | out-file $LogFile -Append
        $Result = 7
    }
    Return ($bla)
}


function send-Mail-User-EWS([string]$UserName, [string]$SRNumber, [string]$Assetsm, [string]$UserMail)
{
    try
    {
        #$UserMail = "Marius.Deffner@wipro.com"

        $from = "svc.itsc@wipro.com"
        $subject = "Information - Austritt Mitarbeiter (ServiceRequest# $($SRNumber))"
        $template = "\\de-frs-201\e$\Heat_Automation\Remove-Employee_Offboarding\Info_UserAustritt\PASSWORT.html"
        #$PathToAttachement = "\\de-frs-201\e$\Heat_Automation\Remove-Employee_Offboarding\Info_UserAustritt\image003.jpg"


        $body = get-Content $template
    
        $replace=[regex]"#NAME#"
        $body = $replace.Replace($body,$Displayname,1)

        $replace=[regex]"#USERNAME#"
        $body = $replace.Replace($body,$Displayname,1)

    
        $replace=[regex]"#ASSETS#"
        $body = $replace.Replace($body,$Assets,1)
        if (-not(Get-Module "Microsoft.Exchange.WebServices"))
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Exchange Module not imported... import it" | out-file $LogFile -Append
            try
            {
                Import-Module -name $dllPath
            }
            catch
            {
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
            
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $ErrorMessage" | out-file $LogFile -Append
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $FailedItem" | out-file $LogFile -Append
            #    exit
            }
        }
        else
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Exchange Module imported correctly" | out-file $LogFile -Append
        }

        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService #-ArgumentList Exchange2010
        $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList  $authUsername, $authUserpassword
        $service.Url = $EWSurl
        $message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
        $message.Subject = $subject
        $message.body = $body
        $message.BccRecipients.Add("Marius.Deffner@wipro.com")
        $message.ToRecipients.Add($UserMail)
        $message.ToRecipients.Add("svc.itsc@wipro.com")
        #$message.CC.Add("svc.itsc@wipro.com")
        #$message.Attachments.AddFileAttachment($PathToAttachement)
        $message.SendAndSaveCopy()
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Message send" | out-file $LogFile -Append
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: sending mail: '$($_.Exception.Message)'" | out-file $LogFile -Append
        Return(7)
    }
}

function send-Mail-User([string]$UserName, [string]$SRNumber, [string]$Assetsm, [string]$UserMail)
{
    try
    {
        $UserMail = "Marius.Deffner@wipro.com"

        $SMTPServer = "ms-az-smtp-gw-1.cellent.int" 
        $from = "itsc.itsc@cellent.de"
        $subject = "Information - Austritt Mitarbeiter (ServiceRequest# $($SRNumber))"
        $template = "\\de-frs-201\e$\Heat_Automation\Remove-Employee_Offboarding\Info_UserAustritt\PASSWORT.html"
        $PathToAttachement = "\\de-frs-201\e$\Heat_Automation\Remove-Employee_Offboarding\Info_UserAustritt\image003.jpg"
        #$template = "\\$($DomainController)\d$\Scripts\Info_Account_Ablaufdatum\PASSWORT.html"
        #$PathToAttachement = "\\$($DomainController)\d$\Scripts\Info_Account_Ablaufdatum\image003.jpg"

        $body = get-Content $template
    
        $replace=[regex]"#NAME#"
        $body = $replace.Replace($body,$Displayname,1)

        $replace=[regex]"#USERNAME#"
        $body = $replace.Replace($body,$Displayname,1)

    
        $replace=[regex]"#ASSETS#"
        $body = $replace.Replace($body,$Assets,1)
    

        $message = New-Object System.Net.Mail.MailMessage
        $message.IsBodyHtml = $true
        $message.subject = $subject
        $message.body = $body
        $message.to.add($UserMail)
        $message.CC.Add("svc.itsc@wipro.com")
        $message.bcc.add("Marius.Deffner@wipro.com")
        $message.from = $from
        $attachment = New-Object System.Net.Mail.Attachment($PathToAttachement)
        $attachment.ContentType.MediaType = "image/jpg"
        $attachment.ContentId = "Attachment"
        $message.attachments.add($attachment)
    
    
        $smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer);
        $smtp.send($message)
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Mail successfully send" | out-file $LogFile -Append
        Return(0)
    }
    catch
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: sending mail: '$($_.Exception.Message)'" | out-file $LogFile -Append
        Return(7)
    }

}

### Begin Main
$body = ""


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Get Assets for User..." | out-file $LogFile -Append
#"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Variables for existance check: LoginID ($($UserID))" | out-file $LogFile -Append

$Assets = Get-Assets -UserName $Displayname
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Prepare Mail for User..." | out-file $LogFile -Append
$Result = send-Mail-User-EWS -UserName $Displayname -SRNumber $SRNumber -Assets $Assets -UserMail $UserMail
#write-host $Assets


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result