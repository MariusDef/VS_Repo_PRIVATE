#################################################################################
#   Author: Marius Deffner, cellent GmbH, 10.10.2019
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
    [string] $ITUserMail,
    [Parameter(Mandatory = $true)]
    [string] $CustomerDisplayName,
    [Parameter(Mandatory = $true)]
    [string] $SRNumber
)


$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0
$DomainController = "de-dco-202.cellent.int"


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Parameter:" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ITUserMail: '$($ITUserMail)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       CustomerDisplayName: '$($CustomerDisplayName)'" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       ServiceRequest: '$($SRNumber)'" | out-file $LogFile -Append
#"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Use Domaincontroller: '$($DomainController)'" | out-file $LogFile -Append

#if ($Displayname.IndexOf("(") -gt 0)
#{
#    $UserID = ($Displayname.Substring($Displayname.IndexOf("(")+1)).trimend(")")
#}


###################################################
####  Functions ###################################
###################################################
#Read-Host -Prompt "Enter PW for SQL HeatService User 'heatservice'" -AsSecureString | ConvertFrom-SecureString | Out-File ("\\de-frs-201\e$\Heat_Automation\Create-Lieferschein\SQLHeatService_cred_tmp.txt")
$Username = "heatservice"
$SecurePassword = Get-Content ("\\de-frs-201\e$\Heat_Automation\Create-Lieferschein\SQLHeatService_cred_tmp.txt") | ConvertTo-SecureString
#$SecurePassword = Get-Content ((split-path -parent $MyInvocation.MyCommand.Definition) + "\SQLHeatService_cred.txt") | ConvertTo-SecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
$UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

$datasource = "de-sql-202\heatprd"
$database = "heatsm"
$connectionstring = "server=$($datasource);database=$($database);User ID=$($Username);Password=$($UnsecurePassword);"


Function Get-Assets([string]$SRNumber)
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
        $query = "SELECT        dbo.CI.AssetID_ITSM AS AssetID, dbo.Manufacturer.Manufacturer, dbo.CI.Model, dbo.CI.SerialNumber_ITSM AS Serialnumber, dbo.CI.Equipment_ITSM AS Equipment
FROM            dbo.CI LEFT OUTER JOIN
                         dbo.Manufacturer ON dbo.CI.ManufacturerLink_RecID = dbo.Manufacturer.RecId LEFT OUTER JOIN
                         dbo.FusionLink ON dbo.CI.RecId = dbo.FusionLink.TargetID LEFT OUTER JOIN
                         dbo.ServiceReq ON dbo.FusionLink.SourceID = dbo.ServiceReq.RecId
WHERE        (dbo.ServiceReq.ServiceReqNumber = $($SRNumber))"
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
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       No CIs for selected SR available" | out-file $LogFile -Append
            $Response = 1
        }
        else
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Yes CIs for selected SR available. Build result table" | out-file $LogFile -Append
            $Response = 2

            $bla += '<table style="font-size:8pt; font-family:arial" border=1 rules=all>'
            $bla += '<tr><th align="left">AssetID</th><th align="left">Manufacturer</th><th align="left">Model</th><th align="left">Serialnumber</th><th align="left">Equipment</th></tr>'
            for ($i=0;$i -lt $Table.Rows.Count;$i++)
            {
                #"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       AssetID: '$($Table.Rows[$i][2])' # Model: '$($Table.Rows[$i][1])'" | out-file $LogFile -Append
                $bla+= "<tr><td>" + $Table.Rows[$i][0] + "</td><td>" + $Table.Rows[$i][1] + "</td><td>" + $Table.Rows[$i][2] + "</td><td>" + $Table.Rows[$i][3] + "</td><td>" + $Table.Rows[$i][4] + "</td></tr>`n"
            }
            $bla += "</table>"
            #write-host $bla
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


function send-Mail-User([string]$ITUserMail, [string]$SRNumber, [string]$Assetsm, [string]$CustomerDisplayName)
{
    try
    {
        $UserMail = "print@cellent.de"

        $SMTPServer = "ms-az-smtp-gw-1.cellent.int" 
        $from = $ITUserMail #"itsc.itsc@cellent.de"
        $subject = "Lieferschein (ServiceRequest# $($SRNumber))"
        $template = "\\de-frs-201\e$\Heat_Automation\Create-Lieferschein\Lieferschein_Endandwender.html"
        #$PathToAttachement = "\\de-frs-201\e$\Heat_Automation\Create-Lieferschein\Lieferschein_Endandwender_files\image001.png"
        #$template = "\\$($DomainController)\d$\Scripts\Info_Account_Ablaufdatum\PASSWORT.html"
        #$PathToAttachement = "\\$($DomainController)\d$\Scripts\Info_Account_Ablaufdatum\image003.jpg"

        $body = get-Content $template
    
        $replace=[regex]"#ITUSER#"
        $body = $replace.Replace($body,$ITUserMail,1)

        $replace=[regex]"#SERVICEREQNUMBER#"
        $body = $replace.Replace($body, "ServiceRequest# $($SRNumber)",1)

        $replace=[regex]"#CUSTOMER#"
        $body = $replace.Replace($body,$CustomerDisplayName,1)

        $replace=[regex]"#DATE#"
        $body = $replace.Replace($body,$((Get-Date).ToShortDateString()),1)

    
        $replace=[regex]"#ASSETS#"
        $body = $replace.Replace($body,$Assets,1)
    

        $message = New-Object System.Net.Mail.MailMessage
        $message.IsBodyHtml = $true
        $message.subject = $subject
        $message.body = $body
        $message.to.add($UserMail)
        #$message.CC.Add("svc_itsc@cellent.de")
        $message.bcc.add("marius.deffner@cellent.de")
        $message.from = $from
        #$attachment = New-Object System.Net.Mail.Attachment($PathToAttachement)
        #$attachment.ContentType.MediaType = "image/jpg"
        #$attachment.ContentId = "Attachment"
        #$message.attachments.add($attachment)
    
    
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

$Assets = Get-Assets -SRNumber $SRNumber
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Prepare Mail for User..." | out-file $LogFile -Append
$Result = send-Mail-User -ITUserMail $ITUserMail -SRNumber $SRNumber -Assetsm $Assets -CustomerDisplayName $CustomerDisplayName
#write-host $Assets


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result