$LogPath = "$env:windir\logs\PW-Reminder_EWS.log"
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

[String]$username = 'sv43448t@wipro.com'
[String]$password = 'S6h}\Xu$'
$EWSurl = 'https://webmail2.wipro.com/ews/exchange.asmx'
$dllPath = "C:\Temp\2.2\Microsoft.Exchange.WebServices.dll"

try{
#LogFile Function
    function write_log 
    {
        param([string]$msg)
        $FileExists = Test-Path $LogPath

        $DateNow = Get-Date -Format [dd.MM.yyyy]-[HH:mm:ss]
     
        if ($FileExists -eq $true)
        {
            try{
                Write-Output "$DateNow | $msg" | Out-File $LogPath -append
               }
               catch{
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
            
                    write-Host "$ErrorMessage"
                    Write-Host "$FailedItem"
                    }
        }
        else
        {
            try{
                New-Item -path $LogPath -ItemType file -ea stop
               }
               catch{
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
            
                    write-Host "$ErrorMessage"
                    Write-Host "$FailedItem"
                    }
            try{
                Write-Output "$DateNow | $msg" | Out-File $LogPath -append
               }
               catch{
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
            
                    write-Host "$ErrorMessage"
                    Write-Host "$FailedItem"
         }
        }
    }
}
catch{
        $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            
            write-Host "$ErrorMessage"
            Write-Host "$FailedItem"
     }

function send-Mail-User([string]$Name, [string]$Days, [string]$MailAddress)
{
    #$MailAddress = "Marius.Deffner@wipro.com"

    $SMTPServer = "ms-az-smtp-gw-1.cellent.int" 
    $from = "svc.itsc@wipro.com"
    $subject = "Ihr cellent-Kennwort läuft demnächst ab!"
    $template = "\\DE-DCO-202\d$\Scripts\Password_Reminder_cellent\PASSWORT.html"
    $PathToAttachement = "\\DE-DCO-202\d$\Scripts\Password_Reminder_cellent\image003.jpg"

    $body = get-Content $template
    
    $replace=[regex]"#NAME#"
    $body = $replace.Replace($body,$Name,1)

    $replace=[regex]"#TAGEN#"
    $body = $replace.Replace($body,$Days,1)

    if (-not(Get-Module "Microsoft.Exchange.WebServices"))
    {
        #write_log "Exchange Module not imported... import it"
        try{
            Import-Module -name $dllPath
        }
        catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            
            write_log "$ErrorMessage"
            write_log "$FailedItem"
            exit
        }
    }
    else
    {
        #write_log  'Exchange Module imported correctly'
    }


    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService #-ArgumentList Exchange2010
    $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList  $username, $password
    $service.Url = $EWSurl
    $message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
    $message.Subject = $subject
    $message.body = $body
    #$message.BccRecipients.Add("Marius.Deffner@wipro.com") | Out-Null
    $message.ToRecipients.Add($MailAddress) |Out-Null
    #$message.Attachments.AddFileAttachment($PathToAttachement)
    $message.SendAndSaveCopy()
    
}

function Check-PasswordDate($Days, [string]$container, [string]$query, $users)
{
    foreach ($user in $users)
    {
        if ($user -ne $null)
        {
            $ts = New-TimeSpan -Start ([datetime]::Now.Date) -End $user.Expirydate
            if($ts.Days -eq $days)
            {
                if($user.extensionAttribute6 -like "*wipro.com")
                {
                    $Mail = $user.extensionAttribute6
                }
                elseif ($user.mail -like "*wipro.com")
                {
                    $Mail = $user.mail
                }
                write_log "Send Mail: $($user.DisplayName) : $($mail) : $($user.ExpiryDate) : $($ts.Days)"
                write-host $user.DisplayName : $mail : $user.ExpiryDate : $ts.Days  
                send-Mail-User -MailAddress $mail -Name $user.DisplayName -Days $days 
            }
            
            
            #| Where {$_.ExpiryDate -eq ([datetime]::Now.Date.AddDays($days))

            #Write-Host $days : $user.Name : $user.mail
            
        }
    }
}

### Begin Main
$container = "OU=Usermanagement,DC=cellent,DC=int"
#$query = "(&(objectcategory=user)(useraccountcontrol=512))"

$users = Get-ADUser -SearchBase $container -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} –Properties "DisplayName", "mail", "msDS-UserPasswordExpiryTimeComputed", "extensionAttribute6" | Select-Object -Property "mail","extensionAttribute6","Displayname",@{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}

Write-Host "Check 1 Day:"
write_log "Check 1 Day:"
Check-PasswordDate -Days 1 -container $container -query $query -users $users

Write-Host "Check 10 Days:"
write_log "Check 10 Days:"
Check-PasswordDate -Days 10 -container $container -query $query -users $users

Write-Host "Check 15 Days:"
write_log "Check 15 Days:"
Check-PasswordDate -Days 15 -container $container -query $query -users $users


Write-Host "Check finished"