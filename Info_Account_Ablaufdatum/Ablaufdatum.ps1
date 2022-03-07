function send-Mail-User([string]$ExternalUserName, [string]$Datum, [string]$ManagerMail, [string]$ManagerName)
{
    #$ManagerMail = "Marius.Deffner@cellent.de"

    $SMTPServer = "ms-az-smtp-gw-1.cellent.int" 
    $from = "ITSC@cellent.de"
    $subject = "Der cellent-Account von '$($ExternalUserName)' läuft demnächst ab!"
    $template = "\\DE-DCO-202\d$\Scripts\Info_Account_Ablaufdatum\PASSWORT.html"
    $PathToAttachement = "\\DE-DCO-202\d$\Scripts\Info_Account_Ablaufdatum\image003.jpg"

    $body = get-Content $template
    
    $replace=[regex]"#MANAGER#"
    $body = $replace.Replace($body,$ManagerName,1)

    $replace=[regex]"#NAME#"
    $body = $replace.Replace($body,$ExternalUserName,1)

    $replace=[regex]"#NAME#"
    $body = $replace.Replace($body,$ExternalUserName,1)

    $replace=[regex]"#NAME#"
    $body = $replace.Replace($body,$ExternalUserName,1)
    
    $replace=[regex]"#DATUM#"
    $body = $replace.Replace($body,$Datum,1)
    

    $message = New-Object System.Net.Mail.MailMessage
    $message.IsBodyHtml = $true
    $message.subject = $subject
    $message.body = $body
    $message.to.add($ManagerMail)
    #$message.bcc.add("Marius.Deffner@cellent.de")
    $message.from = $from
    $attachment = New-Object System.Net.Mail.Attachment($PathToAttachement)
    $attachment.ContentType.MediaType = "image/jpg"
    $attachment.ContentId = "Attachment"
    $message.attachments.add($attachment)
    
    $smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer);
    $smtp.send($message)
}

### Begin Main
$body = ""





$Find = Search-ADAccount -AccountExpiring -timespan "31"

foreach ($item in $find)
{
    #$item
    #"External-Username: " + $item.Name
    #"External-AccountName: " + (get-aduser $item -Properties mailNickname).mailNickname
    #"External-Mail: " + (get-aduser $item -Properties Mail).Mail
    #"AccountExpiration: " + ($item.AccountExpirationDate).ToString("dd.MM.yyyy")
    #"DistinguishedName: " + $item.DistinguishedName
    #"Manager Name: " + (get-aduser(get-aduser(Get-ADUser $item -properties Manager).Manager) -Properties Name).Name
    #"Manager Mail: " + (get-aduser(get-aduser(Get-ADUser $item -properties Manager).Manager) -Properties Mail).mail
    If ($item.DistinguishedName -like "*OU=Internal User*"){$SkipMail = $True}else{$SkipMail = $false}

    if ($SkipMail -ne $True)
    {
        send-Mail-User -ExternalUserName $item.Name -Datum ($item.AccountExpirationDate).ToString("dd.MM.yyyy") -ManagerMail (get-aduser(get-aduser(Get-ADUser $item -properties Manager).Manager) -Properties Mail).mail -ManagerName (get-aduser(get-aduser(Get-ADUser $item -properties Manager).Manager) -Properties Name).Name
    }
    
    #"---------------------------------------"    
}