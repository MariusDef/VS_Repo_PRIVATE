function send-Mail-User([String]$Body)
{
    $MailTo1 = "ITSC.itsc@wipro.com"
    $SMTPServer = "ms-az-smtp-gw-1.cellent.int" 
    $from = "ITSC@cellent.de"
    $subject = "Expired Account overview"
        
    
    $message = New-Object System.Net.Mail.MailMessage
    $message.IsBodyHtml = $false
    $message.subject = $subject
    $message.body = $body
    $message.to.add($MailTo1)
    $message.from = $from
    
    $smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer);
    $smtp.send($message)
}

$expireds = Search-ADAccount -AccountExpired | where {$_.Enabled -eq $true} | Select Name, AccountExpirationdate,DistinguishedName,Enabled,@{Name='Manager';Expression={(Get-ADUser(Get-ADUser $_ -properties Manager).manager).Name}} | fl Name, AccountExpirationdate, Enabled, Manager
    if ($expireds -eq $null){
        $body = "Good morning, `n`nThere are no currently expired accounts."
        }
    else {
        $body ="Good morning, `n`nThe following accounts have been detected as expired yet still active"
        $out = $expireds | Out-String
        $body += $out
        }


$expiring = Search-ADAccount -AccountExpiring -TimeSpan "31" | where {$_.Enabled -eq $true}| Select Name, AccountExpirationdate ,DistinguishedName,Enabled,@{Name='Manager';Expression={(Get-ADUser(Get-ADUser $_ -properties Manager).manager).Name}}  | fl Name, AccountExpirationdate, Enabled, Manager
    if ($expiring -eq $null){
        $body += ""
        $body += ""
        $body += "No accounts are expiring within the next 31 days."
        }
    Else 
    {
        
        $body += "The following accounts were detected as expiring soon:"
        $out = $expiring | Out-String
        $body += $out
    }

#Write-Host $body
send-Mail-User -Body $body