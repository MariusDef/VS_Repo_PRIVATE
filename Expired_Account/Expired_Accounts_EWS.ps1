$LogPath = "$env:windir\logs\Expired_Accounts_EWS.log"
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

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

function send-EWSMail-User
    ( 
        [Parameter(mandatory=$true, Position=0)]
        $body,
        [String]$attachment,
        [String]$domain = "Wipro"
    )
{
    [String]$username = 'sv43448t@wipro.com'
    [String]$password = 'S6h}\Xu$'
    [String]$to = "ITSC.itsc@wipro.com"
    [String]$subject = "Expired/Expiring Account overview"
    $EWSurl = 'https://webmail2.wipro.com/ews/exchange.asmx'
    $dllPath = "C:\Temp\2.2\Microsoft.Exchange.WebServices.dll"
    #$dllPath = """C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"""
    #write_log $dllpath
    #"3"
    if (-not(Get-Module "Microsoft.Exchange.WebServices"))
    {
        write_log "Exchange Module not imported... import it"
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
    #write_log "4"
    $message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
    $message.Subject = $subject
    #$message.Body.BodyType = 'text'
    #write_log "5"
    $message.Body = $body
    #write_log "6"
    $message.ToRecipients.Add($to) | Out-Null
    If ($attachment){$message.Attachments.AddFileAttachment($attachment) | Out-Null}
    #write_log "7"
    $message.Send()
    #write_log "8"
}

#send-EWSMail-User -body 'body `r of `r`n test\r\nmail <br> new <br /> line'
#exit

$expireds = Search-ADAccount -AccountExpired | where {$_.Enabled -eq $true} | Select Name, AccountExpirationdate,DistinguishedName,Enabled,@{Name='Manager';Expression={(Get-ADUser(Get-ADUser $_ -properties Manager).manager).Name}} | fl Name, AccountExpirationdate, Enabled, Manager
    if ($expireds -eq $null){
        $body = 'Good morning, There are no currently expired accounts.'
        $bexpired = $False
        }
    else {
        $expireds | Out-File "C:\temp\Expired-but-enabled.txt" -Force | Out-Null
        $body ='Good morning, the following accounts have been detected as expired yet still active'
        #$out = $expireds + '<br>' | Out-String
        #$body += $out + '<br>'
        $bexpired = $True
        }


$expiring = Search-ADAccount -AccountExpiring -TimeSpan "31" | where {$_.Enabled -eq $true}| Select Name, AccountExpirationdate ,DistinguishedName,Enabled,@{Name='Manager';Expression={(Get-ADUser(Get-ADUser $_ -properties Manager).manager).Name}}  | fl Name, AccountExpirationdate, Enabled, Manager
    if ($expiring -eq $null){
        $body2 += 'No accounts are expiring within the next 31 days.'
        $bexpiring = $False
        }
    Else 
    {
        $expiring | Out-File "C:\Temp\Expiring.txt" -Force | Out-Null
        $body2 = 'The following accounts were detected as expiring soon:<br>'
        #$out = $expiring + '<br>' | Out-String
        #$body += $out + '<br>'

        $bexpiring = $True
    }

#Write-Host $body
#send-Mail-User -Body $body
if (($bexpired -eq $True) -or ($bexpired -eq $False)) 
{
    write_log "Start sending 'expired'..."
    $bsuccess = $True
    do
    {
        try
        {
            if ($bexpired)
            {
                #write_log  "1"
                $body
                send-EWSMail-User -body $body -attachment "c:\temp\Expired-but-enabled.txt"
                #write_log  "2"
            }
            else
            {#write_log "1"
                send-EWSMail-User -body $body
            }
            $bsuccess = $True
            write_log "Send 'expired' successfully"
        }
        catch
        {
            $bsuccess = $false
            write_log  $_
            write_log  "Send 'expired' failed... retry"
        }
        sleep -Seconds 5
    } until ($bsuccess -eq $True)
}

if (($bexpiring -eq $True) -or ($bexpiring -eq $False)) 
{
    write_log "Start sending 'expiring'..."
    $bsuccess = $True
    do
    {
        try
        {
            if ($bexpired)
            {
                send-EWSMail-User -body $body2 -attachment "C:\Temp\Expiring.txt"
            }
            else
            {
                send-EWSMail-User -body $body2
            }
            $bsuccess = $True
            write_log "Send 'expiring' successfully"
        }
        catch
        {
            $bsuccess = $false
            write_log "Send 'expiring' failed... retry"
        }
        sleep -Seconds 5
    } until ($bsuccess -eq $True)
}