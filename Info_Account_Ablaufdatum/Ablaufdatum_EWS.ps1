$LogPath = "$env:windir\logs\Account-AblaufInfo_extern_EWS.log"
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

function send-Mail-User([string]$ExternalUserName, [string]$Datum, [string]$ManagerMail, [string]$ManagerName)
{
    #$ManagerMail = "Marius.Deffner@wipro.com"

    $SMTPServer = "ms-az-smtp-gw-1.cellent.int" 
    $from = "svc.itsc@wipro.com"
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
        write_log  'Exchange Module imported correctly'
    }


    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService #-ArgumentList Exchange2010
    $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList  $username, $password
    $service.Url = $EWSurl
    $message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
    $message.Subject = $subject
    $message.body = $body
    $message.BccRecipients.Add("Marius.Deffner@wipro.com")
    $message.ToRecipients.Add($ManagerMail)
    $message.Attachments.AddFileAttachment($PathToAttachement)
    $message.SendAndSaveCopy()
    
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