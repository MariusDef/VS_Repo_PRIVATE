#$username = "SV43448T@wipro.com"
#$password = "20,Kanalratte"

[System.Net.ServicePointManager]::SecurityProtocol = 'TLS12'

Send-MailMessage -From "svc.itsc@wipro.com" -To "Marius.Deffner@wipro.com" -Subject "testscript" -Body "erst mal nichts" -SmtpServer "Webmail2.wipro.com" -Port "587" -Credential $cred -UseSsl