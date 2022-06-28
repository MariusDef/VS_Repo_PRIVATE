Start-Process "notepad.exe"
$obj = New-Object -ComObject Wscript.Shell  
    
do {
    Start-Sleep -Seconds 180
    $obj.SendKeys('a')
} while (1 -eq 1)