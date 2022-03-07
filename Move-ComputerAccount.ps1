$ComputerName = 'DE-WKS-1005'
$DestOU = 'OU=WIN10,OU=Clients,OU=Administration,DC=cellent,DC=int'

Enter-PSSession -ComputerName "DE-DCO-202"
#cd\
#Import-Module -Assembly "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\ActiveDirectory\ActiveDirectory.psd1"

get-adcomputer $ComputerName | MoveADObject -TargetPath $DestOU

exit