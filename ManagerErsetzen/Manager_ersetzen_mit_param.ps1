<#
'===================================================================
'
'NAME: Manager ersetzen
'
'AUTHOR: Aron Ziegler, cellent GmbH
'DATE:04.10.2017
'
'COMMENT: Tritt ein Manager aus, kann dieses Script verwendet werden,
'          um alle den Manager bei AD Usern zu ersetzen.
'
'Aufruf: Manager ersetzen.ps1
'
'Changes:
'
'
'===================================================================
#>

param(
[Parameter(Mandatory = $true)][string]$Manager,
[string]$newManager)

# Pfad der Log Datei
$LogPath = "$env:windir\logs\Manager ersetzen.txt"

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

#    write_log 'Text'
Write-Host "`n"
write_log "`n"
if ($Manager -and $newManager)
{
#Manager angeben (Logonname/Samaccountname), welcher das Unternehmen verlässt
try{
    #$Manager = Read-Host "Geben Sie den Manager an, welcher das Unternehmen verlässt"
        if(Get-ADUser -Filter {samAccountName -eq $Manager})
        {
        write_log "Erfolgreiche Eingabe des alten Managers $Manager"
        }
        else{
            
            throw "Falsche eingabe"
            }
    }
    catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            
            write-Host "$ErrorMessage"
            Write-Host "$FailedItem"
            Write-Host "Falsche Eingabe des alten Managers, der User $Manager existiert nicht"

            write_log "Falsche Eingabe des alten Managers, der User $Manager existiert nicht"
            #$Manager = Read-Host "Geben Sie den Manager an, welcher das Unternehmen verlässt"
            write_log "Das Skript wird beendet"
            break
            
        }

#Neuer Manager angeben (Logonname/Samaccountname), welcher den alten Manager ersetzt
try{
    #$newManager = Read-Host "Geben Sie den neuen Manager an"
        if(Get-ADUser -Filter {samAccountName -eq $newManager})
        {
        write_log "Erfolgreiche Eingabe des neuen Managers $newManager"
        }
        else{
            throw "Falsche eingabe"
            }
    }
    catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            
            write-Host "$ErrorMessage"
            Write-Host "$FailedItem"
            Write-Host "Falsche Eingabe des neuen Managers, der User $newManager existiert nicht"
            
            write_log "Falsche Eingabe des neuen Managers - $newManager existiert nicht"
            #$newManager = Read-Host "Geben Sie den neuen Manager an"
            write_log "Das Skript wird beendet"
            break
            
         }

try{
#Auslesen der Email Adresse vom neuen Manager
$x = Get-ADUser -Filter {samAccountName -eq $newManager} -Properties Emailaddress| Select -expand emailaddress
write_log "Auslesen der Email Adresse des neuen Managers $x"
}
catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            
            write-Host "$ErrorMessage"
            Write-Host "$FailedItem"
Write-Host "Es ist ein Fehler beim auslesen der Email Adresse aufgetreten"
write_log "Es ist ein Fehler beim auslesen der Email Adresse aufgetreten"
}

try{
#User auslesen mit alten Manager | Werte der betroffenden User ändern
Write-Host "Folgende User haben den alten Manager $Manager hinterlegt, dieser wird mit dem neuen Manager $newManager ausgetauscht"
Get-ADUser -Filter "manager -eq '$Manager'"-Properties UserPrincipalName| Select -expand UserPrincipalName
Get-ADUser -Filter "manager -eq '$Manager'" | Set-ADUser -Manager "$newManager" -Replace @{extensionAttribute2=$x}
write_log "Auslesen der User mit altem Manager + Ändern des Managers und extensionAttribute2"
}
catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            
            write-Host "$ErrorMessage"
            Write-Host "$FailedItem"
Write-Host "Es ist ein Fehler beim ändern der betroffenen User aufgetreten"
write_log "Es ist ein Fehler beim ändern der betroffenen User aufgetreten"
}

 write_log "---------------------------"
 write_log "`n"
 }
 else
 {
 Write-Host "Folgende User haben den Manager $Manager hinterlegt"
 try{
 Get-ADUser -Filter "manager -eq '$Manager'"-Properties UserPrincipalName| Select -expand UserPrincipalName
 write_log "Ausgabe der User, welchen den Manager $Manager hinterlegt haben"
 }
 catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            
            write-Host "$ErrorMessage"
            Write-Host "$FailedItem"
            Write-Host "Es ist ein Fehler beim abrufen der User aufgetreten"
            write_log "Es ist ein Fehler beim abrufen der User aufgetreten"
            write_log "Das Skript wird beendet"
            break
            
 }

 }
 write_log "Das Skript wird erfolgreich beendet"