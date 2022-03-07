#################################################################################
##################################################################################


$ListOnly = $false
$debug = $true
$DomainController = "DE-DCO-202.cellent.int"
$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")

#[pscredential]$oadmin = Get-Credential

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append



$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true
$wb = $xl.Workbooks.Open($MyInvocation.MyCommand.Definition + ".xlsx")
$ws = $wb.Sheets.Item(1)

#write-host "Lastname, Firstname; Titel; Abteilung; Kostenstelle; Manager, Office Location"
#"User;Titel;Abteilung;Kostenstelle;Manager;Office Location"| Out-File "C:\Temp\bla.csv" -Encoding Unicode

for ($i = 2; $i -le 340; $i++) 
{
    $FoundUserAD = $Null
    $NamePreExcel = $ws.Cells.Item($i, 2).Text
    $NamePostExcel = $ws.Cells.Item($i, 1).Text
    $WiproMailExcel = $ws.Cells.Item($i, 3).Text
    
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) " | out-file $LogFile -Append
    

    if (($NamePreExcel -ne "")-and ($NamePostExcel -ne ""))
    {
        $FoundUserAD = Get-ADUser -Filter {(Surname -eq $NamePostExcel) -and (GivenName -eq $NamePreExcel)} -Properties Displayname, SamAccountName -SearchBase "OU=Internal User,OU=Usermanagement,DC=cellent,DC=int" -Server $DomainController
        #"OU=Internal User,OU=Usermanagement,DC=cellent,DC=int"

        if ($debug){write-host $NamePostExcel, $NamePreExcel, $WiproMailExcel}
        if ($debug){write-host $FoundUserAD.Displayname}

    
        if ($FoundUserAD -eq $Null)
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Error: User '$($NamePostExcel), $($NamePreExcel)' not found" | out-file $LogFile -Append
        }
        else
        {
            ### Enable forwarding ################
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) $($NamePostExcel), $($NamePreExcel): Set Value: $($WiproMailExcel)" | out-file $LogFile -Append
            

            if ($ListOnly -eq $false)
            {
                Try
                {
                    $OldValue = (Get-ADUser -Filter {(Surname -eq $NamePostExcel) -and (GivenName -eq $NamePreExcel)} -Properties extensionAttribute6 -Server $DomainController).extensionAttribute6
                     "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Actual WiproMail from AD: '$($OldValue)'" | out-file $LogFile -Append

                    Try
                    {
                        Set-ADUser -Identity $FoundUserAD.SamAccountName -Server $DomainController -Replace @{extensionAttribute6="$($WiproMailExcel)"}
                    }
                    Catch
                    {
                        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())    ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
                    }  



                    $CheckValue = (Get-ADUser -Identity $FoundUserAD.SamAccountName -Properties extensionAttribute6 -Server $DomainController).extensionAttribute6
                    if ($CheckValue -eq $WiproMailExcel)
                    {
                        $Result = 0
                        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())    WiproMail successfully changed" | out-file $LogFile -Append
                    }
                    else
                    {
                        $Result = 7
                        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())    ERROR - Change failed!" | out-file $LogFile -Append
                    }
                }
                Catch
                {
                    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())        ERROR: '$($_.Exception.Message)'" | out-file $LogFile -Append
                }  
                #"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Result: $($Result)" | out-file $LogFile -Append
            }
            
        }
    }
}



$wb.Close()
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Script End ----------------------------------------------------------------------" | out-file $LogFile -Append