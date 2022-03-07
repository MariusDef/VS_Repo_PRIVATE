import-module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length-5) + '\configurationManager.psd1')

$Sitecode2012 = 'CPS'
$Server2012 = 'DE-scc-101'

cd "$($sitecode2012):"



Filter Set-PreStagedState
{
    
    #$_.pkgflags
    if ($_.pkgflags -eq ($_.pkgflags -bor 32)) # Wenn Auto, dann setze Manual
    {
        write-host "Check '$($_.Name)'..."
        write-host "...'$($_.Name)' Auto -> Manual"
        $_.pkgflags = $_.pkgflags - 32 + 16777216
        $_.put()
    }
    elseif ($_.pkgflags -eq ($_.pkgflags -bor 16777216)) # Wenn Manual, dann setzte changes
    {
        #write-host "Check '$($_.Name)'..."
        #write-host "...'$($_.Name)' Manual -> Changes"
        #$_.pkgflags = $_.pkgflags - 16777216 + 16
        #$_.put()
    }
    elseif ($_.pkgflags -ne ($_.pkgflags -bor 16)) # Wenn noch nicht auf 'Changes, weil der Wert nicht migriert werden konnte, dann setze
    {
        write-host "Check '$($_.Name)'..."
        write-host "...'$($_.Name)' original migrated. -> Manual"
        $_.pkgflags = $_.pkgflags + 16777216
        $_.put()
    }
    else
    {
        write-host "Check '$($_.Name)'..."
        write-host "...'$($_.Name)' original migrated. -> Manual"
        $_.pkgflags = $_.pkgflags + 16777216
        $_.put()
    }
}

Filter Set-PreStagedStateApplication
{
    write-host "Set '$($_.LocalizedDisplayName)'..."
    Set-CMApplication -id $_.ci_id -DistributionPointSetting DeltaCopy
}

Get-CMPackage | Set-PreStagedState
Get-CMBootImage | Set-PreStagedState
get-CMOperatingSystemImage | Set-PreStagedState
Get-CMOperatingSystemInstaller | Set-PreStagedState
Get-CMApplication | Set-PreStagedStateApplication
get-CMDriverPackage | Set-PreStagedState
Get-CMSoftwareUpdateDeploymentPackage | Set-PreStagedState
