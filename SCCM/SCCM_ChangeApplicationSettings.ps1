import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

cd csp:


$Applications = Get-CMApplication

foreach ($App in $Applications.LocalizedDisplayName)
{
    $SplitName = $App.split("_")
    Write-Host ($app + " -> Publisher: " + $SplitName[1] + " - & - SoftwareVersion: " + $SplitName[3] + "_" + $SplitName[4])
    
    Set-CMApplication -Name $App -Publisher $SplitName[1] -SoftwareVersion ($SplitName[3] + "_" + $SplitName[4]) -Owner " " -SupportContact " "
}
