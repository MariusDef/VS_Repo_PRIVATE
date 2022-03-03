function Show-CustomMenu
{
    Clear-Host
    Write-Host "================ Select Destination ================"
    
    Write-Host "1: Wähle '1' für D:\"
    Write-Host "2: Wähle '2' für E:\"
    Write-Host "3: Wähle '3' für F:\"
    Write-Host "4: Wähle '4' für C:\Temp"
    Write-Host ""
    Write-host ""
}

Show-CustomMenu
$Selection = Read-Host "Select Zip Destination by number"

switch ($Selection)
{
    '1' {$Laufwerk = "D:\"}
    '2' {$Laufwerk = "E:\"}
    '3' {$Laufwerk = "F:\"}
    '4' {$Laufwerk = "C:\Temp\"}
}

$7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"

if (-not (Test-Path -Path $7zipPath -PathType Leaf)) {
    throw "7 zip file '$7zipPath' not found"
    exit
}

Set-Alias 7zip $7zipPath

$SourceSyteme = "\\winas\wiabt$\IM_NotfallOfflineExport\Systeme"
$sourceDocs = "\\winas\wiabt$\IM_NotfallOfflineExport\Documentations"


#Documents
$Target = $Laufwerk + "WIRTGEN_Docs_" + (get-date -Format 'yyyyMMddHHmm') +".7z"
7zip a -r -mx=0 $Target $sourceDocs 

#Systeme
$Target = $Laufwerk + "WIRTGEN_Systeme_" + (get-date -Format 'yyyyMMddHHmm') +".7z"
7zip a -r -mx=0 $Target $SourceSyteme
