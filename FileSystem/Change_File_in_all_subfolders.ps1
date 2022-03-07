$SearchPath = "d:\SWPool"
$SourceFile = "C:\Users\extdem\Desktop\_COMPANYCODE_MANUFACTURER_PRODUCT_VERSION_RELEASE_LANGUAGE\Wrapper.exe.config"
$DestObjects = Get-ChildItem -Path $SearchPath -Recurse | Where-Object {$_.Name -eq "wrapper.exe.config"}
foreach ($DestObject in $DestObjects)
{
    Write-Host ("Copy from '" + $SourceFile + "' to '" + $DestObject.FullName + "'")
    Copy-Item $SourceFile $DestObject.FullName
}
