import-module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length-5) + '\configurationManager.psd1')

$Sitecode2012 = 'CPS'
$Server2012 = 'DE-SCC-101'

$FilePath = "\\de-tab-1012\prestaged$\"
$DP = "de-scc-201.cellent.int"
$SkipPackages = @()


cd "$($sitecode2012):"


Function Publish-PackageToPKGX
{
    Param
    (
        $PackageID,
        $FileName,
        $Type
    )

    if ($SkipPackages -notcontains $PackageID)

        
    {
        cd C:
    
        #Write-Host "Check '$FileName'..."

        if (-not (Test-Path -Path $FileName))
        {
            Write-Host "Check '$FileName'..."
            Write-host "...Create File"
            cd "$($sitecode2012):"
            switch ($Type)
            {
                1 {Publish-CMPrestageContent -PackageId $PackageID -FileName $FileName -DistributionPointName $DP}
                2 {Publish-CMPrestageContent -BootImageId $PackageID -FileName $FileName -DistributionPointName $DP}
                3 {Publish-CMPrestageContent -OperatingSystemImageId $PackageID -FileName $FileName -DistributionPointName $DP}
                4 {Publish-CMPrestageContent -OperatingSystemInstallerId $PackageID -FileName $FileName -DistributionPointName $DP}
                5 {Publish-CMPrestageContent -DriverPackageId $PackageID -FileName $FileName -DistributionPointName $DP}
                6 {Publish-CMPrestageContent -DeploymentPackageId $PackageID -FileName $FileName -DistributionPointName $DP}
            }   
        }
        else
        {
            #Write-host "...already created"
        }
        cd "$($sitecode2012):"
    }
}


get-CMPackage | foreach{Publish-PackageToPKGX $_.PackageID "$($FilePath)$($_.Name)_$($_.Version)__$($_.PackageID)__V$($_.SourceVersion).pkgx" 1}
Get-CMBootImage | foreach{Publish-PackageToPKGX $_.PackageID "$($FilePath)$($_.Name)_$($_.Version)__$($_.PackageID)__V$($_.SourceVersion).pkgx" 2}
get-CMOperatingSystemImage | foreach{Publish-PackageToPKGX $_.PackageID "$($FilePath)$($_.Name)_$($_.Version)__$($_.PackageID)__V$($_.SourceVersion).pkgx" 3}
Get-CMOperatingSystemInstaller | foreach{Publish-PackageToPKGX $_.PackageID "$($FilePath)$($_.Name)_$($_.Version)__$($_.PackageID)__V$($_.SourceVersion).pkgx" 4}
get-CMDriverPackage | foreach{Publish-PackageToPKGX $_.PackageID "$($FilePath)$($_.Name)_$($_.Version)__$($_.PackageID)__V$($_.SourceVersion).pkgx" 5}
Get-CMSoftwareUpdateDeploymentPackage | foreach{Publish-PackageToPKGX $_.PackageID "$($FilePath)$($_.Name)_$($_.Version)__$($_.PackageID)__V$($_.SourceVersion).pkgx" 6}