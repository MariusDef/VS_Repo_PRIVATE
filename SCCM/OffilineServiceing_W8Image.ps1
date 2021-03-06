#Initialize Variables
$LogFile = "C:\dism\OfflineServiceing.log"
$UpdatesPath = "C:\dism\2015\" 
$MountPath = “c:\dism\mount”
$dism = "dism.exe"
$WimFile ="C:\dism\BUILD-Windows7x64-BDE_2.3.1.wim"
$TargetOS = "Windows6.1"
$counter = 0

#Prepare Log
echo ((Get-Date -Format g) + " - -------------------------------------------") | Out-File $LogFile -Append
echo ((Get-Date -Format g) + " - Check, if '$MountPath' does exist...") | Out-File $LogFile -Append

#Check MountPath
if (!(Test-Path -Path $MountPath)) 
{
    echo ((Get-Date -Format g) + " - ...does not. Create it") | Out-File $LogFile -Append
    mkdir $MountPath
}
else {echo ((Get-Date -Format g) + " - ...does already exists") | Out-File $LogFile -Append}

#Mount Image
echo ((Get-Date -Format g) + " - Mount image ('$WimFile')...") | Out-File $LogFile -Append
& $dism /mount-wim /wimfile:$WimFile /mountdir:$MountPath /index:1 | Out-File $LogFile -Append

#List Updates to add
echo ((Get-Date -Format g) + " - List Updates wit *.cab to include...") | Out-File $LogFile -Append
$UpdateArray = get-childitem "$UpdatesPath" -recurse -force | where-object { $_.FullName -like '*windows6.1*.cab*' }
$lenght = $updateArray.Length

#Add Updates
ForEach ($Updates in $UpdateArray)
{   
    $counter ++
    $strcmd = $Updates.FullName
    echo ((Get-Date -Format g) + " - $counter von $lenght - Check Update: $strcmd") | Out-File $LogFile -Append
    & $dism /image:"$MountPath" /Add-Package /PackagePath:"$strcmd" | Out-File $LogFile -Append
    echo ((Get-Date -Format g)) | out-file $LogFile -Append
}

#Unmount + Commit and Cleanup the image
echo ((Get-Date -Format g) + " - unmounting image...") | Out-File $LogFile -Append
& $dism /Unmount-Wim /Mountdir:$MountPath /commit | Out-File $LogFile -Append
echo ((Get-Date -Format g) + " - cleanup drives...") | Out-File $LogFile -Append
& $dism /Cleanup-Wim | Out-File $LogFile -Append

echo ((Get-Date -Format g) + " - FINISHED!") | Out-File $LogFile -Append