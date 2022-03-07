Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')
cd p01:

$OutFile = "F:\Scripts\Unsorted\AddGlobalExcludeforAppDeployment.log"
$CollToExclude = "P010176B" # @_No_Installations_Exclusion
$Counter = 0

$Collections  = Get-CMDeviceCollection -name "* - Install - *"
#echo $Collection.Name | Out-File $OutFile -Append

foreach ($Collection in $Collections)
{
    $Counter ++

    if (-not(Get-CMDeviceCollectionExcludeMembershipRule -CollectionId $Collection.CollectionID -ExcludeCollectionId $CollToExclude -ErrorAction Continue)) 
    {
        Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $Collection.CollectionID -ExcludeCollectionId $CollToExclude | Out-Null
        echo ((Get-Date -Format g) + " - " + $Counter + ": ADD to Coll " + $Collection.Name) | Out-File $OutFile -Append
    }
    Else
    {
        echo ((Get-Date -Format g) + " - " + $Counter + ": Check collection " + $Collection.Name + " ---> already set") | Out-File $OutFile -Append
    }
}

echo "--------------------------------------------" | Out-File $OutFile -Append