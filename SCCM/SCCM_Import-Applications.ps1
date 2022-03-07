Param
(
    [Parameter(Mandatory = $true)]
    [string] $ApplicationName,
    [Parameter(Mandatory = $true)]
    [ValidateSet("Optional","Core")]
    [string] $ApplicationType,
    [Parameter(Mandatory = $true)]
    [bool] $Lic,
    [Parameter(Mandatory = $true)]
    [ValidateSet("P","V4","V5")]
    [string] $DeployType
)

#Import-Module D:\SCCM\AdminConsole\bin\ConfigurationManager.psd1
import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

$SiteCode = "CSP"
$SiteServer = "WLISCCM002.VPBANK.NET"
$SiteCodeConnection = "CSP:"
$CollFolder = "(B) Software Deployment"
$AppFolder = "(B0) DEVELOPMENT"
#$ApplicationName = "CSP_Adobe_Adobe-Reader_11.0.6_01_ML"
$PathtoSWPool = "\\WLISCCM002\SWPool$\"
$CollComment = "(disabled)"

cd $SiteCodeConnection


function Add-VPNewApplication
{
    pause
    $ApplicationCollectionNameInstall = "B1 - " + $ApplicationName + " - Install"
    $ApplicationCollectionNameUnInstall = "B1 - " + $ApplicationName + " - Un-Install"

    #Create Install Collection
    if (-not(Get-CMDeviceCollection -Name $ApplicationCollectionNameInstall))
    {
        write-host ($ApplicationName + ": Create Collection Install...")
        
        if ($lic -eq $true)
        {
            $CollComment = ($CollComment + "(lic)")
        }
        $ApplicationCollectionNameInstallCreate = New-CMDeviceCollection -LimitingCollectionId SMS00001 -Name $ApplicationCollectionNameInstall -RefreshType Manual -Comment $CollComment | Out-Null
    }

    #Move Inst Collection to Subfolder
    $returnInst = (Get-CMDeviceCollection -Name $ApplicationCollectionNameInstall)
    Move-CMObject -ObjectID $returnInst.CollectionID -CurrentFolderID 0 -TargetFolderID (Get-CMCollFolderID($CollFolder)) -ObjectTypeID 5000 | Out-Null



    #Create Un-Install Collection
    if ($ApplicationType -eq "Optional")
    {
        if (-not(Get-CMDeviceCollection -Name $ApplicationCollectionNameUnInstall))
        {
            write-host ($ApplicationName + ": Create Collection Un-Install...")
            $Schedule = New-CMSchedule -RecurInterval Days -RecurCount 1 -Start "06/11/2014 11:00 PM"
            $ApplicationCollectionNameUnInstallCreate = New-CMDeviceCollection -LimitingCollectionId SMS00001 -Name $ApplicationCollectionNameUnInstall -RefreshType Periodic -RefreshSchedule $Schedule | Out-Null
            #Write-host $ApplicationCollectionNameUnInstallCreate
        }

        #Move Un-Inst Collection to Subfolder
        $returnUninst = (Get-CMDeviceCollection -Name $ApplicationCollectionNameUnInstall)
        Move-CMObject -ObjectID $returnUninst.CollectionID -CurrentFolderID 0 -TargetFolderID (Get-CMCollFolderID($CollFolder)) -ObjectTypeID 5000 | Out-Null


        #Create Query
        $returnUninst = (Get-CMDeviceCollection -Name $ApplicationCollectionNameUnInstall)
        if (-not(Get-CMDeviceCollectionQueryMembershipRule -CollectionId $returnUninst.CollectionID -RuleName $ApplicationName))
        {
            write-host ($ApplicationName + ": Create Query for Uninst-Collection...")
            $QueryRuleName = $ApplicationName
            #$QueryRule = ("select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_VPBANK_CORPPACKAGES on SMS_G_System_VPBANK_CORPPACKAGES.ResourceID = SMS_R_System.ResourceId where SMS_G_System_VPBANK_CORPPACKAGES.KeyName = "CSP_Adobe_Acrobat-Standard_X_01_ML" and SMS_G_System_VPBANK_CORPPACKAGES.Result = "SUCCESS"")
            $QueryRule = ("select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_VPBANK_CORPPACKAGES on SMS_G_System_VPBANK_CORPPACKAGES.ResourceId = SMS_R_System.ResourceId where SMS_G_System_VPBANK_CORPPACKAGES.KeyName = '" + $ApplicationName + "' and SMS_G_System_VPBANK_CORPPACKAGES.Result = 'SUCCESS'")
            Add-CMDeviceCollectionQueryMembershipRule -CollectionId $returnUninst.CollectionID -RuleName $QueryRuleName -QueryExpression $QueryRule | Out-Null
        }
    
        
        #Create Exclude for install collection
        $returnInst = (Get-CMDeviceCollection -Name $ApplicationCollectionNameInstall)
        $returnUninst = (Get-CMDeviceCollection -Name $ApplicationCollectionNameUnInstall)
        if (-not(Get-CMDeviceCollectionExcludeMembershipRule -CollectionId $returnUninst.CollectionID -ExcludeCollectionId $returnInst.CollectionID))
        {
            write-host ($ApplicationName + ": Create Exclude for Uninst Collection...")
            Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $returnUninst.CollectionID -ExcludeCollectionId $returnInst.CollectionID
        }
    }

    #Check for Coll Subfolder
    $Return = (Get-Folder ($CollFolder))
    If ($Return = $false)
    {
        Add-Folder ($CollFolder)
    }
    
    #CreateApplication
    if (-not(Get-CMApplication -Name $ApplicationName))
    {
        echo ($ApplicationName + ": Create application...")
        New-CMApplication -Name $ApplicationName -Description "Created by Script" -LocalizedApplicationName $ApplicationName -AutoInstall $True | Out-Null
        $Jobs_Application = $True
    }
    
    #Create DeploymentType
    if (-not(Get-CMDeploymentType -ApplicationName $ApplicationName))
    {
        echo ($ApplicationName + ": Create deployment type...")
        if ($ApplicationType -eq "Core")
        {
            if ($DeployType -eq "P")
            {
                Add-CMDeploymentType -ApplicationName $ApplicationName  -ScriptInstaller -ManualSpecifyDeploymentType -DeploymentTypeName $ApplicationName -ContentLocation ($PathtoSWPool + $ApplicationName) -AdministratorComment "Created by Script" -InstallationProgram "Wrapper.exe Wrapper_reinst.ini" -LogonRequirementType WhetherOrNotUserLoggedOn -InstallationBehaviorType InstallForSystem -MaximumAllowedRunTimeMinutes "120" -AllowClientsToShareContentOnSameSubnet $True -DetectDeploymentTypeByCustomScript -ScriptContent "echo $true" -ScriptType PowerShell 
            }
        }
        else
        {
        if ($DeployType -eq "P")
            {
                Add-CMDeploymentType -ApplicationName $ApplicationName  -ScriptInstaller -ManualSpecifyDeploymentType -DeploymentTypeName $ApplicationName -ContentLocation ($PathtoSWPool + $ApplicationName) -AdministratorComment "Created by Script" -InstallationProgram "Wrapper.exe Wrapper_reinst.ini" -UninstallProgram "Wrapper.exe Wrapper_deinst.ini" -LogonRequirementType WhetherOrNotUserLoggedOn -InstallationBehaviorType InstallForSystem -MaximumAllowedRunTimeMinutes "120" -AllowClientsToShareContentOnSameSubnet $True -DetectDeploymentTypeByCustomScript -ScriptContent "echo $true" -ScriptType PowerShell 
            }
        }

        #Move Application to Subfolder
        echo ($ApplicationName + ": Move application to subfolder...")
        $return = Get-CMApplication -Name $ApplicationName
        Move-CMObject -ObjectID $return.ModelName -CurrentFolderID 0 -TargetFolderID (Get-CMAppFolderID($AppFolder)) -ObjectTypeID 6000 | out-null
    
        #Distribute Package to all DPs
        echo ($ApplicationName + ": Distribute Binaries to Test DP Group...")
        $Return = (Get-CMDistributionPointGroup | Where-Object {$_.Name -like "*TEST*"})
        Start-CMContentDistribution -ApplicationName $ApplicationName -DistributionPointGroupName ($Return.Name) | Out-Null
        $Jobs_DeplyomentType = $True
    }

    #Create Deployment - Install
    if (-not(Get-CMDeployment -CollectionName $ApplicationCollectionNameInstall))
    {
        echo ($ApplicationName + ": Create deployment for install...")
        Start-CMApplicationDeployment -CollectionName $ApplicationCollectionNameInstall -Name $ApplicationName -AppRequiresApproval $false -PreDeploy $false -TimeBaseOn LocalTime -Comment "Created by Script" -UserNotification DisplayAll -DeployAction Install -DeployPurpose Required -AvaliableDate 2014/06/11 -AvaliableTime 00:00
    }

    if ($ApplicationType -eq "Optional")
    {
        #Create Deployment - Un-Install
        if (-not(Get-CMDeployment -CollectionName $ApplicationCollectionNameUnInstall))
        {
            echo ($ApplicationName + ": Create deployment for Un-install...")
            Start-CMApplicationDeployment -CollectionName $ApplicationCollectionNameUnInstall -Name $ApplicationName -AppRequiresApproval $false -PreDeploy $false -TimeBaseOn LocalTime -Comment "Created by Script" -UserNotification DisplayAll -DeployAction Uninstall -DeployPurpose Required -AvaliableDate 2014/06/11 -AvaliableTime 00:00
        }
    }

    if (($Jobs_Application -eq $True) -or ($Jobs_DeplyomentType -eq $True)) {""}
    if (($Jobs_Application -eq $True) -or ($Jobs_DeplyomentType -eq $True)) {"####" + $ApplicationName + "####"}
    if ($Jobs_Application -eq $True) {echo "- Anlegen eines Icons in der Application"}
    if ($Jobs_DeplyomentType -eq $True) {echo "- Wenn ein MSI in der Installation ist, den ProductCode in den DeploymentType eintragen"}
    if ($Jobs_DeplyomentType -eq $True) {echo "- Erstellen der DetectionMethod"}
    if (($Jobs_Application -eq $True) -or ($Jobs_DeplyomentType -eq $True)) {"#### --------------------- ####"}
    if (($Jobs_Application -eq $True) -or ($Jobs_DeplyomentType -eq $True)) {""}
    if (($Jobs_Application -eq $True) -or ($Jobs_DeplyomentType -eq $True)) {""}

    c:
}

Function Move-CMObject
{
    [CmdLetBinding()]
    Param(
    [Parameter(Mandatory=$True,HelpMessage="Please Enter Object ID")]
              [ARRAY]$ObjectID,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter current folder ID")]
              [uint32]$CurrentFolderID,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter target folder ID")]
              [uint32]$TargetFolderID,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter object type ID")]
              [uint32]$ObjectTypeID              
        )
 
    Try{
        Invoke-WmiMethod -Namespace "Root\SMS\Site_$SiteCode" -Class SMS_objectContainerItem -Name MoveMembers `
        -ArgumentList $CurrentFolderID,$ObjectID,$ObjectTypeID,$TargetFolderID -ComputerName $SiteServer -ErrorAction STOP
    }
    Catch{
        $_.Exception.Message
    }  
}

Function Get-CMAppFolderID
{
    [CmdLetBinding()]
    Param(

    [Parameter(Mandatory=$True,HelpMessage="Please Enter ApplicationName")]
              $FolderName
                 
        )
 
    Try{
        $global:erg = Get-wmiObject -Namespace root\SMS\site_$SiteCode -Query "Select name,containernodeid,objecttype from SMS_ObjectContainerNode WHERE `
         name ='$FolderName' and objecttype ='6000'"
         #write-host $global:erg
    }
    
    Catch{
        $_.Exception.Message
    }  
    #echo $erg.ContainerNodeID
    return $erg.ContainerNodeID
 
}

Function Get-CMCollFolderID
{
    [CmdLetBinding()]
    Param(

    [Parameter(Mandatory=$True,HelpMessage="Please Enter ApplicationName")]
              $FolderName
                 
        )
 
    Try{
        $global:erg = Get-wmiObject -Namespace root\SMS\site_$SiteCode -Query "Select name,containernodeid,objecttype from SMS_ObjectContainerNode WHERE `
         name ='$FolderName' and objecttype ='5000'"
         #write-host $global:erg
    }
    
    Catch{
        $_.Exception.Message
    }  
    #echo $erg.ContainerNodeID
    return $erg.ContainerNodeID
 
}

function Get-Folder {([string]$FolderName)
    $objectType = "5000"
    $FolderExist = $false
    $folderID = GWMI -Namespace "ROOT\SMS\Site_$sitecode" -Query `
    "SELECT ContainerNodeID FROM SMS_ObjectContainerNode WHERE Name = '$($FolderName)' AND ObjectType = '$objectType'"

    If ($folderID -eq "" -Or $folderID -eq $Null)
    {
        $FolderExist = $True
    }
    else 
    { 
        $FolderExist = $False  
    }
    return $FolderExist
}

function Add-Folder {([string]$FolderName)
$objectType = "5000"
     $folderID = GWMI -Namespace "ROOT\SMS\Site_$sitecode" -Query `
    "SELECT ContainerNodeID FROM SMS_ObjectContainerNode WHERE Name = '$($FolderName)' AND ObjectType = '$objectType'"

    If ($folderID -eq "" -Or $folderID -eq $Null)
    {
      $folderClass                     = [WMIClass] "ROOT\SMS\Site_$($sitecode):SMS_ObjectContainerNode"
      $newFolder                       = $folderClass.CreateInstance()
      $newFolder.Name                  = $FolderName
      $newFolder.ObjectType            = $objectType
      $newFolder.ParentContainerNodeid = "0"
      $folderPath                      = $newFolder.Put()

      
    }
    Else
    {
      Write-Host "[ERROR]`t CollectionFolder [$($FolderName)] already exists with ID [$($folderID.ContainerNodeID)]" -foregroundcolor Red
    }

}

Add-VPNewApplication