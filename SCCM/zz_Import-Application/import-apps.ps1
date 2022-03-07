####################################
### Matthias Berroth, cellent AG ###
### Consultant                   ###
### 07 September 2015            ###
### Import-Application           ###
####################################

#Define Variables
$SiteCode = "P00"
$SiteServer = "DEHNM-SCC-PP101.ACAG.AC.ALTANA"
$Comment = "(disabled)"
$Folder = "P00 - (B) Software Deployment"
$PathtoSWPool = '\\DEHNM-SCC-PP101\applications$\'
$IconFileTypes = "*dll","*exe","*jpeg","*jpg","*ico","*png"
$Suffix1 = " - AUTH"
$Suffix2 = " - UNINST"
$CollectionMembershipQuery = 'select *  from  SMS_R_System inner join SMS_G_System_MICROSOFT_Altana_Packages_1_1 on SMS_G_System_MICROSOFT_Altana_Packages_1_1.ResourceId = SMS_R_System.ResourceId where SMS_G_System_MICROSOFT_Altana_Packages_1_1.KeyName = "[PACKAGENAME]" and SMS_G_System_MICROSOFT_Altana_Packages_1_1.Result = "SUCCESS"'
$DistributionPointGroupName = "01 - TEST"
$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$Logfile = "$scriptPath\import-applications.log"
$finaloutput = ""

#Import Module 
import-module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length-5) + '\configurationManager.psd1')
. "$scriptPath\Log-ScriptEvent.ps1"


#Change to SCCM Dir

function select-Applications
{
    c:
    $Application = Get-ChildItem $PathtoSWPool | select-object Name | Out-GridView -PassThru
    Log-ScriptEvent -NewLog $Logfile -Value "$Application selected" -Component $scriptname -Severity 1
    return $Application.Name
}
#functions to check 
function check-Application
{
    param([String]$AppName)
    cd "$($SiteCode):"
    $return = (Get-CMApplication -Name $AppName) -ne $null 
    if($return)
    {
        Log-ScriptEvent -NewLog $Logfile -Value "$AppName already exists" -Component $scriptname -Severity 2
    }
    else
    {
        Log-ScriptEvent -NewLog $Logfile -Value "$AppName does not exist" -Component $scriptname -Severity 1
    } 
    c:  
    return $return
}
function check-DeploymentType
{
    param([String]$AppName)
    cd "$($SiteCode):"
    $return = (Get-CMDeploymentType -ApplicationName $AppName -DeploymentTypeName "Install") -ne $null 
    if($return)
    {
        Log-ScriptEvent -NewLog $Logfile -Value "Deployment Type Install for $AppName already exists" -Component $scriptname -Severity 2
    }
    else
    {
        Log-ScriptEvent -NewLog $Logfile -Value "Deployment Type Install for $AppName does not exist" -Component $scriptname -Severity 1
    } 
    c:
    return $return
}
function check-ApplicationDeployment
{
    param([String]$AppName, [Switch]$Install)
    cd "$($SiteCode):"
    if($Install)
    {
        $CollectionName = "$AppName$Suffix1"
        $Deployment = Get-CMDeployment -CollectionName $CollectionName | Where-Object {$_.SoftwareName -eq $AppName} | Where-Object {$_.DesiredConfigType -eq 1}
        if($Deployment)
        {
            Log-ScriptEvent -NewLog $Logfile -Value "Install Deployment for $AppName on collection $CollectionName already exists" -Component $scriptname -Severity 2
            c:
            return $true
        }
        else
        {
            Log-ScriptEvent -NewLog $Logfile -Value "Install Deployment for $AppName on collection $CollectionName does not exist" -Component $scriptname -Severity 1
            c:
            return $false
        }
    }
    else
    {
        $CollectionName = "$AppName$Suffix2"
        $Deployment = Get-CMDeployment -CollectionName $CollectionName | Where-Object {$_.SoftwareName -eq $AppName} | Where-Object {$_.DesiredConfigType -eq 2}
        if($Deployment)
        {
            Log-ScriptEvent -NewLog $Logfile -Value "Uninstall Deployment for $AppName on collection $CollectionName already exists" -Component $scriptname -Severity 2
            c:
            return $true
        }
        else
        {
            Log-ScriptEvent -NewLog $Logfile -Value "Uninstall Deployment for $AppName on collection $CollectionName does not exist" -Component $scriptname -Severity 1
            c:
            return $false
        }
    }
}
function check-collection
{
    param([String]$AppName, [Switch]$IsAuth)
    cd "$($SiteCode):"
    if($IsAuth)
    {
        $CollectionName = "$AppName$Suffix1"
        $Collection = Get-CMDeviceCollection -Name $CollectionName
        if($Collection)
        {
            Log-ScriptEvent -NewLog $Logfile -Value "$CollectionName already exists" -Component $scriptname -Severity 2
            c:
            return $true
        }
        else
        {
            Log-ScriptEvent -NewLog $Logfile -Value "$CollectionName does not exist" -Component $scriptname -Severity 1
            c:
            return $false
        }
    }
    else
    {
        $CollectionName = "$AppName$Suffix2"
        $Collection = Get-CMDeviceCollection -Name $CollectionName
        if($Collection)
        {
            Log-ScriptEvent -NewLog $Logfile -Value "$CollectionName already exists" -Component $scriptname -Severity 2
            c:
            return $true
        }
        else
        {
            Log-ScriptEvent -NewLog $Logfile -Value "$CollectionName does not exist" -Component $scriptname -Severity 1
            c:
            return $false
        }
    }

}

#Functions to create
function create-application
{
    param([String]$AppName)
    cd "$($SiteCode):"
    if(-not(check-Application -AppName $AppName))
    {
        $SplitedName =  $AppName.Split('_')
        $Version = $SplitedName[1] + '_' + $SplitedName[4]
        $Publisher = $SplitedName[3]
        Log-ScriptEvent -NewLog $Logfile -Value "Start to create $AppName" -Component $scriptname -Severity 1
        Log-ScriptEvent -NewLog $Logfile -Value "Version is $Version" -Component $scriptname -Severity 1
        Log-ScriptEvent -NewLog $Logfile -Value "Publisher is $Publisher" -Component $scriptname -Severity 1
        $IconPath = "$PathtoSWPool$AppName\_Icon"
        Log-ScriptEvent -NewLog $Logfile -Value "Check whether $IconPath exists" -Component $scriptname -Severity 1
        if(Test-Path "$IconPath")
        {
            Log-ScriptEvent -NewLog $Logfile -Value "$IconPath exists." -Component $scriptname -Severity 1
            $Icon = (Get-ChildItem $IconPath).Name
            $i = 0
            foreach($item in $Icon)
            {
                foreach($pattern in $IconFileTypes)
                {
                    if($item -like $pattern)
                    {
                        $i++
                    }
                }
            }
            if($i -eq 1)
            {
                Log-ScriptEvent -NewLog $Logfile -Value "$Icon found in folder $IconPath" -Component $scriptname -Severity 1
                $IconLocationFile = $IconPath + "\" + $Icon
                $UseIcon = $true
            }
            elseif($i -gt 1)
            {
                Log-ScriptEvent -NewLog $Logfile -Value "More than one Icon in $IconPath. Will not set a icon for application." -Component $scriptname -Severity 3
                $UseIcon = $false
            }
            else
            {
                Log-ScriptEvent -NewLog $Logfile -Value "No valid Icon in $IconPath. Will not set a icon for application." -Component $scriptname -Severity 3
                $UseIcon = $false
            }
        }
        Log-ScriptEvent -NewLog $Logfile -Value "Will create application $AppName" -Component $scriptname -Severity 1
        cd "$($SiteCode):"
        try
        {
            if($UseIcon)
            {
                New-CMApplication -Name $AppName -Publisher $Publisher -SoftwareVersion $Version -LocalizedApplicationName $AppName -IconLocationFile $IconLocationFile -AutoInstall $true
                Log-ScriptEvent -NewLog $Logfile -Value "Application $AppName created with Icon" -Component $scriptname -Severity 1
                c:
                return $true
            }
            else
            {
                New-CMApplication -Name $AppName -Publisher $Publisher -SoftwareVersion $Version -LocalizedApplicationName $AppName -AutoInstall $true
                Log-ScriptEvent -NewLog $Logfile -Value "Application $AppName created without Icon" -Component $scriptname -Severity 1
                c:
                return $true
            }
        }
        catch
        {
            $message = $_.Exception.Message
            Log-ScriptEvent -NewLog $Logfile -Value "Failed to create Application" -Component $scriptname -Severity 3
            Log-ScriptEvent -NewLog $Logfile -Value $message -Component $scriptname -Severity 3
            c:
            return $false
        }
    }
    else
    {
        c:
        return $true
    }
}
function Move-Object
{
    param([string]$AppName, [Switch]$IsApp)
    cd "$($SiteCode):"
    if($IsApp)
    {
        $Application = Get-CMApplication -Name $AppName
        $Path = ".\Application\" + $Folder
        try
        {
            Move-CMObject -FolderPath $Path -InputObject $Application
            Log-ScriptEvent -NewLog $Logfile -Value "Moved $AppName to $Folder successfully" -Component $scriptname -Severity 1
            c:
            return $true
        }
        catch
        {
            $message = $_.Exception.Message
            Log-ScriptEvent -NewLog $Logfile -Value "Failed to move $AppName to $Folder" -Component $scriptname -Severity 3
            Log-ScriptEvent -NewLog $Logfile -Value $message -Component $scriptname -Severity 3
            c:
            return $false
        }
    }
    else
    {
        $CollectionName1 = "$AppName$Suffix1"
        $Collection1 = Get-CMDeviceCollection -Name $CollectionName1
        $CollectionName2 = "$AppName$Suffix2"
        $Collection2 = Get-CMDeviceCollection -Name $CollectionName2
        $Path = ".\DeviceCollection\" + $Folder
        try
        {
            Move-CMObject -FolderPath $Path -InputObject $Collection1
            Log-ScriptEvent -NewLog $Logfile -Value "Moved $CollectionName1 to $Folder successfully" -Component $scriptname -Severity 1
            Move-CMObject -FolderPath $Path -InputObject $Collection2
            Log-ScriptEvent -NewLog $Logfile -Value "Moved $CollectionName2 to $Folder successfully" -Component $scriptname -Severity 1
            c:
            return $true
        }
        catch
        {
            $message = $_.Exception.Message
            Log-ScriptEvent -NewLog $Logfile -Value "Failed to move collections to $Folder" -Component $scriptname -Severity 3
            Log-ScriptEvent -NewLog $Logfile -Value $message -Component $scriptname -Severity 3
            c:
            return $false
        }
    }
}
function Create-DeploymentType
{
    param([string]$AppName)
    if(-not(check-DeploymentType -AppName $AppName))
    {
        cd "$($SiteCode):"
        $DeploymentTypePath = "$PathtoSWPool$AppName"
        Log-ScriptEvent -NewLog $Logfile -Value "Create deployment type for application $AppName" -Component $scriptname -Severity 1
        try
        {
            Add-CMDeploymentType -DeploymentTypeName "Install" -ApplicationName $AppName -ScriptInstaller -ContentLocation $DeploymentTypePath -AllowClientsToShareContentOnSameSubnet $true -InstallationProgram "serviceui.exe wrapper.exe wrapper_Inst.ini" -UninstallProgram "serviceui.exe wrapper.exe wrapper_Uninst.ini" -InstallationBehaviorType InstallForSystem -LogonRequirementType WhetherOrNotUserLoggedOn -InstallationProgramVisibility Normal -MaximumAllowedRunTimeMinutes 120 -EstimatedInstallationTimeMinutes 0 -RunInstallationProgramAs32BitProcessOn64BitClient $true -DetectDeploymentTypeByCustomScript -ScriptContent "echo $true" -ScriptType PowerShell 
            Set-CMDeploymentType -ApplicationName $AppName -DeploymentTypeName " - Script Installer" -NewDeploymentTypeName "Install"
            Log-ScriptEvent -NewLog $Logfile -Value "Deployment type for application $AppName created" -Component $scriptname -Severity 1
            Log-ScriptEvent -NewLog $Logfile -Value "Modify detection rule afterwards. Use following information for detection method" -Component $scriptname -Severity 2
            Log-ScriptEvent -NewLog $Logfile -Value "Hive: HKEY_LOCAL_MACHINE" -Component $scriptname -Severity 2
            $key = "Software\Altana\Packages\" + $AppName
            Log-ScriptEvent -NewLog $Logfile -Value "Key: $key" -Component $scriptname -Severity 2
            Log-ScriptEvent -NewLog $Logfile -Value "Value: Result" -Component $scriptname -Severity 2
            Log-ScriptEvent -NewLog $Logfile -Value 'Select "This registry setting must satisfy the following rule to indicate the presence of this application"' -Component $scriptname -Severity 2
            Log-ScriptEvent -NewLog $Logfile -Value "Data Type: String" -Component $scriptname -Severity 2
            Log-ScriptEvent -NewLog $Logfile -Value 'Select "This registry key is associated with a 32-bit application on 64-bit systems"' -Component $scriptname -Severity 2
            Log-ScriptEvent -NewLog $Logfile -Value "Operator: Equals" -Component $scriptname -Severity 2
            Log-ScriptEvent -NewLog $Logfile -Value "Value: SUCCESS" -Component $scriptname -Severity 2
            #Write-Host $finaloutput
            c:
            return $true
        }
        catch
        {
            $message = $_.Exception.Message
            Log-ScriptEvent -NewLog $Logfile -Value "Failed to create deployment type" -Component $scriptname -Severity 3
            Log-ScriptEvent -NewLog $Logfile -Value $message -Component $scriptname -Severity 3
            c:
            return $false
        }
    }
    else
    {
        c:
        return $true
    }  
}
function Create-Collection
{
    param([String]$AppName, [String]$CollDescription, [Switch]$IsAuth)
    if($IsAuth)
    {
        if(-not(check-collection -AppName $AppName -IsAuth))
        {
            cd "$($SiteCode):"
            $CollectionName = "$AppName$Suffix1"
            $CollComment = "$CollDescription" + "(disabled)"
            Log-ScriptEvent -NewLog $Logfile -Value "Start to create $CollectionName" -Component $scriptname -Severity 1
            try
            {
                New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName "All Systems" -Comment $CollComment
                Log-ScriptEvent -NewLog $Logfile -Value "$CollectionName created" -Component $scriptname -Severity 1
                c:
                return $true
            }
            catch
            {
                $message = $_.Exception.Message
                Log-ScriptEvent -NewLog $Logfile -Value "Failed to create Collection" -Component $scriptname -Severity 3
                Log-ScriptEvent -NewLog $Logfile -Value $message -Component $scriptname -Severity 3
                c:
                return $false
            }
        }
        else
        {
            c:
            return $true
        }
    }
    else
    {
        if(-not(check-collection -AppName $AppName))
        {
            cd "$($SiteCode):"
            $CollectionName = "$AppName$Suffix2"
            $ExcludeCollection = "$AppName$Suffix1"
            $CollectionMembershipQuery = $CollectionMembershipQuery.Replace("[PACKAGENAME]",$AppName)
            Log-ScriptEvent -NewLog $Logfile -Value "Start to create $CollectionName" -Component $scriptname -Severity 1
            try
            {
                $currentDate = Get-Date
                $Schedule = New-CMSchedule -RecurInterval Hours -RecurCount 10 -Start $currentDate
                New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName "All Systems" -RefreshType Periodic -RefreshSchedule $Schedule 
                Log-ScriptEvent -NewLog $Logfile -Value "$CollectionName created. Will add Query and Exclusion" -Component $scriptname -Severity 1
                Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $CollectionName -ExcludeCollectionName $ExcludeCollection
                Log-ScriptEvent -NewLog $Logfile -Value "Collection $ExcludeCollection excluded in collection $CollectionName." -Component $scriptname -Severity 1
                Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionName -RuleName $CollectionName -QueryExpression $CollectionMembershipQuery
                Log-ScriptEvent -NewLog $Logfile -Value "Added query membership rule $CollectionName to collection $CollectionName with following expression: $CollectionMembershipQuery" -Component $scriptname -Severity 1
                c:
                return $true
            }
            catch
            {
                $message = $_.Exception.Message
                Log-ScriptEvent -NewLog $Logfile -Value "Failed to create and configure Collection" -Component $scriptname -Severity 3
                Log-ScriptEvent -NewLog $Logfile -Value $message -Component $scriptname -Severity 3
                c:
                return $false
            }
        }
        else
        {
            c:
            return $true
        }
    }

}
function Create-Deployment
{
    param([String]$AppName, [Switch]$Install)
    if($Install)
    {
        if(-not(check-ApplicationDeployment -AppName $AppName -Install))
        {
            cd "$($SiteCode):"
            $CollectionName = "$AppName$Suffix1"
            Log-ScriptEvent -NewLog $Logfile -Value "Start to create deployment on $CollectionName for install application $AppName" -Component $scriptname -Severity 1
            try
            {
                Start-CMApplicationDeployment -Name $AppName -DeployAction Install -CollectionName $CollectionName -DeployPurpose Required
                Log-ScriptEvent -NewLog $Logfile -Value "Deployment on $CollectionName for install application $AppName created successfully" -Component $scriptname -Severity 1
                c:
                return $true
            }
            catch
            {
                $message = $_.Exception.Message
                Log-ScriptEvent -NewLog $Logfile -Value "Failed to create deployment on $CollectionName for install application $AppName" -Component $scriptname -Severity 3
                Log-ScriptEvent -NewLog $Logfile -Value $message -Component $scriptname -Severity 3
                c:
                return $false
            }
        }
        else
        {
            c:
            return $true
        }
    }
    else
    {
        if(-not(check-ApplicationDeployment -AppName $AppName))
        {
            cd "$($SiteCode):"
            $CollectionName = "$AppName$Suffix2"
            Log-ScriptEvent -NewLog $Logfile -Value "Start to create deployment on $CollectionName for uninstall application $AppName" -Component $scriptname -Severity 1
            try
            {
                Start-CMApplicationDeployment -Name $AppName -DeployAction Uninstall -CollectionName $CollectionName -DeployPurpose Required
                Log-ScriptEvent -NewLog $Logfile -Value "Deployment on $CollectionName for uninstall application $AppName created successfully" -Component $scriptname -Severity 1
                c:
                return $true
            }
            catch
            {
                $message = $_.Exception.Message
                Log-ScriptEvent -NewLog $Logfile -Value "Failed to create deployment on $CollectionName for uninstall application $AppName" -Component $scriptname -Severity 3
                Log-ScriptEvent -NewLog $Logfile -Value $message -Component $scriptname -Severity 3
                c:
                return $false
            }
        }
        else
        {
            c:
            return $true
        }
    }
    
}
function distribute-content
{
    param([String]$AppName)
    cd "$($SiteCode):"
    try
    {
        Log-ScriptEvent -NewLog $Logfile -Value "Distribute content of application $AppName to DP group $DistributionPointGroupName" -Component $scriptname -Severity 1
        Start-CMContentDistribution -ApplicationName $AppName -DistributionPointGroupName $DistributionPointGroupName
        Log-ScriptEvent -NewLog $Logfile -Value "Content of application $AppName distributed to DP group $DistributionPointGroupName sucessfully initialized" -Component $scriptname -Severity 1
        c:
        return $true
    }
    catch
    {
        $message = $_.Exception.Message
        Log-ScriptEvent -NewLog $Logfile -Value "Failed to distribute content of $AppName to $DistributionPointGroupName" -Component $scriptname -Severity 3
        Log-ScriptEvent -NewLog $Logfile -Value "Continue neverthless" -Component $scriptname -Severity 2
        Log-ScriptEvent -NewLog $Logfile -Value $message -Component $scriptname -Severity 3
        c:
        return $true
    }  
}

################################## MAIN MAIN MAIN MAIN MAIN MAIN MAIN MAIN MAIN MAIN ############################################################################



$application = select-Applications
$lic = [System.Windows.Forms.MessageBox]::Show("Does the application cost money?" , "LIC" , 4)
if($lic -eq "YES")
{
    $lic = "(lic)"
}
else
{
    $lic=""
}

if(create-application -AppName $application)
{
    if(Move-Object -AppName $application -IsApp)
    {
        if(Create-DeploymentType -AppName $application)
        {
            if(distribute-content -AppName $application)
            {
                if(Create-Collection -AppName $application -CollDescription $lic -IsAuth)
                {
                    if(Create-Collection -AppName $application)
                    {
                        if(Move-Object -AppName $application)
                        {
                            if(Create-Deployment -AppName $application -Install)
                            {
                                if(Create-Deployment -AppName $application)
                                {
                                    Log-ScriptEvent -NewLog $Logfile -Value "$application imported successfully" -Component $scriptname -Severity 1
                                    $key = "Software\Altana\Packages\" + $application
                                    $finaloutput = $finaloutput + "Please configure a detection method for the deployment type install of application $application!`n`n"
                                    $finaloutput = $finaloutput + "Key: $key" + "`n"
                                    $finaloutput = $finaloutput + "Value: Result" + "`n"
                                    $finaloutput = $finaloutput + 'Select "This registry setting must satisfy the following rule to indicate the presence of this application"' + "`n"
                                    $finaloutput = $finaloutput + "Data Type: String" + "`n"
                                    $finaloutput = $finaloutput + 'Select "This registry key is associated with a 32-bit application on 64-bit systems"' + "`n"
                                    $finaloutput = $finaloutput + "Operator: Equals" + "`n"
                                    $finaloutput = $finaloutput + "Value: SUCCESS" + "`n"
                                    [System.Windows.Forms.MessageBox]::Show($finaloutput , "Application Detection method" , 4)
                                    #$finaloutput
                                }
                                else
                                {
                                    Log-ScriptEvent -NewLog $Logfile -Value "The import of $application failed" -Component $scriptname -Severity 3
                                }
                            }
                            else
                            {
                                Log-ScriptEvent -NewLog $Logfile -Value "The import of $application failed" -Component $scriptname -Severity 3
                            }
                        }
                        else
                        {
                            Log-ScriptEvent -NewLog $Logfile -Value "The import of $application failed" -Component $scriptname -Severity 3
                        }
                    }
                    else
                    {
                        Log-ScriptEvent -NewLog $Logfile -Value "The import of $application failed" -Component $scriptname -Severity 3
                    }
                }
                else
                {
                    Log-ScriptEvent -NewLog $Logfile -Value "The import of $application failed" -Component $scriptname -Severity 3
                }
            }
            else
            {
                Log-ScriptEvent -NewLog $Logfile -Value "The import of $application failed" -Component $scriptname -Severity 3
            }
        }
        else
        {
            Log-ScriptEvent -NewLog $Logfile -Value "The import of $application failed" -Component $scriptname -Severity 3
        }
    }
    else
    {
        Log-ScriptEvent -NewLog $Logfile -Value "The import of $application failed" -Component $scriptname -Severity 3
    }
}
else
{
    Log-ScriptEvent -NewLog $Logfile -Value "The import of $application failed" -Component $scriptname -Severity 3
}

