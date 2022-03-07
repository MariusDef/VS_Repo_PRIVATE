Param
(
    [Parameter(Mandatory = $true)]
    [string] $Computername,
    [Parameter(Mandatory = $true)]
    [string] $Domain,
    [Parameter(Mandatory = $true)]
    [string] $UserID,
    [Parameter(Mandatory = $true)]
    [string] $StartTime,
    [Parameter(Mandatory = $true)]
    [string] $EndTime
)


#$Computername = "De-TAB-9991"
#$UserID = "sccmtest1"
#$StartTime = "01.01.1911"
#$EndTime = "31.12.2020"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Replace = "False"

"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Values to write: '$($Domain);$($UserID);$($StartTime);$($EndTime)' - For Object: '$($Computername)'" | out-file $LogFile -Append


Function Check-Parameter
{
    if(-not($Computername -like "DE-[NTUW][BAK][KBS]-????"))
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Computername in wrong format. Quit Script" | out-file $LogFile -Append
        exit
    }
    else
    {
        if(-not($StartTime -like "??.??.????"))
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) StartTime format not correct. User 'DD.MM.YYYY'. Quit Script" | out-file $LogFile -Append
            exit
        }
        else
        {
            if(-not($EndTime -like "??.??.????"))
            {
                "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Endtime format not correct. User 'DD.MM.YYYY'. Quit Script" | out-file $LogFile -Append
                exit
            }
        }
    }
}

Check-Parameter

if (-not($((Get-ADComputer $Computername -Properties "extensionAttribute10").extensionAttribute10)))
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) extensioattribute10 is not in use. Use it first time with value: '$($Domain);$($UserID);$($StartTime);$($EndTime)'" | out-file $LogFile -Append
    Set-ADComputer $Computername -add @{extensionAttribute10="$($Domain);$($UserID);$($StartTime);$($EndTime)"}
}
else
{
    $CurrentValue = $((Get-ADComputer $Computername -Properties "extensionAttribute10").extensionAttribute10)
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) extensioattribute10 is already in use with vaule: '$CurrentValue'" | out-file $LogFile -Append    
    
    if ($CurrentValue -like "*#*")
    {
        $CurrentValue_Rows = $CurrentValue.Split("#")
        foreach ($CurrentValue_Row in $CurrentValue_Rows)
        {
            if ($CurrentValue_Row -like "*;*")
            {
                $CurrentValue_single = $CurrentValue_Row.split(";")
                if ($CurrentValue_single -eq $UserID)
                {
                    $Replace = "True"
                    $Rowtoreplace = $CurrentValue_Row
                }
            }
        }
    }
    elseif ($CurrentValue -like "*;*")
    {
        $CurrentValue_single = $CurrentValue.split(";")
        if ($CurrentValue_single -eq $UserID)
        {
            $Replace = "True"
            $Rowtoreplace = $CurrentValue
        }
    }
        
    if ($Replace -eq "True") 
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) User is already in String. Replace the values..." | out-file $LogFile -Append    
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) From: $($Rowtoreplace)" | out-file $LogFile -Append
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) To     : $($Domain);$($UserID);$($StartTime);$($EndTime)" | out-file $LogFile -Append
        $CurrentValue =  $CurrentValue.Replace($Rowtoreplace,"$($Domain);$($UserID);$($StartTime);$($EndTime)")
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Write to AD: $($CurrentValue)" | out-file $LogFile -Append    
        Set-ADComputer $Computername -replace @{extensionAttribute10=$CurrentValue}
    }
    else
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) User not in List. Add at the end" | out-file $LogFile -Append
        if ($CurrentValue.substring($CurrentValue.length -1, 1) -eq "#")
        {
            $CurrentValue = $CurrentValue + "$($Domain);$($UserID);$($StartTime);$($EndTime)"
        }
        else
        {
            $CurrentValue = $CurrentValue + "#$($Domain);$($UserID);$($StartTime);$($EndTime)"
        }
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) New Value: $($CurrentValue)" | out-file $LogFile -Append
        Set-ADComputer $Computername -replace @{extensionAttribute10=$CurrentValue}
    }
}
