#################################################################################
#   Author: Marius Deffner, cellent GmbH, 17.09.2019
#
#   Possible Exit Codes:
#      0:  Successfull
#      7:  Any job in script failed
#
#
#################################################################################


Param
(
    [Parameter(Mandatory = $true)]
    [string] $UserDisplayname,
    [Parameter(Mandatory = $true)]
    [string] $SR_RecID
    
)

$StartTime = get-date -UFormat "%d.%m.%Y" #"01.01.2017"

$scriptname = $MyInvocation.MyCommand.Name
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$LogFile = ($MyInvocation.MyCommand.Definition + ".log")
$Result = 0

#Read-Host -Prompt "Enter PW for SQL HeatService User 'heatservice'" -AsSecureString | ConvertFrom-SecureString | Out-File ("e:\Heat_Automation\SQLHeatService_cred.txt")
$Username = "heatservice"
$SecurePassword = Get-Content ((split-path -parent $MyInvocation.MyCommand.Definition) + "\SQLHeatService_cred.txt") | ConvertTo-SecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
$UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

$datasource = "de-sql-202\heatprd"
$database = "heatsm"
$connectionstring = "server=$($datasource);database=$($database);User ID=$($Username);Password=$($UnsecurePassword);"


"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) -------------------Start Script ----------------------------------------------------------------------" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Parameter 'UserDisplayname': $($UserDisplayname)" | out-file $LogFile -Append
"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Parameter 'SR_RecID': $($SR_RecID)" | out-file $LogFile -Append
#"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Current used Username: $($Username)" | out-file $LogFile -Append
#"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Current used SecurePassword: $($SecurePassword)" | out-file $LogFile -Append
#"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Current used BSTR: $($BSTR)" | out-file $LogFile -Append
#"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Current used UnsecurePassword: $($UnsecurePassword)" | out-file $LogFile -Append
#"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Current used UnsecurePassword: $((split-path -parent $MyInvocation.MyCommand.Definition) + "\SQLHeatService_cred.txt")" | out-file $LogFile -Append



##############################################################################################
#   open SQL Connection
##############################################################################################
try
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Open SQL Connection..." | out-file $LogFile -Append
    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connectionstring
    $connection.Open()
}
catch
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: opening connection: '$($_.Exception.Message)'" | out-file $LogFile -Append
    $Result = 7
}


##############################################################################################
#   Define Functions
##############################################################################################

function Map-CI
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $CIRecID,
        [Parameter(Mandatory = $true)]
        [string] $SRRecID,
        [Parameter(Mandatory = $true)]
        [string] $AssetID,
        [Parameter(Mandatory = $true)]
        [string] $Model
    )

    $FreeGuid = Get-FreeGuid

    if ((Check-FusionAlreadyExists -CIRecID $CIRecID -SRRecID $SRRecID)-eq $false)
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       CI not mapped to SR. Map it..." | out-file $LogFile -Append
        #write-host $CIRecID
        $QueryUpdateFusion = "insert into FusionLink (RecID, 
						SourceID, 
						SourceLoc, 
						SourceName, 
						sourceBase, 
						TargetID, 
						TargetLoc, 
						TargetName, 
						TargetBase, 
						RelationshipName)
			Values ('$($FreeGuid)',
					'$($SRRecID)',
					'I',
					'ServiceReq',
					'ServiceReq',
					'$($CIRecID)',
					'I',
					'CI.License',
					'CI',
					'ServiceReqAssocCI')"

        try
        {
            $connFusion = $connection.CreateCommand()
            $connFusion.CommandText = $QueryUpdateFusion
            $connFusion.ExecuteNonQuery()
        }
        catch 
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Mapping CI to SR: '$($_.Exception.Message)'" | out-file $LogFile -Append
            $Result = 7
        }
        
    }
    else
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       CI already mapped to SR. Check next, if available..." | out-file $LogFile -Append
    }
}


Function Check-FusionAlreadyExists 
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $CIRecID,
        [Parameter(Mandatory = $true)]
        [string] $SRRecID    
    )

    $query = "select recid from FusionLink where targetname like '%license%' and targetid = '$($CIRecID)' and sourceid = '$($SRRecID)'"
    $conn = $connection.CreateCommand()
    $conn.CommandText = $query

    $Reader = $conn.ExecuteReader()
    $Temp = $Reader.HasRows
    $reader.Close()
    #Write-Host $Temp
    
    return($Temp)
}


Function Get-FreeGuid 
{
    do
    {
        [bool]$Temp = $false
        $GUID = (1..32 | %{ '{0:X}' -f (Get-Random -Max 16) }) -join ''
    
        $query = "select * from FusionLink where recid = '$($GUID)'"
        $conn = $connection.CreateCommand()
        $conn.CommandText = $query

        $Reader = $conn.ExecuteReader()
        $Temp = $Reader.HasRows
        $reader.Close()

    }while($Temp -eq $true)

    return($GUID)
}



##############################################################################################
#   Get all License CIs mapped to the User
##############################################################################################
if($connection.State -eq 'Open')
{
    $query = “SELECT dbo.Employee.DisplayName, dbo.CI.Model, dbo.CI.AssetID_ITSM, dbo.CI.RecId FROM dbo.Employee INNER JOIN dbo.FusionLink ON dbo.Employee.RecId = dbo.FusionLink.SourceID INNER JOIN dbo.CI ON dbo.FusionLink.TargetID = dbo.CI.RecId WHERE (dbo.FusionLink.SourceName = N'Employee') AND (dbo.Employee.DisplayName = '$($UserDisplayname)')  and (dbo.ci.Status = 'active') and (dbo.ci.CIType = 'License')”
   
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString()) Get all mapped License CIs for User: '$($UserDisplayname)'..." | out-file $LogFile -Append
    $conn = $connection.CreateCommand()
    $conn.CommandText = $query

    $Reader = $conn.ExecuteReader()
    $Table = New-Object System.Data.DataTable
    $Table.Load($Reader)
    if ($Table.Rows.Count -eq 0)
    {
        "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       No License CI for user available" | out-file $LogFile -Append
    }
    else
    {
        for ($i=0;$i -lt $Table.Rows.Count;$i++)
        {
            "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       AssetID: '$($Table.Rows[$i][2])' # Model: '$($Table.Rows[$i][1])'" | out-file $LogFile -Append
            Map-CI -CIRecID $Table.Rows[$i][3] -SRRecID $SR_RecID -AssetID $($Table.Rows[$i][2]) -Model $($Table.Rows[$i][1])
        }
    }
    
}



##############################################################################################
#   close SQL Connection
##############################################################################################
try
{
    if($connection.State -eq 'Open')
    {
        $connection.Close()
    }
}
catch
{
    "$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())       Error: Connection cannot beeing closed: '$($_.Exception.Message)'" | out-file $LogFile -Append
    $Result = 7
}




"$((Get-Date).ToShortDateString() + " " + (Get-Date).ToShortTimeString())  Result: $($Result)"  | out-file $LogFile -Append
Write-Output $Result
exit $Result