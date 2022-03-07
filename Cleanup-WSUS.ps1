#Region VARIABLES   

# WSUS Connection Parameters: 
[String]$updateServer = "dehnm-scc-pp101"
[Boolean]$useSecureConnection = $False 
[Int32]$portNumber = 8530


# Initialize Logfile

$LogFile = "f:\scripts\prod\Cleanup-WSUS.log"
[DateTime]$RootDateTime = Get-Date


if ((Get-Item $LogFile).Length -gt 5MB)
{
    Copy-Item $LogFile "$LogFile.bkp"
    Remove-Item $LogFile -Force
}

" " |Out-File $LogFile -Append
" " |Out-File $LogFile -Append
"##########################################################################" |Out-File $LogFile -Append

# Cleanup Parameters: 
# Decline updates that have not been approved for 30 days or more, 
 # are not currently needed by any clients, and are superseded by an aproved update. 

[Boolean]$supersededUpdates = $True 
# Decline updates that aren't approved and have been expired my Microsoft. 
[Boolean]$expiredUpdates = $True 
# Delete updates that are expired and have not been approved for 30 days or more. 
[Boolean]$obsoleteUpdates = $True 
# Delete older update revisions that have not been approved for 30 days or more. 
[Boolean]$compressUpdates = $True 
# Delete computers that have not contacted the server in 30 days or more. 
[Boolean]$obsoleteComputers = $False
# Delete update files that aren't needed by updates or downstream servers. 
[Boolean]$unneededContentFiles = $True   

 #EndRegion VARIABLES  

#Region SCRIPT   
Try
{
    # Load .NET assembly 
    [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")   

    # Connect to WSUS Server 
    $Wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($updateServer,$useSecureConnection,$portNumber)   
    "$(get-date) WSUS:" | Out-File $LogFile -Append
    $Wsus | Out-File $LogFile -Append
    " " | Out-File $LogFile -Append

    # Perform Cleanup 
    $CleanupManager = $Wsus.GetCleanupManager() 
    $CleanupScope = New-Object Microsoft.UpdateServices.Administration.CleanupScope ($supersededUpdates,$expiredUpdates,$obsoleteUpdates,$compressUpdates,$obsoleteComputers,$unneededContentFiles) 
    "$(get-date) CleanUp Scope:" | Out-File $LogFile -Append
    $CleanupScope | Out-File $LogFile -Append
    " " | Out-File $LogFile -Append

    $Result = $CleanupManager.PerformCleanup($CleanupScope)
    "$(get-date) CleanUp Result:" | Out-File $LogFile -Append
    $Result | Out-File $LogFile -Append
    " " | Out-File $LogFile -Append

}
catch
{
    $message = $_.Exception.Message
    "$(get-date) Error: $message" | Out-File $LogFile -Append
}
   
#EndRegion SCRIPT