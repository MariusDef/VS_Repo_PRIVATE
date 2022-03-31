$result = foreach ($gpo in (get-gpo -All | Where-Object {$_.DisplayName -like "WGGPO_MSFT*"} |select-object DisplayName)) {backup-gpo -name $gpo.displayname -path "c:\temp\WGGPO_new"}
$result | out-file c:\temp\WGGPO_new\BackupLog.log

foreach ($gpo in $result) {"import-gpo -backupgponame '$($gpo.Displayname)' -path C:\Temp\WGGPO_new -targetname '$($gpo.Displayname)' -createifneeded" |out-file c:\temp\WGGPO_new\import.ps1 -Append}