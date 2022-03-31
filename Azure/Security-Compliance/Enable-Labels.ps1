Install-Module AzureADPreview -scope currentuser
Import-Module AzureADPreview
Connect-AzureAD

$grpUnifiedSetting = (Get-AzureADDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ)
$template = Get-AzureADDirectorySettingTemplate -Id 62375ab9-6b52-47ed-826b-58e47e0e304b
$setting = $template.CreateDirectorySetting()

$Setting.Values

$Setting["EnableMIPLabels"] = "True"

$Setting.Values

Set-AzureADDirectorySetting -Id $grpUnifiedSetting.Id -DirectorySetting $setting


###################################################################


Install-Module ExchangeOnlineManagement -Scope CurrentUser
Import-Module ExchangeOnlineManagement

Connect-IPPSSession -UserPrincipalName md@morgana1.onmicrosoft.com

Execute-AzureAdLabelSync
