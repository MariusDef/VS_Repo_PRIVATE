$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment








foreach ($Drive in (Get-PSDrive -PSProvider FileSystem))
{
    if (Test-Path -Path "$($Drive.Root)sms")
    {
        if (Test-Path -Path "$($Drive.Root)sms\pkg")
        {
           $tsenv.Value("USEMEDIADIRECTORY") = "$($Drive.Root)sms\pkg" 
           #write-host "$($Drive.Root)sms\pkg" 
        }
    }
    else
    {
        if (Test-Path -Path "$($Drive.Root)_SMSTaskSequence")
        {
            if (Test-Path -Path "$($Drive.Root)_SMSTaskSequence\Packages")
            {
               $tsenv.Value("USEMEDIADIRECTORY") = "$($Drive.Root)_SMSTaskSequence\Packages" 
               #write-host "$($Drive.Root)_SMSTaskSequence\Packages" 
            }
        }
    }
}
