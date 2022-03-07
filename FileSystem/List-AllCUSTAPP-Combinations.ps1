$CompCodes = @(010,300,310,311,312,319,320,321,325,326,330,332,333,335,336,340,342,343,344,345,347,348,354,356,357,360,362,363,368,370,371,372,377,380,382,383,384,397,400,420,421,424,425,426,427,428,429,431,441,509,998,999)
$ComMods = @("NTB", "DSK", "TAB")
$ClientFuncs = @("Technique", "Laboratory", "Office", "Home-Office", "Training")
$ClientArchs = @("32", "64", "32 Native", "64 Native", "W10", "W10 Native")
[int]$i = 0

$Logfile = "f:\Scripts\PROD\List-AllCUSTAPP-Combinations.log"

foreach ($CompCode in $CompCodes)
{
    foreach ($ComMod in $ComMods)
    {
        foreach ($ClientFunc in $ClientFuncs)
        {
            foreach ($ClientArch in $ClientArchs)
            {
                "$($CompCode)_$($ComMod)_$($ClientFunc)_$($ClientArch)" | Out-File $LogFile -Append
                $i += 1
            }
        }

    }
    
    
    #write-host $CompCode
}
"############################################" | Out-File $LogFile -Append
"Summe: $i"  | Out-File $LogFile -Append