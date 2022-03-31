$date = get-date -Day 10

(Get-ADComputer -Filter {LastLogonDate -ge $date} -Properties Name,LastLogonDate -SearchBase "OU=WI,OU=Tier2,OU=ESAE,DC=wirtgen-group,DC=local" | Get-Random -Count 100).name
write-host "-----------------"
(Get-ADComputer -Filter {LastLogonDate -ge $date} -Properties Name,LastLogonDate -SearchBase "OU=PAW,OU=ESAE,DC=wirtgen-group,DC=local" | Get-Random -Count 10).name
write-host "-----------------"
(Get-ADComputer -Filter {LastLogonDate -ge $date} -Properties Name,LastLogonDate -SearchBase "OU=Computer,OU=Wirtgen,DC=wirtgen-group,DC=local" | Get-Random -Count 10).name