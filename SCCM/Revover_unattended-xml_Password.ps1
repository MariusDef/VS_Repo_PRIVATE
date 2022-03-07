$encryptedpwd = "QQBsAGYAaQBuAGcAMQAyADMAQQBkAG0AaQBuAGkAcwB0AHIAYQB0AG8AcgBQAGEAcwBzAHcAbwByAGQAAAA="
[System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($encryptedpwd))
