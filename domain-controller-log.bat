Get-ADUser -Filter * -Properties Name,whenCreated,whenChanged,LastlogonDate,SamAccountName,Enabled -SearchBase "OU=IIM3-Permision,DC=iiml,DC=local" |
    Select-Object Name, whenCreated, whenChanged, LastlogonDate, SamAccountName,
        @{Name="AccountStatus"; Expression={if($_.Enabled -eq $true) {"Active"} else {"Disabled"}}} |
    Export-Csv IIM3-Permision-April.csv -NoTypeInformation