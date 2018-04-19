
Get-MsolAccountSku | Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPACK"} | ForEach-Object {$_.ServiceStatus}