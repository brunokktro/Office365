########################################################### 
# AUTOR  : Bruno Lopes - Axter Soluções em TI 
# DATA    : 24/09/2013
# 
# SOBRE : Este script realiza o "Move-Request" de mailbox
#         da nuvem para o ambiente On-Premises, a partir de 
#         um arquivo de entrada. 
########################################################### 

## Variáveis
$Users = Import-Csv C:\Scripts\user_migration.csv
$Pass = cat C:\Scripts\securestring.txt | convertto-securestring                                                            
$UserCredential = new-object -typename System.Management.Automation.PSCredential -argumentlist $_.Admin ,$Pass


## Testa conteúdo da variável
## Write-Host $Database

$Users | ForEach-Object -Verbose {
$Database = Get-Mailbox $_.UserName | ft Database
New-MoveRequest -OutBound -RemoteTargetDatabase $Database -RemoteHostName $_.HostName -RemoteCredential $UserCredential -TargetDeliveryDomain $_.Domain -Identity $_.Email
}
## Caso algum erro aconteça durante a migração, use o comando abaixo
## Set-UmMailbox $Email -ImListMigrationCompleted $false
