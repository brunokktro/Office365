########################################################### 
# AUTOR  : Bruno Lopes - Axter Soluções em TI 
# DATA    : 24-09-2013
# 
# SOBRE : Este script realiza o "Move-Request" de mailbox
#         da nuvem para o ambiente On-Premises. 
########################################################### 

## Variáveis
$User = $(Read-Host -prompt "Nome Completo")
$Database = Get-Mailbox $User | ft Database
$HostName = $(Read-Host -prompt "URL de acesso ao OWA")
$Cred = get-credential
$Domain = $(Read-Host -prompt "FQDN do Domínio")
$Email = $(Read-Host -prompt "Email do usuário a ser migrado")

## Testa conteúdo da variável
## Write-Host $Database

New-MoveRequest -OutBound -RemoteTargetDatabase $Database -RemoteHostName $HostName -RemoteCredential $Cred -TargetDeliveryDomain $Domain -Identity $Email

## Caso algum erro aconteça durante a migração, use o comando abaixo
## Set-UmMailbox $Email -ImListMigrationCompleted $false
