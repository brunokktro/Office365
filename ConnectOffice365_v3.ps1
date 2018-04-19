########################################################### 
# AUTOR  : Bruno Lopes - Axter Soluções em TI 
# DATA    : 31/05/2014
# 
# SOBRE : Este script realiza a conexão do usuário com o
#         ambiente de execução de script na nuvem, tanto
#         no tenant do Office 365 quanto do Exchange Online.
########################################################### 

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

Import-Module MSOnline
Connect-MsolService -Credential $UserCredential

#Import-Module LyncOnlineConnector
#$Session2 = New-CsOnlineSession -Credential $UserCredential
#Import-PSSession $Session2


#Close the session
## Remove-PSSession $Session
## Get-PSSession | Remove-PSSession