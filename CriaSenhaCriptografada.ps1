########################################################### 
# AUTOR  : Bruno Lopes - Axter Soluções em TI 
# DATA    : 24/09/2013
# 
# SOBRE : Este script cria um arquivo criptografado, com
#         a senha a ser digitada no prompt mostrado. 
########################################################### 

Read-Host -AsSecureString “Digite sua senha do Office 365” | ConvertFrom-SecureString | Out-File C:\Scripts\securestring.pwd

