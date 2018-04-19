## Request New Certificate
New-ExchangeCertificate -GenerateRequest -Path c:\mail_nakamura_local.csr -KeySize 2048 -SubjectName "c=BR, s=PR, l=Curitiba, o=CA Nakamura Inc., ou=IT, cn=mail.nakamura.local" -DomainName autodiscover.nakamura.local, webmail.nakamura.local, LAB00-EX2K7.nakamura.local, LAB00-EX2K7 -PrivateKeyExportable $True

## Install New Certificate
Import-ExchangeCertificate -path C:\Certificate\mail.nakamura.local.pfx | Enable-ExchangeCertificate -Services IMAP, POP, UM, IIS, SMTP

## Enable Services if be installed
Enable-ExchangeCertificate -Services IMAP, POP, IIS, SMTP -thumbprint 9859BB512CC86D670B0480E5D60C342E329A2BDF