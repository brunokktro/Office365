
## Passos para remoção do Hybrid Configuration

Get-OrganizationRelationship | Format-List

## Exchange Online
Remove-OrganizationRelationship -Identity "O365 to On-premises - 2faeb11c-2879-4d97-b997-985196848748"

## Exchange OnPremises
Remove-OrganizationRelationship -Identity "On-premises to O365 - 2faeb11c-2879-4d97-b997-985196848748"

## Exchange OnPremises e Online
Get-FederationTrust

## Exchange OnPremises e Online
Get-RemoteDomain

## Exchange OnPremises
Remove-HybridConfiguration

## Disable DirSync in Ofice 365
Set-MsolDirSyncEnabled -EnableDirsync $false

## Convert Domain (If Use ADFS)
Convert-MsolDomainToStandard -DomainName bragantina.com.br -SkipUserConversion $false -passwordfile c:\password.txt