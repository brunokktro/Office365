## Instalação do RSAT para comunicação com AD

Install-WindowsFeature RSAT-ADDS
Install-WindowsFeature RSAT-Clustering-CmdInterface

## Apenas em casos de florestas legadas de Exchange já implantadas

Setup /PrepareLegacyExchangePermissions


## Novas florestas a partir daqui

Setup /PrepareSchema /IAcceptExchangeServerLicenseTerms

Setup /PrepareAD /OrganizationName:FIORELLA /IAcceptExchangeServerLicenseTerms

Setup /PrepareAllDomains /IAcceptExchangeServerLicenseTerms

Setup /mode:Install /r:CA,MB /IAcceptExchangeServerLicenseTerms /MdbName:MB01 /DbFilePath:D:\ExchData\MBData\MB01.edb /LogFolderPath:D:\ExchData\MBLogs


