Abaixo, vamos citar o nome de cada ferramenta, seu descritivo, e o link de referência para download da mesma.

    ExPerfWiz – script baseado em PowerShell com a função de automatizar a coleta de informações de contadores do PerfMon (Performance Monitor) do Windows Server. Utilizado em situações de troubleshooting de performance e resizing do ambiente de Exchange Server. Vários parâmetros podem ser utilizados para adequar o script, como tempo de duração da execução, intervalos, quantidade de execuções, servidores alvo da coleta, etc. Link: https://github.com/Microsoft/experfwiz

 

    Performance Analysis of Logs (PAL) – ferramenta baseada em GUI (Graphic User Interface) criada para facilitar a leitura e diagnóstico de contadores de performance. Em conjunto com o ExPerfWiz, faz todo o “trabalho duro” de coleta dos contadores e análise, gerando gráficos e tabelas com as informações necessárias para as análises a serem executadas no ambiente. Interessante que o PAL não foi criado exclusivamente para o Exchange Server, mas funciona com IIS, SQL Server, BizTalk, AD entre outros. Link: https://github.com/clinthuffman/PAL

 

    Log Parser Studio – ferramenta também baseada em GUI, criada para auxiliar na leitura e criação de relatórios de logs provenientes do IIS, EventViewer entre outros. Muito utilizada para fazer queries em logs de IIS no Exchange Server, a fim de auxiliar na detecção de informações de conectividade de ActiveSync e OWA. A própria ferramenta já possui uma série de templates pré-criados para estes. Link: https://gallery.technet.microsoft.com/Log-Parser-Studio-cd458765

 

    Exchange VSSTester – script criado para troubleshooting em ambientes de backup do Exchange Server. A idéia é habilitar os logs disponíveis no Event Log Level do Exchange (Event Viewer) e também dos writers do Exchange no Volume Shadow Services. Uma função muito bacana desse script é usar o Disk Shadow, uma opção interessante pra quem precisa liberar espaço em disco utilizado pelos logs, mas não consegue efetuar o backup ou adicionar mais espaço em disco. Essa opção força o purge dos logs das databases. Mais informações sobre nesse artigo. Link: https://github.com/Microsoft/VSSTESTER

 

    Exchange Performance Health Checker – script criado para validar várias configurações documentadas no TechNet pela Microsoft como recomendadas para ambiente de Exchange 2013. São analisadas diversas informações do ambiente, como versão do Exchange, estrutura de hardware do servidor, hypervisor, PageFile, configuração de energia, versão do .NET, NICs, roles do Exchange instaladas e configuração de alta disponibilidade de databases. Link: https://gallery.technet.microsoft.com/Exchange-2013-Performance-23bcca58

 

    Exchange Log Collection – script baseado em PowerShell, criado para unificar toda a coleta de logs do ambiente de Exchange Server em um mesmo arquivo compactado (.zip). Com sua execução, são coletados logs do ambiente de Exchange Server dos mais diversos tipos, como logs de IIS, RpcHTTP, EWSLogs, EASLogs, AutoDiscoverLogs, OWALogs, ADDriverLogs, ClusterLogs entre outros. Link: https://gallery.technet.microsoft.com/Exchange-Log-Collection-8cd2019f

 

    Exchange Server Role Calculator – template do Excel desenvolvido pelo time da MS com o intuito de auxiliar na definição do sizing do ambiente de Exchange Server a ser implantado, tendo em vista informações de inputs, como versões anteriores que já existem, quantidade de mensagens trocadas na organização por dia, por hora, quantidade de usuários, quantidade de dispositivos mobile, performance de discos entre outros. Link: https://gallery.technet.microsoft.com/Exchange-2013-Server-Role-f8a61780

 

    MFCMAPI (Microsoft Foundation Classes MAPI) – Utilitário criado para extrair o máximo das funcionalidades de conexão com o protocolo MAPI, principal protocolo de comunicação de clientes Outlook com os servidores Exchange. Com o MFCMAPI, você consegue se conectar ao core de uma mailbox de usuário e analisar toda sua estrutura interna, como as classes, os atributos e os parâmetros ativos ou não. Em um cenários de troubleshooting, é ideal para remover itens corrompidos ou desativar funções que não respondem aos comandos originais. Este protocolo possui várias características essenciais e é muito poderoso na utilização do Exchange Server, pois é robusto e totalmente voltado ao ambiente de mensageria. Link: https://github.com/stephenegriffin/mfcmapi

 

    Generate Message Profile – script criado para simular valores de conectividade de mailboxes e de troca de mensagens na organização, a fim de seja possível a utilização destes dados baseados em informações bem próximas da realidade do ambiente em questão, para usar em conjunto com o ExRoleCalc e com o CPU Sizing Checker. Link: https://gallery.technet.microsoft.com/Generate-Message-Profile-7d0b1ef4

 

    Exchange CPU Sizing Checker – script baseado em PowerShell, usado para calcular a performance média em que as unidades de processamento dos servidores poderão alcançar, de acordo com o que foi levantado anteriormente com o Exchange Server Role Calc e com o Message Profile. Link: https://gallery.technet.microsoft.com/Exchange-2013-CPU-Sizing-06451c99

 

    Exchange Processor Query Tool – template do Excel criado para obter os valores de taxa de capacidade dos Processors utilizados nos servidores que irão abrigar o ambiente de Exchange Server. Estes valores de referência estão documentados pela MS no TechNet e fazem referência também ao SPEC (Standard Performance Evaluation Corporation), que é uma forma de se realizar benchmark para CPUs. Link: https://gallery.technet.microsoft.com/office/Exchange-Processor-Query-b06748a5

 

    Exchange Server User Monitor – Ferramenta criada pelo time de Produto para conceder ao administrador uma visualização capaz de avaliar a usabilidade dos recursos por cada usuário de forma individual e a sua experiência no Exchange, entendo a utilização de cada serviço pelo mesmo, baseado em coletas de dados em tempo real, como por exemplo, uso de rede, uso de CPU, conexões Outlook, ActiveSync, MAPI, RPC, entre outros. Link: https://www.microsoft.com/en-us/download/details.aspx?id=51101
