# Compilação e homologação manual — fase 2

**Situação observada:** BLOQUEADO para importação, compilação e execução VBA. Excel 16.0, build 20527, x64 (PE 0x8664). O Office devolveu “O acesso de programação ao projeto do Visual Basic não é confiável”, HRESULT 0x800A03EC. Foram importados zero módulos. [Evidência original](./compilacao-probe.json).

Não modificar AccessVBOM, Trust Center, ExecutionPolicy, Registro, locais confiáveis ou política de macros. Os scripts PS1 não foram executados por bypass: os testes COM foram comandos PowerShell inline autorizados. Se a política impedir executar um arquivo de script, registrar BLOQUEADO e usar o roteiro manual permitido; não carregar o arquivo por Invoke-Expression ou ScriptBlock para contornar a política.

## Preparação permitida

1. Criar uma pasta nova sob TEMP, fora do OneDrive, e uma pasta de trabalho **vazia**, sem copiar dados do XLSM original. Salvar com prefixo Fixture_Auditoria_ e extensão XLSM. Nenhuma cópia que contenha estoque real deve ser usada para movimentações.
2. Reproduzir as 17 abas e seus CodeNames usando [fixture-esquema.json](./scripts/fixture-esquema.json). As abas são reais do projeto; não criar uma aba substituta para “Cancelado, Estorno e Devolução”. Ela não existe na base.
3. No editor VBA, a propriedade (Name) de cada planilha é o CodeName; Name é o nome visível da aba. O preparador renomeia CodeNames em duas etapas para evitar colisão entre nomes já gerados pelo Excel.
4. Importar os BAS e a classe de [MDL](./desenvolvimento/MDL/) e os cinco pares FRM/FRX de [Form](./desenvolvimento/Form/). Não importar o XLSM intacto de desenvolvimento como se contivesse as alterações. Manter atributos e eventos. Quando necessário, gerar cópias temporárias Windows-1252 para o importador do VBE, sem mudar a codificação dos fontes de referência; o preparador faz staging com falha explícita em conversão não representável.
5. Código de documento não é classe comum: colar somente o código sem linhas Attribute no módulo de cada planilha correspondente e no módulo EstaPastaDeTrabalho, conforme o esquema. Há um evento existente em Planilha15; conservá-lo, sem criar CommandButton1.
6. Marcar Menu Principal!Z1 com AUDITORIA_VBA_FICTICIO. Importar também AuditoriaListaFake.cls, modTestesAuditoria.bas e modTestesFase2.bas de [scripts](./scripts/). A classe fake testa pesquisa em memória; não substitui controles reais. Inventário esperado, se todas as dependências forem importáveis: 55 componentes (52 de desenvolvimento e três de teste). O XLSM original continua com 50.

O [preparador](./scripts/Preparar-TestesVBA.ps1) foi revisado para BAS/CLS/FRM, pares FRX, CodeNames, documentos e referências, mas **esse caminho revisado não foi executado com VBProject liberado**. Ele para no bloqueio, não executa macros e não atesta compilação. O parâmetro legado RunTests agora recusa execução automática. Não repetir tentativas de VBProject enquanto a mesma restrição permanecer.

## Dependências a conferir no VBE

Em Ferramentas > Referências, registrar nomes, versões, caminhos e qualquer AUSENTE/MISSING. O [esquema](./scripts/fixture-esquema.json) conserva os GUIDs declarados de OLE, Office, Forms, Internet Controls, Common Controls, Search, MSXML 6 e Selenium. Caminhos de caches EXD foram excluídos da tentativa de referência, pois não são instaladores de bibliotecas.

A existência de uma declaração no arquivo não comprova registro nem compatibilidade x64. MSXML 6 e VBScript.RegExp funcionaram nos testes COM; UserForms, MSCOMCTL, SHDocVw, Selenium e controles ActiveX **não foram importados/instanciados nesta fase**. Não instalar, registrar OCX ou remover uma dependência para fazer o projeto compilar sem autorização e análise do uso. Se houver incompatibilidade, registrar o erro exato e parar a homologação dessa funcionalidade.

No designer, confirmar os controles reais, RowSource/ControlSource e propriedades dos cinco formulários. Preencher os intervalos efetivamente apontados pelos RowSources com listas fictícias; não presumir que uma pasta vazia já reproduz essas dependências. Conferir nomes definidos, botões, OnAction, Ribbon e fórmulas UDF. O preparador atual **não reconstrói automaticamente RowSources, nomes definidos, fórmulas originais, vínculos Ribbon nem objetos de planilha**. A fixture só está completa para integração depois dessa verificação.

## Compilação real e registro

1. No VBE, escolher Depurar > Compilar VBAProject.
2. Para cada erro, registrar componente, procedimento, linha destacada e mensagem original; corrigir somente na cópia respaldada. Repetir até o comando concluir sem erro.
3. Registrar versão/arquitetura do Office, lista resolvida de referências, data, nome e SHA-256 da fixture e evidência visual ou registro do operador. Não usar ausência de Problems no VS Code como evidência de compilação.
4. Se o editor ou a política impedir a etapa, manter BLOQUEADO. Não marcar PASS nem executar eventos fiscais para “ver se funciona”.

## Testes VBA preparados

Depois de compilação comprovada e fixture conferida, executar na Janela Imediata:

    ? Auditoria_ExecutarTestesSeguros()
    ? AuditoriaFase2_Executar()

A primeira suíte limpa e usa pequenos intervalos fictícios de Entradas/Entradas Filial. A segunda testa diretamente auxiliares de recuperação e formatação condicional, inclui falhas por etapa, nova tentativa, chave, texto começando com = e recusa de recuperação depois das fronteiras Delete/Save. Ela não chama eventos de gravar dos formulários. Ambas foram apenas revisadas/preparadas; **nenhuma foi executada nesta fase**.

Antes da integração, reproduzir a [matriz de regressão](../../TESTES_VBA.md) nos eventos reais dos formulários, com listas, fórmulas sentinela e XML sintético compatíveis com as regras existentes. Não reutilizar uma fixture após falha de Delete ou Save sem conciliá-la ou recriá-la. Testar ainda falhas na proteção do destino, fórmula matricial/derramada e erro de salvamento. Não forçar a falha em arquivo real.

Para Ribbon, abrir uma segunda pasta fictícia com alterações pendentes, deixá-la ativa e chamar o callback existente com controle Nothing, se a assinatura e o ambiente permitirem. Conferir em disco que somente a fixture proprietária foi salva. O teste COM equivalente já passou; o callback VBA continua pendente.

Para backup, chamar as rotinas somente na fixture, escolher destino novo sob TEMP e verificar SHA-256 dos arquivos fechados, dados, fórmulas, formatos e restauração em outra instância criada para o teste. Excel não permite dois arquivos homônimos na mesma instância. Não executar limpeza automática de versões.

## Critérios antes de homologação

Resolver a decisão de baixa A08 e a conciliação A09, inclusive falhas depois da exclusão/salvamento, edição anual e exclusão em duas abas. Definir a evidência necessária para Cancelado com o responsável. Não basta obter compilação sem erros.

A recuperação implementada é em memória e limitada aos destinos capturados antes de Delete/Save. Não restaura interface, referência estrutural alterada, nomes, objetos, configuração de filtros ou arquivo já salvo, e não sobrevive a encerramento inesperado. Mudanças de estoque e de critérios fiscais continuam condicionadas à decisão explícita do responsável. Nenhuma versão foi promovida.
