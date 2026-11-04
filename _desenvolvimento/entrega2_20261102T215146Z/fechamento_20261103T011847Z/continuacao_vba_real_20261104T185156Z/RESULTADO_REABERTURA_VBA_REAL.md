# Entrega 2 — reabertura real e regressão QueryClose

Etapa concluída em04/11/2026, entre22:04 e22:15UTC. Ponytail5.3.3Full. Sistema NÃO HOMOLOGADO; Entrega2 ainda em fechamento.

## Ponto recuperado

O trabalho anterior está preservado. Git atual main/7c46ea6, árvore limpa no início desta retomada; nenhum commit executado pelo agente. AGENTS vigente relido. A preparação18:51UTC permanece em [preparacao.json](./preparacao.json), com checkpoint privado e hashes. A janela compartilhada e o problema de foco da tentativa anterior não foram tratados como abertura exclusiva aprovada. Nesta retomada não havia Excel operacional aberto; foi criada a instância exclusiva PID7644 e aberta somente a candidata TEMP pelo seletor do Excel. O atalho /x preparado não precisou ser acionado.

## Evidência funcional nova

Antes de qualquer carga manual, a janela imediata retornou o caminho completo correto, C1_UltimoResultado=CACHE_AUDITADO, Saved=False, Workbooks.Count=1, EnableEvents=True. [Registro inicial](./reabertura_vba_real.json). Nenhuma configuração de segurança foi alterada e nenhuma ação de habilitação de macros foi executada pelo agente.

C1_Diagnostico real antes/depois: AUDIT=OK;PROJECTION=VERIFIED;REV=26;MOV=26;RECEIPTS=26;ACTIVE=7;PENDING=0. Não foi necessário CalculateFull nem reconstrução manual nesta sessão. Isso não apaga a tentativa FORMULA_VALUE anterior nem identifica sua causa fina.

Cad_NFE.Show e Cad_NFE_Saida.Show executados na janela imediata; cliques nos X das barras de título acionaram UserForm_QueryClose real. Ambos mostraram apenas Por favor, use o botão Fechar em Ações. Após OK, o botão existente CommandButtonFechar foi acionado por interface. Os dois retornaram normalmente, sem1004; Visible=False conferido. Nenhum campo fiscal foi preenchido ou gravador executado.

A planilha ativa continuava Entradas Filial protegida. U1 permaneceu vazia e ProtectContents=True. EnableEvents, ScreenUpdating e DisplayAlerts=True; Calculation=-4105 antes/depois. A condição que provocava1004 foi, portanto, exercitada após reabrir a revisão corrigida. E2-FECH-07 está resolvido nesse escopo dos dois formulários; não implica validação dos cinco formulários ou de todos os fluxos.

A interface exibiu as17planilhas esperadas e os5UserForms existentes. Depurar > Compilar VBAProject já estava desabilitado no projeto carregado. Esse estado foi observado, mas uma nova invocação explícita de compilação NÃO foi executada. Não alteramos fontes para forçar a habilitação. A compilação real anterior permanece como referência; esta sessão prova a execução das rotinas da revisão42b…, não uma nova edição/compilação.

## Resultados e preservação

[Matriz estruturada](./matriz_vba_real.json): seis verificações dirigidas PASS, com modalidades separadas. As evidências da interface foram transcritas do AX/screenshot CUA; não são logs automáticos da suíte. Compilação explícita nova NÃO EXECUTADA. Regressão dos gravadores Venda/Transferência/Devolução e conflito ACE NÃO EXECUTADA nesta etapa.

[Comparação das17tabelas](./comparacao_banco.json): zero diferenças frente ao snapshot anterior. Banco persistido [capturado](./banco_apos_vba_real.json); WriterLease FREE, 26movimentos/26recibos/66projeções, Pending0. [Três hashes](./integridade_apos_vba_real.json) inalterados:

|Artefato|SHA-256|
|---|---|
|XLSM Final|42b34e7a2632bb4875525b5b1d5de798274bbeda3402153a79bad4e8d0ba1853|
|ACCDB Final|9fc0bfe8001120d2804b8f56836949249606e2030a6af431a71b1757c80267dc|
|Registro de autoridade|b9e0c8aceb4474bb6b74ba727c7064982cfa9371e060b901bdf4b2de61632e2f|

Saved=False foi observado em memória após inicialização; não houve Save nesta etapa e o hash do disco permaneceu igual. Não atribuímos Saved=True nem mascaramos a alteração. Não é uma gravação inesperada em disco.

## Próxima atividade

A candidata permanece aberta, sem UserForm visível e com VBE pronto, na instância exclusiva identificada. Antes do próximo teste, conferir novamente caminho/PID/versão; não usar coordenadas antigas. Próximo defeito: E2-FECH-06, capturando a origem do erro no evento real e a cadeia ADO antes da limpeza. Os seis PASS COM/ADO anteriores não encerram esse defeito.

Permanecem pendentes ENTRY/rascunho órfão, dois gravadores Excel concorrentes, restauração quando registro/origem estão indisponíveis e exceções operacionais sem regra aprovada. Os sete FAIL históricos foram preservados. Entrega3 não liberada.

## Rastreabilidade e reversão

Sem alterações em VBA, fórmulas persistidas, esquema ou fatos. Nenhum XLSM novo foi produzido nesta etapa. Checkpoint privado185156Z preservado; documentação antes dos acréscimos foi copiada pelo script oficial para20261104T221410Z-346758d9. Não substituir a autoridade26 por cópia de banco antigo. Nenhum arquivo operacional aberto/modificado nesta retomada; nenhum commit/push/promoção.
