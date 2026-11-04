# Retomada da Entrega2 — QueryClose e reabertura

Estado em04/11/2026: candidata compilada, corrigida e preservada; Entrega2 ainda em fechamento, sistema NAO HOMOLOGADO. Nenhum arquivo operacional foi modificado.

## Base e preservacao

A referencia R26 anterior foi reaberta com habilitacao manual de macros pelo proprietario. Antes de carregar cache manualmente, C1_UltimoResultado=CACHE_AUDITADO; diagnostico real AUDIT=OK;PROJECTION=VERIFIED;REV=26;MOV=26;RECEIPTS=26;ACTIVE=7;PENDING=0. O XLSM correspondia a b379f42c739100086e1010f776777a90a9acb65d3ef94d8227dcbd1dd45e6e25.

Checkpoint anterior a correcao: C:/Users/MARCOSALEXANDREPERDO/AppData/Local/CodexBackupsPrivados/VBA-Excel/checkpoint_queryclose_20261103T233312Z. Os hashes coincidem com a referencia. Backup das9documentacoes antes das atualizacoes: C:/Users/MARCOSALEXANDREPERDO/AppData/Local/CodexBackupsPrivados/VBA-Excel/20261104T121731Z-ef30ba1a, produzido pelo script oficial de manutencao.

## Defeito comprovado e menor correcao

Ao clicar no X de Cad_NFE, UserForm_QueryClose executava Range("u1") = Range("u1") + 1 na planilha ativa Entradas Filial protegida: erro1004 real, depurado no editor. O mesmo comando existia no evento correspondente de Cad_NFE_Saida. A candidata remove somente essas duas escritas arbitrarias; preserva Cancel=True, mensagem e fechamento pelo botao existente. Nenhum controle, assinatura, regra fiscal, gravador ou esquema foi alterado.

Compilacao real Depurar > Compilar VBAProject executada; menu Compilar ficou desabilitado. Regressao pelos controles reais dos dois formularios: X mostra a orientacao original, botao Fechar oculta sem erro; U1 permaneceu vazia e protegida. EnableEvents, ScreenUpdating e DisplayAlerts=True; Calculation=-4105 antes/depois. Visible=False conferido pela janela imediata. Evidencia de interface transcrita manualmente; nao apresentada como log automatico.

## Pesquisa apos a reabertura

TextBoxPesquisa/ListBox reais selecionaram F2A0001 ativo, NF12003, valor150,3333/data03/11/2026; filial selecionou F2F0002, NF12002, valor456,7891/cor AMARELO QA. Os tokens exibidos sao EpisodeID, conferidos nos fatos ACE: [vinculos](./vinculos_pesquisa_interface.json). A chave da filial e String44 com zero inicial. Isso comprova consulta/seleccao apos reabertura da referencia; nao repete cadastro ou venda sobre a massa final.

## Formula desatualizada — tentativa preservada

Apos editar VBA e usar formularios, C1_Diagnostico retornou FORMULA_VALUE. Application.CalculateFull, seguido do mesmo diagnostico, retornou VERIFIED, sem reconstruir nem modificar fatos. A necessidade de recalc foi comprovada; a causa fina (invalidacao por edicao versus mudanca de data) nao foi isolada e permanece limitada a essa observacao. Nao removemos verificacao de formulas nem mascaramos o FAIL. Salvamento posterior confirmado Saved=True/Workbooks.Count=1.

## Arquivo efetivamente salvo

Novo pacote privado de preservacao: C:/Users/MARCOSALEXANDREPERDO/AppData/Local/VBA-Excel-Entregas/Entrega2_R26_QueryClose_20261104T1221Z_PENDENTE_ACEITE. Office executado apenas no laboratorio TEMP C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/Entrega2_Restauracao_Controlada_20261103T0146Z. O pacote nao e uma nova autoridade; nao abrir sua copia como instalacao.

|Arquivo|SHA256|
|---|---|
|Estoque_Fera_Veiculos_Candidato_Entrega2_Final.xlsm|efd1916395159da3d30ecc99838ff6e0e44dd9e0ca2b13b29570a3b59fe247e0|
|Banco_Candidato_Entrega2_Final.accdb|9fc0bfe8001120d2804b8f56836949249606e2030a6af431a71b1757c80267dc|
|Controle_E2_143C55EC132B48398413268B265117AF.accdb|b9e0c8aceb4474bb6b74ba727c7064982cfa9371e060b901bdf4b2de61632e2f|

72componentes extraidos do XLSM efetivo em fontes_embutidos e vbaProject.bin. Somente2corpos alterados; patch confere com a preparacao. Os5formularios mantem436assinaturas. Os127streams de designer continuam presentes, 43 diferem apos reserializacao pelo Office: nao ha alegacao de igualdade binaria nem de validacao visual integral. Form_referencia e apenas a referencia FRM/FRX; o binario salvo e a evidencia dos designers incorporados. Codificacao das fontes extraidas Windows1252, CRLF preservado.

As17tabelas antes/depois sao iguais, inclusive26fatos/26recibos/66projecoes e Pending0. Hash ACCDB e registro compartilhado inalterados. 39originais reconferidos: zero divergencias.

## Nova sessao depois da correcao

Encerramento normal somente da instancia exclusiva, apos Saved=True e1pasta. Nova instancia Excel aberta; candidata corrigida abriu pelo caminho literal. Office exibiu As macros foram desabilitadas. Nao clicamos Habilitar Conteudo nem alteramos politicas. Inicializacao/fluxo VBA desta nova sessao final: BLOQUEADO. Isso nao anula o PASS da reabertura anterior nem comprova a reabertura funcional da revisao corrigida.

Acao manual necessaria: no XLSM Final ja aberto no TEMP, clicar Habilitar Conteudo na faixa amarela, somente se reconhecer essa candidata sintetica; informar Macros habilitadas na candidata Final corrigida. A continuacao deve ler C1_UltimoResultado antes de qualquer carga manual, executar C1_Diagnostico, confirmar identidade/revisao e pesquisa sem repetir movimentos. Nenhuma permissao para alterar Trust Center ou seguranca global foi solicitada.

## Matriz e limites

[14verificacoes dirigidas](./MATRIZ_RETOMADA_QUERYCLOSE.md):11PASS,2FAIL historicos desta tentativa e1BLOQUEADO. Inclui modalidades estaticas/ADO distintas dos eventos VBA. QueryClose original e FORMULA_VALUE permanecem registrados, com novas evidencias corretivas; nao somar estes resultados como14fluxos operacionais.

Permanecem pendentes: dois gravadores Excel em sessoes independentes; restauracao de desastre sem origem/registro; ENTRY/rascunho orfao; abrangencia completa dos leitores/17abas e excecoes operacionais nao aprovadas. Anulacao/estorno/retificacao fiscal, monetaria ou historica continuam NAO DISPONIVEIS conforme contrato. Abertura fria final depende de habilitacao manual. Portanto Entrega2 nao encerrada, Entrega3 nao liberada e nenhuma promocao autorizada.

## Reversao

Para reverter so o ajuste QueryClose, conservar o pacote novo e usar os dois corpos anteriores do checkpoint sob novo laboratorio autorizado. Nao sobrescrever autoridade26 com banco antigo, nao editar paths/geracao para eleger copia e nao restaurar fatos16 sobre26. Esta correcao nao mudou o banco: seus hashes sao identicos. Respeitar o protocolo vigente de realocacao auditada para qualquer mudanca de local.

Agendamento: nenhum recurso de automacao agendada esta exposto nesta sessao. Nenhuma tarefa foi criada e nao ha promessa de retomada automatica em1hora.
