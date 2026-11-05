# Matriz de regressão — Entrega 2

Escopo sintético, execução 20261102T215146Z. Referência R0: dez execuções PASS. Continuação: 61 execuções, 54 PASS e sete FAIL, 55 IDs distintos. Não somar repetições como novos cenários. O sufixo E2_R3 permaneceu no logger mesmo nas correções R4; horário, número da tentativa e checkpoint determinam a revisão. FAIL históricos não foram removidos.

Fixture de todas as linhas abaixo: Estoque_Fera_Veiculos_Candidato_Operacoes_Corrigidas.xlsm no TEMP Entrega2_20261102T215146Z; banco Banco_Candidato_Operacoes_Corrigidas.accdb. Os [fontes finais](./_desenvolvimento/entrega2_20261102T215146Z/fontes/MDL/modTestesEntrega1.bas), [log bruto](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/execucao/resultados_entrega2.tsv) e [resultados estruturados](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/resultados_estruturados.json) permitem conferir cada assertiva. O rótulo bruto VBA_FUNCIONAL_EVENTOS_REAIS é genérico; a coluna Modalidade abaixo corrige seu alcance por cenário.

## Massa, estado e reprodução

| Procedimentos | Entrada fictícia / ação | Antes e depois / evidência |
|---|---|---|
| E1_Entrada, E1_Venda | A: FIC5A01, NF11001,900; Venda22001,1000, XML válido sintético, chave iniciada em zero | Revisões0→1→2; episódio fechado, histórico preservado; referência_R0 e dumps A/V. |
| E1_Transferencia, E1_Devolucao | B:FIC5B02,11002,700; Transferência22002,800; Devolução33001,800 | Revisões2→5, mesmo episódio/custódia alterada, quatro destinos; dumps T/D R0. |
| E2_AlteracaoInterface | B: cor AZUL depois VERDE; motivo/responsável sintéticos pela interface de interação de teste | Revisões5→6→7; fatos anteriores mantidos, nova projeção e recibo. R1 falhou na reconstrução; correção Call e novo evento PASS. |
| E2_ReentradaInterface | ComboBox Reentrada, pesquisa FIC5A01, ListIndex0, NF11003,650,25,02/11/2026 | Revisão7→8; mesmo VehicleID, novo EpisodeID; dump reentrada_interface. Nenhum VehicleID foi injetado nessa seleção. |
| E1_NovaTransferencia, E2_DevolucaoSucessiva | B: transferência22003 e devolução33003 por800 | Revisões8→9 e10→11; identidade/custódia conferidas. Asserção inicial da devolução errou coluna SQL; evento confirmado não foi repetido, somente consulta corrigida. |
| E2_CanceladoInterface | A reentrado: Cancelado22990 sem XML, valores zero | Revisão9→10; duas saídas, origem ativa e custo conservado. Limite: matriz vinculada; sem aprovação fiscal. |
| E2_LeitoresEFormulas | A antigo900/novo650,25; versão decrementada no token; ListIndex=-1; filtro Venda | Sem movimento; dados deliberadamente distintos, anual/seleção/contagem conferidos. Filtro por helper direto. |
| E2_EntradasInvalidas | Filial: NF vazia/alfabética, valor inválido,31/02/2026,chave curta; placa conhecida sem vínculo | Sem criação de episódio, fato ou recibo; seis recusas antes dos efeitos. |
| E1_FalhaDevolucao | B filial,33002,800, falha APPLY3 durante baixa | Revisão10/recibos10 antes e depois; fatos/versão/custódia revertidos, Pending1; repetição recusada. [Forense](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/D_FALHA_BAIXA/). |
| E1_DestinoDuplicado | Entrada C filial22004/450; transferência para documento já ativo | Entrada confirma revisão12; transferência recusada sem efeito; dump duplicidade_depois. |
| E2_PrecisaoCriar/Consultar | Entrada11987/FX4D001/123,4567 | Revisão13; ACE exato, interface123,46 FAIL; após correção leitor/envio123,4567 PASS. |
| E2_FalhaProjecao | Entrada11988/FCE2X01/321,45, falha PROJECTION1 | Commit revisão14 conservado; Excel parcial; reconstrução verificada e replay sem duplicação; dump falha_projecao_recuperada. |
| E2_SaveEBackup | BeforeSave Cancel=True e botão SalvarTudo | ACE revisão14 conservado, Save recusado detectado, reconstrução/geração reaberta; [backup reaberto](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/backup_verificado/). |
| E2_ReordenacaoESaveAs | Troca de A4:BH4/A5:BH5; SaveAs em diretório inexistente | FORMULA:Entradas:4:12 detectado, reconstrução VERIFIED; erro1004 SaveAs, autoridade intacta. |
| E2_PortabilidadeGuardas | AuthorityPath divergente, BaseID divergente, banco ausente | Primeiro ensaio sem confirmação/fechamento de conexão: dois FAIL de observação. Depois BeginTrans/CommitTrans/fechamento explícito: guardas recusam; configuração restaurada e verificada. Não se alterou segurança. |
| E2_VendaReentrada | A segundo episódio, XML22011/02-11/700; custo650,25 | Revisão14→15; somente episódio selecionado encerrado; primeira passagem mantida; dump venda_reentrada. |
| E2_FalhaAntesCommit | C filial33996/450, APPLY6 | Primeiro FAIL13 da suíte ao serializar snapshot, antes do evento. Após correção, digest de fatos/recibos/versão/custódia igual; Pending1. [Forense](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/ANTES_COMMIT/). |
| E2_ReciboAusente | Remoção ADO controlada de recibo em ramificação sintética | Diagnóstico ERROR5 COUNT_Receipts, escrita contida; sem recriação automática. [Forense](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/RECIBO_AUSENTE/). |
| E2_ErroAposCommit | Entrada11986/FC8D001/80,88, APPLY8 | Revisão15→16, recibo reconhecido, reconstrução completa; dump erro_apos_commit. |

Os dumps UTF-16 contêm tabelas antes/depois e células com tipo, valor, fórmula e NumberFormat. Não capturam todas as propriedades visuais; Ignore é conferido pelo verificador da projeção e testes específicos, não por esse dump. As ramificações falhas foram preservadas antes da retomada do laboratório por checkpoint sem fatos posteriores naquela ramificação. Restaurar o checkpoint nesse ensaio não é uma rotina operacional de conciliação.

## Referência R0 — execução real dos eventos já integrados

| # | ID | Procedimento | Status | Esperado | Observado |
|---|---|---|---|---|---|
| 1 | E1_ENTRADA_EVENTO_R3 | consultar fonte do teste | PASS | 1 recibo, entrada ativa e projecao persistida verificada | ENTRADA_CONFIRMADA; D463125D-4E21-4ADC-A653-C5135779B955;AUDIT=OK;PROJECTION=VERIFIED;REV=1;MOV=1;RECEIPTS=1;ACTIVE=1;PENDING=0 |
| 2 | E1_ENTRADA_VALOR_R3 | consultar fonte do teste | PASS | 900 exatos no ACE e na projecao | ACE=900;EXCEL=900 |
| 3 | E1_V_VALIDO_R3 | consultar fonte do teste | PASS | Venda confirmada, origem preservada e inelegivel, 2 destinos vinculados | SAIDA_CONFIRMADA; 041FEA1A-F45B-4BDC-B86A-C555267381C6;AUDIT=OK;PROJECTION=VERIFIED;REV=2;MOV=2;RECEIPTS=2;ACTIVE=0;PENDING=0 |
| 4 | E1_V_IDENTIDADE_R3 | consultar fonte do teste | PASS | 4 registros de projecao do mesmo episodio e veiculo | AUDIT=OK;PROJECTION=VERIFIED;REV=2;MOV=2;RECEIPTS=2;ACTIVE=0;PENDING=0 |
| 5 | E1_V_CHAVE_TEXTO_R3 | consultar fonte do teste | PASS | 44 digitos com zero inicial como String | 01234567890123456789012345678901234567890123 |
| 6 | E1_ENTRADA_B_R3 | consultar fonte do teste | PASS | Nova entrada 700 exatos, 3 recibos | ENTRADA_CONFIRMADA; EB7637EF-EC21-4A09-9AE1-FCD58F30B9CC;AUDIT=OK;PROJECTION=VERIFIED;REV=3;MOV=3;RECEIPTS=3;ACTIVE=1;PENDING=0 |
| 7 | E1_T_VALIDO_R3 | consultar fonte do teste | PASS | Custodia filial, 4 destinos, mesmo episodio e veiculo | SAIDA_CONFIRMADA; DF69C1F7-9A8F-452C-8301-E2972E0A9BC7;AUDIT=OK;PROJECTION=VERIFIED;REV=4;MOV=4;RECEIPTS=4;ACTIVE=1;PENDING=0 |
| 8 | E1_T_FORMULAS_R3 | consultar fonte do teste | PASS | M/N/O por identidade e valor 800 distinto da entrada 700 | AUDIT=OK;PROJECTION=VERIFIED;REV=4;MOV=4;RECEIPTS=4;ACTIVE=1;PENDING=0 |
| 9 | E1_D_VALIDO_R3 | consultar fonte do teste | PASS | Retorno a matriz no mesmo episodio, quatro destinos preservados | SAIDA_CONFIRMADA; D12B8DD7-20B6-4027-8AEA-2F3BFFB4E328;AUDIT=OK;PROJECTION=VERIFIED;REV=5;MOV=5;RECEIPTS=5;ACTIVE=1;PENDING=0 |
| 10 | E1_D_VINCULOS_R3 | consultar fonte do teste | PASS | 10 projecoes historicas, uma custodia ativa | AUDIT=OK;PROJECTION=VERIFIED;REV=5;MOV=5;RECEIPTS=5;ACTIVE=1;PENDING=0 |

## Continuação — todas as tentativas preservadas

| # / horário | ID / procedimento | Modalidade | Status | Esperado | Observado |
|---|---|---|---|---|---|
| 1 / 2026-11-02 19:51:18 | E2_ALTERAR_EVENTO_E2 — E2_AlteracaoInterface | VBA: evento(s) real(is) e conferência | FAIL | Retificacao cria fato substitutivo e versao, com projecao verificada | RETIFICACAO_CONFIRMADA; 801EEE9C-D915-4055-8BCF-0B57F76376EE;AUDIT=OK;PROJECTION=PENDING_MEMORY_BASE_REVISION_BUILD;REV=6;MOV=6;RECEIPTS=6;ACTIVE=1;PENDING=0 |
| 2 / 2026-11-02 19:51:19 | E2_ALTERAR_IMUTAVEL_E2 — E2_AlteracaoInterface | VBA: conferência após eventos anteriores | PASS | Fato anterior intacto e auditoria consistente | AUDIT=OK |
| 3 / 2026-11-02 19:51:20 | E2_ALTERAR_LEITOR_E2 — E2_AlteracaoInterface | VBA: evento(s) real(is) e conferência | PASS | Pesquisa resolve novo fato do mesmo episodio | AZUL QA |
| 4 / 2026-11-02 19:51:20 | E2_EXCLUSAO_CLASSIFICADA_E2 — E2_AlteracaoInterface | VBA: evento(s) real(is) e conferência | PASS | Anulacao confirmada exige regra; nao apaga fato | EXCLUSAO_CLASSIFICADA; CADASTRO_CONFIRMADO: corrigir dados pelo botao Alterar; anulacao/duplicidade exige motivo e regra aprovada |
| 5 / 2026-11-02 19:56:39 | E2_ALTERAR_EVENTO_E2_R3 — E2_AlteracaoInterface | VBA: evento(s) real(is) e conferência | PASS | Retificacao cria fato substitutivo e versao, com projecao verificada | RETIFICACAO_CONFIRMADA; 4FE11057-79EC-4AAF-9200-4A445A686B43;AUDIT=OK;PROJECTION=VERIFIED;REV=7;MOV=7;RECEIPTS=7;ACTIVE=1;PENDING=0 |
| 6 / 2026-11-02 19:56:40 | E2_ALTERAR_IMUTAVEL_E2_R3 — E2_AlteracaoInterface | VBA: conferência após eventos anteriores | PASS | Fato anterior intacto e auditoria consistente | AUDIT=OK |
| 7 / 2026-11-02 19:56:41 | E2_ALTERAR_LEITOR_E2_R3 — E2_AlteracaoInterface | VBA: evento(s) real(is) e conferência | PASS | Pesquisa resolve novo fato do mesmo episodio | VERDE QA |
| 8 / 2026-11-02 19:56:41 | E2_EXCLUSAO_CLASSIFICADA_E2_R3 — E2_AlteracaoInterface | VBA: evento(s) real(is) e conferência | PASS | Anulacao confirmada exige regra; nao apaga fato | EXCLUSAO_CLASSIFICADA; CADASTRO_CONFIRMADO: corrigir dados pelo botao Alterar; anulacao/duplicidade exige motivo e regra aprovada |
| 9 / 2026-11-02 19:58:17 | E2_REENTRADA_CONTROLES_E2_R3 — E2_ReentradaInterface | VBA: evento(s) real(is) e conferência | PASS | Lista real escolhe episodio encerrado; novo episodio do mesmo VID | ENTRADA_CONFIRMADA; 5265DA42-6241-4218-8F5C-4E70CBB8C89B;AUDIT=OK;PROJECTION=VERIFIED;REV=8;MOV=8;RECEIPTS=8;ACTIVE=2;PENDING=0 |
| 10 / 2026-11-02 19:58:17 | E2_REENTRADA_DECIMAL_E2_R3 — E2_ReentradaInterface | VBA: evento(s) real(is) e conferência | PASS | 650,25 sem conversao por placa | 650,25 |
| 11 / 2026-11-02 20:04:25 | E1_TRANSFERENCIA_SUCESSIVA_E2_R3 — E1_NovaTransferencia | VBA: evento(s) real(is) e conferência | PASS | Novo movimento, mesmo episodio, projecao reconstruida | SAIDA_CONFIRMADA; 16F45BDE-240C-493B-83CC-BA59B7554987;AUDIT=OK;PROJECTION=VERIFIED;REV=9;MOV=9;RECEIPTS=9;ACTIVE=2;PENDING=0 |
| 12 / 2026-11-02 20:05:10 | E2_CANCELADO_EVENTO_E2_R3 — E2_CanceladoInterface | VBA: evento(s) real(is) e conferência | PASS | Cancelado ligado: dois destinos, origem permanece e custo preservado, sem XML | SAIDA_CONFIRMADA; 13D4E738-34F9-4CCD-81B4-A7ACA614D13A;AUDIT=OK;PROJECTION=VERIFIED;REV=10;MOV=10;RECEIPTS=10;ACTIVE=2;PENDING=0 |
| 13 / 2026-11-02 20:05:10 | E2_CANCELADO_HISTORICOS_E2_R3 — E2_CanceladoInterface | VBA: conferência após eventos anteriores | PASS | Duas projecoes da mesma movimentacao | AUDIT=OK |
| 14 / 2026-11-02 20:08:12 | E2_ANUAL_REENTRADA_E2_R3 — E2_LeitoresEFormulas | VBA: conferência após eventos anteriores | PASS | Mesmo VID, episodios distintos, valores900/650,25 e estados distintos | antigos=1;novos=1;AUDIT=OK;PROJECTION=VERIFIED;REV=10;MOV=10;RECEIPTS=10;ACTIVE=2;PENDING=0 |
| 15 / 2026-11-02 20:08:13 | E2_PESQUISA_REENTRADA_E2_R3 — E2_LeitoresEFormulas | VBA: evento(s) real(is) e conferência | PASS | Somente nova passagem ativa, sem primeira ocorrencia por placa | 7A41F294-489E-460C-BE0B-06EDBBA13C4C |
| 16 / 2026-11-02 20:08:13 | E2_SELECAO_OBSOLETA_E2_R3 — E2_LeitoresEFormulas | VBA: evento(s) real(is) e conferência | PASS | Token antigo nao abre gravador autorizado | ENVIO_RECUSADO: Selecao obsoleta; pesquise novamente |
| 17 / 2026-11-02 20:08:14 | E2_LISTINDEX_NEGATIVO_E2_R3 — E2_LeitoresEFormulas | VBA: evento(s) real(is) e conferência | PASS | Nenhuma acao por indice negativo | CONSULTA_POR_ID; 7A41F294-489E-460C-BE0B-06EDBBA13C4C |
| 18 / 2026-11-02 20:08:15 | E2_FILTRO_HISTORICO_E2_R3 — E2_LeitoresEFormulas | VBA: chamada direta do adaptador | PASS | Uma apresentacao por saida; nao duplica com Todas Saidas | COUNT=1 |
| 19 / 2026-11-02 20:08:16 | E2_EXCLUSAO_HISTORICO_E2_R3 — E2_LeitoresEFormulas | VBA: evento(s) real(is) e conferência | PASS | Historico confirmado nao e exclusao de estoque | EXCLUSAO_CLASSIFICADA; HISTORICO_OU_BAIXADO: nao excluir; retificacao substitutiva exige decisao |
| 20 / 2026-11-02 20:08:16 | E2_ESTOQUE_ATIVO_E2_R3 — E2_LeitoresEFormulas | VBA: conferência após eventos anteriores | PASS | Duas custodias ativas por identidade, excluindo baixadas | 2 |
| 21 / 2026-11-02 20:08:16 | E2_ENTRADA_FILIAL_INVALIDA_1_E2_R3 — E2_EntradasInvalidas | VBA: evento(s) real(is) e conferência | PASS | Recusa antes de criar episodio ou gravar fatos | ENTRADA_RECUSADA_OU_PENDENTE: Numero da NF-e invalido |
| 22 / 2026-11-02 20:08:16 | E2_ENTRADA_FILIAL_INVALIDA_2_E2_R3 — E2_EntradasInvalidas | VBA: evento(s) real(is) e conferência | PASS | Recusa antes de criar episodio ou gravar fatos | ENTRADA_RECUSADA_OU_PENDENTE: Valor monetario invalido: invalido |
| 23 / 2026-11-02 20:08:16 | E2_ENTRADA_FILIAL_INVALIDA_3_E2_R3 — E2_EntradasInvalidas | VBA: evento(s) real(is) e conferência | PASS | Recusa antes de criar episodio ou gravar fatos | ENTRADA_RECUSADA_OU_PENDENTE: Data civil invalida |
| 24 / 2026-11-02 20:08:16 | E2_ENTRADA_FILIAL_INVALIDA_4_E2_R3 — E2_EntradasInvalidas | VBA: evento(s) real(is) e conferência | PASS | Recusa antes de criar episodio ou gravar fatos | ENTRADA_RECUSADA_OU_PENDENTE: Chave deve conter exatamente44 digitos |
| 25 / 2026-11-02 20:08:16 | E2_ENTRADA_FILIAL_INVALIDA_5_E2_R3 — E2_EntradasInvalidas | VBA: evento(s) real(is) e conferência | PASS | Recusa antes de criar episodio ou gravar fatos | ENTRADA_RECUSADA_OU_PENDENTE: Numero da NF-e invalido |
| 26 / 2026-11-02 20:08:17 | E2_IDENTIDADE_CONHECIDA_SEM_VINCULO_E2_R3 — E2_EntradasInvalidas | VBA: evento(s) real(is) e conferência | PASS | Placa conhecida nao vincula sozinha | ENTRADA_RECUSADA_OU_PENDENTE: IDENTIDADE_INSUFICIENTE: placa conhecida; selecione Reentrada ou declare Novo veiculo com evidencia |
| 27 / 2026-11-02 20:10:15 | E1_D_FALHA_BAIXA_E2_R3 — E1_FalhaDevolucao | VBA: evento(s) real(is) e conferência | PASS | Rollback dos fatos e da baixa; intencao persistente exige conciliacao | SAIDA_RECUSADA_OU_PENDENTE: PENDENTE_ERRO:7003:Falha durante baixa logica;AUDIT=OK;PROJECTION=VERIFIED;REV=10;MOV=10;RECEIPTS=10;ACTIVE=2;PENDING=1 |
| 28 / 2026-11-02 20:10:16 | E1_D_REPETICAO_PENDENTE_E2_R3 — E1_FalhaDevolucao | VBA: evento(s) real(is) e conferência | PASS | Nao repetir nem prometer recuperacao integral | SAIDA_RECUSADA_OU_PENDENTE: Episodio pendente de conciliacao |
| 29 / 2026-11-02 20:11:30 | E2_D_SUCESSIVA_ERRO_E2_R3 — E2_DevolucaoSucessiva | VBA: evento(s) real(is) e conferência | FAIL | Devolucao apos segunda transferencia | -2147217904:Nenhum valor foi fornecido para um ou mais parâmetros necessários.;SAIDA_CONFIRMADA; 9B281191-FDC7-4485-A349-1C8C0D4EEEFD |
| 30 / 2026-11-02 20:12:07 | E1_T_DUP_DESTINO_E2_R3 — E1_DestinoDuplicado | VBA: evento(s) real(is) e conferência | PASS | Documento ativo no destino recusa transferencia inteira | SAIDA_RECUSADA_OU_PENDENTE: Destino filial ja possui documento ativo; identidade ambigua recusada;AUDIT=OK;PROJECTION=VERIFIED;REV=12;MOV=12;RECEIPTS=12;ACTIVE=3;PENDING=0 |
| 31 / 2026-11-02 20:15:43 | E2_D_SUCESSIVA_CONFERIDA_E2_R3 — E2_ConferirDestinos | VBA: conferência após eventos anteriores | PASS | Evento anterior confirmado, quatro destinos e custodia por episodio | 8293CA3C-2C37-4ABB-B51F-AFBD49593EEA;AUDIT=OK;PROJECTION=VERIFIED;REV=12;MOV=12;RECEIPTS=12;ACTIVE=3;PENDING=0 |
| 32 / 2026-11-02 20:15:44 | E2_ENTRADA_FILIAL_CHAVE_FORMULAS_E2_R3 — E2_ConferirDestinos | VBA: conferência após eventos anteriores | PASS | Cadastro real filial, valor450, chave44 texto, MNO por ID | MATCHES=1;AUDIT=OK |
| 33 / 2026-11-02 20:15:59 | E2_PRECISAO_ENTRADA_E2_R3 — E2_PrecisaoConsultar | VBA: evento(s) real(is) e conferência | FAIL | Leitor preserva quatro casas, sem alterar fato | 123,46 |
| 34 / 2026-11-02 20:16:00 | E2_PRECISAO_ENVIO_E2_R3 — E2_PrecisaoConsultar | VBA: evento(s) real(is) e conferência | FAIL | Envio ao gravador preserva valor exato da autoridade | 123,46 |
| 35 / 2026-11-02 20:17:58 | E2_PRECISAO_ENTRADA_E2_R3 — E2_PrecisaoConsultar | VBA: evento(s) real(is) e conferência | PASS | Leitor preserva quatro casas, sem alterar fato | 123,4567 |
| 36 / 2026-11-02 20:17:59 | E2_PRECISAO_ENVIO_E2_R3 — E2_PrecisaoConsultar | VBA: evento(s) real(is) e conferência | PASS | Envio ao gravador preserva valor exato da autoridade | 123,4567 |
| 37 / 2026-11-02 20:18:03 | E2_C_FALHA_COMPENSACAO_CONTENCAO_E2_R3 — E2_FalhaProjecao | VBA: evento(s) real(is) e conferência | PASS | Commit preservado; Excel parcial nao e apresentado como compensacao | ENTRADA_RECUSADA_OU_PENDENTE: Falha controlada na projecao;AUDIT=OK |
| 38 / 2026-11-02 20:18:03 | E2_APPLICATION_RESTAURADO_E2_R3 — E2_FalhaProjecao | VBA: evento(s) real(is) e conferência | PASS | Estados anteriores restaurados apos erro | EV=Verdadeiro;SU=Verdadeiro;CALC=-4105 |
| 39 / 2026-11-02 20:18:26 | E2_C_FALHA_COMPENSACAO_RECONSTRUIDO_E2_R3 — E2_FalhaProjecao | VBA: evento(s) real(is) e conferência | PASS | Recuperacao da projecao pela autoridade, sem desfazer commit | AUDIT=OK;PROJECTION=VERIFIED;REV=14;MOV=14;RECEIPTS=14;ACTIVE=5;PENDING=0 |
| 40 / 2026-11-02 20:18:43 | E2_REPLAY_EVENTO_E2_R3 — E2_FalhaProjecao | VBA: evento(s) real(is) e conferência | PASS | Mesma operacao confirmada nao duplica fato; replay auditado | ENTRADA_CONFIRMADA; F60DF75F-9A92-49AB-A3C8-B9670F71AA34;AUDIT=OK |
| 41 / 2026-11-02 20:18:50 | E2_SAVE_CANCELADO_REAL_E2_R3 — E2_SaveEBackup | VBA + operação real Excel | PASS | BeforeSave cancela; autoridade nao perde fatos | 5:Save recusado; autoridade preservada, projecao pendente |
| 42 / 2026-11-02 20:19:19 | E2_SAVE_RECUPERADO_E2_R3 — E2_SaveEBackup | VBA + operação real Excel | PASS | Arquivo salvo e geracao reaberta verificados | AUDIT=OK;PROJECTION=VERIFIED;REV=14;MOV=14;RECEIPTS=14;ACTIVE=5;PENDING=0 |
| 43 / 2026-11-02 20:19:21 | E2_BACKUP_BOTAO_REABERTO_E2_R3 — E2_SaveEBackup | VBA: evento(s) real(is) e conferência | PASS | Backup pelo botao, banco reaberto com recibos e lease livre | C:\Users\MARCOSALEXANDREPERDO\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\Entrega2_20261102T215146Z\Backup_C1_DF5D8A85BBCB4D9AB33860A1ABC01F67;AUDIT=OK |
| 44 / 2026-11-02 20:19:23 | E2_REORDENACAO_DETECTADA_E2_R3 — E2_ReordenacaoESaveAs | VBA + operação real Excel | PASS | Referencia fisica alterada nao recebe aprovacao silenciosa | AUDIT=OK;PROJECTION=FORMULA:Entradas:4:12;REV=14;MOV=14;RECEIPTS=14;ACTIVE=5;PENDING=0 |
| 45 / 2026-11-02 20:19:52 | E2_REORDENACAO_RECONSTRUIDA_E2_R3 — E2_ReordenacaoESaveAs | VBA + operação real Excel | PASS | Reconstrucao por IDs e metadados completos | AUDIT=OK;PROJECTION=VERIFIED;REV=14;MOV=14;RECEIPTS=14;ACTIVE=5;PENDING=0 |
| 46 / 2026-11-02 20:19:57 | E2_SAVEAS_INVALIDO_REAL_E2_R3 — E2_ReordenacaoESaveAs | VBA + operação real Excel | PASS | SaveAs invalido nao altera autoridade ou arquivo atual | 1004:O Microsoft Excel não pode acessar o arquivo 'C:\Users\MARCOSALEXANDREPERDO\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\Entrega2_20261102T215146Z\diretorio_inexistente_E2\F6C1EF00'. Há várias razões possíveis:  • O nome do arquivo ou caminho não existe. • O arquivo está sendo usado por outro programa. • A pasta de trabalho que você está tentando salvar tem o mesmo nome de outra  pasta de trabalho aberta no momento. |
| 47 / 2026-11-02 20:23:14 | E2_CAMINHO_DIVERGENTE_E2_R3 — E2_PortabilidadeGuardas | VBA auxiliar + ADO sintético (não evento de movimentação) | FAIL | Nome/copia nao elege autoridade | AUDIT=OK;PROJECTION=VERIFIED;REV=14;MOV=14;RECEIPTS=14;ACTIVE=5;PENDING=0 |
| 48 / 2026-11-02 20:23:18 | E2_BASEID_DIVERGENTE_E2_R3 — E2_PortabilidadeGuardas | VBA auxiliar + ADO sintético (não evento de movimentação) | FAIL | BaseID e linhagem conferidos antes de escrever | AUDIT=OK;PROJECTION=VERIFIED;REV=14;MOV=14;RECEIPTS=14;ACTIVE=5;PENDING=0 |
| 49 / 2026-11-02 20:23:18 | E2_BANCO_AUSENTE_E2_R3 — E2_PortabilidadeGuardas | VBA auxiliar + ADO sintético (não evento de movimentação) | PASS | Banco ausente nao gera outra autoridade | ERROR:-2147467259:Não foi possível encontrar o arquivo 'C:\Users\MARCOSALEXANDREPERDO\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\Entrega2_20261102T215146Z\Banco_Candidato_Operacoes_Corrigidas.accdb'. |
| 50 / 2026-11-02 20:23:27 | E2_GUARDAS_RESTAURADAS_E2_R3 — E2_PortabilidadeGuardas | VBA auxiliar + ADO sintético (não evento de movimentação) | PASS | Vinculo original restaurado no laboratorio, sem alterar fatos | AUDIT=OK;PROJECTION=VERIFIED;REV=14;MOV=14;RECEIPTS=14;ACTIVE=5;PENDING=0 |
| 51 / 2026-11-02 20:26:23 | E2_CAMINHO_DIVERGENTE_E2_R3 — E2_PortabilidadeGuardas | VBA auxiliar + ADO sintético (não evento de movimentação) | PASS | Nome/copia nao elege autoridade | ERROR:5:Copia divergente: caminho da autoridade nao confere |
| 52 / 2026-11-02 20:26:23 | E2_BASEID_DIVERGENTE_E2_R3 — E2_PortabilidadeGuardas | VBA auxiliar + ADO sintético (não evento de movimentação) | PASS | BaseID e linhagem conferidos antes de escrever | ERROR:5:BaseID divergente |
| 53 / 2026-11-02 20:26:23 | E2_BANCO_AUSENTE_E2_R3 — E2_PortabilidadeGuardas | VBA auxiliar + ADO sintético (não evento de movimentação) | PASS | Banco ausente nao gera outra autoridade | ERROR:-2147467259:Não foi possível encontrar o arquivo 'C:\Users\MARCOSALEXANDREPERDO\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\Entrega2_20261102T215146Z\Banco_Candidato_Operacoes_Corrigidas.accdb'. |
| 54 / 2026-11-02 20:26:31 | E2_GUARDAS_RESTAURADAS_E2_R3 — E2_PortabilidadeGuardas | VBA auxiliar + ADO sintético (não evento de movimentação) | PASS | Vinculo original restaurado no laboratorio, sem alterar fatos | AUDIT=OK;PROJECTION=VERIFIED;REV=14;MOV=14;RECEIPTS=14;ACTIVE=5;PENDING=0 |
| 55 / 2026-11-02 20:28:37 | E2_V_VALIDO_REENTRADA_E2_R3 — E2_VendaReentrada | VBA: evento(s) real(is) e conferência | PASS | Venda real encerra somente episodio selecionado, preserva primeira passagem | SAIDA_CONFIRMADA; 1F7B91B8-2685-4C23-8C2D-EB6580BEE853;AUDIT=OK;PROJECTION=VERIFIED;REV=15;MOV=15;RECEIPTS=15;ACTIVE=4;PENDING=0 |
| 56 / 2026-11-02 20:28:38 | E2_V_VALORES_IDENTIDADE_E2_R3 — E2_VendaReentrada | VBA: conferência após eventos anteriores | PASS | VIN650,25 VOUT700, mesmo VID e episodio correto | 7A41F294-489E-460C-BE0B-06EDBBA13C4C;AUDIT=OK |
| 57 / 2026-11-02 20:28:39 | E2_V_BAIXADO_PESQUISA_E2_R3 — E2_VendaReentrada | VBA: conferência após eventos anteriores | PASS | Nenhum episodio vendido aparece como ativo | COUNT=0 |
| 58 / 2026-11-02 20:30:05 | E2_FALHA_COMMIT_ERRO_E2_R3 — E2_FalhaAntesCommit | FAIL da suíte antes do evento | FAIL | Conter erro anterior ao commit | 13:Tipos incompatíveis |
| 59 / 2026-11-02 20:33:34 | E2_FALHA_ANTES_COMMIT_E2_R3 — E2_FalhaAntesCommit | VBA: evento(s) real(is) e conferência | PASS | Rollback de fatos, historico e versao; intencao pendente nao e recuperacao integral | SAIDA_RECUSADA_OU_PENDENTE: PENDENTE_ERRO:7006:Antes do commit;AUDIT=OK;PROJECTION=VERIFIED;REV=15;MOV=15;RECEIPTS=15;ACTIVE=4;PENDING=1 |
| 60 / 2026-11-02 20:35:04 | E2_RECIBO_AUSENTE_CONTIDO_E2_R3 — E2_ReciboAusente | VBA auxiliar + ADO sintético (não evento de movimentação) | PASS | Autoridade inconsistente bloqueada; exige conciliacao, nao recria recibo | ERROR:5:Conciliacao obrigatoria: COUNT_Receipts |
| 61 / 2026-11-02 21:32:56 | E2_ERRO_APOS_COMMIT_E2_R3 — E2_ErroAposCommit | VBA: evento(s) real(is) e conferência | PASS | Commit reconhecido; reconstrucao concluida sem repetir fatos | ENTRADA_CONFIRMADA; 5A5A5330-AA70-4237-96FF-420991C137DC;AUDIT=OK;PROJECTION=VERIFIED;REV=16;MOV=16;RECEIPTS=16;ACTIVE=5;PENDING=0 |

## Sete defeitos históricos — resultados próprios da candidata

| Caso histórico | Correspondência na candidata | Resultado e garantia |
|---|---|---|
| V_VALIDO | E1_Venda R0 e E2_VendaReentrada | PASS dos eventos novos: venda fecha somente episódio selecionado, sem Delete, recibos/históricos/projeção verificados. FAIL original preservado. |
| T_VALIDO | E1_Transferencia R0 e E1_NovaTransferencia | PASS: custódia, quatro destinos, M/N/O e chave por identidade; continuidade do episódio ainda hipótese operacional. |
| D_VALIDO | E1_Devolucao R0 e consulta E2_ConferirDestinos após evento sucessivo | PASS no escopo conferido; primeira asserção SQL da segunda devolução FAIL preservada. |
| D_FALHA_BAIXA | E1_FalhaDevolucao APPLY3 | PASS de rollback transacional dos fatos + PENDENTE_CONCILIACAO. Não é recuperação integral da intenção. |
| T_DUP_DESTINO | E1_DestinoDuplicado | PASS de recusa antes dos efeitos de transferência; não selecionar primeiro registro ambíguo. |
| C_FALHA_BAIXA | Falha de baixa no novo motor ACE / APPLY3 e APPLY6 | Novo mecanismo demonstrou reversão dos fatos e bloqueio de repetição. A antiga classe de compensação não foi corrigida/reexecutada; FAIL histórico permanece. |
| C_FALHA_COMPENSACAO | E2_FalhaProjecao | PASS de reconstrução Excel após commit, sem compensação de fatos confirmados. Não comprova recuperação da classe legada nem falha física. |

## Compilação, inspeções e testes pendentes

| Item | Modalidade | Estado | Evidência / limite |
|---|---|---|---|
| Compilação final | Compilação real VBA pelo VBE | PASS | [Interface após Compile](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/compilacao_final_R4.txt); fontes extraídos do arquivo salvo. |
| Integridade E1/originais | SHA-256 local | PASS | 241 entradas anteriores e41 arquivos de referência sem divergência. |
| Componentes/assinaturas | Análise estática binária/textual | PASS no escopo |69 componentes/17abas/5forms/436 assinaturas mantidas; não equivale a teste funcional de todos os controles. |
| Reabertura interativa E2 | VBA em contexto novo | BLOQUEADO | Aviso de macros aguardando ação manual; gerações de projeção reabertas não substituem esse cenário. |
| Portabilidade outra pasta/TEMP, permissão recusada e instalação | Funcional da candidata | NÃO EXECUTADO | Cópia deslocada é deliberadamente recusada; rebind/instalador integral não comprovado. |
| Conciliação humana auditável com retomada/aborto | Fluxo operacional | NÃO EXECUTADO | Só diagnóstico e contenção integrados; depende de protocolo/decisão. |
| Concorrência de dois gravadores C2 | Dois processos/eventos reais | NÃO EXECUTADO | Testes F10 não foram somados como execução E2. |
| Retificação filial, formulário múltiplo, todos filtros/navegação/relatórios | Eventos/UI | NÃO EXECUTADO integralmente | Revisão estática dos57 pontos não aprova todos os caminhos. |
| XML divergente/malformado e matriz completa de saídas inválidas | Eventos com XML sintético | NÃO EXECUTADO nesta entrega | XMLs válidos dos fluxos executados não substituem esses cenários. |
| Crash/energia/disco cheio/falha física de commit | Falha física | NÃO EXECUTADO | Nenhuma alegação de durabilidade além de Save/reabertura realmente observados. |

Sem testes COM equivalentes ou simulações sendo promovidos a PASS de eventos. [Defeitos e decisões](./REGISTRO_DEFEITOS_ABERTOS.md) impedem homologação e conclusão integral.

## Limite de isolamento dos cenários

Os fluxos normais foram executados em sequência no laboratório E2, com snapshots de estado e checkpoints por fronteira. Os testes de falha partiram de checkpoints conhecidos, preservaram as ramificações forenses e só então retomaram outra ramificação. Não foram executadas sete candidatas físicas independentes, uma para cada caso histórico. Essa exigência de regressão independente permanece pendente; os PASS acima se limitam às execuções e estados documentados.


## Adendo — sessão reaberta, estado atual

A linha BLOQUEADO de reabertura acima registra o estado anterior. Após ação manual do usuário, foi substituída por nova evidência PASS, sem apagar o histórico.

| ID | Procedimento/modalidade | Antes | Esperado | Observado | Resultado | Evidência |
|---|---|---|---|---|---|---|
| E1_SESSAO_REABERTA_E2_R3 | E1_SessaoReaberta; execução real VBA auxiliar, sem evento de movimentação | Projeto E2 fechado anteriormente, reaberto no TEMP; cache automático conferido antes de chamada manual;rev16 | Cache e fórmulas reconstituídos, autoridade/projeção coerentes | AUDIT=OK;PROJECTION=VERIFIED;REV=16;MOV=16;RECEIPTS=16;ACTIVE=5;PENDING=0; hashes inalterados | PASS | [TSV](./_desenvolvimento/entrega2_20261102T215146Z/continuacao_reabertura_20261103T010409Z/resultado_reabertura.tsv), [estado](./_desenvolvimento/entrega2_20261102T215146Z/continuacao_reabertura_20261103T010409Z/evidencias_reabertura.json) |

Acumulado da continuação:62execuções/55PASS/7FAIL,56IDs; referênciaR0 permanece10PASS separados. O cache automático CACHE_AUDITADO e os estados do Excel foram transcritos da janela Imediata, identificados como conferências auxiliares, não novos eventos do produto. O encerramento/reabertura não foi crash ou nova instalação.


## Continuação do fechamento — 03/11/2026

Não substitui resultados históricos. [Relatório/evidências](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_20261103T2147Z/FECHAMENTO_ENTREGA2_EM_ANDAMENTO.md) e [TSV de oito asserções](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_20261103T2147Z/resultados_fechamento_antes_realocacao.tsv).

| ID / procedimento | Modalidade | Resultado observado | Status |
|---|---|---|---|
| F2_INSTALACAO / F2_Baseline | VBA integração | Esquema7, revisão16, auditoriaOK/projeçãoVERIFIED | PASS |
| F2_APP_BASELINE / F2_Baseline | VBA estados Application | True/True/True/-4105 | PASS |
| F2_PRECOMMIT_FATOS / F2_FalhaEConciliacao | Evento real Devolução | APPLY6; fatos/versão/recibos intactos; intenção persistiu | PASS — rollback de fatos |
| F2_REPETICAO_PENDENTE | Evento real | Mesmo OperationID recusado | PASS — RECUSADO, não recuperado |
| F2_OUTRO_ID_PENDENTE | Evento real | Novo OperationID não contornou pendência do episódio | PASS — RECUSADO |
| F2_EXIGE_CONFIRMACAO | API VBA conciliação | Token inválido recusado | PASS |
| F2_CONCILIACAO_SEM_EFEITOS | API VBA com confirmação sintética | ABORTED, recibo1, Pending0, fatos intactos | PASS — CONCILIADO_SEM_EFEITOS |
| F2_CONCILIACAO_IDEMPOTENTE | API VBA | JA_CONCILIADO, sem recibo extra | PASS |
| C3_Realocar | VBA real, transcrição da interface | Geração2/destino novo; origem revogada | PASS no escopo de realocação ao vivo |
| Guarda TEMP da sessão | Reabertura real antes da correção | Somente TEMP recusou caminho da sessão anterior | FAIL preservado |
| Correção Guard10/TempC3 | Diagnóstico/reconstrução VBA na sessão aberta | AUDIT=OK/VERIFIED/rev16 | PASS restrito; frio pendente |
| Compilação da suíte ampliada/corrigida | Compilação real VBE | Sem erro; Compilar desabilitado | PASS |
| Diagnóstico da sessão nova | Execução VBA pela interface | Macros neste projeto estão desativadas | BLOQUEADO |
| F2_Restaurada | Suíte VBA preparada | Não executada nesta revisão; dependência de macros | NÃO EXECUTADO |
| F2_EntradasNovas, Transferencia, Devolucao, Venda, Cancelado, ReentradaNova, FiltrosReais | Eventos reais preparados | Não executados nesta revisão | NÃO EXECUTADO |
| Reabertura da cópia ACCDB preservada | Leitura ADO independente | Mesmas17 tabelas e contagens, rev16 | PASS de preservação; não é teste dos eventos VBA |

Sete casos históricos permanecem com os resultados anteriores. Plano obsoleto, lease órfão, interface completa da conciliação, concorrência integrada, XML restante e backup de desastre continuam pendentes. Correção SQL da suíte foi necessária: KindText está em Movements, não Projections. Não somar compilação, hashes e diagnósticos à contagem de eventos.


## Fechamento revisão 26 — nova evidência, sem apagar a referência

[58 execuções reais detalhadas](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_macros_20261103T2213Z/MATRIZ_EXECUTADA_FECHAMENTO.md): 57 PASS e um FAIL da instrumentação. Os 14 PASS do log de fechamento, 16 PASS da suíte E2 reutilizada e 27 PASS/1 FAIL do módulo final são blocos distintos; não significam 58 operações completas. As oito asserções anteriores do fechamento e os 62 resultados históricos permanecem separados.

| Caso histórico | Evidência atual da candidata | Classificação e limite |
|---|---|---|
| V_VALIDO | F2_V_VALIDO e F2_V_VALIDO_BAIXA, evento real, NF24012/400,2222 | PASS no fluxo candidato; episódio encerrado e histórico por ID, sem Delete |
| T_VALIDO | F2_T_VALIDO, NF24002/300,1111, quatro destinos | PASS no fluxo candidato; mesmo episódio e custódia filial experimental |
| D_VALIDO | F2_D_VALIDO e CF_NOVA_TENTATIVA_CONCILIADA, eventos reais | PASS nos fluxos descritos; quatro destinos, custódia matriz e vínculos |
| D_FALHA_BAIXA | CF_FALHA_BAIXA_LOGICA/APPLY3 + conciliação via diálogo | Efeitos revertidos; intenção inicialmente pendente, depois CONCILIADA_SEM_EFEITOS e nova tentativa confirmada |
| T_DUP_DESTINO | E1_T_DUP_DESTINO_E2_R3 na referência anterior | Recusa candidata já demonstrada; não reexecutada nesta continuação |
| C_FALHA_BAIXA | APPLY3/APPLY6 no novo motor ACE | Reversão dos fatos, conciliação limitada demonstrada; não é execução/correção da antiga classe |
| C_FALHA_COMPENSACAO | E2_FalhaProjecao na referência anterior; Save cancelado e reconstrução reexecutados | Reconstrução da projeção a partir do commit; sem compensação após Delete/commit e sem garantia física |

Os sete FAIL históricos da implementação anterior permanecem. Reabertura final R26: BLOQUEADO por macros; dois gravadores Excel e restauração sem origem/registro: NÃO EXECUTADOS. Não usar esses limites como PASS. Os XMLs divergentes de número/chave são validador direto; demais negativos acionaram o gravador real, três falhando antes de abrir o relatório. [Transcrição da interface XML](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_macros_20261103T2213Z/xml_interface_observada_final.json).


## Retomada de04/11/2026 — QueryClose, revisao26

[Relatorio da retomada](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_reabertura_final_20261103T233016Z/RELATORIO_RETOMADA_QUERYCLOSE.md). Os registros anteriores permanecem historicos. A referencia R26 reabriu com CACHE_AUDITADO e diagnostico VERIFIED; pesquisas reais matriz/filial confirmaram episodios e valores150,3333/456,7891. Foi reproduzido erro1004 no X de Cad_NFE: incremento arbitrario de U1 na planilha ativa protegida. Removida somente essa escrita em QueryClose de Cad_NFE e Cad_NFE_Saida na candidata. Compilacao real e regressao de X/botao Fechar dos dois formularios PASS.17tabelas/ACCDB inalterados;39originais com hashes iguais.

Novo XLSM preservado: SHA256 efd1916395159da3d30ecc99838ff6e0e44dd9e0ca2b13b29570a3b59fe247e0; ACCDB9fc0bfe8001120d2804b8f56836949249606e2030a6af431a71b1757c80267dc. Pacote privado Entrega2_R26_QueryClose_20261104T1221Z_PENDENTE_ACEITE.72componentes;436assinaturas dos5formularios preservadas; designers reserializados, sem alegar identidade binaria.

FORMULA_VALUE apos edicao/formularios preservado como FAIL de tentativa; recalc completo recuperou VERIFIED, sem alterar fatos. Causa fina ainda nao isolada. Arquivo salvo e instancia exclusiva encerrada normalmente. A nova abertura do XLSM corrigido concluiu, mas Office voltou a desabilitar macros: inicializacao/fluxo final BLOQUEADOS ate habilitacao manual. Entrega2 continua em fechamento, NAO HOMOLOGADA; nenhuma liberacao de Entrega3.

[Matriz dirigida](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_reabertura_final_20261103T233016Z/MATRIZ_RETOMADA_QUERYCLOSE.md):14verificacoes,11PASS/2FAIL/1BLOQUEADO. Inclui analise estatica e ADO separados de eventos VBA. Nao somar como14fluxos e nao alterar os sete FAIL legados.


## Retomada controlada — diagnóstico ACE em 04/11/2026, 18:21 UTC

[Evidências e resultado da etapa](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_diagnostico_ACE_20261104T182134Z/RESULTADO_ETAPA_CONTROLADA.md). Ponytail 5.3.3 Full; Git inicial main/b7b9d99 limpo. AGENTS vigente preservado. Apenas o Excel operacional estava aberto; nenhum comando foi enviado a ele. Candidata, ACCDB e registro ativo mantêm seus hashes (42b34e…/9fc0bf…/b9e0c8…). Revisão26, 26movimentos/26recibos e Pending0 confirmados por leitura ADO. Nenhum VBA ou arquivo operacional alterado.

Seis verificações únicas COM/ADO em duas conexões de um processo, somente registro descartável TEMP: seis PASS. A primeira tentativa executou quatro; a segunda repetiu as quatro e acrescentou duas. Não são dez testes únicos nem execução funcional VBA. Conflito capturado com SQLState3218; rollback imediato após a recusa funcionou. SQLState3034 apareceu apenas no rollback após commit deliberado da sondagem. A causa do diagnóstico E2-FECH-06 no evento real continua não isolada, sem correção especulativa.

|IDs novos|Modalidade|Resultado|Limite|
|---|---|---|---|
|ACE_BASELINE, ACE_CONFLITO_RECUSADO, ACE_NOVA_TRANSACAO|COM/ADO, banco descartável|3PASS|Não executam eventos de formulários|
|ACE_CONFLITO_SEM_COMMIT, ACE_LIMPEZA_SEM_COMMIT, ACE_ESTADO_RESTAURADO|COM/ADO, banco descartável|3PASS|Não isolam a mensagem E2-FECH-06 no VBA|
|Inicialização, compilação e evento real nesta sessão|Execução VBA|NÃO EXECUTADO|Candidata não aberta na interface|

Todas as tentativas anteriores e os sete defeitos históricos mantêm seus resultados.


## Reabertura real da candidata corrigida — 04/11/2026, 22:04–22:15 UTC

[Resultado e evidências](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_vba_real_20261104T185156Z/RESULTADO_REABERTURA_VBA_REAL.md). Instância exclusiva PID7644; XLSM42b34e… reaberto com CACHE_AUDITADO automático antes de carga manual. C1_Diagnostico real antes/depois: AUDIT=OK;PROJECTION=VERIFIED;REV=26;MOV=26;RECEIPTS=26;ACTIVE=7;PENDING=0. X e CommandButtonFechar dos dois UserForms executados por interface, sem1004; ambos Visible=False. U1 vazia/protegida e estados Application preservados. Seis verificações dirigidas PASS, com modalidades separadas; não são seis fluxos de estoque.

IDs novos: E2-R26-FINAL-COLD, E2-R26-FINAL-DIAG, E2-R26-FINAL-QC-ENTRADA, E2-R26-FINAL-QC-SAIDA, E2-R26-FINAL-SENTINELA, E2-R26-FINAL-INTEGRIDADE. [Modalidade/esperado/observado/evidência](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_vba_real_20261104T185156Z/matriz_vba_real.json). Compilação explícita nova: NÃO EXECUTADO; menu já desabilitado. Os sete FAIL históricos e a sequência COM/ADO anterior não foram reclassificados.


## Continuação controlada — 05/11/2026: preparação do teste ACE

Preservado o XLSM encontrado com hash cb610aa3e19363d4b66e32a219f33f541b5907c341f122096b644d019f96c9c4; não restaurada versão anterior. Comparação: 72 corpos VBA sem diferenças; células das oito abas divergentes preservadas, com alterações de hash/salt de proteção. Há diferenças de streams binários de formulários/caches, sem nova validação funcional. ACCDB R26 e registro mantêm os hashes anteriores. Criada fixture independente com novas identidades e fatos vazios; comparação DAO confirmou regras/índices e 28 relações nas 18 tabelas. Esses PASS de preparação não equivalem à compilação ou ao evento real.

Gravador real: BLOQUEADO pelo aviso “As macros foram desabilitadas” na nova cópia. Nova compilação: NÃO EXECUTADO. Nenhuma regra fiscal, código VBA ou arquivo operacional foi alterado. Entrega 2 permanece aberta e NÃO HOMOLOGADA.

Evidências, hashes, checkpoint, primeira tentativa revertida e roteiro de captura: [resultado da preparação ACE](_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_ACE_20261105T101001Z/RESULTADO_PREPARACAO_ACE.md). Registros anteriores preservados integralmente.
