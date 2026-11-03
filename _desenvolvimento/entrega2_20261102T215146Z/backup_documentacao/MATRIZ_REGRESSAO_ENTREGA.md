# Matriz de regressão — Entrega 1

**Candidata sintética; resultados novos não alteram os FAIL históricos.** Compilação real PASS. Todos os registros abaixo foram produzidos pelo VBA real no Excel; o procedimento diferencia acionamento de evento de conferência auxiliar. Não são simulações JS nem substitutos COM de macro.

Fixture executada: C:\Users\MARCOS~1\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\Entrega1_20261102T182823Z\Estoque_Fera_Veiculos_Candidato_Integrado.xlsm. Banco preferencial no mesmo diretório. Linha de base definitiva preservada no [checkpoint revisão 8](./_desenvolvimento/entrega1_20261102T182823Z/checkpoints/10_base_revisao8_sem_pendencias.accdb).

[Log bruto UTF-16](./_desenvolvimento/entrega1_20261102T182823Z/evidencias/execucao/resultados_entrega1.tsv), [JSON estruturado](./_desenvolvimento/entrega1_20261102T182823Z/evidencias/resultados_estruturados.json), [estados antes/depois](./_desenvolvimento/entrega1_20261102T182823Z/evidencias/execucao/), [fontes da suíte efetivamente salvos](./_desenvolvimento/entrega1_20261102T182823Z/fontes/MDL/modTestesEntrega1.bas).

## Execuções, sem apagar falhas

| Execução | ID | Procedimento/modalidade | Resultado registrado | Observado |
|---|---|---|---|---|
| 1 | E1_ENTRADA_EVENTO | E1_Entrada → Cad_NFE.CommandButton_Gravar_Click | FAIL | ENTRADA_RECUSADA_OU_PENDENTE: Erro de definição de aplicativo ou de definição de objeto;AUDIT=OK;PROJECTION=UNIDENTIFIED_ROW;REV=1;MOV=1;RECEIPTS=1;ACTIVE=1;PENDING=0 |
| 2 | E1_V_VALIDO | E1_Venda → Cad_NFE_Saida.CommandButton_Gravar_Click | FAIL | SAIDA_RECUSADA_OU_PENDENTE: Valor de origem diverge do banco;AUDIT=OK;PROJECTION=VERIFIED;REV=1;MOV=1;RECEIPTS=1;ACTIVE=1;PENDING=0 |
| 3 | E1_V_IDENTIDADE | E1_Venda → conferência ADO VBA | FAIL | AUDIT=OK;PROJECTION=VERIFIED;REV=1;MOV=1;RECEIPTS=1;ACTIVE=1;PENDING=0 |
| 4 | E1_V_CHAVE_TEXTO | E1_Venda → conferência Excel VBA | PASS | 01234567890123456789012345678901234567890123 |
| 5 | E1_ENTRADA_ERRO_R2 | E1_Seguranca → leitor/auxiliar VBA real | FAIL | 58:O arquivo já existe;PROJECAO_VERIFICADA; 7F9FB807-EC54-4AB7-ACBD-23695DA7445B |
| 6 | E1_ENTRADA_EVENTO_R2 | E1_Entrada → Cad_NFE.CommandButton_Gravar_Click | PASS | ENTRADA_CONFIRMADA; 58BD7A5F-ACE9-43EF-95AD-555CE72B0AA2;AUDIT=OK;PROJECTION=VERIFIED;REV=1;MOV=1;RECEIPTS=1;ACTIVE=1;PENDING=0 |
| 7 | E1_ENTRADA_VALOR_R2 | E1_Entrada → consulta ADO e Excel VBA | PASS | ACE=900;EXCEL=900 |
| 8 | E1_V_VALIDO_R2 | E1_Venda → Cad_NFE_Saida.CommandButton_Gravar_Click | FAIL | SAIDA_CONFIRMADA; 79F0BB27-489A-482F-A226-142EB62B1175;AUDIT=OK;PROJECTION=BASE_REVISION_BUILD;REV=2;MOV=2;RECEIPTS=2;ACTIVE=0;PENDING=0 |
| 9 | E1_V_IDENTIDADE_R2 | E1_Venda → conferência ADO VBA | PASS | AUDIT=OK;PROJECTION=BASE_REVISION_BUILD;REV=2;MOV=2;RECEIPTS=2;ACTIVE=0;PENDING=0 |
| 10 | E1_V_CHAVE_TEXTO_R2 | E1_Venda → conferência Excel VBA | PASS | 01234567890123456789012345678901234567890123 |
| 11 | E1_ENTRADA_EVENTO_R3 | E1_Entrada → Cad_NFE.CommandButton_Gravar_Click | PASS | ENTRADA_CONFIRMADA; 7651663B-DE5F-4E77-9EAA-BF8BD7D6A5B4;AUDIT=OK;PROJECTION=VERIFIED;REV=1;MOV=1;RECEIPTS=1;ACTIVE=1;PENDING=0 |
| 12 | E1_ENTRADA_VALOR_R3 | E1_Entrada → consulta ADO e Excel VBA | PASS | ACE=900;EXCEL=900 |
| 13 | E1_V_VALIDO_R3 | E1_Venda → Cad_NFE_Saida.CommandButton_Gravar_Click | PASS | SAIDA_CONFIRMADA; 1DC32E56-2E6D-4995-8252-1DF6DDCD491E;AUDIT=OK;PROJECTION=VERIFIED;REV=2;MOV=2;RECEIPTS=2;ACTIVE=0;PENDING=0 |
| 14 | E1_V_IDENTIDADE_R3 | E1_Venda → conferência ADO VBA | PASS | AUDIT=OK;PROJECTION=VERIFIED;REV=2;MOV=2;RECEIPTS=2;ACTIVE=0;PENDING=0 |
| 15 | E1_V_CHAVE_TEXTO_R3 | E1_Venda → conferência Excel VBA | PASS | 01234567890123456789012345678901234567890123 |
| 16 | E1_ENTRADA_B_R3 | E1_Transferencia → evento de entrada | PASS | ENTRADA_CONFIRMADA; 272C4CBB-7562-4A92-ABF7-5A66FB8A6D49;AUDIT=OK;PROJECTION=VERIFIED;REV=3;MOV=3;RECEIPTS=3;ACTIVE=1;PENDING=0 |
| 17 | E1_T_VALIDO_R3 | E1_Transferencia → evento de saída matriz | PASS | SAIDA_CONFIRMADA; 61B8991E-80FF-43A9-8A51-2D80ED2405D6;AUDIT=OK;PROJECTION=VERIFIED;REV=4;MOV=4;RECEIPTS=4;ACTIVE=1;PENDING=0 |
| 18 | E1_T_FORMULAS_R3 | E1_Transferencia → conferência Excel VBA | PASS | AUDIT=OK;PROJECTION=VERIFIED;REV=4;MOV=4;RECEIPTS=4;ACTIVE=1;PENDING=0 |
| 19 | E1_D_VALIDO_R3 | E1_Devolucao → evento de saída filial | PASS | SAIDA_CONFIRMADA; C81C3462-AF9F-49C5-BBC5-B9EE5E05B2C8;AUDIT=OK;PROJECTION=VERIFIED;REV=5;MOV=5;RECEIPTS=5;ACTIVE=1;PENDING=0 |
| 20 | E1_D_VINCULOS_R3 | E1_Devolucao → conferência ADO VBA | PASS | AUDIT=OK;PROJECTION=VERIFIED;REV=5;MOV=5;RECEIPTS=5;ACTIVE=1;PENDING=0 |
| 21 | E1_ENTRADA_NFE_VAZIA_R3 | E1_Seguranca → evento entrada | PASS | ENTRADA_RECUSADA_OU_PENDENTE: Numero da NF-e invalido |
| 22 | E1_ENTRADA_VALOR_INVALIDO_R3 | E1_Seguranca → evento entrada | PASS | ENTRADA_RECUSADA_OU_PENDENTE: Valor monetario invalido: abc |
| 23 | E1_ENTRADA_DATA_INVALIDA_R3 | E1_Seguranca → evento entrada | PASS | ENTRADA_RECUSADA_OU_PENDENTE: Data civil invalida |
| 24 | E1_ENTRADA_DUPLICADA_R3 | E1_Seguranca → evento entrada | PASS | ENTRADA_RECUSADA_OU_PENDENTE: Numero da NFe ja cadastrado no estoque deste local |
| 25 | E1_PESQUISA_BAIXADO_R3 | E1_Seguranca → leitor/auxiliar VBA real | PASS | COUNT=0 |
| 26 | E1_TOKEN_OBSOLETO_R3 | E1_Seguranca → leitor/auxiliar VBA real | PASS | ENVIO_RECUSADO: Selecao obsoleta; pesquise novamente |
| 27 | E1_LISTINDEX_NEGATIVO_R3 | E1_Seguranca → leitor/auxiliar VBA real | PASS | ListIndex=-1 |
| 28 | E1_REORDENACAO_DETECTADA_R3 | E1_Seguranca → leitor/auxiliar VBA real | PASS | AUDIT=OK;PROJECTION=FORMULA:Entradas:4:12;REV=5;MOV=5;RECEIPTS=5;ACTIVE=1;PENDING=0 |
| 29 | E1_REORDENACAO_RECONSTRUIDA_R3 | E1_Seguranca → leitor/auxiliar VBA real | PASS | AUDIT=OK;PROJECTION=VERIFIED;REV=5;MOV=5;RECEIPTS=5;ACTIVE=1;PENDING=0 |
| 30 | E1_SAVE_CANCELADO_REAL_R3 | E1_Seguranca → leitor/auxiliar VBA real | PASS | 5:Save recusado; autoridade preservada, projecao pendente |
| 31 | E1_RECONSTRUCAO_APOS_SAVE_R3 | E1_Seguranca → leitor/auxiliar VBA real | PASS | AUDIT=OK;PROJECTION=VERIFIED;REV=5;MOV=5;RECEIPTS=5;ACTIVE=1;PENDING=0 |
| 32 | E1_BACKUP_REABERTO_R3 | E1_Seguranca → leitor/auxiliar VBA real | PASS | C:\Users\MARCOS~1\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\Entrega1_20261102T182823Z\Backup_C1_DD7DBA137BAD453F95A7705CECA6B472;AUDIT=OK |
| 33 | E1_REENTRADA_EPISODIO_R3 | E1_Reentrada → helper de identidade + evento entrada | PASS | ENTRADA_RECUSADA_OU_PENDENTE: Chave invalida;AUDIT=OK;PROJECTION=VERIFIED;REV=6;MOV=6;RECEIPTS=6;ACTIVE=2;PENDING=0 |
| 34 | E1_REENTRADA_VALOR_R3 | E1_Reentrada → helper de identidade + evento entrada | PASS | 650,25 |
| 35 | E1_HISTORICO_EPISODIO_R3 | E1_Reentrada → helper de identidade + evento entrada | PASS | CLOSED |
| 36 | E1_T_DUP_DESTINO_R3 | E1_DestinoDuplicado → eventos entrada filial/saída matriz | PASS | SAIDA_RECUSADA_OU_PENDENTE: Destino filial ja possui documento ativo; identidade ambigua recusada;AUDIT=OK;PROJECTION=VERIFIED;REV=7;MOV=7;RECEIPTS=7;ACTIVE=3;PENDING=0 |
| 37 | E1_TRANSFERENCIA_SUCESSIVA_R3 | E1_NovaTransferencia → evento saída matriz | FAIL | SAIDA_RECUSADA_OU_PENDENTE: Valor monetario invalido: ;AUDIT=OK;PROJECTION=VERIFIED;REV=7;MOV=7;RECEIPTS=7;ACTIVE=3;PENDING=0 |
| 38 | E1_TRANSFERENCIA_SUCESSIVA_R3 | E1_NovaTransferencia → evento saída matriz | PASS | SAIDA_CONFIRMADA; 7B2B1ECF-42A9-4A47-8EA9-4DD970DEEAD4;AUDIT=OK;PROJECTION=VERIFIED;REV=8;MOV=8;RECEIPTS=8;ACTIVE=3;PENDING=0 |
| 39 | E1_D_FALHA_BAIXA_R3 | E1_FalhaDevolucao → evento saída filial | PASS | SAIDA_RECUSADA_OU_PENDENTE: PENDENTE_ERRO:7003:Falha durante baixa logica;AUDIT=OK;PROJECTION=VERIFIED;REV=8;MOV=8;RECEIPTS=8;ACTIVE=3;PENDING=1 |
| 40 | E1_D_REPETICAO_PENDENTE_R3 | E1_FalhaDevolucao → evento saída filial | PASS | SAIDA_RECUSADA_OU_PENDENTE: Episodio pendente de conciliacao |
| 41 | E1_TRANSACAO_FALHA_1_R3 | E1_FalhasEntrada → evento entrada | PASS | ENTRADA_RECUSADA_OU_PENDENTE: PENDENTE_ERRO:7001:Falha apos movimento;AUDIT=OK;PROJECTION=VERIFIED;REV=8;MOV=8;RECEIPTS=8;ACTIVE=3;PENDING=1 |
| 42 | E1_TRANSACAO_FALHA_2_R3 | E1_FalhasEntrada → evento entrada | PASS | ENTRADA_RECUSADA_OU_PENDENTE: PENDENTE_ERRO:7002:Falha apos projecao;AUDIT=OK;PROJECTION=VERIFIED;REV=8;MOV=8;RECEIPTS=8;ACTIVE=3;PENDING=2 |
| 43 | E1_TRANSACAO_FALHA_6_R3 | E1_FalhasEntrada → evento entrada | PASS | ENTRADA_RECUSADA_OU_PENDENTE: PENDENTE_ERRO:7006:Antes do commit;AUDIT=OK;PROJECTION=VERIFIED;REV=8;MOV=8;RECEIPTS=8;ACTIVE=3;PENDING=3 |
| 44 | E1_POS_COMMIT_RECONSTRUIDO_R3 | E1_FalhasEntrada → evento entrada | PASS | ENTRADA_CONFIRMADA; 628C139F-205B-42CC-B17B-DAAC2F735DB6;AUDIT=OK;PROJECTION=VERIFIED;REV=9;MOV=9;RECEIPTS=9;ACTIVE=4;PENDING=3 |

## Interpretação e critérios

- 44 asserções executadas, 38 PASS e 6 FAIL. As quatro primeiras pertencem à primeira base monetária incorreta; R2 contém o FAIL de arquivo de dump existente e o FAIL de reconstrução da venda. As evidências forenses foram preservadas.
- R3: 34 execuções / 33 IDs únicos. A transferência sucessiva falhou por preenchimento apagado no Change; após correção foi repetida com PASS. Não ocultar a execução anterior.
- E1_REENTRADA_EPISODIO_R3 PASS atesta duas identidades de episódio e apenas uma ativa. O texto observado inclui falha de confirmação da projeção. **Não usar esse PASS como aprovação integral da reentrada inicial.** A correção de chave opcional foi seguida por reconstrução e operações nas revisões 7/8 com confirmação VERIFIED, evidenciadas na consulta final do ACE.
- E1_REORDENACAO_DETECTADA utiliza troca deliberada de fórmulas/campos e exige rejeição; não prova equivalência de um Sort nativo preservando Ignore. A recuperação demonstrada é reconstrução integral por identidade.
- E1_SAVE_CANCELADO_REAL usa clsSaveFase9/WorkbookBeforeSave, não Err.Raise antes do Save. A geração seguinte foi salva, reaberta pelo VBA e verificada.
- E1_TRANSACAO_FALHA_1/2/6 e E1_D_FALHA_BAIXA: **efeitos transacionais revertidos + PENDENTE_CONCILIACAO**, não recuperação operacional integral. Os rascunhos/intenção permanecem. E1_D_REPETICAO_PENDENTE comprova recusa, não recuperação.
- E1_POS_COMMIT_RECONSTRUIDO usa Err.Raise após CommitTrans bem-sucedido. Não simula nem comprova falha física de commit, crash ou energia.
- Os dumps incluem dados, tipos, fórmulas e NumberFormat; Ignore é verificado pelo verificador C1. Não alegar dump completo de todos os objetos/estilos do Excel.

## Regressões históricas exigidas

| Referência histórica | Nova evidência na candidata | Limite |
|---|---|---|
| V_VALIDO | E1_V_VALIDO_R3 PASS | Concluído no cenário sintético; regressão negativa completa pendente |
| T_VALIDO | E1_T_VALIDO_R3 PASS | Fórmulas M/N/O e valor discriminante 700/800 verificados |
| D_VALIDO | E1_D_VALIDO_R3 PASS | Mesmo episódio e quatro destinos da devolução |
| D_FALHA_BAIXA | E1_D_FALHA_BAIXA_R3 PASS de contenção | Pendente de conciliação, não recuperação integral |
| T_DUP_DESTINO | E1_T_DUP_DESTINO_R3 PASS de recusa | Sem associação automática entre veículos |
| C_FALHA_BAIXA | NÃO EXECUTADO no contexto legado de compensação C | ACE tem novo mecanismo; não reclassificar teste legado |
| C_FALHA_COMPENSACAO | NÃO EXECUTADO | Falha integral anterior preservada; não há promessa de compensação automática |

## Verificações adicionais e pendências

| Verificação | Modalidade | Estado | Evidência/limite |
|---|---|---|---|
| Depurar > Compilar VBAProject | Compilação real VBA | PASS | Comando desabilitado após execução; menu salvo |
| 69 componentes / 17 CodeNames / eventos dos cinco forms | Inspeção estática binária | PASS no escopo contado | Inventário; designer embutido não é byte a byte idêntico |
| Originais e backup F10 | SHA-256 | PASS | 39 originais e manifesto inicial |
| Banco definitivo revisão 8 | ADO COM direto, não macro | PASS | READY, VERIFIED, lease FREE, zero Pending, 8 recibos |
| Reabertura interativa do XLSM | Execução funcional VBA auxiliar | PASS na continuação | Inicialização automática e execução 47 em processo novo; FAIL anterior preservado |
| Preparar-CopiaSintetica.ps1 | Script local | BLOQUEADO | PSSecurityException; não houve bypass |
| Restauração completa em outra pasta com rebind | VBA/integração | NÃO EXECUTADO | Backup reaberto/auditado, mas isso não comprova restauração completa |
| Commit físico recusado, crash, energia, corrupção | Falha física | NÃO EXECUTADO | Nenhuma simulação substitui esses testes |
| Concorrência com gravadores C1 em dois processos | Integração | NÃO EXECUTADO | Fencing herdado F10 não dispensa nova regressão dos gravadores |
| Cancelado, edição/retificação e leitores completos do fluxo D | Funcional | NÃO EXECUTADO/BLOQUEADO | Gravadores incompatíveis não foram habilitados |
| XML divergente, negativos completos de venda, backup com falha de disco | Funcional | NÃO EXECUTADO | XMLs válidos sintéticos utilizados nos eventos; demais suítes históricas preservadas |

## Continuação — verificação de reabertura

O quadro de 44 execuções acima é a linha de base preservada. O log acumulado contém agora 47 asserções: 40 PASS e 7 FAIL. Os testes abaixo não reexecutaram as movimentações.

| Identificação | Procedimento/modalidade | Esperado | Observado | Resultado | Evidência |
|---|---|---|---|---|---|
| Execução 45 / E1_SESSAO_REABERTA_R3 | E1_SessaoReaberta, VBA auxiliar real | Diagnóstico VERIFIED na revisão 8 | ERROR:5:Copia divergente: caminho da autoridade nao confere | FAIL preservado | resultado_antes_correcao.tsv |
| Execução 46 / E1_SESSAO_REABERTA_R3 | Mesmo procedimento após correção na sessão aberta | AUDIT OK, projeção VERIFIED, sem Pending | REV=8;MOV=8;RECEIPTS=8;ACTIVE=3;PENDING=0 | PASS | resultados_apos_correcao.tsv |
| E1_ALIAS_83_LONGO | C1_MesmoArquivoAutoridade, Janela Imediata VBA | True | Verdadeiro | PASS | aliases-vba.jpg; transcrição da interface |
| E1_ALIAS_COPIA_DISTINTA | Mesmo auxiliar | False | Falso | PASS | Mesmo screenshot; não integra TSV |
| E1_ALIAS_AUSENTE | Mesmo auxiliar | False | Falso | PASS | Mesmo screenshot; não integra TSV |
| E1_COMPILACAO_ALIAS | Depurar > Compilar VBAProject | Ausência de erro e comando desabilitado após compilar | Confirmado pelo menu | PASS | compilacao-caminhos.txt |
| Execução 47 / E1_NOVA_INSTANCIA / E1_SESSAO_REABERTA_R3 | Excel PID 50212; conferência VBA auxiliar | Inicialização e diagnóstico sem carregamento manual | CACHE_AUDITADO; AUDIT=OK;PROJECTION=VERIFIED;REV=8;MOV=8;RECEIPTS=8;ACTIVE=3;PENDING=0 | PASS após habilitação manual | sessao-fria-inicializacao.jpg, sessao-fria-resultado.jpg e resultados_sessao_fria.tsv |

[Evidências da continuação](./_desenvolvimento/entrega1_20261102T182823Z/continuacao_reabertura_20261102T205915Z/). O campo de modalidade do TSV é genérico da suíte; E1_SessaoReaberta é conferência auxiliar real e não aciona evento de movimentação. A execução 46 ocorreu na sessão aberta; a execução 47 comprova fechamento/reabertura normal em novo processo após habilitação manual. Nenhuma delas comprova crash, durabilidade física ou recuperação operacional de Pending.
