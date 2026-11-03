# Matriz de regressão — Fase 9

Somente execução experimental sintética. **65 asserções VBA: 64 PASS e 1 FAIL preservado; quatro benchmarks VBA PASS e três testes COM de concorrência entre processos PASS, separados.** Sondagem inicial ADO e hashes não aumentam essa contagem. Nenhum evento operacional principal foi aprovado nesta fase.

Resultado PASS/FAIL descreve a asserção; Tratamento descreve o efeito sobre o estado (RECUPERADO, RECUSADO, PENDENTE_CONCILIACAO etc.). Rótulos CONFIRMADO/VERIFICADO são operações normais, não classes de recuperação. F9_PRECISAO_TEXTO gravou literalmente RECUPERADO no log, mas demonstra somente tipagem/precisão; não contabilizar esse rótulo como recuperação de falha.

## Dados, origem e estado comparado

Ciclo:VID_A/EP_A1 entrada1000,1234 e venda1370,9876; reentrada VID_A/EP_A2 entrada1500,1111, transferências1940,2222/2500,4444 e devolução2100,3333; VID_B/EP_B1 entrada2200,4321. NF88001 se repete sem fundir episódios. Chave sintética44 dígitos com zero inicial. Data2026-11-02. Custódia global é hipótese experimental, não política aprovada.

F9_Core registra alterações BeforeCanon/AfterCanon e ciclo_final.tsv. F9_Faults guarda12 tabelas antes/depois de cada cenário falha_1..8 em falha_{n}_antes/depois.tsv e compara assinatura tipada de Episodes/Movements/Projections/Receipts/ChangeLog/Meta. Operations/Attempts/Pending são intenção durável anterior e ficam deliberadamente diferentes; não anunciar rollback de toda a intenção. Projeções foram comparadas por PID/EP/VID e valores discriminantes; não houve snapshot integral das17 abas do produto. [Diretório de evidências](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/).

## core_f9.tsv

Modalidade:execução funcional VBA experimental no Excel real. [TSV original](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/core_f9.tsv).

| Cenário | Procedimento | Fixture | Resultado | Tratamento | Esperado | Observado |
|---|---|---|---|---|---|---|
| F9_LEASE_PRIMEIRO | F9_Core | ciclo.accdb | PASS | RECUSADO | Primeiro escritor recebe fence1 | 1 |
| F9_LEASE_SEGUNDO | F9_Core | ciclo.accdb | PASS | RECUSADO | Segundo nao adquire | Lease ativo |
| F9_SESSAO_LEITURA | F9_Core | ciclo.accdb | PASS | RECUSADO | Gravacao recusada pelo provedor | -2147467259:A operação deve usar uma consulta atualizável. |
| F9_GUID_LOCAL | F9_Core | ciclo.accdb | PASS | VERIFICADO | 100 GUIDs unicos gerados; sem garantia matematica universal | 100 |
| F9_ENTRADA_A1 | F9_Core | ciclo.accdb | PASS | CONFIRMADO | Entrada confirmada | EP_A1 |
| F9_ENTRADA_B1 | F9_Core | ciclo.accdb | PASS | CONFIRMADO | Entrada distinta | EP_B1 |
| F9_VENDA_A1 | F9_Core | ciclo.accdb | PASS | CONFIRMADO | Venda encerra episodio | CLOSED |
| F9_REENTRADA | F9_Core | ciclo.accdb | PASS | CONFIRMADO | Novo EP, mesmo VehicleID | VID_A |
| F9_TRANSFERENCIA | F9_Core | ciclo.accdb | PASS | CONFIRMADO | Mesmo EP / FILIAL | FILIAL |
| F9_DEVOLUCAO | F9_Core | ciclo.accdb | PASS | CONFIRMADO | Mesmo EP / MATRIZ | MATRIZ |
| F9_NOVA_TRANSFERENCIA | F9_Core | ciclo.accdb | PASS | CONFIRMADO | Mesmo EP versao4 | 4 |
| F9_IDEMPOTENCIA | F9_Core | ciclo.accdb | PASS | RECUSADO | Recibo original sem repetir efeitos | JA_CONCLUIDA |
| F9_PAYLOAD | F9_Core | ciclo.accdb | PASS | RECUSADO | Payload diferente recusado | RECUSADO_PAYLOAD |
| F9_VERSAO | F9_Core | ciclo.accdb | PASS | RECUSADO | Versao obsoleta recusada | RECUSADO_VERSAO |
| F9_COPIA_DIVERGENTE | F9_Core | ciclo.accdb | PASS | RECUSADO | BaseID divergente nao escreve | RECUSADO_BASE |
| F9_PK_VEICULO | F9_Core | ciclo.accdb | PASS | RECUSADO | PK recusa duplicado | -2147467259:As alterações solicitadas para a tabela não foram bem-sucedidas, já que criariam valores duplicados de índice, chave primária ou relação. Altere os dados no campo ou nos campos que contêm dados duplicados, remova o índice, ou redefina-o para permitir entradas duplicadas, e tente novamente. |
| F9_FK_EPISODIO | F9_Core | ciclo.accdb | PASS | RECUSADO | FK recusa orfao | -2147467259:Não é possível adicionar ou alterar registros, pois é necessário que eles tenham um registro relacionado na tabela 'Vehicles'. |
| F9_PK_PROJECAO | F9_Core | ciclo.accdb | PASS | RECUSADO | ProjectionID duplicado recusado | -2147467259:As alterações solicitadas para a tabela não foram bem-sucedidas, já que criariam valores duplicados de índice, chave primária ou relação. Altere os dados no campo ou nos campos que contêm dados duplicados, remova o índice, ou redefina-o para permitir entradas duplicadas, e tente novamente. |
| F9_PRECISAO_TEXTO | F9_Core | ciclo.accdb | PASS | RECUPERADO | Currency com quatro casas e chave String de 44 caracteres | 1000,1234;String |
| F9_M3_AMBIGUO | F9_Core | ciclo.accdb | PASS | PENDENTE_CONCILIACAO | Legado ambiguo nao vinculado | AMBIGUO;NULL |
| F9_SAVE_RECUSADO | F9_Core | ciclo.accdb | PASS | PENDENTE_CONCILIACAO | Save recusado nao desfaz commit / recibo duravel no banco | Saved=Falso;BeforeSave=Verdadeiro |
| F9_DISCO_VS_MEMORIA | F9_Core | ciclo.accdb | PASS | RECUPERADO | Alteracao cancelada ausente em disco |  |
| F9_RECONSTRUCAO | F9_Core | ciclo.accdb | PASS | RECUPERADO | 7 projecoes/2 ativos reconstruidos | 2 |
| F9_REORDENACAO | F9_Core | ciclo.accdb | PASS | RECUPERADO | Contagem por PID mantida apos sort | 2 |
| F9_REABERTURA_CONEXAO | F9_Core | ciclo.accdb | PASS | RECUPERADO | 7 recibos e registros verificados apos fechar/reabrir | 7 |
| F9_FENCING | F9_Core | ciclo.accdb | PASS | RECUSADO | Escritor antigo perde autoridade | RECUSADO_FENCE |

## falhas_f9.tsv

Modalidade:execução funcional VBA experimental no Excel real. [TSV original](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/falhas_f9.tsv).

| Cenário | Procedimento | Fixture | Resultado | Tratamento | Esperado | Observado |
|---|---|---|---|---|---|---|
| F9_FALHA_1 | F9_Faults | falha_1.accdb | PASS | RECUPERADO | Efeitos transacionais revertidos; intencao pendente conservada | PENDENTE_ERRO:7001:Falha apos movimento;IGUAL=Verdadeiro |
| F9_FALHA_REPLAY_1 | F9_Faults | falha_1.accdb | PASS | PENDENTE_CONCILIACAO | Novo ID nao contorna intencao | PENDENTE_CONCILIACAO |
| F9_FALHA_2 | F9_Faults | falha_2.accdb | PASS | RECUPERADO | Efeitos transacionais revertidos; intencao pendente conservada | PENDENTE_ERRO:7002:Falha apos projecao;IGUAL=Verdadeiro |
| F9_FALHA_REPLAY_2 | F9_Faults | falha_2.accdb | PASS | PENDENTE_CONCILIACAO | Novo ID nao contorna intencao | PENDENTE_CONCILIACAO |
| F9_FALHA_3 | F9_Faults | falha_3.accdb | PASS | RECUPERADO | Efeitos transacionais revertidos; intencao pendente conservada | PENDENTE_ERRO:7003:Falha durante baixa logica;IGUAL=Verdadeiro |
| F9_FALHA_REPLAY_3 | F9_Faults | falha_3.accdb | PASS | PENDENTE_CONCILIACAO | Novo ID nao contorna intencao | PENDENTE_CONCILIACAO |
| F9_FALHA_4 | F9_Faults | falha_4.accdb | PASS | RECUPERADO | Efeitos transacionais revertidos; intencao pendente conservada | PENDENTE_ERRO:-2147467259:Não é possível adicionar ou alterar registros, pois é necessário que eles tenham um registro relacionado na tabela 'Operations'.;IGUAL=Verdadeiro |
| F9_FALHA_REPLAY_4 | F9_Faults | falha_4.accdb | PASS | PENDENTE_CONCILIACAO | Novo ID nao contorna intencao | PENDENTE_CONCILIACAO |
| F9_FALHA_5 | F9_Faults | falha_5.accdb | PASS | RECUPERADO | Efeitos transacionais revertidos; intencao pendente conservada | PENDENTE_ERRO:7005:Antes do recibo;IGUAL=Verdadeiro |
| F9_FALHA_REPLAY_5 | F9_Faults | falha_5.accdb | PASS | PENDENTE_CONCILIACAO | Novo ID nao contorna intencao | PENDENTE_CONCILIACAO |
| F9_FALHA_6 | F9_Faults | falha_6.accdb | PASS | RECUPERADO | Efeitos transacionais revertidos; intencao pendente conservada | PENDENTE_ERRO:7006:Antes do commit;IGUAL=Verdadeiro |
| F9_FALHA_REPLAY_6 | F9_Faults | falha_6.accdb | PASS | PENDENTE_CONCILIACAO | Novo ID nao contorna intencao | PENDENTE_CONCILIACAO |
| F9_FALHA_7 | F9_Faults | falha_7.accdb | PASS | RECUPERADO | Efeitos transacionais revertidos; intencao pendente conservada | PENDENTE_ERRO:3246:Um objeto de conexão não pode ser fechado explicitamente durante uma transação.;IGUAL=Verdadeiro |
| F9_FALHA_REPLAY_7 | F9_Faults | falha_7.accdb | PASS | PENDENTE_CONCILIACAO | Novo ID nao contorna intencao | PENDENTE_CONCILIACAO |
| F9_APOS_COMMIT | F9_Faults | falha_8.accdb | PASS | PENDENTE_CONCILIACAO | Commit e recibo persistem; Excel ainda desatualizado | COMMIT_CONFIRMADO_PROJECAO_PENDENTE |
| F9_ORIGEM_ALTERADA | F9_Faults | origem_alterada.accdb | PASS | PENDENTE_CONCILIACAO | Origem mudou apos preparar; sem efeitos novos | PENDENTE_ORIGEM_ALTERADA |
| F9_HISTORICO_ALTERADO | F9_Faults | historico_alterado.accdb | PASS | PENDENTE_CONCILIACAO | Recibo antigo nao valida historico alterado | PENDENTE_HISTORICO |
| F9_LEASE_ORFAO | F9_Faults | lease_orfao.accdb | PASS | PENDENTE_CONCILIACAO | Conexao encerrada nao libera lease persistido; sem TTL | 07A302A5-BD12-471E-B04A-2F51BA8730BA |

## projecao_f9.tsv

Modalidade:execução funcional VBA experimental no Excel real. [TSV original](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/projecao_f9.tsv).

| Cenário | Procedimento | Fixture | Resultado | Tratamento | Esperado | Observado |
|---|---|---|---|---|---|---|
| F9_PROJECAO_INTERROMPIDA | F9_ProjectionFaults | falha_8.accdb | PASS | PENDENTE_CONCILIACAO | Autoridade integra; projecao parcial nao autoriza | 7100:Falha de projecao apos cabecalho |
| F9_RECONSTRUIR_APOS_COMMIT | F9_ProjectionFaults | falha_8.accdb | PASS | RECUPERADO | 8 projecoes e2 ativos apos reconstruir | 8;2 |
| F9_SORT_VALORES_DISTINTOS | F9_ProjectionFaults | falha_8.accdb | PASS | RECUPERADO | PID/EP/VID/valores/formulas discriminantes preservados | 8 |
| F9_TOKEN_LISTBOX_OBSOLETO | F9_ProjectionFaults | falha_8.accdb | PASS | RECUSADO | Tokenversao3 nao autoriza4 | VERSAO_OBSOLETA |
| F9_SAVE_CAMINHO_INVALIDO | F9_ProjectionFaults | falha_8.accdb | PASS | PENDENTE_CONCILIACAO | Erro real de Save nao desfaz autoridade | 1004:O Microsoft Excel não pode acessar o arquivo 'C:\Users\MARCOS~1\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase9_20261102_142753Z\PAI_INEXISTENTE\4DEB0E00'. Há várias razões possíveis:<CR><LF><CR><LF>• O nome do arquivo ou caminho não existe.<CR><LF>• O arquivo está sendo usado por outro programa.<CR><LF>• A pasta de trabalho que você está tentando salvar tem o mesmo nome de outra <CR><LF>pasta de trabalho aberta no momento. |

## complemento_f9.tsv

Modalidade:execução funcional VBA experimental no Excel real. [TSV original](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/complemento_f9.tsv).

| Cenário | Procedimento | Fixture | Resultado | Tratamento | Esperado | Observado |
|---|---|---|---|---|---|---|
| F9_M0_M4_PROVENIENCIA | F9_Complemento | migracao.accdb | PASS | PENDENTE_CONCILIACAO | Duas correspondencias explicitas euma ambigua sem vinculo | 2 CONFIRMADO;1 AMBIGUO;SHA256=38de109d962cb44f45cab4934ea957be1480c362cad00da4b7d9538ddddc2218 |
| F9_FK_COMPOSTA_EP_MOV | F9_Complemento | migracao.accdb | PASS | RECUSADO | Movimento de A1 nao pode projetar episodio B1 | -2147467259:Não é possível adicionar ou alterar registros, pois é necessário que eles tenham um registro relacionado na tabela 'Movements'. |
| F9_PROJECAO_REVISAO_DIVERGENTE | F9_Complemento | Projecao_Inspecao_Metadata.xlsm | PASS | RECUSADO | Revisao de outra autoridade nao autoriza | RECUSADO_REVISAO |
| F9_METADATA_APOS_SORT_ANTES | F9_Complemento | Projecao_Inspecao_Metadata.xlsm | FAIL | FAIL | Ignore/formato preservados alem da associacao | RECUSADO_IGNORE |
| F9_METADATA_RECONSTRUIDO | F9_Complemento | Projecao_Inspecao_Metadata.xlsm | PASS | RECUPERADO | Metadado reparado; identidade e valores conferidos | VERIFICADA |
| F9_EDICAO_PROJECAO_RECUSADA | F9_Complemento | Projecao_Inspecao_Metadata.xlsm | PASS | RECUSADO | Edicao local divergente nao autoriza | RECUSADO_CONTEUDO |
| F9_COMMIT_CONEXAO_FECHADA | F9_Complemento | falha_8.accdb | PASS | RECUSADO | API recusa commit sem conexao; I/O commit nao exercitado | 3704:Operação não permitida quando o objeto está fechado. |

## sessao_fria_f9.tsv

Modalidade:execução funcional VBA experimental no Excel real. [TSV original](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/sessao_fria_f9.tsv).

| Cenário | Procedimento | Fixture | Resultado | Tratamento | Esperado | Observado |
|---|---|---|---|---|---|---|
| F9_FRIO_CICLO | F9_Cold | ciclo.accdb | PASS | RECUPERADO | 7 efeitos/recibos persistidos em novo processo | 7 |
| F9_FRIO_EPISODIOS | F9_Cold | ciclo.accdb | PASS | RECUPERADO | A1 encerrado; A2 ativo4; mesmo VID | CLOSED/ACTIVE/4 |
| F9_FRIO_BENCH_100 | F9_Cold | bench_100.accdb | PASS | RECUPERADO | Reconstruir do banco frio sem usar XLSM salvo | COUNT=100;SECONDS=0,481664000020828;FMT=@ |
| F9_FRIO_BENCH_1000 | F9_Cold | bench_1000.accdb | PASS | RECUPERADO | Reconstruir do banco frio sem usar XLSM salvo | COUNT=1000;SECONDS=0,61371129998588;FMT=@ |
| F9_FRIO_BENCH_5000 | F9_Cold | bench_5000.accdb | PASS | RECUPERADO | Reconstruir do banco frio sem usar XLSM salvo | COUNT=5000;SECONDS=1,13452700001653;FMT=@ |
| F9_FRIO_BENCH_10000 | F9_Cold | bench_10000.accdb | PASS | RECUPERADO | Reconstruir do banco frio sem usar XLSM salvo | COUNT=10000;SECONDS=1,91194749998976;FMT=@ |
| F9_FRIO_BACKUP | F9_Cold | backup_ciclo.accdb | PASS | RECUPERADO | Backup fechado reaberto integro | 7 |
| F9_FRIO_ORFAO | F9_Cold | lease_orfao.accdb | PASS | PENDENTE_CONCILIACAO | Lease orfao persiste; sem remocao por tempo | ACTIVE |
| F9_FRIO_INTENCAO | F9_Cold | falha_1.accdb | PASS | PENDENTE_CONCILIACAO | Intencao confirmada antes de falha permanece | PREPARED/PENDING |

## Concorrência COM em processos independentes

Não são testes VBA nem simulação; Excel manteve a conexão/transação e outro processo ADO executou a tentativa.

| Cenário | Resultado | Processo/esperado/observado | Evidência |
|---|---|---|---|
| F9_PROCESSO_SEGUNDO | PASS | Affected=0;Fence=1;PID=26732 | [JSON](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/concorrencia_segundo_processo.json) |
| F9_PROCESSO_LOCK_REAL | PASS | Não foi possível atualizar; atualmente bloqueado. | [JSON](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/concorrencia_lock_real_corrigido.json) |
| F9_RECONEXAO_APOS_LIBERACAO | PASS | Affected=1;Fence=2;PID=42856 | [JSON](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/concorrencia_reconexao.json) |

O primeiro concorrencia_lock_real.json é INVALIDO: conexão não abriu e o auxiliar rotulou erroneamente PASS para objeto fechado. Preservado e excluído; reteste requer Opened=True e erro de bloqueio.

## Benchmark comparativo

Uma execução por volume. Fonte:bancos sintéticos completos de relações essenciais, sem replay dos gravadores principais. Seed em lote prepara Vehicles/Episodes/Operations/Movements/Projections diretamente; não mede cada operação normal de negócio. Commit_META mede um UPDATE de revisão, não a duração de todo movimento. Carga da ListBoxLista real (14 colunas) sem clicar/renomear eventos privados. Fórmulas nativas3n já foram calculadas na reconstrução; CALC_ALL é recálculo explícito subsequente, não cálculo frio do catálogo operacional.

| n | Seed | Query lote | Validação lote | Construir índice | Candidato200 | Verificada200 | Lista completa | ListBox | Reconstrução | Recalc posterior | Save | Reabrir | Commit_META | Total |
|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| 100 | 0.0952 | 0.0060 | 0.0042 | 0.0008 | 0.0002 | 1.6615 | 0.0089 | 0.0045 | 0.0492 | 0.0003 | 0.0831 | 0.4307 | 0.0041 | 2.9581 |
| 1000 | 0.7986 | 0.0142 | 0.0364 | 0.0020 | 0.0005 | 1.6650 | 0.0497 | 0.0064 | 0.1276 | 0.0004 | 0.1120 | 0.4423 | 0.0045 | 3.8594 |
| 5000 | 3.9697 | 0.0490 | 0.1717 | 0.0077 | 0.0006 | 1.6611 | 0.2297 | 0.0211 | 0.5430 | 0.0008 | 0.1626 | 0.5300 | 0.0040 | 7.9854 |
| 10000 | 7.9105 | 0.1590 | 0.3400 | 0.0155 | 0.0007 | 1.7162 | 0.5041 | 0.0365 | 0.9963 | 0.0012 | 0.2426 | 0.6096 | 0.0041 | 13.1593 |

Todos os quatro retornaram PASS de quantidade/contagem ativa. Tempos em segundos. F8 lista de 10 mil: 9,7135 s versus F9 0,5041 s (~19,3x neste caminho); F9 consulta verificada200 1,7162s versus F8 0,8083s: não houve melhora em todos os caminhos. F9 usa uma leitura relacional em lote e valida tipos/canônicos/FKs; F8 validava células/formatos/Ignore individualmente. Apresentação/Ignore da projeção é contrato separado; não afirmar benchmark de garantias idênticas ou ganho universal. Índice candidato200 não é consulta autorizada.

Memória do PID45652:antes1000≈529,7MiB e depois10000≈540,6MiB de working set; não medidos pico ou vazamento prolongado;100 sem amostra própria. Arquivos mem_antes1000/mem_depois1000/mem_depois5000/mem_depois10000.json guardam valores exatos. Sessão fria reconstruiu10mil em1,9119s.

## Falhas, limites e referências preservadas

F9_METADATA_APOS_SORT_ANTES permanece FAIL de Ignore; F9_METADATA_RECONSTRUIDO é prova corretiva distinta, sem apagar o resultado. Guarda complementar falhou antes dos testes por procurar marcador no nome8.3; correção mínima manteve limite TEMP e exclusão OneDrive. Primeiras importações duplicadas foram removidas após checkpoint;62 componentes no núcleo,63 após complemento.

| Referência / cenário | Classificação nesta fase |
|---|---|
| V_VALIDO,T_VALIDO,D_VALIDO,D_FALHA_BAIXA,T_DUP_DESTINO | FAIL F5 preservados; eventos reais NÃO EXECUTADOS em F9 |
| C_FALHA_BAIXA,C_FALHA_COMPENSACAO | FAIL integral F6 preservados; rollback ACE não corrige o VBA anterior |
| Falha real de I/O durante CommitTrans, energia/crash/disco cheio, RollbackTrans falhar | NÃO EXECUTADO; erro3246 foi Close anterior, erro3704 foi API em conexão fechada |
| Snapshot de projeção durante escrita concorrente/revogação durante Prepare | NÃO EXECUTADO; lacuna do candidato |
| Diário externo truncado/órfão novo | NÃO EXECUTADO; F7/F8 preservados. Diário F9 é tabela transacional; falha FK e intenção fria foram executadas |
| Ack persistente de projeção após Save | NÃO IMPLEMENTADO/NÃO EXECUTADO; recibo de negócio já precede Excel |
| Fiscal/Cancelado/XML/DETRAN/Selenium | NÃO EXECUTADO e sem alteração de regras |
| Portão 2, dados reais, troca de7 Delete/gravadores/fórmulas operacionais | NÃO AUTORIZADO/NÃO EXECUTADO |

[Durabilidade](./TESTES_DURABILIDADE_FASE9.md), [concorrência](./TESTES_CONCORRENCIA_FASE9.md), [reconstrução](./PROTOCOLO_RECONSTRUCAO_EXCEL_FASE9.md).
