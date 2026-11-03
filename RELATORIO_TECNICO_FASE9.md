# Relatório técnico — Fase 9

**Laboratório transacional executado; Portão 2 NÃO aprovado; sistema principal NÃO HOMOLOGADO.** ACE 16 disponível demonstrou transações internas, idempotência, controle de versões, contenção entre processos e reconstrução Excel em sessão fria. Isso não resolve os sete FAIL oficiais anteriores nem autoriza integrar os gravadores.

Execução 20261102_142753Z, [evidências](./_auditoria/fase9_20261102_142753Z/). Doze documentos obrigatórios lidos integralmente, todos disponíveis. Entrega definitiva F8 R2 conferida pelo SHA256 a64b1a81bfba47000d7f69f353ef7fa697746f78bded65f53685c5dee4a0fa3d; 2.268 itens do manifesto final sem divergência. 43 backups novos verificados antes de alterações. Originais/FRM/FRX/CodeNames/eventos e 60 componentes herdados preservados. Não foi reiniciada auditoria geral.

## Entrega e execução

[Fixture experimental final](./_auditoria/fase9_20261102_142753Z/entrega/Fixture_Persistencia_Fase9_EXPERIMENTAL_NAO_HOMOLOGADA.xlsm), SHA256 **fbb127500d504e6ac9bf3f857dfc936f1f0e43449b3549a2473fff2a5b56386d**, 63 componentes/17 abas. Três componentes novos somente no laboratório: clsSaveFase9,modPersistenciaFase9,modVerificacaoFase9. [Fontes finais realmente incorporados](./_auditoria/fase9_20261102_142753Z/prototipos/revisao_final/). Os bancos, projeções e testes são sintéticos, arquivados após execução exclusivamente em TEMP fora do OneDrive.

Importação/compilação pelo VBE nativo; menu Compilar desabilitado ao final e execução efetiva das macros, sem VBProjectCOM. Primeiras duplicatas experimentais removidas depois de checkpoint; correção pontual da guarda do caminho curto no complemento. Duas instâncias próprias Excel PID 45652/PID 48356 fechadas normalmente; janela original 8589906 preservada. Estados finais True/True/-4105/True para EnableEvents/ScreenUpdating/Calculation/DisplayAlerts, com 1 workbook de teste em cada encerramento. Nenhuma instalação, configuração de segurança, consulta fiscal, commit/push/publicação ou operação real. [Sessões transcritas da interface](./_auditoria/fase9_20261102_142753Z/evidencias/sessoes-office-fase9.json).

## Evidência principal

65 asserções funcionais VBA: 64 PASS e 1 FAIL de metadado preservado, com prova corretiva distinta PASS. Quatro benchmarks VBA PASS e três experiências COM em processos independentes PASS, separados. Não somar sondagens/hashes como testes de negócio nem o PASS inválido do primeiro auxiliar de contenção. [Matriz completa](./MATRIZ_REGRESSAO_FASE9.md), [consolidado estruturado](./_auditoria/fase9_20261102_142753Z/resultados-consolidados-fase9.json).

O ciclo Entrada→Venda→Reentrada→Transferência→Devolução→Transferência gerou 7 movimentos/recibos, mesmo veículo em dois episódios, notas repetidas e chave String de 44 caracteres com zero inicial. Episódio A1 encerrado e A2 ativo; custódia global é hipótese provisória . Sete falhas pré-commit restauraram assinatura dos dados de negócio; intenção preparada em transação anterior permaneceu pendente e bloqueou ID novo. Falha real de FK no ChangeLog reverteu a operação. F9_FALHA_7 parou no Close com 3246, sem testar CommitTrans; commit em conexão fechada 3704 é recusa de API, não falha de disco.

Após commit, fatos/recibo persistem independentemente do Excel. Interrupção da projeção, Save cancelado por evento e SaveAs com pai inexistente 1004 não desfizeram a autoridade. Reconstrução por IDs passou; sessão nova recuperou 7 registros/recibos e reconstruiu quatro volumes. Backup de banco fechado reabriu íntegro. Não existe transação única ACCDB/XLSM, nem prova de energia/crash/disco cheio/rollback falhar.

Dois processos testaram exclusividade: segunda aquisição recusada; transação aberta impediu UPDATE conflitante com erro “atualmente bloqueado”; após rollback/liberação, nova sessão recebeu Fence2. Lease órfão permaneceu ACTIVE em sessão fria e exigiu conciliação. Fencing recusou proprietário antigo no teste controlado. Não demonstra todo interleaving de preparação/revogação ou snapshot sob escrita concorrente.

## FAIL novo e riscos que não foram escondidos

F9_METADATA_APOS_SORT_ANTES:Sort nativo preservou os vínculos/valores conferidos, mas perdeu Ignore. Verificador recusou; reconstrução do metadado com nova validação passou. O FAIL permanece registrado. Portanto não aprovar classificação manual irrestrita ou apenas olhar ausência de#REF!. Projecao_Inspecao_Metadata foi depois salva com adulteração intencional para teste negativo; não é referência íntegra.

O esquema 1 usa FKs separadas; vínculo composto EP/Movement foi reforçado/testado apenas em migracao.accdb. SchemaNo 2 uniforme, imutabilidade/constraints completas, auditoria de todos os replays, validação de datas/documentos arbitrários, hash de payload/recibo criptográfico, ack persistente de projeção, proteção contra SQL fora do protocolo e snapshot concorrente continuam pendentes. Os leitores operacionais e57 procedimentos F7 não foram migrados. Precisão monetária no Excel foi provada apenas na massa sintética.

## Desempenho e decisão técnica

Em10mil:listagem validada em lote0,5041s, preenchimento ListBox real0,0365s, reconstrução0,9963s, Save0,2426s, reabertura0,6096s; sessão fria reconstrução1,9119s. F8 lista: 9,7135 s era por célula; a nova comparação é de caminho relacional em lote, com apresentação derivada separadamente.200 consultas verificadas F9 demoraram1,7162s, mais que0,8083s na F8. Não confundir candidato de índice com autorização e não anunciar aceleração universal. Memória working set≈529,7→540,6MiB na janela1000–10000, sem medida de pico ou teste prolongado.

Recomendar **ACE 16/ADODB-ADOX x64 como candidato experimental disponível**, não como escolha definitiva do produto. DLL 16.0.20527.20016 do Office Click-to-Run foi realmente carregada; ACE 12 também abriu; Jet 4 não abriu em x64. DLL SQLite e driver SQL não foram tomados como motor pronto. Licença/implantação e ambiente destino continuam decisão P2-05. O redistribuível2016 sem suporte não deve ser instalado como atalho. [Comparação, suporte e fontes oficiais](./COMPARATIVO_ARMAZENAMENTO_FASE9.md).

## Próximo portão

[Esquema e lacunas](./ESQUEMA_TRANSACIONAL_FASE9.md), [durabilidade](./TESTES_DURABILIDADE_FASE9.md), [concorrência](./TESTES_CONCORRENCIA_FASE9.md), [reconstrução](./PROTOCOLO_RECONSTRUCAO_EXCEL_FASE9.md), [ensaio M0–M4](./PLANO_MIGRACAO_M0_M4_FASE9.md) e [decisões](./DECISOES_PENDENTES_FASE9.md) formam o pacote de preparação.

Antes de integrar, consolidar candidato/schema com constraints, fencing da preparação, snapshot, verificador integral de projeções/recibos e protocolo de conciliação aprovado; executar falhas reais de fronteira e restauração específica. P2-01 aP2-11 não foram aprovados. V_VALIDO,T_VALIDO,D_VALIDO,D_FALHA_BAIXA,T_DUP_DESTINO,C_FALHA_BAIXA,C_FALHA_COMPENSACAO permanecem FAIL das implementações anteriores . Sete Delete, lógica fiscal e arquivo operacional permanecem intactos.

Restauração:originais não precisam ser restaurados. Para repetir, copiar mestre final/insumos para TEMP novo, manter nome Fixture_Persistencia_Fase9_Mestre.xlsm e marcador exigido ou revisar somente guardas experimentais em nova cópia, compilar pelo VBE e usar bancos/resultados com nomes novos. Nunca sobrescrever evidências. Backup documental e checkpoint preservam reversibilidade; após fatos novos em eventual implantação, conciliar em vez de restaurar cegamente arquivos antigos.

## Integridade final e limites do ambiente

Rechecados 39 originais e 43 backups, todos íntegros. Dos 2.268 caminhos da linha de base F8, 2.261 continuam idênticos e sete receberam somente os acréscimos autorizados de documentação/.gitignore, com prefixos antigos preservados. Nenhum arquivo de dados/fonte anterior foi sobrescrito. Os 17 CodeNames e os 60 componentes herdados conferem. [Conferência detalhada](./_auditoria/fase9_20261102_142753Z/verificacao-final-fase9.json), [manifesto final](./_auditoria/fase9_20261102_142753Z/manifesto-final-fase9.json) e [digest](./_auditoria/fase9_20261102_142753Z/manifesto-final-fase9.sha256).

O executável node não está no PATH do terminal; a análise/hashes foram executados no runtime Node REPL já disponível ao agente, sem instalar ferramentas. O bloqueio histórico de scripts PowerShell não foi repetido ou contornado; as sondagens COM desta fase foram comandos locais autorizados, sem VBProject. Inspeção estática/Problems não substitui a compilação e a execução VBA evidenciadas pela interface.
