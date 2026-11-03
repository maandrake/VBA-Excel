# Matriz de regressão — Fase 10

## Resultado por modalidade

- **121 asserções VBA únicas:121 PASS,0 FAIL na última execução de cada caso.** Execução real de laboratório compilado. Não inclui benchmarks.
- Benchmark VBA:12 amostras PASS (três por volume); volume100 repetido após a guarda final. Volumes maiores conservam a revisão indicada em cada medida.
- Operação COM entre processos: quatro experiências de concorrência PASS e seis verificações frias PASS. Não substituem execução funcional VBA.
- Sessão fria VBA: seis verificações BLOQUEADAS por política do Office; o procedimento F10_Frio não chegou a produzir log de testes.
- Histórico de tentativas reprovadas preservado. Não somar seus resultados com os finais como se fossem cenários únicos adicionais.

Compilação real realizada pelo Editor VBA antes da execução das suítes e após a correção final. Conferência estática de66 componentes,17 CodeNames e63 corpos herdados não é compilação. Eventos operacionais principais não foram alterados/testados nesta fase.

## Dados, origem e estados

Mestre: [entrega experimental](./_auditoria/fase10_20261102_155920Z/entrega/Fixture_Integridade_Fase10_EXPERIMENTAL_NAO_HOMOLOGADA.xlsm). Office executado exclusivamente em TEMP. ACCDBs e arquivos exatos de cada cenário estão em [execucao_temp](./_auditoria/fase10_20261102_155920Z/execucao_temp/). Fontes:[prototipos](./_auditoria/fase10_20261102_155920Z/prototipos/).

Ciclo de referência: veículos fictícios F10A001/F10B001, três episódios (dois do mesmo VehicleID), sete movimentos/recibos, dois ativos, NF-e88001 repetida em contextos distintos, chave01234567890123456789012345678901234567890123, data2026-11-02 e valores discriminantes de quatro casas. Bases de falha são cópias independentes do ciclo fechado. Antes/depois dos efeitos: arquivos tx_*_before.tsv/after.tsv; demais estados: ciclo_f10_final, migrado_final, constraints_final, intencao_*_final e projecao_autoridade_final, além das próprias bases preservadas. Snapshot10 não inclui a tabela adicional ProjectionAcks: os ACKs permanecem no ACCDB, no arquivo confirmado e nas consultas frias JSON.

Cada linha abaixo identifica procedimento, fixture, expectativa e observação. Entradas e passos exatos estão no procedimento correspondente; [registros estruturados](./_auditoria/fase10_20261102_155920Z/evidencias/resultados_vba_finais.json) conservam origem e limitações. Dados de contrato (SHAabc, datas inválidas, chaves curtas, valores extremos, IDs ausentes) são separados dos dados do ciclo nos próprios testes F10_Schema.

|ID|Procedimento / fixture|Esperado|Observado / tratamento|Status|
|---|---|---|---|---|
|F10_SHA256|F10_Schema / CryptoAPI|SHA256 UTF8 conhecido|ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad / VERIFICADO|PASS|
|F10_SCHEMA_EQUIVALENTE|F10_Schema / novo.accdb;legado_f9.accdb|Novo e migrado possuem esquema/regras identicos|8698a2b7f44693510141e9fe75fd123fe41b4d99f7c60a03c1f9d2a99625d26b;8698a2b7f44693510141e9fe75fd123fe41b4d99f7c60a03c1f9d2a99625d26b / VERIFICADO|PASS|
|F10_MIGRADO_AUDIT|F10_Schema / legado_f9.accdb|7 movimentos/recibos auditados|OK / VERIFICADO|PASS|
|F10_MIGRACAO_REPLAY|F10_Schema / legado_f9.accdb|Reaplicar nao duplica ou altera fatos|OK / RECUSADO|PASS|
|F10_DATAS|F10_Schema / Contrato|Calendario civil ISO estrito|2024-02-29;2026-02-29;02/11/2026;2026-13-01 / RECUSADO|PASS|
|F10_MONEY|F10_Schema / Contrato|Quatro casas exatas; sem arredondar quinta casa ou overflow|1000.1234;1.12345;9999999999999999 / RECUSADO|PASS|
|F10_ID|F10_Schema / Contrato|Novos IDs GUID; legado so migracao|GUID;vazio;PLACA123 / RECUSADO|PASS|
|F10_FK_EP_MOV|F10_Schema / legado_f9.accdb|Nao associar B1 ao movimento A1|-2147467259:O registro não pode ser excluído ou alterado porque a tabela 'Receipts' inclui registros relacionados a ele. / RECUSADO|PASS|
|F10_FK_MOV_OP|F10_Schema / legado_f9.accdb|Episodio do movimento coincide com operacao|-2147467259:O registro não pode ser excluído ou alterado porque a tabela 'Projections' inclui registros relacionados a ele. / RECUSADO|PASS|
|F10_FK_CURRENT|F10_Schema / legado_f9.accdb|CurrentPID de outro episodio recusado|-2147467259:Não é possível adicionar ou alterar registros, pois é necessário que eles tenham um registro relacionado na tabela 'Projections'. / RECUSADO|PASS|
|F10_FK_VEHICLE_HISTORY|F10_Schema / legado_f9.accdb|Historico impede trocar veiculo do episodio|-2147467259:O registro não pode ser excluído ou alterado porque a tabela 'Receipts' inclui registros relacionados a ele. / RECUSADO|PASS|
|F10_DELETE_PARENT|F10_Schema / legado_f9.accdb|Nao excluir episodio com fatos|-2147467259:O registro não pode ser excluído ou alterado porque a tabela 'Operations' inclui registros relacionados a ele. / RECUSADO|PASS|
|F10_STATE_RULE|F10_Schema / legado_f9.accdb|Estado fora do dominio recusado no motor|-2147467259:Contrato sintetico F10: Episodes / RECUSADO|PASS|
|F10_KEY_SHORT|F10_Schema / legado_f9.accdb|Chave curta recusada no motor|-2147467259:Contrato sintetico F10: Projections / RECUSADO|PASS|
|F10_EMPTY_ID_ENGINE|F10_Schema / legado_f9.accdb|ID vazio recusado|-2147467259:Contrato sintetico F10: Vehicles / RECUSADO|PASS|
|F10_ORPHAN|F10_Schema / legado_f9.accdb|FK recusa veiculo inexistente|-2147467259:Não é possível adicionar ou alterar registros, pois é necessário que eles tenham um registro relacionado na tabela 'Vehicles'. / RECUSADO|PASS|
|F10_ENTRADA|F10_Core / ciclo_f10.accdb|Entrada A1|CONFIRMADA / CONFIRMADO|PASS|
|F10_ENTRADA_B|F10_Core / ciclo_f10.accdb|Entrada B|CONFIRMADA / CONFIRMADO|PASS|
|F10_VENDA|F10_Core / ciclo_f10.accdb|Encerra A1|CONFIRMADA / CONFIRMADO|PASS|
|F10_REENTRADA|F10_Core / ciclo_f10.accdb|Mesmo VID novo episodio|CONFIRMADA / CONFIRMADO|PASS|
|F10_TRANSFERENCIA|F10_Core / ciclo_f10.accdb|Preserva episodio; filial|CONFIRMADA / CONFIRMADO|PASS|
|F10_DEVOLUCAO|F10_Core / ciclo_f10.accdb|Mesmo episodio; matriz|CONFIRMADA / CONFIRMADO|PASS|
|F10_TRANSFERENCIA_2|F10_Core / ciclo_f10.accdb|Custodia filial; versao4|CONFIRMADA / CONFIRMADO|PASS|
|F10_REPLAY_AUDITADO|F10_Core / ciclo_f10.accdb|Mesmo recibo; nova tentativa auditada sem efeitos|JA_CONCLUIDA / RECUSADO|PASS|
|F10_REPLAY_PAYLOAD|F10_Core / ciclo_f10.accdb|Payload divergente auditado|RECUSADO_PAYLOAD / RECUSADO|PASS|
|F10_VERSAO_OBSOLETA|F10_Core / ciclo_f10.accdb|ExpectedVersion3 nao autoriza4|RECUSADO_VERSAO / RECUSADO|PASS|
|F10_BASE_DIVERGENTE|F10_Core / ciclo_f10.accdb|Outra BaseID nao escreve|RECUSADO_BASE / RECUSADO|PASS|
|F10_CICLO_AUDIT|F10_Core / ciclo_f10.accdb|7 movimentos/recibos;2 ativos;trilha integra|OK / VERIFICADO|PASS|
|F10_LISTBOX_OBSOLETA|F10_Core / ciclo_f10.accdb|Token antigo re-resolvido|VERSAO_OBSOLETA / RECUSADO|PASS|
|F10_CHAVE_CURRENCY|F10_Core / ciclo_f10.accdb|Zero inicial/44 e escala4|01234567890123456789012345678901234567890123 / VERIFICADO|PASS|
|F10_BACKUP_HASH|F10_Core / backup_ciclo_f10.accdb|Banco fechado copiado SHA256 igual|7325cdb9a19a24510e23fa4f6ea3af5baea12f1583606e4b914675810261e106 / VERIFICADO|PASS|
|F10_TX_-1|F10_Transacoes / transacao_antes1.accdb|Efeitos iguais; intencao continua pendente; nao e falha fisica|RECUSADO_ANTES_BEGIN;AUDIT=OK / RECUSADO|PASS|
|F10_TX_NOVO_ID_-1|F10_Transacoes / transacao_antes1.accdb|Novo ID nao contorna pendencia|PENDENTE_CONCILIACAO / PENDENTE_CONCILIACAO|PASS|
|F10_TX_-2|F10_Transacoes / transacao_antes2.accdb|Efeitos iguais; intencao continua pendente; nao e falha fisica|PENDENTE_ERRO:7012:Depois de BeginTrans;AUDIT=OK / RECUPERADO|PASS|
|F10_TX_NOVO_ID_-2|F10_Transacoes / transacao_antes2.accdb|Novo ID nao contorna pendencia|PENDENTE_CONCILIACAO / PENDENTE_CONCILIACAO|PASS|
|F10_TX_1|F10_Transacoes / transacao_1.accdb|Efeitos iguais; intencao continua pendente; nao e falha fisica|PENDENTE_ERRO:7001:Falha apos movimento;AUDIT=OK / RECUPERADO|PASS|
|F10_TX_NOVO_ID_1|F10_Transacoes / transacao_1.accdb|Novo ID nao contorna pendencia|PENDENTE_CONCILIACAO / PENDENTE_CONCILIACAO|PASS|
|F10_TX_2|F10_Transacoes / transacao_2.accdb|Efeitos iguais; intencao continua pendente; nao e falha fisica|PENDENTE_ERRO:7002:Falha apos projecao;AUDIT=OK / RECUPERADO|PASS|
|F10_TX_NOVO_ID_2|F10_Transacoes / transacao_2.accdb|Novo ID nao contorna pendencia|PENDENTE_CONCILIACAO / PENDENTE_CONCILIACAO|PASS|
|F10_TX_3|F10_Transacoes / transacao_3.accdb|Efeitos iguais; intencao continua pendente; nao e falha fisica|PENDENTE_ERRO:7003:Falha durante baixa logica;AUDIT=OK / RECUPERADO|PASS|
|F10_TX_NOVO_ID_3|F10_Transacoes / transacao_3.accdb|Novo ID nao contorna pendencia|PENDENTE_CONCILIACAO / PENDENTE_CONCILIACAO|PASS|
|F10_TX_4|F10_Transacoes / transacao_4.accdb|Efeitos iguais; intencao continua pendente; nao e falha fisica|PENDENTE_ERRO:-2147467259:Você deve digitar um valor no campo 'ChangeLog.BeforeCanon'.;AUDIT=OK / RECUPERADO|PASS|
|F10_TX_NOVO_ID_4|F10_Transacoes / transacao_4.accdb|Novo ID nao contorna pendencia|PENDENTE_CONCILIACAO / PENDENTE_CONCILIACAO|PASS|
|F10_TX_5|F10_Transacoes / transacao_5.accdb|Efeitos iguais; intencao continua pendente; nao e falha fisica|PENDENTE_ERRO:7005:Antes do recibo;AUDIT=OK / RECUPERADO|PASS|
|F10_TX_NOVO_ID_5|F10_Transacoes / transacao_5.accdb|Novo ID nao contorna pendencia|PENDENTE_CONCILIACAO / PENDENTE_CONCILIACAO|PASS|
|F10_TX_6|F10_Transacoes / transacao_6.accdb|Efeitos iguais; intencao continua pendente; nao e falha fisica|PENDENTE_ERRO:7006:Antes do commit;AUDIT=OK / RECUPERADO|PASS|
|F10_TX_NOVO_ID_6|F10_Transacoes / transacao_6.accdb|Novo ID nao contorna pendencia|PENDENTE_CONCILIACAO / PENDENTE_CONCILIACAO|PASS|
|F10_TX_7|F10_Transacoes / transacao_7.accdb|Efeitos iguais; intencao continua pendente; nao e falha fisica|PENDENTE_ERRO:3246:Um objeto de conexão não pode ser fechado explicitamente durante uma transação.;AUDIT=OK / RECUPERADO|PASS|
|F10_TX_NOVO_ID_7|F10_Transacoes / transacao_7.accdb|Novo ID nao contorna pendencia|PENDENTE_CONCILIACAO / PENDENTE_CONCILIACAO|PASS|
|F10_TX_APOS_COMMIT|F10_Transacoes / transacao_8.accdb|Oitavo recibo confirmado; projecao ainda nao salva|COMMIT_CONFIRMADO_PROJECAO_PENDENTE;AUDIT=OK / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_DETECTA_1|F10_AuditoriaAdversarial / audit_1.accdb|Alteracao controlada detectada; nao autentica contra adulteracao coordenada|COUNT_Receipts / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_BLOQUEIO_DURAVEL_1|F10_AuditoriaAdversarial / audit_1.accdb|Replay nao produz efeitos e congela autoridade|PENDENTE_INTEGRIDADE:COUNT_Receipts / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_FRONTEIRA_FRIA_1|F10_AuditoriaAdversarial / audit_1.accdb|Reabertura preserva exigencia de conciliacao|RECONCILE / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_DETECTA_2|F10_AuditoriaAdversarial / audit_2.accdb|Alteracao controlada detectada; nao autentica contra adulteracao coordenada|RECEIPT_HASH / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_BLOQUEIO_DURAVEL_2|F10_AuditoriaAdversarial / audit_2.accdb|Replay nao produz efeitos e congela autoridade|PENDENTE_INTEGRIDADE:RECEIPT_HASH / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_FRONTEIRA_FRIA_2|F10_AuditoriaAdversarial / audit_2.accdb|Reabertura preserva exigencia de conciliacao|RECONCILE / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_DETECTA_3|F10_AuditoriaAdversarial / audit_3.accdb|Alteracao controlada detectada; nao autentica contra adulteracao coordenada|HISTORY / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_BLOQUEIO_DURAVEL_3|F10_AuditoriaAdversarial / audit_3.accdb|Replay nao produz efeitos e congela autoridade|PENDENTE_INTEGRIDADE:HISTORY / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_FRONTEIRA_FRIA_3|F10_AuditoriaAdversarial / audit_3.accdb|Reabertura preserva exigencia de conciliacao|RECONCILE / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_DETECTA_4|F10_AuditoriaAdversarial / audit_4.accdb|Alteracao controlada detectada; nao autentica contra adulteracao coordenada|CHANGELOG / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_BLOQUEIO_DURAVEL_4|F10_AuditoriaAdversarial / audit_4.accdb|Replay nao produz efeitos e congela autoridade|PENDENTE_INTEGRIDADE:CHANGELOG / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_FRONTEIRA_FRIA_4|F10_AuditoriaAdversarial / audit_4.accdb|Reabertura preserva exigencia de conciliacao|RECONCILE / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_DETECTA_5|F10_AuditoriaAdversarial / audit_5.accdb|Alteracao controlada detectada; nao autentica contra adulteracao coordenada|COMMAND_HASH / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_BLOQUEIO_DURAVEL_5|F10_AuditoriaAdversarial / audit_5.accdb|Replay nao produz efeitos e congela autoridade|PENDENTE_INTEGRIDADE:COMMAND_HASH / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_FRONTEIRA_FRIA_5|F10_AuditoriaAdversarial / audit_5.accdb|Reabertura preserva exigencia de conciliacao|RECONCILE / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_DETECTA_6|F10_AuditoriaAdversarial / audit_6.accdb|Alteracao controlada detectada; nao autentica contra adulteracao coordenada|COUNT_ChangeLog / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_BLOQUEIO_DURAVEL_6|F10_AuditoriaAdversarial / audit_6.accdb|Replay nao produz efeitos e congela autoridade|PENDENTE_INTEGRIDADE:COUNT_ChangeLog / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_FRONTEIRA_FRIA_6|F10_AuditoriaAdversarial / audit_6.accdb|Reabertura preserva exigencia de conciliacao|RECONCILE / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_DETECTA_7|F10_AuditoriaAdversarial / audit_7.accdb|Alteracao controlada detectada; nao autentica contra adulteracao coordenada|ATTEMPT_HASH / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_BLOQUEIO_DURAVEL_7|F10_AuditoriaAdversarial / audit_7.accdb|Replay nao produz efeitos e congela autoridade|PENDENTE_INTEGRIDADE:ATTEMPT_HASH / PENDENTE_CONCILIACAO|PASS|
|F10_AUDIT_FRONTEIRA_FRIA_7|F10_AuditoriaAdversarial / audit_7.accdb|Reabertura preserva exigencia de conciliacao|RECONCILE / PENDENTE_CONCILIACAO|PASS|
|F10_FK_ISOLADA_OP_EP|F10_AuditoriaAdversarial / constraints_f10.accdb|Novo movimento sem filhos; FK composta recusa episodio errado|-2147467259:Não é possível adicionar ou alterar registros, pois é necessário que eles tenham um registro relacionado na tabela 'Operations'. / RECUSADO|PASS|
|F10_PLANO_MID_IMUTAVEL|F10_AuditoriaAdversarial / constraints_f10.accdb|Movimento diferente do reservado recusado|-2147467259:Não é possível adicionar ou alterar registros, pois é necessário que eles tenham um registro relacionado na tabela 'Operations'. / RECUSADO|PASS|
|F10_CONTROLE_POSITIVO_FINAL|F10_AuditoriaAdversarial / constraints_f10.accdb|Regra aceita dados validos; rejeicao anterior nao mascara FKs|NOT_REJECTED / VERIFICADO|PASS|
|F10_CALENDARIO_ENGINE_FINAL|F10_AuditoriaAdversarial / constraints_f10.accdb|Data real invalida recusada|-2147467259:Contrato sintetico F10: Projections / RECUSADO|PASS|
|F10_RECIBO_DUPLICADO|F10_AuditoriaAdversarial / constraints_f10.accdb|PK/unique impedem recibo duplicado|-2147467259:As alterações solicitadas para a tabela não foram bem-sucedidas, já que criariam valores duplicados de índice, chave primária ou relação. Altere os dados no campo ou nos campos que contêm dados duplicados, remova o índice, ou redefina-o para permitir entradas duplicadas, e tente novamente. / RECUSADO|PASS|
|F10_PROJECAO_INTEGRAL|F10_Projecoes / projecao_f10.accdb|7 registros/2 ativos e metadados/fatos/recibos/formulas|VERIFIED / RECUPERADO|PASS|
|F10_PROJECAO_ADULTERACAO_1|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|MONEY / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_2|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|FORMULA / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_3|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|NUMBER_FORMAT / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_4|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|IGNORE / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_5|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|DERIVED_RECEIPT / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_6|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|IDENTITY_DUP_MISSING / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_7|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|BASE_REVISION_LINEAGE_SEAL / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_8|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|BASE_REVISION_LINEAGE_SEAL / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_9|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|DERIVED_RECEIPT / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_10|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|IGNORE / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_11|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|IDENTITY_VALUE_TYPE_2 / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_12|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|BUILD_STATE / RECUSADO|PASS|
|F10_PROJECAO_ADULTERACAO_13|F10_Projecoes / projecao_f10.accdb|Verificador recusa divergencia sem escrever na autoridade|TOTAL_FORMULA / RECUSADO|PASS|
|F10_SORT_METADATA_DETECTADO|F10_Projecoes / projecao_f10.accdb|Reproduzir perda de Ignore do Sort; FAIL F9 permanece|IGNORE / RECUSADO|PASS|
|F10_SORT_RECONSTRUIDO|F10_Projecoes / projecao_f10.accdb|Reconstruir e conferir vinculos discriminantes/metadados|VERIFIED / RECUPERADO|PASS|
|F10_ACK_ANTES_SAVE|F10_Projecoes / C:\Users\MARCOS~1\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase10_20261102_155920Z\Projecao_F10_Confirmada.xlsm|Sem arquivo nao confirmar|PENDING_FILE / PENDENTE_CONCILIACAO|PASS|
|F10_SAVE_SEM_ACK|F10_Projecoes / C:\Users\MARCOS~1\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase10_20261102_155920Z\Projecao_F10_Confirmada.xlsm|Save sozinho nao confirma projecao|INTENT / PENDENTE_CONCILIACAO|PASS|
|F10_ACK_REABERTO|F10_Projecoes / C:\Users\MARCOS~1\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase10_20261102_155920Z\Projecao_F10_Confirmada.xlsm|Reabrir somente leitura; verificar e persistir hash exato|CONFIRMED:f98e65232ad0527fccaadee4f76ddaa9951b430acc863243ef97295c6f9efe71 / RECUPERADO|PASS|
|F10_SAVE_CANCELADO_REAL|F10_Projecoes / C:\Users\MARCOS~1\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase10_20261102_155920Z\Projecao_F10_SaveCancelado.xlsm|BeforeSave cancela chamada real|Observado=Verdadeiro;Saved=Falso / PENDENTE_CONCILIACAO|PASS|
|F10_ACK_SAVE_CANCELADO|F10_Projecoes / C:\Users\MARCOS~1\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase10_20261102_155920Z\Projecao_F10_SaveCancelado.xlsm|Nao confirmar geracao apenas em memoria|PENDING_VERIFY:BUILD_ID / PENDENTE_CONCILIACAO|PASS|
|F10_SAVEAS_INVALIDO|F10_Projecoes / projecao_f10.accdb|Erro real1004 sem perder fatos/recibos|1004:O Microsoft Excel não pode acessar o arquivo 'C:\Users\MARCOS~1\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase10_20261102_155920Z\PARENT_MISSING\09EC7E00'. Há várias razões possíveis:<CR><LF><CR><LF>• O nome do arquivo ou caminho não existe.<CR><LF>• O arquivo está sendo usado por outro programa.<CR><LF>• A pasta de trabalho que você está tentando salvar tem o mesmo nome de outra <CR><LF>pasta de trabalho aberta no momento. / PENDENTE_CONCILIACAO|PASS|
|F10_REBUILD_INTERROMPIDO|F10_Projecoes / projecao_f10.accdb|Cabecalho parcial nao autoriza; autoridade igual|7100:BASE_REVISION_LINEAGE_SEAL / PENDENTE_CONCILIACAO|PASS|
|F10_BACKUP_REABERTO|F10_BackupRestauracao / restauracao_independente_f10.accdb|Backup fechado restaura esquema/fatos/recibos/trilha|OK / RECUPERADO|PASS|
|F10_BACKUP_RECONSTRUIDO|F10_BackupRestauracao / restauracao_independente_f10.accdb|Reconstrucao integral do backup restaurado|VERIFIED / RECUPERADO|PASS|
|F10_RESTAURACAO_FATOS_POSTERIORES|F10_BackupRestauracao / restauracao_independente_f10.accdb;transacao_8.accdb|Copia antiga nao apaga oitavo fato; preservar ambas|OLD=7;NEW=8;NAO_PROMOVIDO / PENDENTE_CONCILIACAO|PASS|
|F10_PREPARE_COM_DISPUTA|F10_ConcorrenciaPreparar / concorrencia_f10.accdb|Revogacao concorrente nao atravessa transacao de preparacao|PREPARADA / RECUSADO|PASS|
|F10_OLD_OWNER_COMMIT|F10_ConcorrenciaObsoleto / concorrencia_f10.accdb|A antigo nao confirma apos B obter nova geracao|RECUSADO_FENCE / RECUSADO|PASS|
|F10_NOVO_DONO_NAO_HERDA_INTENCAO|F10_ConcorrenciaObsoleto / concorrencia_f10.accdb|B nao conclui automaticamente intencao de A|PENDENTE_OWNER_CHANGED / PENDENTE_CONCILIACAO|PASS|
|F10_NOVO_DONO_NOVO_ID|F10_ConcorrenciaObsoleto / concorrencia_f10.accdb|Novo ID nao contorna intencao orfa|PENDENTE_CONCILIACAO / PENDENTE_CONCILIACAO|PASS|
|F10_SNAPSHOT_COM_DISPUTA|F10_FronteiraExecutar / fronteira_snapshot_f10.accdb|Snapshot coerente sob lease e fronteira transacional|Revision=7;Count=7 / VERIFICADO|PASS|
|F10_APPLY_COM_DISPUTA|F10_FronteiraExecutar / fronteira_apply_f10.accdb|Confirmacao serializada contra revogacao concorrente|CONFIRMADA / RECUPERADO|PASS|
|F10_INTENCAO_BASE_1|F10_IntencaoIncompleta / intencao_1_f10.accdb|Intencao integra antes da adulteracao|PREPARADA / VERIFICADO|PASS|
|F10_INTENCAO_DIVERGENTE_1|F10_IntencaoIncompleta / intencao_1_f10.accdb|Diagnosticar vinculo/hash/origem incompleta|PREPARED_LINK / RECUSADO|PASS|
|F10_INTENCAO_CONCILIAR_1|F10_IntencaoIncompleta / intencao_1_f10.accdb|Persistir bloqueio por divergencia sem restauracao ficticia|PENDENTE_INTEGRIDADE:PREPARED_LINK / PENDENTE_CONCILIACAO|PASS|
|F10_INTENCAO_BASE_2|F10_IntencaoIncompleta / intencao_2_f10.accdb|Intencao integra antes da adulteracao|PREPARADA / VERIFICADO|PASS|
|F10_INTENCAO_DIVERGENTE_2|F10_IntencaoIncompleta / intencao_2_f10.accdb|Diagnosticar vinculo/hash/origem incompleta|PREPARED_LINK / RECUSADO|PASS|
|F10_INTENCAO_CONCILIAR_2|F10_IntencaoIncompleta / intencao_2_f10.accdb|Persistir bloqueio por divergencia sem restauracao ficticia|PENDENTE_INTEGRIDADE:PREPARED_LINK / PENDENTE_CONCILIACAO|PASS|
|F10_INTENCAO_BASE_3|F10_IntencaoIncompleta / intencao_3_f10.accdb|Intencao integra antes da adulteracao|PREPARADA / VERIFICADO|PASS|
|F10_INTENCAO_DIVERGENTE_3|F10_IntencaoIncompleta / intencao_3_f10.accdb|Diagnosticar vinculo/hash/origem incompleta|PREPARED_HASH / RECUSADO|PASS|
|F10_INTENCAO_CONCILIAR_3|F10_IntencaoIncompleta / intencao_3_f10.accdb|Persistir bloqueio por divergencia sem restauracao ficticia|PENDENTE_INTEGRIDADE:PREPARED_HASH / PENDENTE_CONCILIACAO|PASS|
|F10_INTENCAO_BASE_4|F10_IntencaoIncompleta / intencao_4_f10.accdb|Intencao integra antes da adulteracao|PREPARADA / VERIFICADO|PASS|
|F10_INTENCAO_DIVERGENTE_4|F10_IntencaoIncompleta / intencao_4_f10.accdb|Diagnosticar vinculo/hash/origem incompleta|PREPARED_LINK / RECUSADO|PASS|
|F10_INTENCAO_CONCILIAR_4|F10_IntencaoIncompleta / intencao_4_f10.accdb|Persistir bloqueio por divergencia sem restauracao ficticia|PENDENTE_INTEGRIDADE:PREPARED_LINK / PENDENTE_CONCILIACAO|PASS|
|F10_FK_DENTRO_TRANSACAO|F10_FKTransacao / fk_transacao_f10.accdb|Rejeicao da FK Vehicles; rollback e estado conferidos|-2147467259:Não é possível adicionar ou alterar registros, pois é necessário que eles tenham um registro relacionado na tabela 'Vehicles'.;AUDIT=OK / RECUPERADO|PASS|

## Evidências COM e impedimentos

[frio_com_f10.json](./_auditoria/fase10_20261102_155920Z/evidencias/frio_com_f10.json), [prepare](./_auditoria/fase10_20261102_155920Z/evidencias/prepare_worker_prepare_done.json), [nova geração](./_auditoria/fase10_20261102_155920Z/evidencias/worker_b_fence.json), [snapshot](./_auditoria/fase10_20261102_155920Z/evidencias/snapshot_worker_prepare_done.json), [Apply](./_auditoria/fase10_20261102_155920Z/evidencias/apply_worker_prepare_done.json), [bloqueio frio original](./_auditoria/fase10_20261102_155920Z/evidencias/bloqueio_sessao_fria.json).

|Cenário ainda não demonstrado|Status|Motivo / condição de continuação|
|---|---|---|
|F10_FRIO_DADOS_RECIBOS, F10_FRIO_ACK_PERSISTIDO, F10_FRIO_ACK_PENDENTE, F10_FRIO_CONCILIACAO_DURAVEL, F10_FRIO_INTENCAO, F10_FRIO_LEASE_ORFAO|BLOQUEADO|Office recusou execução fria; COM posterior tem escopo próprio|
|Erro físico retornado de CommitTrans; energia; crash; disco cheio; corrupção|NÃO EXECUTADO|Não foi produzida falha física segura; erro3246 ocorre antes de CommitTrans|
|Backup aberto consistente; falha real de rollback; takeover administrativo de lease órfão|NÃO EXECUTADO|Sem mecanismo aprovado/demonstrado|
|Concorrência em rede e SQL externo hostil; autenticação criptográfica da trilha|NÃO EXECUTADO|Modelo/controles específicos pendentes|
|ACK adulterado coordenadamente, corrida de arquivo, duas cópias autoritativas com mesmo BaseID|NÃO EXECUTADO|Limites de protocolo ainda abertos|
|Migração interrompida e retomada automática|NÃO EXECUTADO|Reaplicação de migração concluída foi aprovada; interrupção DDL exige checkpoint|

## Falhas preservadas e correções

R1: erro1004 na atribuição Ignore por intervalo; corrigido por célula. R2: Annual persistida como texto por NumberFormat=@, dois resultados derivados FAIL e interrupção94; corrigido General e exigência de baseline válida. Os PASS negativos dessa R2 não são considerados prova suficiente porque a baseline era inválida.

ANTES_PREPARED_intencoes_f10.tsv: oito FAIL em auditoria/conciliação de intenções incompletas; o caso1 permitiu nova preparação. ANTES_CUSTODIA_intencoes_f10.tsv: dois FAIL remanescentes por PlaceCode não conferido. Correção final acoplou PreparedAudit10 ao auditor e conferiu local do episódio contra CurrentPID. Resultado final12 PASS; arquivos reprovados e fontes anteriores preservados.

Testes transacionais foram repetidos após essa mudança porque o auditor é chamado por Prepare/Apply; os19 casos finais passaram. Não multiplicar essa repetição na contagem de IDs únicos. A FK isolada em transação foi criada porque o teste inicial de histórico falhou primeiro em Required.

V_VALIDO, T_VALIDO, D_VALIDO, D_FALHA_BAIXA, T_DUP_DESTINO, C_FALHA_BAIXA e C_FALHA_COMPENSACAO continuam **FAIL históricos do produto**, não reexecutados nem resolvidos por este laboratório. Sete Delete e57 leitores preservados. Sistema NÃO HOMOLOGADO.
