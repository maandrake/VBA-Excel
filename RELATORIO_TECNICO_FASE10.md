# Relatório técnico — Fase 10

**Candidato experimental com pendências. Sistema principal NÃO HOMOLOGADO. Portão2 não executado.** O laboratório consolidou schema4 e demonstrou, no escopo ensaiado, rollback interno, identidade/versionamento, fencing, trilha verificada e reconstrução da projeção. A suíte fria VBA final foi bloqueada pela política do Office; há limites adicionais antes de qualquer integração.

## Preservação e entrega

Execução20261102_155920Z. Os13 documentos obrigatórios foram lidos integralmente, sem reiniciar a auditoria geral. A entrega final F9 foi conferida: fbb127500d504e6ac9bf3f857dfc936f1f0e43449b3549a2473fff2a5b56386d. Foram verificados2.417 arquivos do manifesto anterior sem divergência, incluindo39 originais;50 backups novos foram criados e novamente verificados antes dos adendos documentais. Evidências e bancos anteriores foram preservados.

Entrega:[Fixture_Integridade_Fase10_EXPERIMENTAL_NAO_HOMOLOGADA.xlsm](./_auditoria/fase10_20261102_155920Z/entrega/Fixture_Integridade_Fase10_EXPERIMENTAL_NAO_HOMOLOGADA.xlsm),2.562.750 bytes, SHA-256 **7d1f119124e71a3314601e8c41475db22de9429010aa5424b1652d655b38b7e2**. Área arquivada:[_auditoria/fase10_20261102_155920Z](./_auditoria/fase10_20261102_155920Z/). Office executou somente no TEMP sintético identificado no manifesto inicial; o arquivo arquivado no OneDrive não foi aberto no Office.

A fixture contém66 componentes:63 herdados e três módulos experimentais novos. Os63 corpos herdados permanecem equivalentes, preservando strings e lógica; a comparação desconta capitalização ajustada pelo VBE, atributos VB_Base gerados e espaços/quebras de linha. As fontes novas conferem com o VBA efetivamente salvo. As17 planilhas/CodeNames foram preservadas. Os FRM/FRX externos originais têm hashes preservados; o Office reserializou streams internos dos designers, portanto não se declara identidade binária dos designers internos. Nenhum controle foi editado intencionalmente; Cad_NFE/ListBox foi carregado no MSForms real nos benchmarks, sem certificar todos os layouts/eventos.

## Alterações estritamente experimentais

|Componente / procedimento|Problema e causa|Alteração / dependências|Evidência e risco residual|Reversão|
|---|---|---|---|---|
|modPersistenciaFase10 / Migrate10, CreateDB10|Restrições F9 não uniformes; ponte CurrentPID parcialmente nula rejeitava NEW|Schema4 único em base nova/migrada, FKs compostas, regras DAO, ponte CurrentEpisodeID anulável e recibos/canon versionados|16 asserções de esquema; testes isolados adicionais; DDL interrompido ainda exige checkpoint|Base sintética F9 e checkpoints preservados|
|modPersistenciaFase10 / Prepare10, Apply10, LockWriter10|Leitura prévia de lock não bastaria; intenção não pode ser herdada por novo dono|Revalidação e contagem afetada dentro de transação, CAS de versão e vínculo Owner/Fence da tentativa|Disputas em Prepare/Apply/snapshot e recusa do token antigo; somente escritores cooperantes|Módulo anterior e bases independentes|
|modCenariosFase10 / AuditBatch10; modPersistenciaFase10 / Audit10|Auditoria de fatos não bastava para intenções PREPARED incompletas|Hashes/cadeia/antes-depois por leitura em bloco; auditor chama PreparedAudit10|Adulterações26 PASS; guarda nova12 PASS após FAIL reproduzidos; não autentica SQL hostil|Checkpoint antes da guarda e logs reprovados|
|modFronteirasFase10 / PreparedAudit10|Pending/tentativa inicial ausente ou payload preparado alterado passava; custódia corrente não conferida|Exige vínculos completos, envelope/fonte e local do episódio igual à projeção CurrentPID; próxima preparação persiste RECONCILE|ANTES_PREPARED oito FAIL; ANTES_CUSTODIA dois FAIL; final12 PASS; contenção não restaura registros|Fontes anteriores e ACCDBs preservados|
|modCenariosFase10 / BuildFull10, VerifyFull10|Ignore de Range falhou; formato textual impediu cálculo de Annual|Ignore por célula da chave; General para fórmula e @ para documentos; baseline deve ser VERIFIED|23 PASS finais; R1/R2 preservados; contrato cobre projeção sintética de19 colunas|Fontes/checkpoints R1/R2|
|modCenariosFase10 / IntentAck10, ConfirmAck10|Save não equivale a arquivo verificado|INTENT separado de BUILT e confirmação após Save/reabertura/verificação/hash|ACK normal, Save cancelado, SaveAs inválido e rebuild interrompido; corrida de arquivo/ACK hostil pendentes|Preservar confirmação e arquivo; reconstruir nova geração|

Nenhuma alteração em gravadores reais,57 leitores, sete Delete, regras fiscais ou dados operacionais. Os arquivos BAS novos usam Windows-1252/CRLF. O nome externo modLaboratorioFase10.bas contém VB_Name=modPersistenciaFase10, documentado para não importar uma duplicata com nome inventado.

## Resultados comprovados

- **121 asserções VBA únicas PASS**, zero FAIL na última execução de cada caso: esquema16; ciclo15; transações19; auditoria26; projeção23; restauração3; concorrência/fencing/fronteiras6; intenção12; FK transacional1.
- Benchmark:12 amostras VBA PASS em100/1.000/5.000/10.000, com três repetições;100 repetido após a guarda final. Medidas maiores pertencem à revisão anterior à guarda adicional, sem reclassificação silenciosa.
- Quatro experiências de concorrência COM entre processos PASS e seis verificações frias COM PASS. A segunda sessão legítima ensaiou escrita diagnóstica, não movimentação completa do produto.
- Seis verificações frias VBA **BLOQUEADAS**. Instância quente PID52068 encerrada normalmente; nova PID41728 recusou macros. Conteúdo ativo bloqueado, sem habilitação específica disponível naquele estado. Não foi alterada nenhuma política. PID42180 do usuário permaneceu ativo.
- Compilação real no Editor VBA e execução das rotinas experimentais comprovadas na instância quente após a correção final. Inspeção estática não foi usada como substituto.

A matriz detalha IDs, procedimentos, fixtures, esperado, observado e evidências. Os121 PASS são do laboratório e não resolvem os sete FAIL históricos do produto. Falhas de desenvolvimento R1/R2 e da guarda intermediária permanecem arquivadas, com causa e correção.

## O que foi demonstrado

Bases novas e sintéticas migradas têm o mesmo fingerprint estrutural. O ciclo usa dois episódios do mesmo veículo, NF-e repetida e valores diferentes, sem inferir vínculo por placa/linha. O motor confirmou operações e recibos em transações internas; erros anteriores ao commit restauraram os efeitos examinados, mantendo intenção de transação anterior. Novo OperationID não contornou pendência.

A projeção foi reconstruída por identidade, validada integralmente no contrato sintético, salva e reaberta antes do ACK no banco. Sort nativo perdeu Ignore, foi detectado e reconstrução corrigiu os metadados. As leituras COM após fechar a instância quente recuperaram fatos, recibos, hash do arquivo confirmado e estados pendentes. Não houve execução funcional VBA fria.

## Limites que impedem integração/homologação

1. Os sete FAIL históricos, a política de baixa, os leitores/fórmulas e os eventos reais continuam pendentes; Portão2 não aprovado.
2. Sessão fria VBA bloqueada. Falha física de CommitTrans, energia, crash, disco cheio e corrupção NÃO EXECUTADOS. Erro3246 de Close com transação ativa é anterior ao commit.
3. Imutabilidade contra SQL externo e adulteração coordenada de hashes não demonstrada; atributos cadastrais mutáveis, como placa, ainda não têm histórico completo de retificação no laboratório. MAC/assinatura/ACL e governança dependem de decisão.
4. Autoridade de duas cópias ACCDB com mesmo BaseID, integridade/imutabilidade completa dos ACKs e corrida de arquivo entre verificação/confirmação exigem protocolo adicional.
5. Recusas anteriores à existência de OperationID válido são evidenciadas nos testes, sem trilha durável completa de todas as recusas no banco.
6. Lease/intenção órfãos e estado RECONCILE exigem conciliação; não há takeover administrativo automático nem recuperação integral por bloqueio.
7. Backup fechado foi conferido; cópia a quente consistente e recuperação DDL interrompida não demonstradas. Base nova e migração concluída idempotente não provam retomada de migração parcial.
8. ACE local funcionou no ambiente atual; distribuição, licença e compatibilidade no destino permanecem PENDENTES. Nenhum componente foi instalado.

## Critérios de entrega

|Critério|Situação|
|---|---|
|Schema consolidado/versionado e constraints|Demonstrado no laboratório; regras de protocolo não são todas constraints físicas|
|Concorrência/fencing|Demonstrado nas três fronteiras e token antigo; não em todos os caminhos externos/futuros|
|Auditoria/idempotência|Demonstrado nos cenários; limites de autenticação/recusas e retificações permanecem|
|Verificador e confirmação persistente|Demonstrados na sessão quente; verificação fria VBA bloqueada; ACK/corrida pendentes|
|Recuperação/backup/restauração|Demonstrados para erros controlados/backup fechado; sem resistência física comprovada|
|Benchmark|Quatro volumes, três medidas; revisões e metodologia explicitadas|
|Integração operacional/homologação|Não autorizada e não realizada|

O laboratório está entregue com evidências e pendências delimitadas; não se declara atendimento integral das garantias de produção.

## Continuação objetiva da sessão fria

Na instância isolada, abrir apenas o mestre TEMP indicado no manifesto. Se o próprio Office oferecer Habilitar Conteúdo para essa cópia sintética, habilitar somente o documento. Se continuar indicando conteúdo ativo bloqueado e exigir Central de Confiabilidade/política, parar e registrar a decisão específica de segurança antes de alterar qualquer configuração. No Editor VBA: Depurar → Compilar VBAProject; na janela Imediata executar Call F10_Frio. Devolver mensagem F10_COLD_DONE e o arquivo frio_f10.tsv com seis resultados, ou a mensagem original. Não repetir suítes que criam arquivos já existentes.

## Documentos e restauração

[Esquema](./ESQUEMA_CONSOLIDADO_FASE10.md), [concorrência](./TESTES_CONCORRENCIA_FASE10.md), [transações](./TESTES_TRANSACOES_FASE10.md), [recuperação](./TESTES_RECUPERACAO_FASE10.md), [projeções](./VERIFICACAO_PROJECOES_FASE10.md), [benchmark](./BENCHMARK_FASE10.md), [matriz](./MATRIZ_REGRESSAO_FASE10.md), [decisões](./PACOTE_DECISAO_INTEGRACAO_FASE10.md). Backups:[backup](./_auditoria/fase10_20261102_155920Z/backup/). Reverter somente o laboratório por checkpoint conhecido; com fatos posteriores preservar ambas as versões e conciliar, sem restauração cega.

[Roteiro de reprodução e retomada da sessão fria](./_auditoria/fase10_20261102_155920Z/REPRODUCAO_FASE10.md).
