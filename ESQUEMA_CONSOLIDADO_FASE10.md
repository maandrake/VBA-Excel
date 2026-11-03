# Esquema consolidado — Fase 10

Laboratório sintético; esquema 4. Nenhuma migração do produto foi autorizada.

O catálogo efetivo contém 13 tabelas: as 12 entidades da Fase 9 e ProjectionAcks, que distingue intenção de reconstrução e confirmação de arquivo persistido. Pending representa a pendência de conciliação; não foi criada tabela redundante Reconciliation.

SHA-256 estrutural observado em base nova e migrada: `8698a2b7f44693510141e9fe75fd123fe41b4d99f7c60a03c1f9d2a99625d26b`. Inclui colunas, chaves, índices, regras DAO e Required; não é uma assinatura autenticada.

## Catálogo observado

|Tabela|Colunas efetivas|Relações estrangeiras|
|---|---|---|
|Attempts|AttemptHash, AttemptID, ExpectedVersion, FenceNo, OperationID, OutcomeText, OwnerID, PayloadHash, RequestCanon, StateAfter, StateBefore|FK_AO → Operations|
|ChangeLog|AfterCanon, AuditHash, AuditID, BeforeCanon, OperationID, RevisionNo|FK_LO → Operations|
|Episodes|CurrentEpisodeID, CurrentPID, EpisodeID, PlaceCode, StatusText, VehicleID, VersionNo|FK_CURRENT_EP → Projections; FK_EV → Vehicles|
|Meta|BaseID, ID, IntegrityReason, LineageID, RevisionNo, SchemaHash, SchemaNo, WriteState||
|Movements|EpisodeID, KindText, MovementID, OperationID|FK_M_PLAN → Operations; FK_ME → Episodes; FK_MO → Operations; FK_MO_EP → Operations|
|Operations|AccessKey, AmountIn, AmountOut, CivilDate, CommandCanon, CommandHash, CommandVersion, DocNumber, EpisodeID, ExpectedVersion, KindText, OperationID, PayloadText, PlannedMID, PlannedPID, SourceCanon, SourceHash, StatusText|FK_OE → Episodes|
|Pending|EpisodeID, OperationID, ReasonText|FK_NE → Episodes; FK_NO → Operations; FK_PENDING_EP → Operations|
|ProjectionAcks|AuthoritySeal, BaseID, BuildID, FileHash, FilePath, LineageID, RevisionNo, StatusText||
|Projections|AccessKey, AmountIn, AmountOut, CivilDate, ContentCanon, DocNumber, EpisodeID, MovementID, PlaceCode, ProjectionID|FK_P_PLAN → Operations; FK_PE → Episodes; FK_PM → Movements; FK_PM_EP → Movements|
|Provenance|DecisionText, EpisodeID, EvidenceText, LegacyRef, SourceHash|FK_PROV_EP → Episodes|
|Receipts|BaseID, EffectCanon, EpisodeID, MovementID, OperationID, PreviousHash, ProjectionID, ReceiptHash, ReceiptID, RevisionNo, VehicleID|FK_R_MOV → Movements; FK_R_PROJ → Projections; FK_R_VEHICLE → Episodes; FK_RO → Operations|
|Vehicles|Plate, VehicleID||
|WriterLease|FenceNo, ID, OwnerID, StatusText||

## Contratos e limites

- Novos IDs internos: GUID de 36 caracteres, imutável no protocolo. IDs legados conhecidos são preservados apenas na migração sintética. Placa, NF-e e linha não são identidade global.
- NF-e documental: número como texto, chave exatamente 44 dígitos; zero inicial preservado. O exemplo sintético não comprova autorização fiscal.
- Currency no VBA/ADO/ACE: escala de quatro casas, sem aceitar quinta casa silenciosamente. O limite experimental de magnitude 1.000.000.000 em Prepare10 protege a representação na projeção Excel; depende de decisão operacional antes de eventual integração.
- Data civil ISO yyyy-mm-dd, calendário real validado; não se usa NumberFormat regional como identidade.
- ExpectedVersion exige comparação com a versão corrente do episódio; RevisionNo é sequência da autoridade. CommandVersion 1 identifica a representação legada migrada e 2 inclui os IDs reservados de movimento/projeção. Não confundir CommandVersion com SchemaNo=4.
- PK, unicidade por operação/revisão e FKs compostas vinculam episódio/operação/movimento/projeção. Atualizações de pais referenciados e exclusões parentais são recusadas sem cascata.
- CurrentEpisodeID é ponte anulável: episódio NEW possui CurrentPID e ponte nulos; após entrada ambos estão presentes e a ponte coincide com EpisodeID. Evita a rejeição indevida de novo episódio observada no protótipo inicial.
- Regras DAO verificam Required, domínios de estado, revisões, calendário e dígitos da chave. Nem toda regra de protocolo é uma constraint física.
- Hashes detectam divergência acidental/adulteração não coordenada. Acesso SQL direto pode alterar fatos e recalcular hashes; não há ACL específica, MAC, assinatura ou âncora externa implementados. IDs sem dependentes não são fisicamente imutáveis contra SQL externo.

## Migração experimental

Migrate10 aplica a evolução 1→4 em cópia de uma base sintética F9 conhecida e em base nova criada pelo mesmo bootstrap. Backfill de comandos, recibos e tentativas é explícito; valores legados não são convertidos em evidência inexistente. Reaplicação sobre esquema 4 completo verifica auditoria e não duplica fatos. DDL interrompido não foi demonstrado como retomável: preservar a cópia forense e reconstruir um candidato a partir do checkpoint.

Não copiar indiscriminadamente o antigo migracao.accdb. A compatibilidade de estrutura foi medida entre novo.accdb e legado_f9.accdb desta execução, além de ciclo F10 com GUIDs novos. Os primeiros testes negativos foram complementados por inserções isoladas: uma rejeição causada por dependentes não prova, sozinha, a FK que se pretendia testar.

## Evidências

[Catálogo ADOX efetivo](./_auditoria/fase10_20261102_155920Z/evidencias/schema_catalogo_f10.json), [matriz](./MATRIZ_REGRESSAO_FASE10.md) e [fontes experimentais](./_auditoria/fase10_20261102_155920Z/prototipos/). Os arquivos são arquivados ao fim da execução.
