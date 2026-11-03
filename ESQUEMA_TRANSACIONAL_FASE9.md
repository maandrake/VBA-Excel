# Esquema transacional experimental — Fase 9

Esquema 1 executado somente em bancos sintéticos; reforço de FK composta ensaiado em migracao.accdb. [SQL proposto](./_auditoria/fase9_20261102_142753Z/prototipos/esquema-proposto-v2.sql), [fontes realmente incorporados](./_auditoria/fase9_20261102_142753Z/prototipos/revisao_final/). Nenhuma aba ou coluna nova no sistema principal.

| Tabela | Identidade e relação | Finalidade / limite |
|---|---|---|
| Meta | ID1, BaseID, SchemaNo, RevisionNo | Autoridade local da fixture, revisão CAS; valor não torna cópia arbitrária legítima |
| WriterLease | ID1, OwnerID, FenceNo, StatusText | Escritor exclusivo pelo protocolo; sem TTL/liberação automática |
| Vehicles | VehicleID PK; Plate atributo | Mesmo veículo possui episódios distintos; não fundir por placa |
| Episodes | EpisodeID PK, VehicleID FK, VersionNo, Status,CurrentPID,Place | Custódia global experimental: T/D interna mantém episódio; venda fecha; reentrada cria outro |
| Operations | OperationID PK, EpisodeID FK, ExpectedVersion,kind,quantias,CommandCanon,SourceCanon,PlannedPID/MID,status | Intenção confirmada em transação anterior aos efeitos; repetição/payload/pendência |
| Attempts | AttemptID PK,OperationID FK,OwnerID,FenceNo | Tentativa preparada; protótipo não grava cada replay recusado como tentativa nova |
| Movements | MovementID PK,OperationID unique/FK,EpisodeID FK,kind | Um movimento por operação neste modelo reduzido |
| Projections | ProjectionID PK,EpisodeID FK,MovementID FK,quantias,data,documento,chave,local,ContentCanon | Fato histórico persistente; posição no Excel é derivada. Episódio indica CurrentPID elegível |
| Receipts | ReceiptID PK,OperationID unique/FK,RevisionNo,EffectCanon | Confirmado na mesma transação dos efeitos; não depende do Save Excel |
| ChangeLog | AuditID counter,OperationID FK,BeforeCanon,AfterCanon | Pré/pós estado mínimo; ensaio de erro FK no diário reverte transação |
| Pending | EpisodeID PK/FK,OperationID FK,Reason | Nova intenção não contorna pendência |
| Provenance | LegacyRef PK,SourceHash,Decision,EpisodeID opcional,Evidence | Duas relações sintéticas declaradas; ambíguo permanece NULL |

## Tipos e serialização

IDs opacos TEXT40, GUIDs novos por CoCreateGuid para projeções/movimentos/recibos/sessões/tentativas; teste100 GUIDs distintos. VehicleID/EpisodeID/OperationID de casos usam literais conhecidos. O gerador/validação universal desses IDs ainda precisa contrato de produção e índice único; não foi provada impossibilidade de colisão. PK/FK não tornam todo atributo automaticamente imutável contra SQL direto.

Chave TEXT44 e String na leitura, zero inicial preservado. Número de NF-e TEXT20 repetido88001 em contextos distintos, sem ser chave global. Data civil TEXT10 ISO explícito2026-11-02 no plano de teste, sem dependência regional. A API reduzida não aceita datas/documentos arbitrários; validação completa desses campos depende da integração futura e não foi inventada. Valores CURRENCY/VBA Currency, quatro casas; precisão observada1000,1234/1370,9876, sem mudar tolerâncias fiscais. Excel exibe Double; a prova cobre os intervalos sintéticos, não todo o limite Currency.

Canonicalização textual tipada/comprimento, quantias em unidades10^-4 e IDs/plano/envelope. ContentCanon detecta alteração nos campos cobertos, não autentica pessoa nem assina XML. Não chamar selo textual de SHA-256. SHA-256 é usado nos arquivos/manifestos e na proveniência. O hash de payload criptográfico e a assinatura do recibo ainda são requisitos do candidato produtivo.

## Fronteiras e insuficiências explícitas

Prepare confirma intenção+Attempt+Pending; Apply bloqueia linha de lease, confere fonte/versão e grava movimento, projeção, episódio, ChangeLog, revisão e recibo numa transação. Erro antes do commit executa rollback e compara todos os dados de negócio, mantendo intenção pendente. Encerramento/crash durante confirmação não foi testado. Repetição íntegra retorna recibo; mudanças de valor foram recusadas. O replay atual verifica canônico da projeção e não é verificador universal de todos os FKs/logs/recibos contra adulteração coordenada.

Esquema 1 possui FKs separadas de episódio/movimento; reforço composto foi criado e testado somente no banco migracao.accdb: unique(EpisodeID,MovementID) e FK composta na projeção. A publicação de SchemaNo 2 e a aplicação uniforme aos escritores/benchmarks **não foram executadas**. CurrentPID, veículo ativo único, imutabilidade de IDs/fatos e relações completas de cadeia/predecessor dependem de validações/constraints adicionais. Origem histórica de cada efeito deve ser explicitada antes de M4/M7.

A revisão do protocolo também precisa tornar a preparação e a revogação de lease mutuamente exclusivas na mesma fronteira; a leitura de Own9 na preparação não demonstra todos os interleavings com uma revogação concorrente. Nenhum resultado do laboratório autoriza integrar este núcleo como gravador definitivo.

Compatível provisoriamente com o [contrato F7](./CONTRATO_IDENTIDADE_FASE7.md), com as reduções acima; sem inferência fiscal ou migração de legado real.
