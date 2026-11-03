# Testes de transações — Fase 10

Escopo: VBA real compilado no Excel de laboratório, ADO/ACE 16, bases sintéticas independentes. Fontes e logs em [_auditoria/fase10_20261102_155920Z](./_auditoria/fase10_20261102_155920Z/). A [matriz](./MATRIZ_REGRESSAO_FASE10.md) discrimina cada asserção e suas evidências.

## Fronteiras efetivamente exercitadas

Prepare10 confirma a intenção, tentativa e Pending numa primeira transação. Apply10 usa outra transação para movimento, projeção, atualização CAS do episódio, histórico, revisão e recibo. Uma falha em Apply não apaga a intenção já persistida.

F10_Transacoes executou 19 asserções aprovadas: nove fronteiras anteriores ao commit com conferência de efeitos e bloqueio de novo OperationID, mais a falha posterior ao commit. A comparação DataSig10 cobre Episodes, Movements, Projections, Receipts, ChangeLog e Meta. Operations, Attempts e Pending são examinados separadamente; não se promete retorno literal da base inteira ao estado anterior à preparação.

|Ponto|Falha observada|Estado posterior|Classificação|
|---|---|---|---|
|-1|Recusa antes de BeginTrans|Sete fatos/recibos anteriores intactos; intenção pendente|RECUSADO|
|-2|Erro artificial 7012 após BeginTrans|Efeitos revertidos e auditados; intenção permanece|RECUPERADO dos efeitos|
|1|Erro artificial 7001 após movimento|Mesmo conjunto anterior de fatos|RECUPERADO dos efeitos|
|2|Erro artificial 7002 após projeção|Sem novo movimento/projeção confirmado|RECUPERADO dos efeitos|
|3|Erro artificial 7003 durante baixa lógica sintética|Episódio, vínculos e revisão anteriores restaurados|RECUPERADO dos efeitos|
|4|Erro real de Required em ChangeLog.BeforeCanon|Rollback conferido; esse erro não prova a FK pretendida|RECUPERADO dos efeitos|
|5|Erro artificial 7005 antes do recibo|Sem recibo da operação; intenção persiste|RECUPERADO dos efeitos|
|6|Erro artificial 7006 antes de chamar CommitTrans|Fatos anteriores intactos|RECUPERADO dos efeitos|
|7|ACE recusou Close com transação ativa, erro 3246|Rollback conferido pela rotina de erro|RECUPERADO dos efeitos|
|8|Interrupção artificial após CommitTrans|Oitavo fato/recibo confirmado; projeção ainda ausente|PENDENTE_CONCILIACAO da projeção|

O erro 3246 acontece antes da chamada de CommitTrans: não é falha física de commit. Erro de I/O retornado por CommitTrans, disco cheio, energia, corrupção e crash não foram provocados. Nenhuma instância do usuário foi encerrada.

O cenário de FK dentro de transação foi separado do ponto 4 para impedir que Required mascarasse a rejeição referencial; seu resultado está na matriz final, junto ao teste F10_FK_DENTRO_TRANSACAO.

## Recibos, comandos e tentativas

O ciclo usa duas entradas do mesmo veículo e uma entrada de outro veículo, sete movimentos e duas posições ativas. Entrada → venda → reentrada → transferência → devolução → transferência preserva VehicleID e cria EpisodeID distinto na reentrada. A continuidade de EpisodeID entre custódias é hipótese experimental, sujeita a aprovação.

Cada operação reserva MovementID e ProjectionID. O recibo inclui BaseID, revisão, episódio, veículo, movimento e projeção, efeito canônico e hash do recibo anterior. ChangeLog guarda antes/depois e hash. Tentativas de replay de operação conhecida guardam RequestCanon, ExpectedVersion, resultado, transição, proprietário e fencing. Replay idêntico retorna JA_CONCLUIDA sem duplicar fatos; payload divergente é recusado e auditado.

F10_AuditoriaAdversarial aprovou 26 asserções: sete adulterações, respectivas recusas e persistência de RECONCILE após reabrir a conexão, além de cinco provas dirigidas de constraints. Não confundir essa reabertura de conexão com sessão fria de outro Excel, registrada separadamente.

## Limites obrigatórios

- Recusa antes de existir OperationID válido ainda é registrada no log de ensaio, sem inventar uma operação para satisfazer a FK de Attempts. Um registro durável de todas as recusas exige contrato adicional.
- Hash SHA-256 sem segredo detecta divergência; não autentica um administrador que possa alterar dados e hashes conjuntamente.
- FKs impedem diversas mudanças de pais referenciados, mas não oferecem imutabilidade completa de todos os fatos contra SQL externo.
- O protocolo serializa escritores cooperantes; não é proteção contra escrita externa fora do protocolo.
- Banco e Save do Excel não formam uma transação conjunta.

Os sete FAIL históricos e os sete Delete do produto permanecem inalterados. PASS deste laboratório não aprova os eventos operacionais.
