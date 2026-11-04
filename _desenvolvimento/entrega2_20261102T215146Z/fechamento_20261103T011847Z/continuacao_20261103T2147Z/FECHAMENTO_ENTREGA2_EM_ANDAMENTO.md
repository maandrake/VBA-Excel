# Fechamento da Entrega 2 — continuação em 03/11/2026

**EM ANDAMENTO. NÃO HOMOLOGADA. A Entrega 3 não foi iniciada.** Compilação real aprovada; a nova sessão permanece bloqueada para execução de macros. O nome preferencial dos binários contém Final, mas não constitui aceite de fechamento.

## Referência preservada

Entrega 2 congelada, revisão 16: XLSM 856c1c148451bae188d5539b3a344371579275268d7475dd535e1ade89ee39d0; ACCDB 1390a79c2b36fcea13843bb211dcb56ffbfe3fe8523e2200f70f32e1dbc35581. Incluída a continuação de reabertura 20261103T010409Z. Os 39 originais foram novamente conferidos sem divergência em [integridade_originais.json](./integridade_originais.json).

Novos backups de manutenção pelo script do projeto: C:/Users/MARCOSALEXANDREPERDO/AppData/Local/CodexBackupsPrivados/VBA-Excel/20261103T214644Z-4a4b072c. O script aceita somente arquivos relativos ao repositório; os binários de laboratório em TEMP receberam checkpoint adicional por cópia exclusiva e conferência SHA-256, no mesmo armazenamento privado, em checkpoint_lab_20261103T2147Z. Nenhum checkpoint anterior foi descartado.

## Implementação efetiva desta continuação lógica

1. modFechamentoEntrega2: registro de autoridade externo compartilhado e vínculo por AuthorityID, BaseID, linhagem, geração e caminhos canônicos. O portão transacional do registro é conferido na aquisição e nas escritas protegidas por WriterLease. Não há eleição automática por nome ou cópia. A proteção cobre os clientes compatíveis ensaiados; não protege contra SQL externo ou cópias adulteradas do próprio registro.
2. Conciliação experimental: plano selado com estado anterior, responsável, justificativa e confirmação específica; nova tentativa e recibo persistidos. Somente intenção PREPARED não ENTRY comprovadamente sem efeitos pode ser encerrada como ABORTED. Movimentos confirmados não são desfeitos; evidências originais permanecem. A guarda Pending retirada é preservada no snapshot do recibo. Intenções ENTRY/rascunhos, lease órfão e estados incertos permanecem bloqueados.
3. Realocação explícita por C3_Realocar, com hash esperado, destino vazio, responsável e motivo. Foi executada em VBA, preservando revisão 16. A origem passou a RECONCILE/AUTHORITY_RELOCATED e seu diagnóstico retornou Autoridade revogada por geracao. O registro passou para geração 2. [Transcrição identificada da interface](../realocacao_observada.json).
4. Correção de portabilidade: as guardas comparavam o caminho com o TEMP variável da sessão. Uma nova sessão recebeu outro diretório e foi recusada. Passaram a validar o TEMP local canônico do usuário (LOCALAPPDATA/Temp), delimitando o separador, mantendo marcador sintético, nome da candidata, rejeição de OneDrive e vínculo exato da autoridade. Após a correção, reconstrução e diagnóstico reais retornaram AUDIT=OK;PROJECTION=VERIFIED;REV=16;MOV=16;RECEIPTS=16;ACTIVE=5;PENDING=0 na sessão aberta. Isso não substitui o reteste frio, agora bloqueado pelo Office.
5. Suíte modTestesFechamentoE2 ampliada e incorporada. Corrigida antes de executar a consulta do teste do filtro: KindText pertence a Movements, não a Projections. Não houve modificação de regras fiscais ou dos eventos dos cinco UserForms.

## Resultados realmente observados

O [TSV anterior à realocação](./resultados_fechamento_antes_realocacao.tsv) tem oito asserções PASS: instalação do contrato 7; estados Application; rollback dos fatos antes do commit; recusa de repetição pendente; recusa com outro OperationID para o mesmo episódio; exigência de confirmação; conciliação sem efeitos; replay idempotente do plano. Não somar essas asserções aos 62 resultados históricos como se fossem novos cenários operacionais completos.

A falha pré-commit foi provocada pelo evento real de Devolução, com FIC5C03, NF34001, valor450. Fatos, versão e recibos ficaram iguais; houve intenção pendente. A API de conciliação recebeu confirmação sintética da suíte, não uma decisão humana pela interface. Produziu um recibo de conciliação, sem alterar a revisão de fatos. O botão/diálogo completo de conciliação ainda não foi exercitado.

A compilação foi efetivamente acionada no Editor VBA após a ampliação e após a correção SQL da suíte. Não apareceu erro e Compilar VBAProject ficou desabilitado. [Evidência transcrita](./compilacao_consulta_corrigida.txt). Extração do XLSM salvo contou 71 componentes, 17 planilhas e cinco UserForms. [Comparação](./comparacao_fontes.json). Corpos canônicos dos cinco formulários e quantidades de procedimentos conferem com a E2 congelada; não é alegação de igualdade binária dos designers (52/127 streams idênticos, demais reserializados pelo Office). Nenhum controle foi criado ou renomeado.

O Office recusou o diagnóstico na nova sessão com a mensagem original: As macros neste projeto estão desativadas. [Bloqueio](./bloqueio_macros_nova_sessao.json). Não houve mudança de Trust Center, AccessVBOM, ExecutionPolicy, Registro ou habilitação automática.

Leitura ADO independente do banco persistido: revisão16, esquema7/contrato3, READY, geração2, WriterLease FREE, seis veículos, sete episódios, 17 operações (uma ABORTED), 16 movimentos/projeções/recibos/ChangeLog, 40 SheetProjections, 19 Attempts, zero Pending, um plano/recibo de conciliação e 24 ProjectionAcks. ProjectionState armazenado VERIFIED é um metadado persistido; não representa novo PASS funcional da sessão bloqueada. [Estado](./estado_persistido_atual.json).

A célula Entradas!E3 contém #N/A herdado da massa de teste. A origem foi localizada no modTestesAuditoria, que atribui TESTE_2 e CVErr(2042) à linha3. Ela fica antes da região de projeção iniciada na linha4 e não foi apagada nem tratada como prova de fórmula operacional quebrada. [Inspeção XML](./erros_cache_xml.json). A validação integral continua exigindo identidade, valores, fórmulas e formatos.

## Artefatos efetivamente preservados

Pacote privado: [pasta de preservação](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/VBA-Excel-Entregas/Entrega2_Fechamento_20261103T2147Z_PENDENTE_REGRESSAO/>). Não executar Office nessa cópia de preservação; utilizar a autoridade TEMP indicada abaixo.

| Arquivo | SHA-256 |
|---|---|
| Estoque_Fera_Veiculos_Candidato_Entrega2_Final.xlsm | 8b25d17ae74fe646f13a5d834d51382182d8579337633418ad68985325f98305 |
| Banco_Candidato_Entrega2_Final.accdb | f73f68ff144567d25e9af54a5f213cd1d95cf2583299bdb87301a8d3d67b8d14 |
| Controle_E2_143C55EC132B48398413268B265117AF.accdb | 267460b51138ac31ddb278571f5643a93eed8e7e68642fd84905d3fdcd29bd54 |

O registro de controle é dependência adicional, não outra fonte de fatos de estoque. A cópia não assume autoridade por ser reaberta ou renomeada. Fontes efetivamente embutidos e vbaProject.bin foram preservados; códigos de UserForms/documentos extraídos não são novos pares FRM/FRX importáveis. O designer completo permanece no XLSM e no vbaProject.bin; pares de referência anteriores permanecem intactos.

## Pendências essenciais

- Habilitação manual de macros para a cópia restaurada; regressão fria e eventos da suíte ampliada.
- Concorrência pelos gravadores após o novo portão; matriz completa de falhas da conciliação, plano obsoleto, órfãos e divergências.
- Restauração de backup quando a autoridade anterior estiver indisponível; não existe autorização para recriar cegamente registro ou lease. Relocação ao vivo ensaiada não equivale a recuperação após desastre.
- Regressão restante XML inválido/divergente, leitores reais, ADJUST e sete defeitos históricos na revisão final.
- Motivo exato de Workbook.Saved=False ainda não isolado. Workbook_Open chama carregamento de cache; hashes históricos não mudaram sem Save. Recalcular em memória é distinto de gravar no disco; não foi mascarado atribuindo Saved=True.
- Anulação/estorno, retificação documental/monetária/histórica, custódia, Cancelado fiscal e políticas de logs/retensão dependem das regras do responsável. Casos não aprovados continuam indisponíveis, não PASS de implementação.

Não há transação conjunta ACE/XLSM nem prova de recuperação de energia, corrupção física ou encerramento inesperado. Bloquear uma repetição não é recuperar uma operação. Os sete FAIL históricos permanecem preservados.

## Retomada sem reiniciar auditoria

Abrir somente a candidata em TEMP: C:/Users/MARCOS~1/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/Entrega2_Restauracao_Controlada_20261103T0146Z/Estoque_Fera_Veiculos_Candidato_Entrega2_Final.xlsm. O proprietário deve habilitar o conteúdo dessa cópia manualmente, sem modificar políticas globais. Antes de novos efeitos, conferir cache/diagnóstico e executar F2_Restaurada na revisão16.

Depois, por etapas e com conferência do banco a cada execução: F2_EntradasNovas, F2_Transferencia, F2_Devolucao, F2_Venda, F2_Cancelado, F2_ReentradaNova e F2_FiltrosReais. Os XML são sintéticos; relatórios modais devem ser observados e fechados legitimamente. Não repetir automaticamente um evento que tenha produzido estado incerto. Não reexecutar F2_FalhaEConciliacao indiscriminadamente: a intenção34001 já está ABORTED e seu registro é evidência preservada.

Conferência independente da preservação: 17 tabelas reabertas por ADO e comparadas integralmente ao estado capturado; conteúdo idêntico. O SHA-256 UTF-8 do recibo de conciliação também confere. Esses resultados não ativam a cópia como autoridade nem substituem execução VBA. [Verificação](./verificacao_copia_e_recibo.json).
