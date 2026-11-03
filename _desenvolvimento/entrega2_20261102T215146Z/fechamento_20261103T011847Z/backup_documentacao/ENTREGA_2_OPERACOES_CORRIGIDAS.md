# Entrega 2 — operações corrigidas na candidata sintética

Execução 20261102T215146Z, 02/11/2026. **Candidata incremental compilada e testada; NÃO HOMOLOGADA. A conclusão integral da Entrega 2 permanece condicionada às pendências abaixo.** Não houve modificação do XLSM original, integração em produção ou dados reais.

## Artefatos efetivamente produzidos

- [Candidata XLSM](./_desenvolvimento/entrega2_20261102T215146Z/entrega/Estoque_Fera_Veiculos_Candidato_Operacoes_Corrigidas.xlsm), SHA-256 856c1c148451bae188d5539b3a344371579275268d7475dd535e1ade89ee39d0.
- [Banco ACE sintético](./_desenvolvimento/entrega2_20261102T215146Z/entrega/Banco_Candidato_Operacoes_Corrigidas.accdb), SHA-256 1390a79c2b36fcea13843bb211dcb56ffbfe3fe8523e2200f70f32e1dbc35581.
- [Fontes finais extraídos](./_desenvolvimento/entrega2_20261102T215146Z/fontes/), [matriz de execução](./MATRIZ_REGRESSAO_ENTREGA2.md), [evidências](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/), [manifesto](./_desenvolvimento/entrega2_20261102T215146Z/entrega/manifesto_artefatos.json).

A cópia executada está no TEMP, no diretório Entrega2_20261102T215146Z do laboratório vscode-agent-codex-sandbox-JTJOQt. As cópias do workspace são preservação: não executar Office no OneDrive. AuthorityPath e BaseID impedem que uma cópia deslocada assuma autoridade. O [manual](./MANUAL_INSTALACAO_RESTAURACAO.md) informa o caminho executável e os limites de instalação.

## Base preservada e estado final

A revisão final da Entrega 1, incluindo continuacao_reabertura_20261102T205915Z, foi conferida: XLSM 5b2a0d55fcea077dd034ace4c4892bbc9fdd46c69760d7d318d2162f54771749, ACCDB 3198e07699f644f64c325a9fb517470b64950aa5a262cc354629ed61f1503dc3. Os 241 arquivos do manifesto anterior conferem. Os 39 originais e os dois artefatos finais E1 também mantêm hashes. [Conferência](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/verificacao_manifesto_entrega1.json).

A candidata conserva 69 componentes, 17 planilhas, cinco UserForms e os 18 módulos de documento/CodeNames. A comparação verificou 436 assinaturas existentes dos formulários, sem alteração. Há oito componentes com alterações funcionais nesta entrega e seis com normalizações de capitalização/atributos do VBE. Designers foram reserializados pelo Office: 127 streams com tamanhos iguais, 52 hashes iguais. Não se declara identidade binária nem equivalência integral de todas as propriedades; os pares FRM/FRX de referência permanecem preservados e o vbaProject.bin final está incluído.

ACE 16, esquema 6, contrato 2, BaseID 143C55EC-132B-4839-8413-268B265117AF, revisão 16. Consulta ADO direta: seis veículos, sete episódios, 16 movimentos, 16 recibos, 17 tentativas, 40 SheetProjections, zero Pending, WriterLease FREE, WriteState READY e ProjectionState VERIFIED. O último diagnóstico VBA foi AUDIT=OK;PROJECTION=VERIFIED;REV=16;MOV=16;RECEIPTS=16;ACTIVE=5;PENDING=0. [Estado persistente](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/ace_final_revisao16.json).

## Funcionalidades e correções

| Componente | Implementação e causa tratada | Evidência / limite |
|---|---|---|
| modCandidatoIntegrado | Contexto por tipo, ObjPtr e página; chave por instância. Elimina colisão simples de contexto entre páginas/formulários. | Contextos múltiplos simultâneos não foram ensaiados integralmente. |
| Cad_NFE / camada C2 | Compra, Reentrada e Novo veículo nos controles existentes. Pesquisa de reentrada lista episódios encerrados; seleção explícita transporta VehicleID e cria novo EpisodeID. Placa conhecida sem vínculo é recusada. | E2_REENTRADA_CONTROLES e IDENTIDADE_CONHECIDA_SEM_VINCULO. O teste escolheu a ListBox real, sem injetar VehicleID. Novo veículo com placa coincidente/evidência humana ainda requer regressão. |
| Cad_NFE / C2_AlterarEntrada | Alteração por ViewID/EpisodeID/PID/versão; ADJUST cria fato substitutivo com motivo/responsável e mantém o anterior imutável. | Retificação AZUL e VERDE no mesmo episódio, recibos e leitor conferidos. Limitada a metadados; documento, valores, placa/data/chave e históricos encerrados continuam bloqueados para regra específica. |
| Exclusões dos formulários | Classificação explícita de cadastro confirmado, histórico/baixado e pendência. Sem Delete ou ClearContents como baixa. | Botões reais recusaram anulação sem regra e exclusão histórica. Isso é bloqueio seguro, não exclusão implementada para todos os casos. |
| Cad_NFE_Saida | Operação definida antes dos campos e restauração da origem por token antes/depois de Change. | Segunda transferência e devolução após transferência passaram; custo, episódio e custódia conferidos. |
| modCandidatoIntegrado | Formatação de seleção/envio passou de duas para quatro casas. O ACE tinha 123,4567, mas a interface arredondava para 123,46. | Dois FAIL reproduzidos e duas novas verificações PASS com 123,4567. Não houve alteração de regra fiscal. |
| modPersistenciaFase10 / modCenariosFase10 | Contrato 6 admite ADJUST e CANCEL, auditoria, ponteiro atual e transições coerentes; esquema anterior preservado antes da migração. | Auditoria final OK. DDL em etapas não possui garantia transacional integral. |
| Cancelado matriz vinculado | Duas projeções de saída, valores zerados conforme ramo legado e origem permanece ativa. Sem XML nessa hipótese. | Botão real e dois históricos conferidos. Cancelado sem origem/filial permanece bloqueado; não há homologação fiscal. |
| UserFormFiltro / pesquisa | Compra e históricos por IDs, histórico apresentado uma vez, token revalidado, negativos e seleção obsoleta tratados. | Pesquisa e ListBox reais; teste do filtro chamou C2_Filtrar diretamente e depois eventos de seleção/exclusão. Não aprova todos os botões do filtro. |
| Diagnóstico / backup | Consulta de Pending e rascunhos órfãos sem mutação; botão Salvar Tudo usa backup pareado C1. | Backup reaberto, auditado e lease livre. Não há decisão automática de conciliação. |

O primeiro erro da migração/reconstrução revelou MONEY_FORMAT: a verificação de HISTORY_OUT havia perdido o ramo de data; foi restaurado. O tratamento de erro passou a conferir State antes de fechar a conexão. A primeira retificação confirmou no ACE, mas não reconstruiu por chamada interpretada como rótulo VBA; Call explícito corrigiu. Os resultados falhos estão preservados.

## Execução e evidências

A referência R0 foi executada em base E2 nova: dez asserções PASS de Entrada/Venda/Transferência/Devolução pelos controles existentes. A continuação tem 61 execuções de asserção: 54 PASS e sete FAIL preservados, 55 IDs distintos. Repetições não são novos cenários. Os FAIL incluem três erros do aparato/metodologia (consulta de coluna inexistente após devolução, conversão de snapshot antes de executar o evento e visibilidade de alterações não confirmadas em duas guardas), além das falhas técnicas de retificação/precisão corrigidas. A [matriz](./MATRIZ_REGRESSAO_ENTREGA2.md) conserva cada tentativa e a modalidade efetiva.

A compilação real final foi acionada no Editor VBA após as correções. O menu voltou com Compilar VBAProject desabilitado e nenhum erro. [Registro da interface](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/compilacao_final_R4.txt). Não se usou análise estática como compilação. O projeto salvo foi extraído e contado novamente.

A reabertura interativa ficou inicialmente BLOQUEADA no aviso de macros. Após habilitação manual pelo usuário, a inicialização automática e a suíte VBA de reabertura passaram com revisão16/projeçãoVERIFIED. O adendo abaixo contém a execução e a evidência; o bloqueio anterior permanece preservado no checkpoint. Gerações de projeção abertas pelo próprio VBA são uma modalidade distinta desse teste interativo.

## Integridade e recuperação

Falha controlada durante baixa da devolução e antes do commit reverteu fatos, quatro destinos, versão e custódia. A intenção persistiu em Pending e a repetição foi recusada. As bases forenses foram preservadas antes de retomar outra ramificação de laboratório por checkpoint conhecido; isso não constitui conciliação operacional nem liberação automática de Pending.

Falha durante projeção após commit preservou o fato e o recibo; reconstrução pela autoridade recuperou o Excel. Replay do mesmo evento não duplicou o movimento. Save cancelado pelo evento real BeforeSave e SaveAs inválido foram reproduzidos. Erro injetado após commit reconheceu o recibo e reconstruiu sem repetir. Recibo removido artificialmente em base sintética provocou bloqueio de auditoria; não foi recriado automaticamente.

Não há transação conjunta ACE/XLSM, rollback garantido após commit, nem prova contra energia, corrupção física ou SQL externo. As falhas da classe de compensação legada permanecem históricas; os fluxos ACE usam outro mecanismo. Os sete FAIL oficiais anteriores não foram alterados retroativamente.

## Pendências para concluir a Entrega 2

1. Executar os ensaios de instalação/restauração em outra pasta; a reabertura no laboratório existente passou, mas o pacote não é portátil por simples cópia.
2. Implementar o encerramento/retomada de intenções mediante decisão auditável de conciliação; o diagnóstico atual é somente leitura.
3. Completar regressão dos leitores/relatórios remanescentes, controles não exercitados, XML inválido/divergente, todas as combinações de saída inválida e concorrência dos gravadores integrados.
4. Decidir anulação/estorno de fatos confirmados e retificação documental/monetária/histórica. Manter esses casos bloqueados.
5. Validar operacionalmente episódio/custódia, Cancelado, política de logs/retenção e limites de recuperação.

A candidata é software executável produzido, não somente relatório. Seus resultados representam os fluxos e fronteiras descritos; não justificam declarar a aplicação operacionalmente completa ou homologada. [Defeitos abertos](./REGISTRO_DEFEITOS_ABERTOS.md) e [restauração](./MANUAL_INSTALACAO_RESTAURACAO.md) delimitam o uso.


## Continuação — reabertura interativa aprovada (02/11/2026,22:06)

Depois da habilitação manual das macros pelo usuário, o projeto E2 foi selecionado explicitamente no VBE, pois a Entrega1 também estava aberta. Antes de qualquer carregamento manual, C1_UltimoResultado retornou CACHE_AUDITADO junto ao caminho completo da candidata E2 no TEMP. A execução real de C1_Diagnostico retornou AUDIT=OK;PROJECTION=VERIFIED;REV=16;MOV=16;RECEIPTS=16;ACTIVE=5;PENDING=0.

A suíte existente E1_SessaoReaberta foi então executada no projeto E2 e registrou E1_SESSAO_REABERTA_E2_R3 PASS no TSV. É um teste VBA auxiliar de cache/fórmulas, sem movimentação e sem acionamento de botão de estoque. O rótulo genérico do logger não altera essa modalidade. Conferidos17planilhas, EnableEvents/ScreenUpdating/DisplayAlerts=True e Calculation=-4105. Workbook.Saved=False após inicialização/cálculo; o arquivo não foi salvo novamente nesta conferência.

Os SHA-256 do XLSM e ACCDB permaneceram exatamente iguais aos entregues. Não foi necessário alterar fontes ou repetir operações. A base continua revisão16,16movimentos/recibos,5ativos e0Pending. [Evidências da continuação](./_desenvolvimento/entrega2_20261102T215146Z/continuacao_reabertura_20261103T010409Z/), [resultado estruturado](./_desenvolvimento/entrega2_20261102T215146Z/continuacao_reabertura_20261103T010409Z/evidencias_reabertura.json) e [TSV acumulado](./_desenvolvimento/entrega2_20261102T215146Z/continuacao_reabertura_20261103T010409Z/resultados_entrega2_com_reabertura.tsv).

O log acumulado agora contém62execuções de asserção:55PASS/7FAIL preservados,56IDs. A referênciaR0 de10PASS continua separada. As consultas imediatas de inicialização e estados não foram somadas como novas asserções. O bloqueio de reabertura E2-09 foi superado nesse escopo. Portabilidade/restauração em outra pasta, conciliação operacional e cobertura remanescente continuam abertas; sistema NÃO HOMOLOGADO. A instância também mantém outros projetos; não se declara teste de nova instalação, crash ou energia.
