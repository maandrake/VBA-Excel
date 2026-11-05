# Retomada controlada — 05/11/2026

## Estado confirmado

- Git: main, 5911e89, limpo antes desta execução. AGENTS vigente relido; Ponytail 5.3.3 Full disponível e aplicado. Nenhum commit, push ou alteração de segurança.
- Última etapa validada continua sendo a reabertura VBA real e os eventos QueryClose de dois formulários em 04/11: [resultado anterior](../continuacao_vba_real_20261104T185156Z/RESULTADO_REABERTURA_VBA_REAL.md). Nenhum gravador foi executado na presente etapa.
- XLSM encontrado com SHA-256 cb610aa3e19363d4b66e32a219f33f541b5907c341f122096b644d019f96c9c4, salvo em 04/11/2026 22:21:59 UTC. Não foi restaurado o hash anterior 42b34e7a2632bb4875525b5b1d5de798274bbeda3402153a79bad4e8d0ba1853.
- Os 72 corpos VBA embutidos não mudaram. Nas oito abas divergentes, os XMLs diferem somente no hash e salt da proteção; as células não mudaram. A normalização ocorreu somente na comparação em memória. O arquivo não foi editado. Streams binários de alguns formulários e caches VBA também diferem; isso NÃO equivale à validação funcional ou à prova de FRX idêntico. Não se atribuiu autoria ao salvamento. [Comparação](comparacao_xlsm_salvo.json).
- ACCDB R26 e registro de autoridade continuam com os hashes 9fc0bfe8001120d2804b8f56836949249606e2030a6af431a71b1757c80267dc e b9e0c8aceb4474bb6b74ba727c7064982cfa9371e060b901bdf4b2de61632e2f.
- Checkpoint verificado dos três arquivos em C:/Users/MARCOSALEXANDREPERDO/AppData/Local/CodexBackupsPrivados/VBA-Excel/checkpoint_ACE_20261105T101001Z — [manifesto da preparação](verificacao_retomada.json).

## Laboratório independente produzido

C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/Entrega2_Restauracao_Controlada_ACE_20261105T101001Z_A2

- XLSM: Estoque_Fera_Veiculos_Candidato_Entrega2_Final.xlsm.
- Banco: Banco_Candidato_Entrega2_Final.accdb. Registro: Controle_Laboratorio_ACE.accdb.
- BaseID: 8C04F762-BA01-4B87-A968-4A41315133C4; linhagem: 5DA22109-CFE3-4171-8C49-8DD108468DB1; AuthorityID: AD971B45-2343-437C-A776-886EFA51CA6C. São identidades novas, e não uma segunda autoridade da R26.
- Esquema 7, contrato existente preservado, revisão 0, lease FREE, zero veículos/episódios/movimentos/recibos/Pending. Projeção PENDING; as abas ainda contêm a cópia sintética anterior e exigem reconstrução explícita antes de qualquer ensaio de escrita. Não interpretar a projeção inicial como atualizada.
- [Preparador reproduzível](Preparar-LaboratorioACE.ps1), [resultado e hashes](laboratorio.json), [estado anterior ao VBA](banco_antes_vba.json). O preparador recusa diretório existente e qualquer checkpoint com hash divergente. Executado com assertions.
- Primeira tentativa, sem sufixo A2, falhou porque Meta.IntegrityReason não aceita Null. A transação foi revertida: 26 movimentos e 26 recibos permaneceram na cópia descartável, revisão 26; nenhuma inicialização foi confirmada. Evidência preservada em [snapshot A1](preparacao_A1_rollback.json). A correção do script conserva o valor existente em vez de inventar outro valor. A1 não é uma autoridade instalada e não deve ser aberta no Office.
- Comparação DAO somente leitura: regras, índices e 28 relações das 18 tabelas iguais à referência. [Evidência](comparacao_esquema.json), [inventário](esquema_laboratorio.json). Não é compilação VBA.

## Bloqueio efetivamente observado

Instância de teste criada nesta execução: PID 44464, janela 6883988, sem outra instância Excel no início. A cópia A2 foi aberta pela interface; o Office mostrou “AVISO DE SEGURANÇA — As macros foram desabilitadas”. Não houve clique automático em Habilitar Conteúdo, alteração de Trust Center ou sondagem VBProject.

Ação mínima do usuário: clicar em Habilitar Conteúdo apenas na candidata A2 já aberta e informar Habilitado. A habilitação anterior de outro caminho não foi presumida para esta cópia.

## Próximo ensaio dirigido — ainda NÃO EXECUTADO

1. Confirmar ThisWorkbook.FullName, PID e que só existe a fixture A2 na instância. Confirmar BaseID/linhagem e diagnóstico; executar CandidatoReconstruir somente em A2 para construir a projeção vazia. Guardar antes/depois.
2. No Editor VBA, colocar breakpoint na primeira instrução do handler Bad de C3_AdquirirPortao, antes de On Error Resume Next ou RollbackTrans. A chamada é Cad_NFE.CommandButton_Gravar_Click → C1_GravarEntrada → Acquire10 → C3_AdquirirPortao → Exec10/ADODB.Command.Execute.
3. Preparar os controles reais usando [comandos de arranjo](preparacao_formulario_immediate.txt); conferir nomes no designer antes de preencher. O arranjo não executa o gravador. Acionar o botão real para executar o evento.
4. Reutilizar Segurar-Registro-Sintetico.ps1 da continuação de 03/11, apontando EXCLUSIVAMENTE para o registro A2. Ele mantém a escrita sem commit por até 180 segundos e reverte ao liberar. Não usar o registro R26.
5. No breakpoint, capturar Err.Number, Description, Source, etapa, who, estado de R/c e cada item de R.Errors (Number, Description, Source, NativeError, SQLState) ANTES de executar qualquer limpeza. Não executar BeginTrans para sondar a transação. A fase ativa deve ser inferida pelas fronteiras efetivamente alcançadas e pelo resultado de cada Begin/Commit/Rollback. Neste caminho, OperationID/AttemptID ainda não foram criados; não inventar IDs.
6. Acompanhar a propagação por Acquire10 e C1_GravarEntrada, registrando também qualquer erro secundário de rollback, a mensagem final e o fechamento das conexões. Não confundir 3218, 3034 e estado incerto.
7. Liberar o processo de lock pelo protocolo da flag; comparar movimentos/recibos/Pending/WriterLease/AuthorityControl e testar a operação subsequente somente se não houver estado incerto.
8. Corrigir apenas a causa demonstrada; então compilar no Editor e executar a regressão afetada. Uma correção em Acquire10 alcança seus escritores, conciliação e realocação; uma correção no portão alcança todos os escritores protegidos. Não aplicar patch geral de Resume Next.

## Classificação

| Verificação | Modalidade | Resultado | Evidência |
|---|---|---|---|
| Preservação dos três arquivos encontrados | SHA-256 / cópia verificada | PASS | verificacao_retomada.json |
| Comparação dos 72 corpos VBA e células | Inspeção binária/textual, não compilação | PASS no escopo descrito | comparacao_xlsm_salvo.json |
| Inicialização sintética A2 | Script ADO, não macro | PASS | laboratorio.json, banco_antes_vba.json |
| Esquema/índices/relações preservados | DAO somente leitura | PASS | comparacao_esquema.json |
| Gravador real e cadeia original do conflito | Evento VBA real | BLOQUEADO por macros desabilitadas | observação da interface |
| Nova compilação VBA | Editor VBA | NÃO EXECUTADO | depende da liberação e de alteração justificada |
| E2-FECH-08, E2-11, dois Excel e restauração sem origem | Ensaios dirigidos | NÃO EXECUTADO nesta etapa | nenhuma conclusão nova |

E2-FECH-06 permanece aberto; E2-FECH-07 mantém a classificação anterior nos dois formulários. Nenhum PASS de preparação substitui evento funcional. Entrega 2 não concluída; sistema não homologado. Os sete FAIL históricos não foram reclassificados.
