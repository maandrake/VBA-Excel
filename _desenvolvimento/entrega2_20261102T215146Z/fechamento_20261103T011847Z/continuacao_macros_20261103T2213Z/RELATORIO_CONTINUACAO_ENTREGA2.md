# Fechamento da Entrega 2 — continuação de 03/11/2026

**EM FECHAMENTO; NÃO HOMOLOGADA. Entrega 3 ainda não liberada.** A revisão26 existe como XLSM executável compilado e ACCDB integrado. A reabertura final em nova instância foi impedida pelo aviso de macros desabilitadas; o clique deve ser manual. Não confundir essa pendência final com o teste frio inicial aprovado na revisão16 após a realocação.

## Preservação

A Entrega2 congelada permanece com XLSM856c1c148451bae188d5539b3a344371579275268d7475dd535e1ade89ee39d0 e ACCDB1390a79c2b36fcea13843bb211dcb56ffbfe3fe8523e2200f70f32e1dbc35581. Conferidos39originais sem divergência. [Manifesto original](./integridade_originais_final.json), [referência congelada](./integridade_entrega2_congelada.json). Backups pelo script do projeto:20261103T221904Z-c2ced208 e20261103T230700Z-9ddfa761 em C:/Users/MARCOSALEXANDREPERDO/AppData/Local/CodexBackupsPrivados/VBA-Excel/. Checkpoints binários adicionais em TEMP/armazenamento privado, pois Backup-Manutencao aceita somente entradas relativas ao repositório. Nenhum original ou backup anterior foi sobrescrito.

## Estado persistido

Leitura independente ADO/ACE do banco, depois dos eventos: esquema7/contrato3, BaseID/linhagem143C55EC-132B-4839-8413-268B265117AF, geração2, revisão26, READY, WriterLeaseFREE, ProjectionStateVERIFIED. O campo persistido VERIFIED foi complementado por diagnóstico real do VBA com AUDIT=OK;PROJECTION=VERIFIED;REV=26;MOV=26;RECEIPTS=26;ACTIVE=7;PENDING=0.

| Entidade | Quantidade |
|---|---|
| Meta | 1 |
| WriterLease | 1 |
| CandidateSettings | 1 |
| AuthorityBinding | 1 |
| Operations | 28 |
| Episodes | 10 |
| Vehicles | 8 |
| Movements | 26 |
| Projections | 26 |
| SheetProjections | 66 |
| Receipts | 26 |
| ChangeLog | 26 |
| Attempts | 32 |
| Pending | 0 |
| ReconciliationPlans | 3 |
| ReconciliationReceipts | 2 |
| ProjectionAcks | 37 |

[Estado completo antes/depois dos cenários](./estado_apos_xml_concorrencia.json), [estado intermediário após ciclo](./estado_apos_ciclo.json). As tabelas conservam valores tipados, payloads, tentativas, recibos e snapshots. Não se apagaram as duas operações ABORTED ou seus recibos de conciliação.

## Execução real

Esta continuação contém58execuções de asserção:57PASS/1FAIL preservado, em três logs. São asserções, não58fluxos completos ou58cenários independentes. Oito PASS da sessão anterior do fechamento permanecem separados. As62execuções históricas da Entrega2 não foram reclassificadas.

-14PASS em resultados_fechamento: reabertura fria da realocação na revisão16, aliases8.3/longo, entradas matriz/filial, duplicidade, Transferência, Devolução, Venda, baixa lógica, Cancelado, reentrada e três botões reais do UserFormFiltro.
-16PASS em resultados_entrega2: ADJUST matriz, imutabilidade, leitor, exclusão classificada, Save cancelado por BeforeSave real, reconstrução, backup reaberto, cinco entradas inválidas, placa conhecida sem vínculo, reordenação detectada e reconstruída, SaveAs inválido.
-28execuções em resultados_final_e2:27PASS/1FAIL. Falha APPLY3 real, plano obsoleto recusado, ADJUST filial discriminante, diálogo completo de conciliação, replay abortado recusado, nova tentativa após conciliação, metadados por identidade, tokens, cópias de banco recusadas, dez XML e concorrência.

A [matriz desta sessão](./MATRIZ_EXECUTADA_FECHAMENTO.md) contém cada ID, modalidade, esperado, observado, horário e fixture. As suites acionam eventos privados pelos controles reais existentes; isso não equivale a um usuário clicar manualmente cada botão. O diálogo de conciliação e os relatórios XML foram manipulados pela interface real e suas transcrições estão identificadas.

## Dados discriminantes e integridade

Entrada matriz F2A0001/NF12001/123,4567; filial F2F0002/NF12002/456,7891/chave44iniciada emzero. Transferência24002/300,1111 →Devolução34011/300,1111 →Venda24012/400,2222. Mesmo VehicleID/EpisodeID durante custódia experimental; venda encerrou episódio. Reentrada12003/150,3333 foi escolhida pela ListBox real e gerou novo EpisodeID para o mesmoVehicleID, sem VID injetado. Cancelado24990 manteve origem e produziu os dois históricos do ramo existente. ADJUST filial mudou BRANCO para AMARELOQA, preservando fato anterior e criando recibo/fato substitutivo.

Conferidas66SheetProjections por PID→EP→VID, incluindo chaves String44, NumberFormat@ e Ignore. Leitores conferiram token obsoleto, pesquisa vazia e seleção ListIndex=-1; o último ensaio é restrito à ausência de navegação imediata ao atribuir-1, não prova geral de autorização de escrita por esse estado. Não se declarou equivalência integral de todos os relatórios das17abas.

## Conciliação auditável

APPLY3 durante baixa da Devolução34021 reverteu fatos, quatro destinos, custódia e versão; manteve intenção Pending. Um ADJUST em outro episódio elevou revisão24→25 e invalidou plano anterior. O resolvedor recusou o plano obsoleto sem liberar Pending. CandidatoConciliarIntencao foi acionado e preenchido pela interface com responsável/motivo sintéticos, confirmação explícita Sim e novo plano. ResultadoCONCILIADO_SEM_EFEITOS: intenção ABORTED, recibo de conciliação persistido, sem inventar movimento. Replay do OperationID abortado foi recusado; somente depois a nova tentativa34021 foi confirmada uma vez, revisão26 e quatro destinos coerentes.

OperationIDAA80FD0A-5A9A-4113-921C-12C9046D1A26; EpisodeID739E3127-121B-4700-890D-98C1892B3121. [Diálogo transcrito](./conciliacao_interface.json). A decisão foi executada pelo agente no laboratório autorizado, não é aceite operacional de pessoa real. PREPARED/ENTRY com rascunho, recibo ausente, histórico divergente, commit incerto e lease órfão não são liberados por esse procedimento. Bloqueio não é recuperação.

## XML e concorrência

Dez casos recusados sem aumentar revisão, versão ou recibos. Malformado/númeroinválido/chaveinválida falharam no carregador com MsgBox antes de abrir relatório; o rótulo genérico do logger não implica relatório nesses três. Operação incompatível, valor, tributo ausente e somatório foram rejeitados com frmRelatorioXML real. Comparação de número/chave divergentes chamou o validador real diretamente com DOM e formulário: o carregador normal preenche esses campos a partir do XML, portanto não se afirma execução do gravador nesses dois testes. XML válido duplicado passou no relatório e foi recusado no gravador. [Transcrições](./xml_interface_observada_final.json). Regras fiscais e tolerâncias não foram alteradas; documentos são fictícios e sem comprovação fiscal.

O evento real GravarFilial enfrentou lock transacional ACE mantido por outro processo e foi recusado sem criar CFCC001/NF12061 ou alterar fatos. O processo externo fez rollback, OwnerFinalnull e geração2. [Evidência](./concorrencia_processo_externo.json). **Não foram duas instâncias de Excel gravando**; o ensaio cobre um gravador VBA contra processo ADO independente. A mensagem de conflito retornou erro de transação pouco claro e requer diagnóstico melhor. Não prova todos os interleavings ou resistência a SQL externo.

## Correções da instrumentação e compilação

Novo modRegressaoFinalE2 incorporado e compilado no Editor VBA. O FAIL CF_LEITORES_ERRO1004 ocorreu ao ler Ignore em folha protegida pelo teste. Corrigido somente o helperCI com restauração da proteção; nova CF_PROJECOES_IDENTIDADE_TIPOS passou66projeções. O produto já continha IgnoredC1 para esse caso. Primeira preparação XML foi interrompida normalmente porque o provedor de teste era definido depois da seleção e abriu formulário modal; nenhum XML/evento de escrita havia sido executado. Ordem corrigida somente na suíte e tentativa preservada.

Projeto salvo:72componentes/17planilhas/5UserForms. Corpos dos cinco forms e436assinaturas de procedimentos permanecem iguais à referência anterior do fechamento, excluído somente GUIDVB_Base reserializado pelo Office; os demais71componentes são iguais nesse critério, acrescido o módulo de teste. Não se afirma igualdade binária dos designers. [Inventário](./inventario_final_salvo.json), [comparação](./comparacao_72_componentes_final.json), [formulários](./preservacao_formularios_final.json), [compilação real observada](./compilacao_final_observada.txt). Nenhum gravador/regra fiscal foi alterado nesta sessão.

## Saved e reabertura final

Após Save real:SavedTrue; depois CandidatoCarregarCache:True; depois C1_Diagnostico:False. VerifyC1 usa IgnoredC1 e Unprotect/Protect para ler Errors(xlNumberAsText).Ignore, alterando estado em memória. SHAantes/depoisfd35ca202643e679d4ac3fdf41566743b1026f7d8dcfe985a3dd15d2b0278962idêntico, sem gravação inesperada. Não atribuído Saved=True para esconder mudança. [Experimento](./investigacao_saved.json).

Novo Save explícito final,17abas eSavedTrue. Havia somenteuma pasta aberta; a instância de teste foi encerrada normalmente porApplication.Quit, sem processo do usuário encerrado. Nova instância aberta e candidata exata reaberta; Office exibiu macros desabilitadas. Inicialização automática da revisão26 e fluxo pós-reabertura estão **BLOQUEADOS**, aguardando clique manual. [Antes do encerramento](./antes_sessao_fria_final.json).

## Artefatos produzidos

[Pacote privado](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/VBA-Excel-Entregas/Entrega2_R26_20261103T2310Z_PENDENTE_ACEITE/>), somente preservação; executar Office apenas na autoridade TEMP autorizada.

| Arquivo | SHA-256 |
|---|---|
| Estoque_Fera_Veiculos_Candidato_Entrega2_Final.xlsm | b379f42c739100086e1010f776777a90a9acb65d3ef94d8227dcbd1dd45e6e25 |
| Banco_Candidato_Entrega2_Final.accdb | 9fc0bfe8001120d2804b8f56836949249606e2030a6af431a71b1757c80267dc |
| Controle_E2_143C55EC132B48398413268B265117AF.accdb | b9e0c8aceb4474bb6b74ba727c7064982cfa9371e060b901bdf4b2de61632e2f |

Fontes72CP1252/CRLF, vbaProject.bin exato, paresFRM/FRX de referência e manifesto foram incluídos. Reimportação desses pares/textos em projeto novo não foi ensaiada; o XLSM já contém o projeto compilado. O nomeFinal não representa aceite. Cópia do banco/registro não estabelece autoridade; não abrir as cópias privadas como instalação operacional.

## Pendências essenciais e sequência de conclusão

1.Habilitar conteúdo manualmente na candidata TEMP em nova sessão, registrar C1_UltimoResultado antes de qualquer loader manual, diagnosticar e executar consulta funcional pelo controle real; conferir estadosApplication e hashsemSave.
2.Ensaiar recuperação integral de backup em outro local quando origem/registro indisponíveis; a realocação ao vivo aprovada e o backup reaberto são garantias diferentes. Não restaurar backup24 sobre fatos26.
3.Ampliar concorrência para dois gravadores Excel integrados e perda de autorização; o lock ADO atual é prova parcial.
4.Completar exceções críticas de identidade/novoveículocomplacacoincidente e relatórios não exercitados; manter hipóteses operacionais explícitas.
5.Anulação/estorno/retificação fiscal,monetária,histórica e órfãos deENTRY continuam NÃO DISPONÍVEIS sem regra e mecanismo aprovado.

Os cinco critérios demonstrados não autorizam declarar a Entrega2 encerrada com essas pendências essenciais. Não iniciado novo ciclo de auditoria, fase ou migração real. Nenhum commit,push,instalação,alteração de segurança ou promoção.
