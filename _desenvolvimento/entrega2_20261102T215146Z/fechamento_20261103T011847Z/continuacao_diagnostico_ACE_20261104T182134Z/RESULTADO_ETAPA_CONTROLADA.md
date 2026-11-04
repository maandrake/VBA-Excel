# Entrega 2 — diagnóstico dirigido de conflito ACE

Execução em 04/11/2026, 18:21–18:24 UTC. Ponytail 5.3.3 Full. Etapa limitada ao defeito E2-FECH-06; sem alteração de VBA, esquema, fatos ou arquivos operacionais. Entrega 2 permanece em fechamento, NÃO HOMOLOGADA.

## Retomada e preservação

Raiz Git confirmada: VBA-Excel, branch main, commit b7b9d99; árvore inicialmente limpa. AGENTS.md vigente lido e preservado. A pasta continuacao_habilitada_20261104T1713Z existe, mas estava vazia, inclusive sem arquivos ocultos ou subdiretórios. Nenhum resultado foi deduzido dessa pasta.

O último avanço funcional comprovado continua sendo a correção mínima de QueryClose nos dois formulários, sua compilação e os testes reais anteriores. A reabertura funcional da revisão corrigida ainda não foi comprovada nesta execução.

A interface mostrou apenas Estoque Fera Veículos 2026.xlsm, o arquivo operacional. Nenhum comando foi enviado a essa janela. O teste independente abaixo utilizou somente ADO e cópias descartáveis no TEMP; não automatizou Excel ou VBProject.

[Checkpoint e caminhos](./checkpoint.json): cópias do XLSM, ACCDB e registro compartilhado, com SHA-256 conferido antes do ensaio. Checkpoint privado: `%LOCALAPPDATA%/CodexBackupsPrivados/VBA-Excel/checkpoint_diagnostico_ACE_20261104T182134Z`. Backup oficial dos sete documentos antes de acrescentar estes resultados: `20261104T182134Z-dd3488e7`. A primeira versão do script foi preservada pelo mesmo mecanismo em `20261104T182309Z-d41f2c51` antes de acrescentar a segunda sequência.

As cópias Registro_Descartavel.accdb não foram instaladas, realocadas nem eleitas autoridades. Nenhum cliente Excel foi configurado para utilizá-las. Os campos de caminho copiados não foram seguidos pelo script.

## Teste executado e resultado

Procedimento: [Testar-DiagnosticoACE.ps1](./Testar-DiagnosticoACE.ps1). Provedor real Microsoft.ACE.OLEDB.16.0, duas conexões ADODB no mesmo processo PowerShell, UPDATE parametrizado na linha AuthorityControl ID=1. CommandTimeout do comando fixado em dois segundos para limitar o ensaio; não representa reprodução integral do Cmd10 nem de seus timeouts no Excel.

| Verificação única | Esperado | Observado | Resultado |
|---|---|---|---|
| ACE_BASELINE | Uma linha atualizada sem concorrente; rollback | Uma linha, rollback concluído | PASS |
| ACE_CONFLITO_RECUSADO | B não escreve enquanto A mantém a linha bloqueada | Erro ACE, SQLState 3218, HRESULT -2147467259, cerca de 2,20 s | PASS |
| ACE_NOVA_TRANSACAO | Nova transação pode adquirir a linha depois do rollback de A | Uma linha atualizada e revertida | PASS |
| ACE_CONFLITO_SEM_COMMIT | A sequência de erro usada pela candidata também recusa B | SQLState 3218 | PASS |
| ACE_LIMPEZA_SEM_COMMIT | Rollback imediatamente após UPDATE recusado funciona | Rollback concluído, sem erro | PASS |
| ACE_ESTADO_RESTAURADO | Todos os campos da linha iguais antes/depois | Igualdade completa do JSON dos campos | PASS |

São seis verificações únicas COM/ADO, não seis eventos VBA nem seis fluxos de estoque. A [primeira tentativa](./resultado_inicial.json) executou quatro delas; a [segunda](./resultado_limpeza.json) repetiu essas quatro e acrescentou duas para conferir a limpeza. Não somar as tentativas como dez cenários diferentes.

Mensagem original capturada antes de qualquer limpeza: “Não foi possível atualizar; atualmente bloqueado.” SQLState 3218, erro nativo -72156238, origem Microsoft Access Database Engine.

Uma sondagem deliberada tentou CommitTrans após o UPDATE recusado e, em seguida, RollbackTrans. O commit da transação de B sem efeitos teve sucesso; o rollback posterior retornou SQLState 3034/HRESULT -2147168242: “Você tentou confirmar ou repetir uma transação sem antes iniciá-la.” Esse erro era esperado pela ordem artificial da sondagem. **Não reproduz o fluxo de erro da candidata e não prova falha física de commit.**

Na sequência equivalente à candidata, UPDATE recusado → RollbackTrans, o rollback funcionou. Portanto não há evidência para atribuir E2-FECH-06 simplesmente a esse rollback ou aplicar uma correção especulativa. Acquire10 e C3_AdquirirPortao já conservam número e descrição antes da limpeza. A origem da mensagem exibida no evento real continua NÃO ISOLADA; o defeito não foi encerrado.

## Integridade após execução

[Hashes antes/depois](./integridade_apos.json): os três arquivos ativos permaneceram idênticos.

| Arquivo | SHA-256 |
|---|---|
| XLSM Final | 42b34e7a2632bb4875525b5b1d5de798274bbeda3402153a79bad4e8d0ba1853 |
| ACCDB Final | 9fc0bfe8001120d2804b8f56836949249606e2030a6af431a71b1757c80267dc |
| Registro de autoridade | b9e0c8aceb4474bb6b74ba727c7064982cfa9371e060b901bdf4b2de61632e2f |

[Leitura persistente ADO](./banco_apos.json): esquema7, revisão26, READY, WriterLease FREE, 26movimentos, 26recibos, 66projeções, Pending0. Essa leitura não executa Audit10, inicialização ou compilação VBA.

[Conferência estática dirigida](./fontes_conferidas.json): 72componentes encontrados no XLSM; modFechamentoEntrega2, modPersistenciaFase10 e modCandidatoIntegrado correspondem às fontes preservadas do pacote QueryClose. Foi conferido apenas esse escopo, sem repetir a auditoria geral. O hash42b… é uma gravação posterior à referência efd…; não foi revertida nem sobrescrita.

## Próxima atividade e bloqueio concreto

Próximo teste: na candidata isolada, verificar ThisWorkbook.FullName, cache automático e C1_Diagnostico; depois reproduzir E2-FECH-06 pelo botão real, capturando etapa, Err.Number/Source e ADODB.Errors antes da limpeza. Somente então decidir a menor correção e compilar/regredir os chamadores afetados.

Nesta execução não existe janela da candidata de laboratório. Para disponibilizá-la sem tocar no Excel operacional: iniciar manualmente `excel /x`, abrir o XLSM Final no diretório TEMP registrado em checkpoint.json e habilitar o conteúdo somente se o Office exigir e o proprietário reconhecer a candidata sintética. Não alterar Trust Center, AccessVBOM ou políticas globais. A rotina de inicialização deve ser conferida antes de carregar cache manualmente.

Compilação, eventos reais, dois processos Excel e reabertura funcional desta sessão: NÃO EXECUTADOS. A janela de laboratório ausente impede a próxima validação de interface; não equivale a um novo bloqueio de segurança observado. Os sete FAIL históricos, ENTRY órfão, restauração sem registro/origem e exceções operacionais não aprovadas permanecem nas classificações anteriores.

## Reversão

Não houve alteração funcional para reverter. Preservar as duas tentativas, o banco descartável e os checkpoints. As alterações de documentação são acréscimos; as versões anteriores constam do backup oficial. Nunca restaurar banco antigo sobre revisão26 nem eleger cópia do registro por nome/localização. Nenhum commit, push, merge, instalação ou promoção foi executado.
