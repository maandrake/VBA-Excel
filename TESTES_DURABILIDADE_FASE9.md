# Durabilidade e recuperação — Fase 9

Execução real de VBA/ACE/Excel nas fixtures TEMP. [Matriz por cenário](./MATRIZ_REGRESSAO_FASE9.md), [snapshots antes/depois](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/). PASS de contenção é separado de recuperação de dados.

| Ponto / hipótese | Evidência observada | Classificação / alcance |
|---|---|---|
| Antes da escrita | Versão/base/payload inválidos recusados; assinatura sem efeitos no replay | RECUSADO; não prova todos os campos fiscais |
| Durante destinos | F9_FALHA_1/2 após movimento/projeção; assinatura dos dados recomposta | RECUPERADO no banco; intenção separada continua pendente |
| Durante baixa lógica | F9_FALHA_3 após mudança do episódio; rollback comparado | RECUPERADO transacional experimental; Delete real não foi alterado |
| Durante diário | F9_FALHA_4 insere OperationID órfão no ChangeLog; erro real FK e rollback | RECUPERADO; diário dentro do mesmo banco, não append externo |
| Antes do recibo/commit | F9_FALHA_5/6 falhas artificiais explicitamente identificadas | RECUPERADO antes do commit; não falha física do provedor |
| Tentativa de fechar conexão na transação | F9_FALHA_7 retornou3246: conexão não pode ser fechada durante transação; rollback depois passou | RECUPERADO dos efeitos; CommitTrans não chegou a ser chamado nesse cenário |
| Commit em conexão já fechada | Complemento retornou3704 | RECUSADO pela API; não é teste de erro de I/O durante commit |
| Após commit/antes Excel | F9_APOS_COMMIT conservou8 fatos/recibos; projeção ainda antiga | PENDENTE_CONCILIACAO da projeção; autoridade confirmada |
| Projeção parcial | Falha7100 após cabeçalho; autoridade igual | PENDENTE_CONCILIACAO; reconstrução seguinte RECUPERADO8 registros/2 ativos |
| Durante Save | BeforeSave cancelou chamada real; Saved=False; alteração T3 só em memória sumiu ao reabrir | PENDENTE da apresentação;7 recibos do banco preservados |
| Erro de Save | SaveAs para pai inexistente retornou1004; banco igual | PENDENTE da projeção; caminho inválido não equivale a disco cheio |
| Após Save/antes recibo | Na arquitetura C o recibo da autoridade antecede Excel; não existe novo recibo de negócio dependente do Save | Ack persistente de projeção NÃO IMPLEMENTADO/ NÃO EXECUTADO; não alegar commit conjunto |
| Histórico editado | Canônico da projeção/recibo divergiu; não repetir | PENDENTE_CONCILIACAO; dados alterados não foram restaurados automaticamente |
| Origem mudou após preparar | PENDENTE_ORIGEM_ALTERADA, nenhum efeito novo | PENDENTE_CONCILIACAO |
| Compensação/rollback falha | Não foi produzida falha real de RollbackTrans ou restauração de disco | NÃO EXECUTADO; C_FALHA_COMPENSACAO/C_FALHA_BAIXA F6 continuam FAIL |
| Energia/crash/disco cheio | Nenhuma interrupção forçada | NÃO EXECUTADO; sem substituição por simulação |

## Sessão fria real

Primeiro Excel PID45652 fechado normalmente. Nova instância /x PID48356 abriu o mesmo mestre salvo e conexões novas; F9_Cold executou nove verificações PASS. Ciclo7 registros/7 recibos, A1CLOSED/A2ACTIVE versão4, backup fechado reaberto íntegro; quatro bancos reconstruídos sem usar os XLSM salvos. Intenção PREPARED/Pending e lease órfão ACTIVE persistiram; bloqueio não é recuperação. [TSV frio](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/sessao_fria_f9.tsv).

Arquivos de dados/evidência pré-existentes conferidos por hash antes/depois, excluídos locks transitórios ~$ do Office. O mestre mudou entre a captura inicial e o encerramento por salvamento final da sessão quente; essa captura não isola alteração por sessão fria. Não alegar igualdade binária do mestre nessa janela. Os bancos e projeções referenciados mantiveram seus hashes. [Conferência](./_auditoria/fase9_20261102_142753Z/evidencias/sessao-fria-hashes.json).

Propriedades User Commit Sync=True e Asynchable Commit=False foram lidas, não alteradas. Fechar/reabrir demonstra persistência nesses encerramentos normais; não demonstra flush físico contra energia ou falha do dispositivo. Backup é cópia de banco fechado, verificada e reaberta, não certificação de backup on-line.
