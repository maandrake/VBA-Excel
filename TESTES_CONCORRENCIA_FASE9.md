# Concorrência e escritor único — Fase 9

Política experimental, não segurança por papel do produto. Duas sessões ADO em um Excel e processos independentes foram testados; não houve kill, expiração de lock ou alteração de segurança.

| Cenário | Origem / fixture | Esperado e observado | Resultado / tratamento |
|---|---|---|---|
| Primeiro escritor | F9_Core/ciclo.accdb | Acquire retorna Fence1 | PASS; autorização experimental |
| Segunda conexão | F9_Core/ciclo.accdb | UPDATE condicional não adquire enquanto ACTIVE | PASS/RECUSADO |
| Somente leitura | ADO Mode=Read/ciclo.accdb | UPDATE recusado: consulta não atualizável | PASS/RECUSADO pelo provedor |
| Escritor antigo | F9_Core/ciclo.accdb | Após mudança explícita de Owner/Fence no laboratório, token antigo recebe RECUSADO_FENCE | PASS/RECUSADO; não prova todos os interleavings |
| Versão obsoleta | ExpectedVersion2 contra4 | RECUSADO_VERSAO | PASS/RECUSADO |
| Contorno por novo OperationID | F9_Faults, sete bancos com intenção pendente | Novo ID não remove Pending | PASS/PENDENTE_CONCILIACAO |
| Lease órfão | Conexão proprietária fechada normalmente sem Release; lease_orfao.accdb | Lease ACTIVE persiste e novo Acquire retorna0, inclusive em sessão fria | PASS/PENDENTE_CONCILIACAO. Não foi crash |
| Processo independente | Excel PID45652 mantém lease; PowerShell PID26732 tenta Acquire no mesmo banco | Affected0, proprietário/fence original observados | PASS/RECUSADO; [JSON](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/concorrencia_segundo_processo.json) |
| Contenção real de transação | Excel atualiza Meta e mantém transação aberta; PowerShell PID49412 tenta UPDATE | Conexão aberta=True, erro “Não foi possível atualizar; atualmente bloqueado.” após 2,2663 s | PASS/RECUSADO; [JSON](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/concorrencia_lock_real_corrigido.json) |
| Liberação e reconexão | VBA RollbackTrans/Release; PowerShell PID42856 | Acquire1, Fence2, revisão voltou0 | PASS/RECUPERADO; [JSON](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/concorrencia_reconexao.json) |

A primeira tentativa externa de contenção usou propriedades de conexão não aceitas e falhou antes de Open. O auxiliar gravou indevidamente PASS para “objeto fechado”; esse resultado é **INVÁLIDO e excluído da contagem**. Arquivo original preservado: [tentativa inválida](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/concorrencia_lock_real.json). O reteste exige Opened=True e mensagem de bloqueio do provedor. Nenhuma falha de conexão foi convertida em prova de lock.

## Garantia efetivamente exercitada

WriterLease usa OwnerID+FenceNo; Acquire faz UPDATE onde FREE, Apply revalida dentro da transação e toca a linha de lease antes dos efeitos. Release confere proprietário/fence. Não há prazo que apague lock. O núcleo bloqueou segunda sessão, pendência e token revogado nos casos testados. Não protege contra alguém que abre ACCDB e executa UPDATE fora do protocolo; ACL/autenticação e desenho de acesso ainda precisam aprovação.

A preparação lê Own9 dentro da transação, mas não foi submetida a todos os interleavings com revogação; reforçar bloqueio/CAS na preparação antes do candidato produtivo. Snapshot coerente da leitura em lote enquanto um escritor confirmado altera revisões NÃO EXECUTADO. As medições foram feitas com dados estáveis e escritor controlado. Não anunciar isolamento serializável integral, tolerância a split-brain, lease roubado seguro ou recuperação de processo morto.

Reconciliação de órfão: congelar escrita, conferir proprietário, sessão/processo, BaseID, revisão, Pending e recibos; responsável decide retomada em cópia; CAS de Owner/Fence avança token e impede escritor antigo. Não remover por timeout. A revogação explícita foi demonstrada em laboratório sem pendência, não autoriza revogar automaticamente um escritor do usuário.
