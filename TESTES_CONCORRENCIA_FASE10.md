# Concorrência e fencing — Fase 10

## Ambiente e fronteiras

Excel/VBA real PID52068, ACE16 x64 e processos PowerShell/ADO independentes. Bases sintéticas em TEMP. Os ensaios executaram duas conexões e dois processos concorrentes; não é uma simulação JavaScript. Os resultados COM não são classificados como execução de um segundo gravador VBA.

Acquire10 usa UPDATE condicional de lease livre, contagem afetada=1 e nova geração. Prepare10, Apply10 e SnapshotFull10 revalidam proprietário/fencing mediante UPDATE da linha de lease **dentro da transação**. Apply também vincula a intenção à tentativa PREPARED do mesmo proprietário e geração; não basta uma leitura prévia do lock.

## Interleavings observados

|Caso|Sequência e estado anterior|Esperado|Observado e estado posterior|Evidência|
|---|---|---|---|---|
|Disputa durante Prepare|A mantém transação e linha de lease; B abre outra conexão e tenta revogar|B não atravessa a fronteira de preparação|PASS; PID44372 abriu e tentou; ACE bloqueou por2,231s; A retornou PREPARADA e auditou OK|prepare_worker_prepare_done.json;concorrencia_vba_f10.tsv|
|Nova geração B|A libera; B adquire FREE com CAS|Uma linha adquirida e token novo|PASS; PID42952, Fence3; aquisição/revalidação/escrita diagnóstica afetaram1 linha cada|worker_b_fence.json|
|A obsoleto|A tenta confirmar com Fence2 após B Fence3|Recusa sem efeitos|PASS; RECUSADO_FENCE e DataSig igual|fencing_vba_f10.tsv|
|B não herda intenção|Novo dono tenta Apply da intenção de A|Exigir conciliação|PASS; PENDENTE_OWNER_CHANGED; novo OperationID também PENDENTE_CONCILIACAO|fencing_vba_f10.tsv|
|Disputa durante snapshot|VBA segura lease/revisão enquanto B disputa a linha|Snapshot coerente e B recusado|PASS; PID48224 bloqueado2,147s; snapshot com revisão7 e sete registros|snapshot_worker_prepare_done.json;fronteira_snapshot_vba_f10.tsv|
|Disputa durante Apply|VBA segura transação de confirmação; B disputa lease|O escritor obsoleto não revoga no meio dos efeitos|PASS; PID35952 bloqueado2,120s; A confirmou oitavo movimento e auditoria OK|apply_worker_prepare_done.json;fronteira_apply_vba_f10.tsv|
|Lease órfão em processo novo|Conexão B termina normalmente sem liberar lease, com intenção pendente|Não liberar por relógio|PASS COM; ACTIVE persistiu na reabertura pelo PID25576|frio_com_f10.json|

Os arquivos de evidência estão em [evidencias](./_auditoria/fase10_20261102_155920Z/evidencias/). O código de sincronização é LockWriter10/hook10; a janela de espera tem35s e não apaga locks. Os workers fazem rollback da tentativa de disputa. A liberação de A é explícita; não foi takeover por prazo.

## Garantias e limitações

A escrita legítima de B neste ensaio foi um metadado diagnóstico, após adquirir/revalidar a geração e comparar revisão7. Não é uma movimentação operacional completa executada pelo processo B. A intenção órfã impede novo movimento do mesmo episódio até conciliação. Nenhuma resolução administrativa automática foi implementada.

Os snapshots são serializados pelo protocolo cooperante de escritor único. Não foi comprovado isolamento serializável geral do ACE, snapshots multiversão, ausência de todos os deadlocks, concorrência em rede/OneDrive ou imunidade a SQL externo que ignore WriterLease. Um lease persistente não é um bloqueio obrigatório contra um usuário com acesso direto ao arquivo.

Encerramento usado: normal e identificado; não crash ou energia. Não foi encerrada a instância do usuário. Os tokens antigos foram recusados no protocolo ensaiado; proteção em todos os futuros gravadores/leitores reais depende de integração e regressão ainda não autorizadas.

A suíte fria VBA foi BLOQUEADA pela mensagem do Office: “As macros neste projeto estão desativadas”. A tela Informações confirmou conteúdo ativo bloqueado e não apresentou habilitação específica disponível naquele estado. Nenhuma política foi alterada nem houve bypass. Os seis testes frios previstos no VBA permanecem BLOQUEADOS; seis leituras COM posteriores têm modalidade própria.
