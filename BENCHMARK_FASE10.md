# Benchmark — Fase 10

VBA real/Excel x64/ACE16; dados sintéticos; três repetições por volume, todas PASS. Tempo em segundos: média [mínimo–máximo], não somente melhor resultado. [Medições brutas](./_auditoria/fase10_20261102_155920Z/evidencias/benchmark_medidas.json).

Os volumes100/1.000/5.000/10.000 usam schema4 completo, com comandos, tentativas, movimentos, projeções, ChangeLog e recibos encadeados. A carga inicial é uma transação por volume: SEED e COMMIT_SEED são uma única medida repetida nas três linhas, não três commits de carga independentes.

**Revisão:** 1.000/5.000/10.000 foram medidos antes da guarda adicional PreparedAudit10; 100 foi repetido com a guarda final, preservando a referência anterior. A nova guarda não foi reavaliada em três repetições nos volumes maiores. Não atribuir silenciosamente os tempos antigos ao auditor final.

## Consultas e lista

|Registros|Consulta de fatos em bloco|Auditoria integral|Snapshot auditado|Índice +200 candidatos|20 resoluções verificadas|Lista de snapshot + MSForms|
|---|---|---|---|---|---|---|
|100|0,00567 [0,00524–0,00639]|0,331 [0,329–0,334]|0,376 [0,367–0,388]|0,00096 [0,00084–0,00110]|0,174 [0,164–0,185]|0,00062 [0,00035–0,00081] + 0,00395 [0,00374–0,00421]|
|1000|0,018 [0,017–0,019]|1,916 [1,874–1,955]|1,980 [1,971–1,998]|0,00256 [0,00239–0,00274]|0,172 [0,168–0,174]|0,00483 [0,00413–0,00537] + 0,00773 [0,00741–0,00826]|
|5000|0,122 [0,110–0,143]|9,578 [9,348–9,771]|10,110 [9,883–10,376]|0,012 [0,011–0,013]|0,204 [0,177–0,218]|0,027 [0,022–0,031] + 0,028 [0,023–0,033]|
|10000|0,258 [0,255–0,259]|18,505 [18,160–18,789]|19,241 [19,110–19,443]|0,021 [0,020–0,022]|0,181 [0,175–0,188]|0,049 [0,043–0,055] + 0,048 [0,045–0,049]|

O índice é candidato, sem autorização de mutação. Resolve20 confere identidade/versão dos alvos e não substitui a auditoria integral da autoridade. A lista parte de snapshot já auditado: o tempo de materialização não inclui novamente o custo de Snapshot. O binding foi feito no controle real Cad_NFE.ListBoxLista, sem acionar gravadores.

## Projeção e persistência

|Registros|Reconstrução incl. snapshot/cálculo inicial|Verificação integral|Save|Reabertura na mesma instância|Commit diagnóstico|
|---|---|---|---|---|---|
|100|0,410 [0,398–0,422]|0,372 [0,368–0,375]|0,091 [0,085–0,100]|0,280 [0,273–0,291]|0,00768 [0,00714–0,00852]|
|1000|2,153 [2,147–2,160]|2,030 [2,004–2,079]|0,111 [0,108–0,113]|0,354 [0,310–0,433]|0,00844 [0,00784–0,00901]|
|5000|10,847 [10,534–11,115]|10,079 [9,890–10,389]|0,218 [0,198–0,258]|0,465 [0,403–0,564]|0,00982 [0,00738–0,013]|
|10000|21,085 [20,940–21,219]|19,558 [19,474–19,684]|0,271 [0,256–0,278]|0,536 [0,510–0,567]|0,00747 [0,00657–0,00863]|

RECALC é um segundo Calculate sem alterações pendentes; não representa o custo de recalcular uma planilha suja. A reconstrução contém o cálculo inicial. A reabertura medida é real, mas dentro da mesma instância; não é startup frio do Excel. O commit diagnóstico atualiza um metadado sob lease, sem fingir que equivale ao commit de uma movimentação completa.

|Registros|Carga inicial única|Commit da carga único|Bytes do ACCDB fechado|
|---|---|---|---|
|100|0,684|0,00186|1720320|
|1000|5,285|0,011|13135872|
|5000|27,195|0,00628|63377408|
|10000|55,437|0,081|126578688|

## Comparação e gargalos

F8 mediu aproximadamente9,71s para materializar/validar lista10.000 com leituras repetidas por célula. F10 materializa o array em cerca0,05s e vincula MSForms em cerca0,05s, porém exige antes um snapshot auditado de aproximadamente19s. Não anunciar ganho de autorização integral de9,71s para0,05s. F9 usava outro conjunto de garantias; seus tempos também não são diretamente equivalentes aos deste esquema e cadeia de recibos.

O gargalo atual é auditoria integral repetida, com hashes e reconstrução canônica de todo o histórico, executada em Snapshot, Build e Verify. Otimização futura: cache de snapshot auditado por BaseID/revisão/esquema e invalidação comprovada dentro do protocolo; reduzir criação de contextos criptográficos e passagens redundantes. Não implementado para evitar reduzir garantias sem nova regressão. A fórmula anual nativa usa COUNTIF em faixa limitada, nunca coluna inteira; seu custo deve ser separado em futura medição com planilha suja.

## Memória e limites

Amostras externas do PID52068: antes de1.000/depois de100, WorkingSet605.786.112 bytes e Private437.182.464; após os volumes maiores e repetição100, WorkingSet612.728.832 e Private483.303.424. São fotografias de memória do processo, não pico, consumo exclusivo por volume ou prova de ausência de vazamento. Pico contínuo e série temporal por cenário: NÃO EXECUTADO. A instância foi encerrada normalmente; a instância do usuário foi preservada.

A sessão fria VBA final foi BLOQUEADA pelo Office. Reabertura COM em processo separado confirmou dados e recibos, sem substituir o benchmark ou a verificação funcional do VBA frio.
