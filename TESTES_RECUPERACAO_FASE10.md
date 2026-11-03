# Recuperação e restauração — Fase 10

## Garantias demonstradas e estados

RECUPERADO significa que o estado relevante foi conferido após rollback ou reconstrução. COMPENSADO exige a reversão explícita de efeitos identificados. RECUSADO significa impedir a operação. PENDENTE_CONCILIACAO mantém o estado diagnóstico e proíbe repetição insegura; não equivale a recuperação integral.

|Situação|Conduta experimental|Limite|
|---|---|---|
|Falha antes de BeginTrans|Recusar sem efeitos novos|Intenção anteriormente preparada permanece|
|Erro antes do commit de Apply|Rollback dos efeitos; conferir fatos/recibos/versões|Intenção é de transação anterior; não repetir automaticamente|
|Commit confirmado e Excel ausente|Reconstruir a projeção a partir da autoridade|Não reexecutar a movimentação|
|Save cancelado ou SaveAs inválido|Manter recibo do fato no banco e confirmação de arquivo pendente|Não marcar BUILT em memória como arquivo confirmado|
|Projeção adulterada|Recusar confirmação, reconstruir e revalidar|Preservar arquivo divergente como evidência|
|Recibo, histórico, comando ou tentativa divergente|Gravar Meta.WriteState=RECONCILE na próxima preparação e bloquear escritas|A guarda não restaura conteúdo adulterado|
|Lease ou intenção órfãos|Exigir decisão de recuperação; não liberar por idade|Sem procedimento administrativo de takeover implantado|
|Resultado de commit incerto|Preservar arquivos; consultar recibo/OperationID em nova sessão; conciliar antes de repetir|Falha física de commit não foi ensaiada|
|Falha de compensação/crash/energia|Suspender dependências e conciliar|Recuperação integral não comprovada nas fases anteriores|

## Backup e restauração ensaiados

F10_Core fechou a conexão do ciclo de sete movimentos e copiou o ACCDB com igualdade SHA-256: 7325cdb9a19a24510e23fa4f6ea3af5baea12f1583606e4b914675810261e106. A cópia independente foi reaberta, auditada e usada para reconstrução integral da projeção (F10_BackupRestauracao). Os três testes finais passaram.

O cenário com fatos posteriores preservou a cópia antiga com sete movimentos e a base transacao_8.accdb com oito. A rotina não promoveu a cópia antiga nem apagou o fato posterior. O diagnóstico PENDENTE_CONCILIACAO é a resposta segura, sem alegar merge automático.

Backup com conexões abertas: análise somente. Não há prova de consistência de uma simples cópia a quente. Cópia bruta aberta, compactação durante uso, crash, falha de energia e disco cheio: NÃO EXECUTADO. Nenhuma instalação ou configuração de segurança foi alterada.

## Roteiro de restauração do laboratório

1. Preservar o ACCDB/XLSM suspeito e seus hashes; suspender escritores dependentes.
2. Copiar a entrega e o backup verificado para outra pasta TEMP, sem sobrescrever evidências. Nunca executar Office diretamente no arquivo arquivado no OneDrive.
3. Conferir manifesto, esquema, BaseID, linhagem, revisão, recibos, tentativas, Pending e WriterLease.
4. Se existirem fatos posteriores ao backup, manter ambas as fontes e obter decisão de conciliação. Não restaurar cegamente uma versão antiga.
5. Reconstruir uma projeção nova e verificar o contrato integral, inclusive chave, fórmulas, tipos, formatos, Ignore, identidades e totais.
6. Salvar, fechar, reabrir somente leitura e conferir. Registrar confirmação na autoridade identificando BuildID, BaseID, revisão, linhagem, caminho e hash exatos.
7. Liberar uma nova sessão somente por procedimento aprovado; não presumir que tempo decorrido revogou o escritor anterior.

O mestre experimental tem guarda de nome e TEMP. Para reproduzir, usar uma cópia chamada Fixture_Integridade_Fase10_Mestre.xlsm em pasta TEMP com o marcador VBA_Fase10_20261102_155920Z, junto aos dados sintéticos. Os ensaios criam nomes novos e recusam sobrescrita. Para repetir uma suíte, criar um novo diretório de reprodução com os insumos mínimos; preservar os resultados anteriores.

## Decisões ainda necessárias

Fonte autoritativa e localização controlada; responsabilidade pelo lease órfão; recibo de reconciliação; tratamento de divergência após Save; ACL, privacidade e retenção; backup consistente e testes físicos; contrato de episódio/custódia; migração dos leitores e gravadores do produto. Ver [pacote de decisão](./PACOTE_DECISAO_INTEGRACAO_FASE10.md).
