# Proposta de armazenamento — Fase 8

Estado: decisão pendente; Portão 2 NÃO autorizado. Comparação arquitetural, sem instalação nem implantação. Os benchmarks desta fase medem o repositório experimental em XLSM; não comparam motores de banco.

## Alternativas

| Critério | A — auxiliar no XLSM | B — dados/diário locais separados | C — banco local transacional |
|---|---|---|---|
| Autoridade proposta | Um XLSM identificado por BaseID e revisão | Um manifesto de geração + dados + diário, nunca a cópia mais recente por nome | Banco é autoridade; Excel é projeção reconstruível |
| Integridade | Validação de IDs, versões e selos pela aplicação; edição manual pode contorná-la | Exige protocolo de gerações, parser estrito, recibos e conciliação | Restrições únicas, referências e atualização por versão dentro do banco; dependem do esquema e motor aprovados |
| Durabilidade | Depende de Save/reabertura; intenção não salva pode desaparecer | Intenção anterior aos efeitos, fechamento e confirmação durável a demonstrar; gravação de texto não prova flush contra energia | Commit do motor a validar com driver, disco e modo de journal escolhidos |
| Concorrência | Exigir escritor único; cópia do arquivo não é lock confiável | Lock exclusivo com proprietário/sessão; lock órfão exige conciliação, não remoção automática | Transações e bloqueios do motor; ainda definir escritor único da camada operacional e controle otimista |
| Backup | Copiar geração fechada, hash, abrir e validar; preservar recibos | Backup coordenado do conjunto; diário sem dados ou dados sem diário são divergência | Backup consistente próprio do motor + versão de esquema; não copiar arquivo aberto sem protocolo validado |
| Restauração | Não restaurar cegamente após novos movimentos | Restaurar geração coerente e reconciliar operações posteriores | Restaurar autoridade e reaplicar/reconciliar operações comprovadas; reconstruir projeções |
| Versões/cópias divergentes | BaseID + revisão + assinatura do conteúdo e linhagem | Manifesto assinado por hash e encadeamento de gerações | Versão de schema e de agregados, OperationID único, hash de payload e recibo |
| Performance | Medição real nesta fase; selos célula a célula custam mais que índice | Depende de tamanho, leitura incremental, parser e fsync; NÃO EXECUTADO | Índices e transações prometem melhor separação, mas benchmark do motor NÃO EXECUTADO |
| Complexidade | Menor inicialmente; maior risco de confundir planilha com autoridade | Alta: implementar protocolo de arquivos, detecção de truncamento e recuperação | Alta na migração; menor dependência da posição das células após modelagem correta |
| Compatibilidade VBA | Nativa; 17 abas e leitores legados ainda exigem adaptação | FSO/Win32 possíveis, sem garantia de persistência só por Close | Driver COM/ADO ou outra interface aprovada, compatível com Office 64 bits; disponibilidade não demonstrada nesta fase |
| Instalação | Nenhuma nova | Nenhuma necessariamente; dependências a verificar | Motor/provedor/versões precisam aprovação; Access instalado não prova driver adequado nem licença de distribuição |
| Privacidade/permissões | ACL da pasta e acesso ao XLSM; proteção de aba não é autenticação | ACL de todos os arquivos e backups; diário mínimo, sem XML integral | ACL/credenciais/backup do banco; contas e auditoria conforme motor; sem credenciais em VBA |
| Migração | IDs somente após M2/M3 e em cópia aprovada | IDs e protocolo de autoridade devem migrar juntos | Importar staging, validar correspondências, contagens e constraints; ambíguos em quarentena |

## Recomendação condicional

Preferir C como arquitetura alvo: identidade, movimentos, idempotência e recibos na mesma autoridade transacional; Excel como interface/projeção. Não escolher nem instalar o motor nesta etapa. Aprovar primeiro custódia, esquema, escritor único e requisitos de recuperação; depois provar o motor local disponível em laboratório. A é admissível apenas como transição restrita e explicitamente limitada; B pode servir de diário de contenção, mas exige mais engenharia de durabilidade e conjuntos coerentes. Nenhuma alternativa está aprovada para o produto.

Uma transação do banco não inclui automaticamente o Save do XLSM. A proposta evita essa suposição: commit confirma o estado autoritativo; a projeção recebe uma revisão/recibo e pode ser reconstruída. Se a planilha falhar após commit, bloquear novos comandos dependentes da projeção até sua reconciliação. Não compensar um commit autoritativo apagando células ao acaso.

A documentação oficial descreve transações ADO no objeto Connection e condicionadas ao provedor: [Microsoft — BeginTrans/CommitTrans/RollbackTrans](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/begintrans-committrans-and-rollbacktrans-methods-ado). Inferência arquitetural desta proposta: isso não demonstra transação conjunta com um arquivo Excel.

## Aceite antes da escolha

Demonstrar chave única de OperationID no escopo BaseID, payload canônico, CAS por ExpectedVersion, referências íntegras, exclusão lógica, commit/recibo, bloqueios concorrentes, falha no Save da projeção, reabertura fria, backup/restauração e privacidade. Testes de energia, crash, disco cheio e dois escritores permanecem NÃO EXECUTADOS. Definir RPO/RTO e retenção com o responsável; nenhum prazo foi inventado.

Ver [protocolo](./CONTRATO_PERSISTENCIA_FASE8.md), [plano M0–M9](./PLANO_IMPLANTACAO_FASE8.md) e [decisões](./PACOTE_APROVACAO_PORTAO2.md).
