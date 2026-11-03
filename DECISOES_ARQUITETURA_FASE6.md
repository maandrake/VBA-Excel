# Decisões de arquitetura — Fase 6

Situação: **aguardam aprovação específica; nenhuma decisão abaixo está implicitamente aprovada**. O escopo autorizado nesta execução permitiu somente protótipos sintéticos e documentação. A versão principal mantém seus fontes e os sete Delete. O sistema permanece NÃO HOMOLOGADO.

| Decisão | Proposta técnica concreta | Alternativas / consequência | Evidência e critério para aprovar |
|---|---|---|---|
| D01 — Identidade e política de baixa | Adotar IDs imutáveis separados para veículo, episódio, movimentação e projeção; baixa lógica somente com todos os leitores migrados | A: Delete mantém quebra/troca silenciosa. B: limpar preserva endereço mas perde conteúdo e reutilização reatribui referência. C: preserva vínculos no modelo, exige esquema e migração conjunta | [Identidade](./ARQUITETURA_IDENTIDADE_FASE6.md), [impacto](./MATRIZ_IMPACTO_BAIXA_FASE6.md) e [regressão](./MATRIZ_REGRESSAO_FASE6.md). Aprovar semântica de episódio/custódia, tratamento de reentrada, reserva de armazenamento e adaptação dos leitores antes de integrar. |
| D02 — Históricos e conciliação | Vincular cada histórico por ProjectionID/MovementID/EpisodeID; correções versionadas, sem primeira coincidência por NF-e | Históricos ambíguos ficam bloqueados para edição/baixa dependente. Reparo exige plano exato e aprovação; não compensar Delete/Save estrutural automaticamente | Falhas D_FALHA_BAIXA e C_FALHA_COMPENSACAO; mutação pós-conclusão. Aprovar critérios de correspondência e responsável pela conciliação. |
| D03 — Cancelamento fiscal | Manter comportamento atual; separar estado interno e evidência fiscal | Nenhuma exigência nova de evento XML foi implementada. Registro interno não comprova autorização fiscal | C_CANCELADO testa apenas o modelo interno. Definir evidências e regra fiscal pelo responsável antes de qualquer alteração no ramo Cancelado. |
| D04 — Logs: destino, privacidade e retenção | Minimizar conteúdo, restringir acesso e coordenar backup do diário com a base | Aba no XLSM, arquivo local separado ou armazenamento relacional aprovado possuem riscos distintos. Sem retenção automática nesta fase | [Recuperação](./ARQUITETURA_RECUPERACAO_FASE6.md). Aprovar dados mínimos, localização, responsáveis, permissões, backup e prazo; não copiar documentos fiscais completos por conveniência. |
| D05 — Ambiente, compilação e dependências | Manter importação/compilação manual assistida pelo Editor VBA já comprovada | AccessVBOM/Trust Center continuam sem alteração. Compilação experimental não homologa o produto | [Evidência R2](./_auditoria/fase6_20261101_224949Z/evidencias/compilacao-real-r2.json); GUID/API compilou no Office instalado, não em outras arquiteturas. Futuro pacote integrado precisa compilar e executar eventos reais. |
| D06 — Fonte autoritativa e durabilidade operacional | Definir escritor único inicialmente, recibo persistente, bloqueio de pendências e protocolo de conciliação | Diário TSV experimental não oferece ACID, lock multiwriter, flush garantido ou resistência à adulteração. Banco pode garantir transação interna, não commit conjunto com Excel | Aprovar armazenamento e autoridade, política para cópias divergentes, envelopes/pré-imagens, exclusão mútua e plano de testes de Save/crash/recuperação antes de promover. |

## Pacote recomendado para uma próxima implementação

Aprovar D01 e D02 como contrato técnico conjunto, com C condicionada a migração dos leitores/fórmulas. Selecionar D04/D06 antes de tornar o diário persistente parte do produto. Não autorizar apenas troca dos comandos Delete, pois pesquisas, contagens e fórmula anual por placa ainda podem mostrar episódios errados.

A aprovação deve dizer explicitamente: modelo de episódio em transferências/devoluções; onde persistir IDs e diário; como tratar registros antigos sem vínculo; como editar histórico; quais pendências bloqueiam operação; quem concilia; e quais leitores/fórmulas entram no mesmo corte. Não basta aprovar genericamente “melhorar integridade”.

## Limites de validação para a próxima fase

Exigir testes dos eventos reais de cadastro/saída/edição/exclusão migrados, mesmos cinco FAIL obrigatórios, pesquisa após baixa, reentrada e histórico, reordenação, arquivos divergentes, falha real de Save no mecanismo escolhido e recuperação persistente. Crash, energia, multiwriter e corrupção permanecem NÃO EXECUTADOS nesta fase. Falhas em recuperação integral e em leitores legados continuam FAIL, ainda que a guarda de conciliação passe.

Aprovação deste documento não promoverá automaticamente um XLSM para produção. A promoção continua exigindo autorização expressa separada.
