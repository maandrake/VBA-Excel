# Plano de correção — Fase 3

Execução 20261031_213925Z. Continuação das Fases 1 e 2, sem reauditoria geral ou promoção.

## Base e escopo confirmado

39 originais e backups da Fase 1 íntegros; 155 arquivos da primeira execução e 146 da segunda preservados; 46 backups da Fase 2 conferidos. A cópia inicial de 42 arquivos coincidiu integralmente com a entrega da Fase 2. Novo backup de 69 arquivos: 42 de desenvolvimento, 19 scripts e oito documentos. [Manifesto inicial](./_auditoria/fase3_20261031_213925Z/manifesto-inicial-fase3.json).

Somente [clsOperacaoPlanilha.cls](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) foi alterada entre os 34 fontes do sistema. Os outros 33, os cinco FRX, os formulários, o XML, as regras fiscais e os sete Delete ficaram idênticos à Fase 2. São 665 procedimentos exportados, sem mudança de assinatura. O XLSM base e sua cópia permanecem intactos e sem revisão incorporada.

## Trabalho e situação

| Frente | Resultado atual | Critério ainda necessário |
|---|---|---|
| Integração/compilação | BLOQUEADO: Office recusou VBProject, 0x800A03EC; zero importações | Operador importar/compilar com as proteções atuais e registrar evidência real. |
| Fixture | Esqueleto real de 17 abas e pacote CP1252/FRX preparado sob TEMP | Conferir documentos, CodeNames, referências, listas, fórmulas, objetos e Ribbon. |
| A09 ciclo de vida, C3-01 | Uso único do contexto; recusa de recuperação inativa, sem falso sucesso | Suíte VBA ampliada, ainda bloqueada. |
| A09 metadado, C3-02 | Snapshot/restauração de Ignore do erro de número como texto | Quatro casos COM passaram; executar a classe VBA. |
| A09 fronteiras | 15 experiências COM: dez recompõem campos; quatro exigem conciliação e uma tem recuperação incompleta | Não generalizar para macro; resolver/decidir conciliação após Delete/Save e crash. |
| Edição/exclusão anual | Referências ambíguas documentadas; código preservado | Responsável definir vínculo único e política histórica antes de alterar. |
| A08 | A, B e C comparadas com sentinelas; nenhuma política integrada | Aprovar identidade, política e migração antes de implementar. |
| A14/A15 | 12 cenários de parser/modelo passaram; fonte fiscal intacto | Executar ValidarXMLNFe e relatório real com documento fiscal fictício completo. |
| Backup | Cinco cenários COM, duas versões e restauração, incluindo falhas de destino/escrita | Executar as três rotinas VBA e validar interface/destino operacional. |
| A13/A20 | Análise e decisões propostas; sem mudança fiscal ou de logs reais | Decisão do responsável conforme [pendências](./PENDENCIAS_DECISAO.md). |

## Correções rastreáveis

### C3-01 — ciclo de vida da compensação, subitem A09

**Arquivo/procedimentos:** classe, Iniciar e Recuperar. **Causa:** contexto sem snapshots podia ser reiniciado depois de confirmar/finalizar; flags antigas permaneciam. Recuperar devolvia sucesso sem ter iniciado uma operação.

**Alteração:** mIniciada persiste após o primeiro início válido e recusa reutilização, inclusive vazia. Recuperar sem contexto ativo informa que não executou recuperação. Os quatro consumidores existentes já recuperam antes de Finalizar. Não muda sequência de estoque nem fronteiras Delete/Save.

**Dependências/testes:** Excel/Collection e estados Application; contratos estáticos passaram e AuditoriaFase3_CicloVida foi preparada na suíte existente. Runtime VBA BLOQUEADO. **Risco:** consumidor externo não mapeado que reutilize contexto vazio agora recebe erro; nenhum dos quatro eventos faz isso.

### C3-02 — metadado da chave, subitem A09

**Arquivo/procedimentos:** classe, RegistrarDestino e RestaurarCelula. **Causa:** ClearContents/atribuição reseta Ignore, embora valor e NumberFormat sejam restaurados. O ensaio COM demonstrou True → False na compensação anterior.

**Alteração:** sexto item do snapshot captura Errors(xlNumberAsText).Ignore; restauração ocorre depois de valor/fórmula/NumberFormat. Não converte chave para número nem altera tributo. **Dependências:** enumeração Excel já usada pelo projeto. **Testes:** quatro tipos de célula por COM equivalente, todos PASS; suíte VBA existente ampliada e ainda bloqueada. **Risco:** comportamento em outra versão/arquitetura exige runtime; falha desse setter será tratada como falha da célula, sem prometer recuperação completa.

[Blocos antes/depois](./_auditoria/fase3_20261031_213925Z/alteracoes-fase3.json), [regressão do metadado](./_auditoria/fase3_20261031_213925Z/a09-metadado-corrigido-com.json), [contratos estáticos](./_auditoria/fase3_20261031_213925Z/verificacao-estatica-fase3.json). Reverter somente a classe da nova cópia a partir do backup, junto com a versão correspondente do teste/pacote. Backups e execuções anteriores não devem ser sobrescritos.

## Recomendação para A08

Recomenda-se **C, baixa lógica por identidade estável, como direção de projeto condicionada**. A exclusão física quebrou referências; a limpeza preservou fórmulas, mas reutilizar a posição trocou a associação de veículo. C preservou identidade/histórico no protótipo, porém usa mapa apenas em memória, sem persistência ou adaptação da aplicação. Não está pronta para integração.

A coluna A repete valores nos históricos anuais (18 grupos na matriz e um na filial no recorte lido); chave existe como texto de 44 dígitos em apenas 17 dos 42 registros lidos da filial e não é gravada pelo cadastro matriz/anual. Isso impede supor um identificador universal. A pesquisa rápida já atravessa lacunas; os filtros legados de saídas/NCM devem ser tratados separadamente. Detalhes e comparação pelos dez critérios estão no [relatório](./RELATORIO_HOMOLOGACAO_FASE3.md).

## Próximos passos e aprovação

1. Operador executar o [roteiro manual](./_auditoria/fase3_20261031_213925Z/COMPILACAO_MANUAL.md) e devolver compilação/referências reais.
2. Definir identidade de entrada/episódio de estoque e ligação ao histórico, respondendo às pendências.
3. Aprovar política A08 e plano de migração com diagnóstico de ambiguidades, sem criar campos automaticamente.
4. Executar a [matriz funcional](./MATRIZ_TESTES_FASE3.md), inclusive erros dentro da recuperação e nas fronteiras.
5. Conciliar pendências, avaliar riscos e só então solicitar promoção expressa.

Sem commit, push, publicação, instalação, mudança de segurança, retenção automática ou uso de dados reais nos testes.
