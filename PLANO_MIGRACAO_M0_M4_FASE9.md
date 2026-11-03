# Ensaio e plano de migração M0–M4 — Fase 9

Somente laboratório autorizado. Implantação no produto depende de aprovação nominal P2. Segue o [plano reversível F7](./PLANO_MIGRACAO_REVERSIVEL_FASE7.md); baixa/leitores/gravadores continuam fora desta execução.

| Etapa | Entrada e insumo | Ensaio observado | Saída necessária para integração |
|---|---|---|---|
| M0 | Base/autoridade/responsável/janela, esquema/fontes/FRX/manifestos, backup coerente |2.268 hashes F8 conferidos;43 backups novos; mestre final F8 copiado; banco sintético fechado copiado e reaberto em nova sessão | Responsável e fonte autoritativa aprovados; restauração completa repetível; congelamento/escritor demonstrado no ambiente destino |
| M1 | Inventário de leitores/escritores e dependências |60 componentes F8 preservados;3 experimentais novos;17 abas herdadas;57 leitores F7 reutilizados como mapa, sem nova auditoria geral | Completar dependências desconhecidas, catálogos de fórmulas e interfaces do candidato aprovado;57 não prova completude universal |
| M2 | Esquema versionado, motor/provedor e contrato de custódia aprovados |12 tabelas sintéticas schema 1; PK/FK/unique; reforço composto testado em migracao.accdb; proposta v2 preparada | Aprovar P2-01/02/04/05/06/07/08/09/10; consolidar constraints e SchemaNo 2, snapshots concorrentes e acesso; nenhum schema operacional criado |
| M3 | Mapa de proveniência com hash/endereço original/fingerprint/evidência/decisão |legado_sintetico.tsv SHA25638de109d962cb44f45cab4934ea957be1480c362cad00da4b7d9538ddddc2218; duas correspondências explicitamente conhecidas e uma ambígua | Cada vínculo legado exige evidência/responsável. Mesmo veículo/placa/NF não permite fundir episódios. Zero/múltiplos candidatos ficam pendentes |
| M4 | Correspondências aprovadas e staging reversível |SYN_A1→EP_A1 eSYN_A2→EP_A2, mesmo VID e NF mas episódios distintos; SYN_AMBIGUO sem EpisodeID; FK composta recusa movimento de A1 ligado a B1 | IDs imutáveis/proveniência completa, relações e órfãos verificados, repetição de migração sem duplicidade, revisão de esquema publicada. Nenhum registro real recebeu ID |

A fonte sintética e o hash foram produzidos/conferidos em Node; o valor de hash foi passado explicitamente ao VBA. O ensaio grava as correspondências conhecidas em fixture, não é algoritmo genérico de inferência/migração do legado. [Insumo](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/legado_sintetico.tsv), [teste](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/complemento_f9.tsv).

## Pacote mínimo de eventual autorização

Escolher nominalmente identidade/custódia, histórico, motor/interface/licença, autoridade, escritor, conciliação, dados mínimos/retencão e limites de recuperação. Autorizar candidato sintético e etapas delimitadas; não inferir M5–M9, dados reais ou produção. Reforçar o núcleo antes de M7: constraints completas, validadores, fencing na preparação, recibo/efeitos íntegros, estado persistente de projeção, restauração e falhas de commit.

## Reversão

Sem fatos novos: restaurar conjunto compatível M0 com hashes, esquema, banco, diário/recibos e fontes. Migração parcial: não promover, preservar staging/proveniência e descartar somente candidato após conservar evidências. Histórico ambíguo: manter pendência, não restaurar vínculo inventado. Autoridade divergente: quarentena e conciliação por BaseID/linhagem/revisões/recibos. Com fatos novos: preservar ambos conjuntos, extrair e conciliar fatos posteriores e reconstruir projeções; nunca restaurar cegamente XLSM/banco antigo.
