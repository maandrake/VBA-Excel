# Plano de implantação preparado — Fase 8

Somente preparação do Portão 1. M0–M9 seguem o [plano F7](./PLANO_MIGRACAO_REVERSIVEL_FASE7.md). Nenhuma etapa operacional abaixo está implicitamente aprovada ou executada. O experimento não fornece um pacote pronto para substituir o XLSM.

| Etapa | Entrada exigida | Insumo preparado / trabalho necessário | Saída objetiva e portão |
|---|---|---|---|
| M0 — congelamento | Responsável, autoridade, escopo e janela aprovados | Checklist de conjunto: XLSM, fontes, FRM/FRX, CodeNames, esquema, diário/recibos, versões, SHA-256; backup coerente fora da sincronização | Responsável identificado; escritor exclusivo demonstrado; backup reaberto e restauração ensaiada. F8 preservou apenas laboratório |
| M1 — inventário | M0 íntegro | Reutilizar 57 procedimentos nominais F7, 17 abas e catálogo de fórmulas; completar nomes, tabelas, validações, relatórios, Ribbon/OnAction e chamadas dinâmicas do candidato | Cada leitor/escritor com dono, identidade atual/proposta e teste. Dependência desconhecida bloqueia o corte afetado; 57 não é prova de completude universal |
| M2 — esquema | Decisões de identidade, custódia, baixa, história, armazenamento, autoridade e privacidade assinadas | DDL/esquema versionado, constraints e interfaces; não ocupar coluna supostamente livre | Aprovação explícita do layout e motor. Hoje BLOQUEADO por decisões, sem IDs no principal |
| M3 — correspondências | Esquema e proveniência aprovados | Tabela de correspondências por BaseID/hash/aba/endereço/fingerprint; evidências por vínculo | CONFIRMADO com responsável; AMBIGUO/SEM_EVIDENCIA ficam pendentes. Proibido unir somente por placa/NF-e ou escolher primeiro |
| M4 — IDs/vínculos | Correspondências aprovadas | Staging isolado; Vehicle/Episode/Movement/Projection/Operation imutáveis; versões; vínculos históricos e mapa reversível | Zero colisões/órfãos nos elegíveis, episódio/custódia coerente, repetição da migração não duplica. IDs sintéticos F8 não migraram legado |
| M5 — leitores em sombra | M4 validado e principal preservado | Integrar em candidato aprovado adaptadores de pesquisa, filtros, token/ListBox, formulário, navegação e contagem; revalidar identidade antes da ação | Ativo/baixado/lacunas/reentrada/duplicado/obsoleto/reordenação aprovados nos controles reais. F8 só prova adaptadores e carga de controle, não todos os eventos legados |
| M6 — fórmulas/relatórios | Leitores e projeções estáveis | Substituir placa/posição por episódio/movimento; controlar unicidade; fórmulas auxiliares e anuais, totais, formatos/Ignore, validações e condicional | Nenhuma associação silenciosa após classificar/mover/reusar. Comparar identidade e valores distintos; recalcular todo o catálogo real. F8 mede até 200 fórmulas nativas, não dezenas de milhares |
| M7 — gravadores/recuperação | Protocolo autoritativo e M5/M6 aprovados | Adaptar eventos reais sem renomear; envelope completo, CAS, intenção durável, efeitos, recibos e conciliação | Compilar no Excel; testar Venda/T/D/Cancelado, cada destino e fronteira, Save, replay e sessão fria. Recuperação integral F6 ainda FAIL; guardas não satisfazem este portão |
| M8 — política de baixa | Aprovação expressa e M1–M7 concluídos | Substituir sete Delete somente no candidato, conforme política aprovada e leitores migrados | Novas evidências resolvem V_VALIDO, T_VALIDO, D_VALIDO, D_FALHA_BAIXA e T_DUP_DESTINO; zero erro silencioso; nenhuma gravação habilitada pelo leitor antigo |
| M9 — homologação/promoção | Sem críticos essenciais e decisões fechadas | Comparativo final, backup/restauração, riscos aceitos, matriz de eventos reais e autorização separada de promoção | Assinatura do responsável; distribuição/substituição operacional requer autorização própria. NÃO HOMOLOGADO hoje |

## Contratos provisórios de identidade

VehicleID identifica o veículo interno comprovado; placa é atributo revisável. EpisodeID identifica uma passagem/custódia segundo decisão específica: proposta preserva episódio na transferência/devolução interna e cria nova projeção/local/versão; reentrada após encerramento cria novo episódio. MovementID é fato individual com origem/destino e episódio explícitos. ProjectionID identifica uma projeção, não a linha. OperationID identifica comando lógico; AttemptID identifica tentativa; ExpectedVersion evita escrita com leitura antiga. Os cinco IDs de domínio são imutáveis; correção é revisão/fato substitutivo com proveniência. Não reutilizar IDs nem inferir vínculo ausente. A custódia proposta continua pendente de aprovação.

Gerador futuro deve produzir IDs opacos, verificar colisão na autoridade, ter índice único e lidar com cópias divergentes. O laboratório usa literais sintéticos para facilitar provas; não é prova nova de colisão de GUID. Legado sem evidência permanece consultável em modo restrito, sem autorização para mutação ambígua.

## Reversão por situação

| Situação | Procedimento | Proibição |
|---|---|---|
| Sem novas movimentações | Congelar candidato, conferir hashes e restaurar conjunto M0 compatível (autoridade/projeção/diário/schema) | Misturar planilha antiga com diário novo |
| Com novas movimentações | Bloquear escrita, preservar ambos conjuntos, extrair fatos e recibos posteriores, conciliar por ID, preparar revisão de retorno e reconstruir projeções | Restaurar cegamente XLSM antigo, perdendo operações |
| Migração parcial | Não promover; identificar etapas/geração, validar staging e mapa de proveniência; descartar candidato apenas após preservar evidências | Executar leitores antigos sobre esquema novo parcialmente preenchido |
| Histórico ambíguo | Manter pendência e consulta restrita; responsável decide vínculo com evidências | Completar chave/placa/episódio inventados ou usar primeira ocorrência |
| Autoridade divergente | Quarentena, comparar BaseID/linhagem/revisão/recibos/hashes; decisão de conciliação | Eleger arquivo pelo nome/data ou somar cópias sem análise |

## Restauração do laboratório F8

Os originais não precisam de reversão. O checkpoint e o manifesto inicial preservam a referência F7 e todos os documentos anteriores. Para repetir, copiar a entrega/insumos para TEMP novo, conferir SHA-256, manter os marcadores exigidos pelas guardas ou adaptar exclusivamente o componente de teste em nova revisão; importar componentes pelo VBE, compilar e executar cenários em cópias novas. Não executar evidências dentro do OneDrive. Não sobrescrever TSV/fixtures anteriores. Resultados de cada revisão são imutáveis no arquivo de evidências.

O catálogo de erros e o [pacote de decisões](./PACOTE_APROVACAO_PORTAO2.md) definem o que ainda impede integrar. Nenhuma nova alteração fiscal foi proposta.
