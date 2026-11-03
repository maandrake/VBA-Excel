# Contrato de identidade — Fase 7

Estado: especificação implementável proposta, Portão 1. Integração, esquema e semântica de custódia dependem de aprovação explícita. Não modifica o produto.

## Entidades e ciclo de vida

| ID | Criação / escopo / imutabilidade | Relações e término |
|---|---|---|
| VehicleID | GUID local ao confirmar identidade física comprovada; imutável, nunca derivado da placa | 1:N EpisodeID. Identificação física exige evidência/proveniência; placa/chassi/documento são atributos verificáveis, nenhum isoladamente impõe fusão automática. Correção de identidade exige decisão registrada, sem reaproveitar ID. |
| EpisodeID | GUID ao aceitar uma nova passagem pelo estoque; imutável | Pertence a um VehicleID; relaciona entrada, custódias, movimentos e históricos. Venda encerra episódio; reentrada comprovada cria outro, com predecessor explícito. |
| MovementID | GUID reservado na preparação de uma operação aceita | Fato imutável após confirmação; várias projeções ligadas. Correção cria fato/revisão substitutiva com Supersedes, motivo e responsável, preservando anterior. Não reutilizar ID de movimento abortado para outro fato. |
| ProjectionID | GUID reservado por efeito do plano, antes da escrita | Um registro lógico de uma projeção; não é a linha. Cada aba/histórico tem ID próprio e FKs de episódio/movimento. Mudança de endereço não muda ID. Não reutilizar ID/linha para substituir entidade histórica. |
| OperationID | GUID da intenção persistida antes do primeiro efeito | Mesmo comando e envelope canônico preservam ID entre tentativas; recibo concluído íntegro retorna resultado existente. Payload diferente recusa. OperationID não é AttemptID; nova tentativa recebe AttemptID, sem nova intenção artificial. |

Todos os cinco IDs são imutáveis e não têm significado fiscal. BaseID imutável e SchemaVersion/Revision são requisitos de autoridade; não são autorização para criar colunas. GUID por CoCreateGuid já exercitado na F6; conferir unicidade no domínio persistido escolhido e recusar colisão, formato inválido ou ausência. Não há promessa matemática de impossibilidade de colisão. Campos sem ID ficam pendentes de conciliação, nunca resolvidos por primeira ocorrência.

## Custódia — decisão D01-C específica

Recomendação: EpisodeID representa a passagem global pelo estoque da organização; transferência e devolução interna preservam EpisodeID e VehicleID, criam MovementID e ProjectionIDs de custódia/históricos novos. Apenas uma projeção de custódia atual por episódio; anterior permanece consultável historicamente e inelegível para nova saída. Transferência não é reentrada após venda.

Alternativa: episódio por estabelecimento, com predecessor/sucessor e relação de transferência explícita. A decisão depende da definição operacional do responsável. Os ensaios F6 usam a primeira alternativa, sem aprová-la. Devolução de cliente após venda encerrada cria novo episódio se houver nova admissão; não confundir com o botão Devolução filial existente. Nenhuma regra fiscal foi inferida dessas distinções. Cancelado mantém o comportamento atual até D03.

## Resolução e elegibilidade

Resolver(ProjectionID, EpisodeID, VehicleID, ExpectedVersion, BaseID) deve retornar exatamente um registro ou erro explícito AUSENTE, AMBIGUO, DIVERGENTE, OBSOLETO ou PENDENTE. LinhaDica é cache descartável, não fallback. Localização atual é resolvida novamente ao confirmar a seleção. Verificar unicidade da projeção em todas as partições permitidas, vínculos, revisão e envelope. ListIndex/ActiveCell/Find/placa/NF-e somente ajudam a navegação ou a reunir candidatos. Nunca autorizam escrita sozinhos.

Pesquisa operacional exige episódio ATIVO, custódia correspondente, projeção atual única, ausência de pendência e versão conferida. Consulta histórica tem finalidade explícita e pode exibir encerrados, sem permitir saída automática. Não considerar toda linha preenchida ativa. Contagem e totais usam exatamente a mesma elegibilidade; inconsistência gera erro visível, não desaparecimento silencioso.

## Envelope, versões e documentos

Definir envelope versionado de todos os campos que influenciam efeitos: IDs/base, tipo, custódia origem/destino, versão, dados monetários/datas convertidos, documento e itens confirmados, plano de projeções e regras de cálculo vigentes. Canonicalização por nomes/ordem/tipos/comprimentos, distinção vazio/null/zero e precisão/escala aprovadas. Hash SHA-256 sobre bytes de codificação definida; hash não é assinatura/autenticidade. F6 usa apenas selo textual A:L e payload parcial: insuficiente para produção.

Chave NF-e é String de exatamente 44 dígitos quando preenchida, formato @ antes da atribuição, Ignore preservado. Chave/documento/item ficam relacionados ao episódio/fato, não o substituem. Documento ausente não autoriza adivinhar vínculo nem dispensar validação fiscal existente. Mesmo número em contextos distintos não funde episódios.

ExpectedVersion é comparado antes de gravar e antes da confirmação sob escritor exclusivo; incremento apenas na transição confirmada. Recibo idempotente de operação já concluída pode devolver resultado mesmo com versão corrente posterior, desde que envelope original e efeitos sejam íntegros; não aplicar novamente o comando obsoleto. Versão otimista sozinha não trava duas instâncias.

## Legado e históricos

Mapa de proveniência: hash da base, aba/endereço original, fingerprint tipado, evidências e decisão. Endereço é localizador da migração, nunca identidade definitiva. Classificar CONFIRMADO, SEM_EVIDENCIA ou AMBIGUO; somente confirmado recebe vínculo aprovado. GUID novo não prova que duas linhas pertençam ao mesmo episódio. Não usar placa/NF-e coincidentes como critério suficiente. Históricos têm ProjectionID próprio e FK para MovementID/EpisodeID; correções preservam revisão anterior e trilha.

Migração, autoridade e leitores são inseparáveis da baixa. Critério de aceite: nenhuma referência passa silenciosamente a outro veículo/episódio; ausências e duplicidades produzem erro/bloqueio explícito. Os protótipos seguintes demonstrarão subconjuntos desse contrato; não demonstram o contrato integral ou os eventos reais migrados.
