# Arquitetura de identidade — Fase 6

Execução: 20261101_224949Z. Situação: **proposta experimental; integração no produto NÃO APROVADA**. A08/A09 continuam críticos. Não houve migração de dados operacionais nem alteração dos sete Delete do produto.

## Referência e limites

A fixture da Fase 5 foi localizada e verificada: SHA-256 f989289554a5c081cef8bd13de2e0b6f0e48ec4c1ba1dd5e75c349eff458f72c, 56 componentes e 17 planilhas. As 518 entradas de seu manifesto foram conferidas antes das alterações documentais; os 39 originais também coincidiram. Os arquivos das cinco execuções anteriores foram inventariados para nova verificação final. Veja [manifesto inicial](./_auditoria/fase6_20261101_224949Z/manifesto-inicial-fase6.json).

O modelo foi implementado exclusivamente em fixtures sintéticas. O protótipo possui cinco abas auxiliares e usa U:X para vínculos sintéticos. **Esses endereços não constituem proposta de ocupação de colunas do produto**: as 17 abas precisam de reserva de esquema aprovada. Campos, abas e estados abaixo são conceitos propostos, não componentes existentes da aplicação.

## Contrato de identidade

| Identificador proposto | Significado e cardinalidade | Regra |
|---|---|---|
| VehicleID | Um veículo físico; várias passagens ao longo do tempo | Não derivar da placa; origem da identificação e revisão humana devem ser registradas. Chassi ou outros atributos somente após confirmar disponibilidade e qualidade. |
| EpisodeID | Uma passagem identificada pelo estoque | Nova entrada depois de uma venda encerrada gera novo episódio, relacionado ao anterior quando comprovado. |
| MovementID | Um evento sobre o episódio | Entrada, venda, transferência, devolução ou outra operação aprovada; imutável após confirmação. |
| ProjectionID / HistoryID | Um registro em uma projeção ou histórico | Uma movimentação pode ter várias projeções; cada uma tem ID próprio e vínculo explícito. A linha é somente uma dica de localização. |
| IntentID / OperationID | Um comando de negócio aceito | Preservado entre tentativas da mesma intenção; não recalcular a partir da NF-e. No protótipo os dois nomes compartilham o mesmo GUID. |
| AttemptID | Uma execução da intenção | Novo em cada tentativa; vinculado à operação original. |
| Revision / ExpectedVersion | Versão esperada do episódio/projeção | Comparação antes de gravar; divergência exige recarregar e conciliar. Não substitui trava de concorrência. |

Relações: VehicleID 1:N EpisodeID; EpisodeID 1:N MovementID; MovementID 1:N ProjectionID. OperationID vincula entrada validada, tentativa, movimentação e todos os destinos esperados. A projeção registra EpisodeID, VehicleID, MovementID, OperationID e seu próprio ID. Não se procura novamente o histórico apenas por placa ou NF-e.

Transferência matriz/filial e devolução foram modeladas experimentalmente como continuidade do mesmo episódio, com mudança de custódia e nova projeção ativa. Essa semântica precisa de decisão D01: se a operação real exigir episódios distintos, o contrato deverá usar relações predecessor/sucessor explícitas. O protótipo não aprova essa escolha.

Cancelado preserva no ensaio os dois destinos de saída e a permanência da entrada. O estado interno não comprova evento fiscal e não altera D03.

## Geração, persistência e resolução

O protótipo gera GUID local por CoCreateGuid/StringFromGUID2, com declarações compatíveis com VBA7. Verifica a existência antes de aceitar e recusa colisão artificial; isso reduz risco, sem prometer impossibilidade matemática de colisão. GUID não codifica informação fiscal nem certifica a identidade de um legado ambíguo.

No produto, gerar o ID uma única vez ao preparar a intenção, persistir antes das projeções e usar índices de unicidade no armazenamento aprovado. Se a geração ou a persistência falhar, não iniciar a movimentação. Identificador aceito é imutável; jamais reutilizar após baixa ou descarte de linha. Cópias de trabalho precisam carregar a identificação da base e sua revisão, para que um comando de outra base não seja aceito.

Resolver uma origem exige: episódio existente uma única vez; veículo declarado; episódio ativo na custódia esperada; projeção atual única; vínculo entre todos os IDs; conteúdo e versão conferidos. Zero ou múltiplos candidatos bloqueiam. Não usar primeiro Find, primeira placa, primeiro número, ListIndex ou ActiveCell como identidade. A seleção da interface deve transportar o ID oculto e a versão e resolver novamente ao confirmar.

Na revisão R2, um selo textual canônico captura tipos, Value2, Formula, NumberFormat e Ignore de A:L e é conferido em todas as projeções do episódio antes da operação e antes de aceitar replay de uma operação concluída. Esse selo é uma verificação de coerência do protótipo, **não é assinatura, hash criptográfico nem proteção contra adulteração**. A captura não cobre todo o formulário e precisa evoluir para um envelope versionado completo, aprovado e minimizado. Não se deve esconder diferenças de dados alterando também o selo.

As chaves sintéticas mantêm 44 caracteres, zero inicial, tipo String, NumberFormat @ e Ignore. Ausência de chave não impede identificar internamente um episódio já comprovado; não dispensa a validação XML ou fiscal vigente no produto.

## Unicidade, reentrada e históricos

Número de NF-e, placa e chave não são chaves internas universais. Número pode repetir em contextos distintos, placa pode retornar e chaves podem estar ausentes. Regras fiscais de duplicidade continuam separadas das restrições de identidade. Uma tentativa repetida só pode ser reconhecida pelo OperationID persistido e pelo mesmo conteúdo canônico/versão/efeitos esperados.

Mesmo OperationID com conteúdo diferente: recusar e registrar divergência. Operação concluída com efeitos íntegros: retornar o recibo existente, sem novos destinos. Operação pendente ou efeito alterado: conciliação antes de repetir, inclusive com novo OperationID. Compensação comprovada permite nova tentativa vinculada à mesma intenção, preservando o histórico de tentativas.

Histórico não deve ser alterado por posição ou primeira coincidência. Proposta D02: correção identificada, com revisão anterior preservada e relação Supersedes, motivo, responsável e revisão. Não apagar silenciosamente a evidência anterior. O protótipo bloqueia histórico ambíguo; não migrou CommandButtonAlterar_Click nem implementou edição versionada no produto.

## Migração reversível proposta

1. Congelar a revisão de referência e conferir hashes; criar cópia sintética de ensaio e mapeamento de proveniência somente leitura.
2. Inventariar cada registro legado por base, aba, endereço original, hash/conteúdo e contexto disponível. Esse endereço é evidência da migração, não identidade futura.
3. Construir candidatos de VehicleID e EpisodeID com dados efetivamente existentes. Comparar datas, operação, veículo, origem/destino e documentos disponíveis; classificar correspondência confirmada, ausente ou ambígua.
4. Submeter zero/múltiplas correspondências e duplicidades históricas à conciliação humana. Não preencher lacunas por suposição nem tratar um GUID novo como prova de vínculo.
5. Versionar o mapa aprovado, registrar decisão/proveniência e atribuir IDs aos registros confirmados. Legados não resolvidos ficam bloqueados para movimentação dependente, com consulta histórica preservada.
6. Em uma revisão conjunta, adaptar todos os escritores, leitores, fórmulas, listagens, totais, backups e conciliação; testar todos os fluxos. Proibir execução mista de gravador antigo sobre base com nova política.
7. Fazer corte controlado com checkpoint, mapa de versões e reconciliação de quantidades, valores, fórmulas e vínculos. A autorização de corte é separada desta Fase 6.
8. Retorno antes de qualquer operação nova: restaurar checkpoint completo e mapa anterior. Depois de operações novas: primeiro exportar/reconciliar os fatos e construir restauração compatível; simplesmente abrir o XLSM antigo pode perder movimentações.

Não houve aplicação desse procedimento a dados reais.

## Critérios de aceite para futura integração

Exigir: nenhuma associação silenciosa; leitores retornando somente episódios elegíveis; históricos por ID; reentrada sem herdar a venda anterior por placa; filtros atravessando lacunas; fórmulas resolvendo IDs; travas/versões e conciliação; diário e backup persistentes aprovados; ensaios reais dos UserForms migrados e dos sete pontos de baixa; referência original preservada. Reordenação física, edição manual e abertura de cópias divergentes precisam ser detectadas ou proibidas pelo contrato.

Os resultados e as limitações executadas estão em [protótipos](./PROTOTIPOS_INTEGRIDADE_FASE6.md) e [matriz](./MATRIZ_REGRESSAO_FASE6.md). A proposta de recuperação está em [arquitetura operacional](./ARQUITETURA_RECUPERACAO_FASE6.md). Aprovações pendentes: [decisões](./DECISOES_ARQUITETURA_FASE6.md).
