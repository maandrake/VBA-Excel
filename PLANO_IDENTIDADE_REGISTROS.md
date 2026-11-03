# Plano de identidade dos registros — proposta da Fase 4

**PROPOSTA NÃO APROVADA E NÃO IMPLEMENTADA.** Nenhum campo, coluna, tabela, status, identificador, leitor ou Delete foi alterado. A aprovação arquitetural condicional da estratégia C não autoriza sua implantação. D01 e D02 continuam abertas em [PENDENCIAS_DECISAO](./PENDENCIAS_DECISAO.md).

## Base e problema a resolver

A Fase 3 encontrou 18 grupos de números repetidos no histórico matriz e um no filial; somente 17 das 42 entradas filiais têm chave de 44 dígitos armazenada como texto. Esses são resultados históricos de leitura OOXML local, não novas consultas nem certificação fiscal. [Estatísticas](./_auditoria/fase3_20261031_213925Z/identidade-existente.json). Não foram exportados valores reais de nota, placa ou chave.

Os escritores atuais não persistem chave completa em todos os cadastros/históricos. Placa pode mudar ou reaparecer; número da NF-e não inclui emitente/série; chave identifica um documento, que pode conter vários itens/veículos, e não determina por si uma passagem pelo estoque. Código de produto e linha física não demonstram identidade de veículo/episódio. Dados incompletos ou iguais não autorizam fundir registros.

## Contrato conceitual a aprovar

| Conceito proposto | Contrato desejado | Evidências e limites |
|---|---|---|
| Veículo | Identidade de longo prazo independente de placa e de localização | Confrontar atributos disponíveis e documentos; chassi/VIN, se efetivamente disponível e válido, pode apoiar a identificação, sem presumir sua existência nas colunas atuais. Não derivar de COD/placa isolados. |
| Passagem pelo estoque | Um episódio de ingresso até sua conclusão operacional aprovada | Identificador técnico opaco, imutável e não reciclado, criado somente após aprovar armazenamento/migração. Não é chave fiscal, número da nota nem posição. |
| Documento/item de ingresso | Evidência de origem do episódio, preservada sem coerção numérica | Quando disponível: chave textual de 44 dígitos + item correspondente + vínculo confirmado com veículo; ausência de chave mantém evidência incompleta, não gera chave inventada. |
| Movimentação matriz/filial | Relação explícita entre origem, destino e documento/etapa | Não inferir relação por igualdade de linha/número/placa. Decidir se transferência interna mantém episódio global e cria segmentos de custódia ou exige episódios relacionados; recomendação é episódio global com segmentos quando a regra operacional confirmar continuidade. |
| Retorno do veículo | Novo ingresso após saída definitiva, com ligação ao episódio anterior quando comprovada | Não reutilizar a identidade da passagem encerrada. Diferenciar retorno de veículo de reenvio/repetição técnica da mesma operação. Devolução que desfaz ou inicia episódio exige decisão do responsável. |
| Histórico | Relação com episódio e operação, separada da posição atual | Definir se registro anual é evidência imutável, projeção ou ambos. Edição/cancelamento por nova versão/evento é proposta, não regra atual. Manutenção/exclusão do anual precisa aprovação explícita. |
| Tentativa de operação | Identidade própria para distinguir nova tentativa de novo fato | Futuro registro persistente pode permitir detectar repetição; precisa conteúdo, local, concorrência, retenção e conciliação aprovados. Identificador em memória não garante idempotência após crash. |

Um UUID/GUID poderia representar a identidade técnica depois de aprovação, sem codificar placa/chave. Sua unicidade deve ser verificada, e eventual colisão recusada. Não foi gerado nem inserido identificador na aplicação. Não usar hash de todos os campos como identidade: editar um campo mudaria a identidade e duplicidades continuariam ambíguas.

## Seleção de candidatos e ambiguidades

1. Preservar os valores originais, tipos e proveniência. Normalização serve a comparação; não sobrescreve documento original. Chaves preenchidas continuam String de 44 dígitos, inclusive zeros; nunca completar uma chave desconhecida com zeros.
2. Separar documento/item, veículo e passagem. Comparar evidências disponíveis, emitente/destinatário e relações temporais/operacionais quando constarem dos dados; esses atributos não são declarados existentes apenas por estarem no contrato proposto.
3. Produzir candidatos, com motivos explícitos. Zero candidatos significa pendência; mais de um significa ambiguidade. Mesmo um único candidato estatístico precisa satisfazer o contrato aprovado. Não escolher primeiro Find, maior linha, data mais recente ou maior semelhança.
4. Igualdade de chave sem igualdade do item/veículo não basta; mesmo veículo com novo documento pode ser novo episódio. Duplicação exata pode ser erro ou reprocessamento: preservar ambos até decisão.
5. Chave ausente/incompleta, placa divergente ou conflito entre abas impede migração automática desse vínculo. O responsável valida com evidência local autorizada; não executar consultas fiscais/DETRAN para suprir lacunas.

## Migração proposta, separada da operação atual

**M0 — aprovação:** definir a semântica de ingresso/transferência/retorno/cancelamento, identidade e vínculo anual (D01/D02), sem introduzir regra fiscal (D03). Escolher armazenamento: extensão explícita do modelo atual ou repositório auxiliar autorizado. As opções têm custo de compatibilidade, concorrência, segurança e backup; nenhuma foi criada.

**M1 — diagnóstico preservado:** gerar snapshot somente leitura e manifesto de uma cópia autorizada, registrar por aba/célula a proveniência e fórmulas relevantes, manter originais e tipos. Os números da Fase 3 não substituem reconciliação na data da migração. Não executar esta coleta de dados reais nesta fase.

**M2 — classificação:** formar propostas de vínculo/episódio e fila explícita de pendências. Registrar duplicados e ausência de evidências sem fundir, excluir, ocultar ou inventar dados. Identidade não aprovada permanece bloqueada para operações que dependam dela.

**M3 — homologação fictícia:** implementar esquema aprovado somente em fixture, adaptar consumidores e conferir contagens/valores/fórmulas/chaves antes/depois. Usar cópia versionada e ensaiar reversão antes da migração operacional.

**M4 — migração validada:** depois de autorização específica, atribuir as identidades aprovadas, gerar correspondência antiga→nova com versão e proveniência, manter ambiguidades explícitas e auditar cardinalidades. Uma linha antiga pode participar de relações distintas; não impor artificialmente 1:1 entre todas as abas.

**M5 — corte controlado:** adaptar todos os escritores/leitores na mesma revisão. Evitar operação mista em que um escritor use posição e outro identidade. Liberação e eventual retorno à versão anterior exigem conciliar operações posteriores; reverter só o arquivo pode perder fatos legítimos. Nenhuma etapa M1–M5 foi executada em dados reais.

## Adaptação dos consumidores

| Consumidor atual | Ajuste futuro condicionado | Critério de aceite |
|---|---|---|
| Pesquisa rápida de entradas | Já varre arrays até última linha usada e atravessa lacunas. Aplicar política aprovada de elegibilidade sem confundir linha com identidade; linha permanece endereço de navegação transitório. | Busca encontra registros após lacunas; seleção revalida identidade e versão antes de editar; mudança de posição não abre outro episódio. |
| Resultado ListBox | Hoje guarda origem na coluna 0, linha na 12 e chave na 13. Não alterar silenciosamente o contrato usado por modPreencherFormulario/AbrirRegistroDaLista. | Aprovar transporte de identidade/versão, adaptar todos os consumidores e testar ListIndex=-1, seleção antiga, sem resultado e duplicados. |
| Filtros legados de saídas/NCM | Loops While param na primeira célula vazia; política de lacunas exige revisão separada. NCM não é automaticamente um estoque com baixa. | Cada filtro conserva cabeçalho, ordem e critérios aprovados; pesquisa rápida não recebe correção desnecessária de lacunas. |
| PrimeiraVagaGravacao | Reutiliza célula A realmente Empty. Reutilização física não pode reciclar identidade nem reatribuir histórico antigo. | Nova ocupação tem episódio correto; referências antigas permanecem ligadas ao episódio anterior. |
| ProximaLinhaGravacao/históricos | Acrescenta após última A. Política de anual precisa vínculo definido. | Edição atinge exatamente o registro aprovado, recusa ambiguidade e preserva evidência anterior conforme contrato. |
| Alterar/ExcluirFilial | Remover dependência da seleção anual anterior/primeiro número encontrado somente após aprovar correspondência. | Duplicados e origem ausente não geram escrita/exclusão automática; falha entre abas exige conciliação. |
| Sete Delete | Manter literais atuais até decisão; migrar em conjunto com fórmulas/leitores quando aprovado. | Nenhum #REF!, referência a outro veículo, registro invisível ou perda lateral/histórica. |
| Fórmulas e nomes | Inventariar referências físicas, buscas por placa/número, UDF, nomes e dependências fora do alvo. Não trocar tudo por uma nova busca sem provar cardinalidade. | Reordenação, lacunas e retorno preservam relações e totais; ambiguidades são detectadas. |
| XML, backup, exportação e Ribbon | XML fornece evidência do documento, não inventa episódio. Backup inclui futuro armazenamento e versão da correspondência. Exportações não perdem identidade. | Restaurar arquivo e vínculos na mesma versão; preservar privacidade e chave textual. |

A referência de implementação continua sendo [MAPA_PROJETO](./MAPA_PROJETO.md); a tabela é especificação futura, não alteração efetiva.

## Casos necessários antes de aprovação

| Caso fictício | Resultado exigido na futura implementação | Estado nesta fase |
|---|---|---|
| Mesmo número, emitentes/séries diferentes | Sem fusão automática | NÃO EXECUTADO |
| Mesma NF-e com dois veículos/itens | Evidências/documento compartilháveis; episódios/itens distintos conforme regra | NÃO EXECUTADO |
| Mesmo veículo vendido e depois readquirido | Novo episódio relacionado, sem ressuscitar o anterior | NÃO EXECUTADO |
| Transferência matriz→filial e retorno interno | Relação de custódia aprovada, sem dupla contagem | NÃO EXECUTADO |
| Chave ausente/incompleta ou erro de tipo | Pendência explícita; sem gerar chave fictícia em migração real | NÃO EXECUTADO |
| Dois candidatos históricos ou seleção anual errada | Operação recusada para decisão; não usar primeiro Find | NÃO EXECUTADO |
| Lacuna/reutilização/reordenação | Pesquisa e vínculo histórico corretos; fórmulas preservadas | NÃO EXECUTADO |
| Repetição após falha parcial/crash | Reconhecer operação persistida e conciliar; sem promessa ACID | NÃO EXECUTADO |
| Backup/restauração e reversão | Identidades, relações e arquivo de dados da mesma versão | NÃO EXECUTADO |

Os protótipos COM anteriores continuam como evidência do risco, sem novos ensaios equivalentes. Para autorizar integração, o responsável deve aprovar D01/D02 e o esquema concreto, e a implementação futura precisa compilar e passar pelos casos acima. **O plano não resolve A08/A09 por si só.**
