# Plano de migração reversível — Fase 7

**Plano proposto, Portão 2 não aprovado. Nenhum dado real foi migrado.** O corte da baixa depende de leitores, fórmulas, históricos e persistência adaptados no mesmo candidato.

## Pacotes, saídas e portões

| Etapa | Atividade / artefato exigido | Critério para avançar | Reversão |
|---|---|---|---|
| M0 — congelar | Definir responsável, autoridade, janela sem movimentações; copiar conjunto aprovado com SHA-256, versões, esquema, fontes, FRM/FRX e CodeNames | Base íntegra, backup reaberto/testado e escritor exclusivo | Nada operacional foi alterado; descartar candidato mantendo evidências |
| M1 — inventário | Exportação local controlada de registros/células, fórmulas compartilhadas, nomes, tabelas, controles, OnAction, UDF, listas e integrações; mapa nominal de leitores | Todos os leitores/escritores conhecidos possuem dono, critério e regressão; dependência desconhecida bloqueia o corte afetado | Restaurar checkpoint; inventário somente leitura |
| M2 — esquema | Aprovar armazenamento de IDs/fatos/diário e esquema versionado; não ocupar colunas presumidas livres | D01-C, D02 e D04/D06 específicos aprovados | Migração de esquema somente em cópia; principal preservado |
| M3 — correspondências | Gerar mapa de proveniência por hash da base/aba/endereço/fingerprint; apresentar candidatos e evidências sem vincular automaticamente | CONFIRMADO com decisão registrada; SEM_EVIDENCIA/AMBIGUO ficam pendentes de conciliação | Mapa reversível, jamais apagar linhas de origem ou evidências |
| M4 — identidade/histórico | Atribuir IDs aos vínculos aprovados; cada projeção própria e FK de episódio/movimento; não fundir por placa/NF-e | Unicidade, órfãos zero nos elegíveis, versões e associação anual verificadas | Antes de novos fatos, restaurar conjunto completo M0; mapa de decisão permanece evidência |
| M5 — leitores em sombra | Implementar repositório/índice e adaptadores; comparação somente leitura com dados de referência; pesquisa rápida/filtros/ListBox/preenchimento/navegação | Ativos/baixados/lacunas/duplicados/reentrada/reordenação e seleção obsoleta aprovados; divergências justificadas | Reativar pacote anterior apenas contra base anterior compatível |
| M6 — fórmulas/relatórios | Migrar referências físicas e COUNTIF por placa para IDs; elegibilidade comum em contagens/totais; adequar auxiliares da filial, validações e formatação condicional | Reordenação/reutilização não reatribuem veículo; erros explícitos para 0/múltiplos, tipos/formatos/Ignore conservados | Restaurar fórmulas/layout e dados compatíveis do checkpoint; não copiar fórmula velha sobre base nova |
| M7 — escrita/recuperação | Adaptar cadastro, edição, exclusão, Venda, Transferência e Devolução; OperationID/envelope, plano completo, versões, snapshots/recibos e conciliação | Compilação real, eventos reais, falhas em cada fronteira, Save e sessão fria do protocolo escolhido | Antes de fatos novos, M0; depois, exportar/reconciliar fatos novos antes de qualquer retorno |
| M8 — baixa/corte | Somente agora ativar política aprovada e substituir sete Delete no candidato, nunca indiscriminadamente | Cinco FAIL F5 resolvidos por novas evidências dos eventos; nenhuma dependência antiga habilitada para escrita | Corte suspenso diante de falha; base em quarentena, plano de reversão aprovado |
| M9 — homologação/promoção | Responsável revisa relatório, diferenças e risco residual; autorização separada de distribuição/produção | Regressão essencial integral e decisões fechadas; sem críticos desconhecidos | Restauração coordenada autoridade/projeção/diário; não abrir cópia velha como autoridade |

## Novos registros e legado

Novos registros recebem IDs uma vez, antes dos efeitos e sob fonte autoritativa aprovada. Identidade física do veículo é comprovada ou permanece pendente; nova passagem após encerramento tem novo EpisodeID. Não reutilizar ID de registro excluído/linha vazia. Transferências e devoluções usam a semântica de custódia aprovada, nunca uma inferência por documento.

Legados sem chave, com placa repetida, nNF repetido ou histórico múltiplo conservam consulta e proveniência. Para associar veículo/episódio/movimento, exigir evidências do contexto, datas, origem/destino/documentos e decisão do responsável. Zero ou múltiplos candidatos bloqueiam movimentação/edição dependente. Um GUID recém-gerado não resolve ambiguidade histórica. Não escolher a primeira ocorrência nem completar dados ausentes inventados.

Histórico corrigido cria revisão/fato substitutivo e trilha, conforme D02; não atualizar silenciosamente a linha anual apontada por ActiveCell. Snapshot da célula precisa conferir ID/versão antes de restaurar; o endereço original pode ter outro ocupante.

## Contratos de integração

Interfaces propostas: ResolverProjecao(BaseID, ProjectionID, EpisodeID, VehicleID, ExpectedVersion); ConsultarEstoque(proposito, local, filtro); PrepararOperacao(OperationID, envelope); ConfirmarResultado(recibo); Conciliar(planoAprovado). São contratos propostos, não funções existentes de produção. Retornos explícitos para ausente, ambíguo, baixado, obsoleto e pendente; nada de fallback para primeiro Find.

ListBox preserva apresentação atual, mas transporta token por adaptador compatível com cada formulário. Não renomear eventos/controles. Busca por NF-e/placa mostra candidatos e contexto; ação revalida identidade atual. Leitor histórico é separado do estoque elegível e não alimenta gravação sem validação.

Fórmulas por ID devem receber dependências explícitas e faixas delimitadas; a UDF experimental não varre colunas inteiras. Índices em memória são invalidados por Revision e por mudança de fonte, não apenas por endereço. Quando a tabela é regenerada, fórmulas e formatos derivados acompanham a projeção inteira. Benchmark em volume real sanitizado e versões de Office suportadas ainda é necessário; nenhum ganho de desempenho foi medido nesta fase.

## Comparação antes/depois e reversão

Guardar matrizes por identidade de: valores/tipos, documentos/chaves String44, fórmulas/resultados, NumberFormat/Ignore, vínculos/versões e contagens ativas por custódia. Comparar quantidades/valores de históricos sem deduplicar silenciosamente; separar erro esperado de fonte antiga de regressão nova. Inspecionar todos os destinos de cada movimento e selecionar registros por IDs, não por ordem física. Datas exigem representação persistente estável: a F7 encontrou normalização de formato embutido ao reabrir.

Reversão sem fatos novos: verificar hashes do checkpoint e restaurar conjunto compatível completo, com autoridade/revisão/diário. Com fatos novos: bloquear escrita, preservar ambas as versões, exportar fatos e recibos, conciliar e produzir revisão de retorno; restauração cega do XLSM antigo perde operações. Não existe desfazer automático garantido depois de Delete/Save.

Neste portão, reversão é simples: remover da área de trabalho apenas o candidato experimental se desejado e restaurar documentação do backup após conferir hashes; preservar evidências. Nenhum principal precisa ser revertido porque não foi alterado. Para repetir, copiar os insumos arquivados para TEMP novo, importar a suíte pelo VBE e verificar guardas/caminhos; não executar evidências dentro do OneDrive.
