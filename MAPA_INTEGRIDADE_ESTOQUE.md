# Mapa de integridade do estoque — Fase 5

**A08/A09 críticos; nenhuma política de baixa modificada.** Este mapa complementa o [contrato de identidade proposto](./PLANO_IDENTIDADE_REGISTROS.md) e a [matriz real](./MATRIZ_REGRESSAO_FASE5.md). Identidade estável continua condicionada à aprovação D01/D02. Nenhuma coluna, tabela, status ou identificador foi criado.

## Escritas observadas nos eventos reais

| Fluxo / evento em Cad_NFE_Saida | Sequência observada na fixture | Fronteira e resultado |
|---|---|---|
| Venda / CommandButton_Gravar_Click | Saidas → Todas Saidas → baixa em Entradas → Save | A:N deslocado; referências diretas quebram e indireta passa ao veículo seguinte. FAIL |
| Transferência / CommandButton_Gravar_Click | Saidas → Entradas Filial → Entradas Anual Filial → Todas Saidas → baixa em Entradas → Save | Destino L7 String/44/@/Ignore e M7/N7 corretos; origem e referências inconsistentes. FAIL |
| Devolução / CommandButton_GravarFilial_Click | Saidas Filial → Todas Saidas Filial → Entradas → Entradas Anual → baixa em Entradas Filial → Save | A:M deslocado; N/O não acompanham a identidade. FAIL |
| Cancelado / CommandButton_Gravar_Click | Saidas → Todas Saidas → Save | Entrada permanece; sem XML pelo comportamento atual. PASS funcional restrito; D03 fiscal pendente |

As quatro gravações da Devolução precedem a exclusão da origem. Quando Delete falhou com erro 1004 após proteção de teste, os destinos permaneceram gravados e a origem também. A mensagem exigiu conciliação. Não houve segunda tentativa.

## Sete exclusões preservadas

As linhas abaixo são as do fonte VBA embutido da base Fase 5, incluindo atributos; podem diferir da numeração do editor. [Comparação dos fontes](./_auditoria/fase5_20261101_202114Z/evidencias/comparacao-fontes-embutidos.json).

| Fonte / procedimento | Linha e comando / intervalo | Vínculo e risco |
|---|---|---|
| [Cad_NFE](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm), CommandButton_Excluir_Click | 140, Selection.Delete Shift:=xlUp; A:N | Find por número; auxiliares fora da largura não se deslocam |
| Cad_NFE, CommandButton_ExcluirFilial_Click | 211, c.Resize(1,13).Delete; A:M em Entradas Filial | Find por número; N/O e referências externas ficam expostos |
| Cad_NFE, CommandButton_ExcluirFilial_Click | 217, cAnual.Resize(1,16).Delete; A:P anual filial | Segundo Find independente por número; não prova mesmo registro histórico; falha entre exclusões |
| [Cad_NFE_Saida](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm), CommandButton_ExcluirFilial_Click | 60, Selection.EntireRow.Delete | Exclusão física de saída filial; desloca a linha inteira e referências externas |
| Cad_NFE_Saida, CommandButton_GravarFilial_Click | 264, c.Resize(1,13).Delete; A:M filial | Baixa da devolução; reproduzido N4=#REF! e O4 de outro veículo |
| Cad_NFE_Saida, CommandButton_Excluir_Click | 494, Selection.EntireRow.Delete | Exclusão física da saída matriz; vínculos históricos não definidos |
| Cad_NFE_Saida, CommandButton_Gravar_Click | 890, c.Resize(1,14).Delete; A:N matriz | Baixa de Venda/Transferência; referências da demonstração quebram ou passam a outro veículo |

Não houve conversão para EntireRow.Delete nem ampliação de intervalos. Uma instrução comentada de exclusão não foi contada como oitava. As exclusões de edição/manual não foram executadas sobre identidade ambígua.

## Mapa de impacto para futura identidade estável

| Área / componentes e procedimentos | Dependência atual | Adaptação que dependerá de decisão |
|---|---|---|
| Cad_NFE.CommandButton_Gravar_Click / GravarFilial_Click | Vaga física pela coluna A, matriz e histórico; snapshots por Range | Definir passagem do veículo, vínculo dos destinos e reutilização da vaga sem reaproveitar identidade antiga |
| modIntegridadeOperacao.PrimeiraVagaGravacao / ProximaLinhaGravacao | Primeira A vazia / última A + 1 | Diferenciar disponibilidade física de identidade e histórico; não reutilizar linha apenas porque A ficou vazia sem contrato |
| Cad_NFE.CommandButtonAlterar_Click | Find em Entradas A:A por TextBoxPesquisa; limpa A:K e procura primeira vaga. Em Entradas Anual limpa a ActiveCell previamente selecionada sem localizar histórico | Resolver origem e anual inequivocamente antes de qualquer limpeza. Não selecionar automaticamente o primeiro histórico repetido. Não há correção segura sem D02 |
| Cad_NFE.CommandButton_ExcluirFilial_Click | Dois Find separados por número e duas exclusões | Aprovar vínculo matriz/filial/anual e comportamento quando zero, um ou vários candidatos |
| Cad_NFE_Saida gravadores matriz/filial | Origem localizada por NF de entrada; destinos determinados por posição; baixa ocorre após escritas | Exigir identidade aprovada da passagem; pré-validação de correspondência e conciliação após fronteira estrutural |
| MóduloPesquisa.PesquisarEntradasRapido / UltimaLinhaUsada | Percorre array até última A usada, começando na linha 2; não termina na primeira lacuna. Coluna oculta 12 guarda número físico da linha, coluna 0 a aba, 13 a chave filial | Revalidar identidade ao abrir/editar; invalidar resultado se a planilha mudar. Manter travessia de lacunas |
| MóduloPesquisa.AbrirRegistroDaLista; Cad_NFE.ListBoxLista_Change / ListBoxLista2_Change | Navegação usa linha oculta 12; dados exibidos são cópia do resultado. ListIndex negativo é protegido | Não confiar na posição armazenada após inserção, exclusão ou reutilização. Atualizar resultados/revalidar contra identidade |
| MóduloPesquisa.IrParaCelulaNFe | Find em todas as células por número, não exclusivamente coluna A; retorna um candidato | Delimitar domínio e tratar ambiguidade depois do contrato, sem inventar unicidade do nNF |
| Cad_NFE_Saida.PreencheCampos / PreencheCabecalho / PreencheLista e variantes 2 | While .Cells(...).Value <> Empty encerra na primeira lacuna | Com limpeza sem deslocamento, registros posteriores podem ficar invisíveis. Adaptar antes de considerar estratégia B/C integrada |
| Cad_NFE_Saida.ListBoxLista_Change / ListBoxListaFilial_Change | Copiam colunas do resultado e usam valores de NF/veículo; não são identidade estável | Distinguir dados da linha e identidade. **Coluna 12 dessas listas é percentual**, não confundir com linha oculta da pesquisa rápida |
| Cad_NFE.CommandButtonCopiarAutoEntradaFilial_Click / CommandButtonChaveR_Click | Busca na coluna B e varredura até última L, respectivamente | Rever domínios e correspondência da passagem; chave de NF-e não substitui identidade do veículo/item |
| UserFormFiltro: Compra, Entrada, Aguardando, Repasse, Venda, Transferência, Devolução | AutoFilter contextual A3:N500 e navegação End(xlDown) | Limite fixo e navegação atravessam blocos conforme lacunas; não substituir navegação contextual sem teste de interface |
| mdlCopiar.UF_NavegarCelula / MoveActiveCell e botões de cópia/navegação dos formulários | ActiveCell, deslocamento de linha e planilha contextual | Revalidar origem selecionada; movimentação física não é identidade persistente |
| Fórmulas auxiliares e demonstrações | Endereço físico, referência direta/indireta e auxiliares fora do Delete | Migrar somente fórmulas previamente inventariadas; comprovar vínculo com veículo, não só ausência de #REF! |
| Históricos anuais | Número de NF repetível e linha física; seleção anterior pode controlar edição | Contrato explícito de registro histórico, operação e passagem; manter ambíguos separados para decisão |
| Cad_NCM.PreencheCampos / PreencheCabecalho / PreencheLista | Catálogo também encerra na primeira célula vazia | Dependência adjacente, não identidade de estoque; preservar sem remoção automática |

Inventários direcionados: [207 linhas candidatas iniciais](./_auditoria/fase5_20261101_202114Z/evidencias/mapa-identidade-posicao.json) e [procedimentos/linhas complementares](./_auditoria/fase5_20261101_202114Z/evidencias/dependencias-posicao-direcionadas.json). São buscas estáticas, com falsos positivos (por exemplo percentual em ListBox e teste IsEmpty do relatório), **não contagem de defeitos** nem execução funcional desses leitores.

## Provas com sentinelas

A fixture semeou A/B/C nas linhas 4–6 e lacunas 2/3. M = I*2, N = A e O = marcador do veículo. Em Demonstração, Q/R referenciam diretamente NF/placa da origem e S usa INDIRECT sobre a posição F4. Após Delete matriz, Q4/R4 viraram #REF!, S4 retornou placa B e O4 ainda dizia MATRIZ_A. Após Delete filial, N4 ficou #REF!, O4 FILIAL_A e S5 retornou B. Essas fórmulas sentinela são **instrumentação sintética**, não afirmação de que Q:S existem dessa forma no sistema original. Evidenciam riscos diferentes: erro explícito, associação silenciosa errada e auxiliar não deslocado.

Na Transferência, os novos M7 = I7*2 e N7 = A7 permaneceram válidos, e O7 manteve DESTINO_SENTINELA. Isso aprova somente a preservação do destino observado; não aprova o fluxo completo, que falhou na origem.

## Contrato e migração: proposta não integrada

A identidade deverá representar **uma passagem de um veículo/item pelo estoque**, separada da posição da linha e da identidade da NF-e. Placa pode mudar ou retornar; nNF não é globalmente único; uma chave pode representar vários itens e não está completa no legado. A movimentação matriz–filial deve manter uma ligação explícita à passagem de origem, sem tratar automaticamente chave da saída como identidade da compra original.

Antes de criar esquema, o responsável deve aprovar: evidências que resolvem cada passagem; multiplicidade de itens/veículos por NF-e; ligação de cada histórico; novo ingresso do mesmo veículo; regras para registros antigos sem chave; tratamento de zero/múltiplos candidatos; critérios de baixa e reutilização de posição. Registros ambíguos exigem resolução humana, sem escolher primeiro Find, placa ou NF isolada.

Migração proposta: inventário somente de leitura → conjunto de candidatos e ambiguidades → validação responsável → mapeamento reversível legado/passagem/históricos → adaptação dos leitores e fórmulas → execução paralela em cópia sintética e depois homologação autorizada. Não há diário persistente/idempotência comercial implantados. Critério: todas as referências do mapa continuam apontando ao veículo e à passagem corretos após lacuna, transferência, retorno, duplicidade histórica e recuperação. Aprovação da arquitetura não equivale a autorização de migrar dados reais.
