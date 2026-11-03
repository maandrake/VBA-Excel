# Matriz de impacto da baixa — Fase 6

Modalidade: **análise estática dirigida**, somada a provas VBA explicitamente indicadas na [matriz de regressão](./MATRIZ_REGRESSAO_FASE6.md). Escopo: componentes efetivamente extraídos da fixture Fase 5, fontes preservados e fórmulas OOXML das 17 abas. Não é aprovação de uma política nova nem prova de cobertura de todos os chamadores dinâmicos/integrações externas.

## Causas dos cinco FAIL obrigatórios

| ID real Fase 5 | Causa raiz | Regressão Fase 6 |
|---|---|---|
| V_VALIDO | Delete A:N desloca dados; referência direta quebra e INDIRECT mantém endereço apontando ao veículo seguinte | A_V_VALIDO, B_V_VALIDO, C_V_VALIDO; reordenação e leitores legados |
| T_VALIDO | Mesma baixa estrutural da matriz depois de quatro gravações; sucesso do evento não verifica associação | A_T_VALIDO, B_T_VALIDO, C_T_VALIDO |
| D_VALIDO | Delete A:M da filial deixa auxiliares N/O na posição anterior; referências e dados perdem correspondência | A_D_VALIDO, B_D_VALIDO, C_D_VALIDO |
| D_FALHA_BAIXA | Destinos já escritos e fronteira estrutural marcada antes de Delete; proteção provoca 1004, sem compensação segura integral | A_D_FALHA_BAIXA e C_FALHA_BAIXA/COMPENSACAO; bloqueio de repetição |
| T_DUP_DESTINO | Identificação por NF-e/placa não distingue episódio; destino aceita ocorrências ambíguas; origem também sofre Delete | A_T_DUP_DESTINO; C_T_DUP_DESTINO com IDs previamente declarados, mais projeção/histórico ambíguos |

A prova C_T_DUP_DESTINO distingue episódios declarados e número repetido em contextos sintéticos; **não resolve automaticamente a duplicidade legada de mesma placa da Fase 5**. Identidade do veículo e migração continuam exigindo conciliação. As fontes de Fase 5 e seus cinco FAIL permanecem intactos.

## Sete pontos Delete preservados

| Componente/procedimento | Ponto no fonte embutido F5 | Alcance legado | Dependências a migrar em conjunto |
|---|---:|---|---|
| Cad_NFE.CommandButton_Excluir_Click | 140 | Entradas A:N | Projeção de entrada, auxiliares, referências e histórico anual |
| Cad_NFE.CommandButton_ExcluirFilial_Click | 211 | Entradas Filial A:M | N/O, correspondência da origem e histórico |
| Cad_NFE.CommandButton_ExcluirFilial_Click | 217 | Entradas Anual Filial A:P | Segunda busca por NF-e; histórico ambíguo |
| Cad_NFE_Saida.CommandButton_ExcluirFilial_Click | 60 | EntireRow em Saidas Filial | Histórico de saída, fórmulas e seleções físicas |
| Cad_NFE_Saida.CommandButton_GravarFilial_Click | 264 | Devolução: Entradas Filial A:M | Quatro destinos e custódia do episódio |
| Cad_NFE_Saida.CommandButton_Excluir_Click | 494 | EntireRow em Saidas | Histórico matriz e associações |
| Cad_NFE_Saida.CommandButton_Gravar_Click | 890 | Venda/Transferência: Entradas A:N | Duas/quatro gravações, origem e referências |

Os números são linhas dos componentes embutidos extraídos, não necessariamente linhas dos FRM com cabeçalho de designer. Nenhum dos comandos foi alterado nesta fase.

## Contratos dos leitores e escritores

| Grupo | Comportamento atual / impacto | Adaptação exigida se C for aprovada |
|---|---|---|
| Pesquisa rápida | Array até última linha de A; atravessa lacunas. Lista carrega aba, linha oculta 12 e chave 13 | Acrescentar elegibilidade por episódio/custódia e ID/versão. Atravessar lacunas não significa excluir registros baixados. C_LEITOR_LEGADO chama o leitor VBA real e evidencia a falha. |
| Filtros legados de saída/NCM | PreencheCampos/Lista com While até primeira célula vazia | Percorrer projeções válidas sem depender de contiguidade; manter contrato de cada lista. Na lista de saída, coluna 12 pode ser percentual, não identidade/linha. |
| UserFormFiltro | AutoFilter, intervalos fixos e End(xlDown) | Adaptar fonte/critério para visão ativa; preservar navegação contextual intencional. |
| ListBox | ListIndex e dados posicionais; seleção abre célula e formulário | ID e versão ocultos; revalidar a seleção antes de escrever. Não reinterpretar uma coluna em todos os formulários. |
| Cadastro e primeira vaga | Busca/CountIf por NF-e; linha disponível por A vazia | Nova projeção com identidade e política de duplicidade separada. Posição disponível não pode reaproveitar vínculo histórico. |
| Alteração | ActiveCell e busca por número; histórico anual pode usar célula anterior | Resolver ProjectionID/HistoryID único e versão; bloquear 0/múltiplos; correção versionada aprovada. |
| Exclusão em duas abas | Find independente por NF-e em cada aba | Um plano de efeitos explícito com IDs vinculados; não remover primeiro candidato. |
| Transferência/Devolução | Gravações sucessivas e baixa física da origem | Mesmo contrato de operação, custódia e vínculos; recibo e pendência de conciliação. |
| Auxiliares de cópia/navegação | ActiveCell/Offset carregam contexto | Guardar ID selecionado e resolver endereço atual; ActiveCell pode continuar para navegação visual. |
| Snapshots e recuperação | Range físico é capturado/restaurado | Conferir identidade/versão da célula antes de aplicar pré-imagem; não restaurar sobre registro que ocupou o endereço depois. |

## Inventário nominal dos pontos examinados

São 50 candidatos da busca dirigida e sete contratos complementares confirmados, **não 57 defeitos**. As linhas abaixo referem-se ao texto embutido da fixture, com evidência em [leitores-identidade-posicao.json](./_auditoria/fase6_20261101_224949Z/evidencias/leitores-identidade-posicao.json). Cada procedimento deve ser revisado no corte de migração; uso de ActiveCell/Find isolado não prova defeito.

| Procedimento | Linha embutida inicial | Dependência observada |
|---|---:|---|
| [clsOperacaoPlanilha.RegistrarDestino](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | 44 | ReferenciaFisica |
| [clsOperacaoPlanilha.RestaurarCelula](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | 118 | ReferenciaFisica |
| [modIntegridadeOperacao.ProximaLinhaGravacao](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/modIntegridadeOperacao.bas) | 19 | LinhaFisica |
| [mdlCopiar.UF_CopiarCelula](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/mdlCopiar.bas) | 39 | ActiveCell |
| [mdlCopiar.UF_AtualizaLabels](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/mdlCopiar.bas) | 62 | ActiveCell |
| [mdlCopiar.MoveActiveCell](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/mdlCopiar.bas) | 111 | ActiveCell, LinhaFisica |
| [MóduloPesquisa.AbrirRegistroDaLista](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/MóduloPesquisa.bas) | 219 | ListIndex |
| [MóduloPesquisa.UltimaLinhaUsada](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/MóduloPesquisa.bas) | 260 | LinhaFisica |
| [MóduloPesquisa.IrParaCelulaNFe](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/MóduloPesquisa.bas) | 303 | Find |
| [Cad_NCM.ListBoxLista_Change](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NCM.frm) | 69 | ListIndex |
| [Cad_NCM.TextBoxFiltro_Change](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NCM.frm) | 157 | ListIndex |
| [Cad_NCM.PreencheCampos](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NCM.frm) | 169 | Lacunas |
| [Cad_NCM.PreencheCabecalho](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NCM.frm) | 185 | Lacunas |
| [Cad_NCM.PreencheLista](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NCM.frm) | 201 | ListIndex, Lacunas |
| [Cad_NFE.CommandButton_Excluir_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 117 | Find, ActiveCell, ReferenciaFisica, LinhaFisica |
| [Cad_NFE.CommandButton_ExcluirFilial_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 189 | Find, ReferenciaFisica |
| [Cad_NFE.CommandButton_Gravar_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 267 | CountIf, ReferenciaFisica, ActiveCell, Lacunas |
| [Cad_NFE.CommandButton_GravarFilial_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 404 | CountIf, ReferenciaFisica, ActiveCell, Lacunas |
| [Cad_NFE.CommandButtonAlterar_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 663 | Find, ActiveCell, ReferenciaFisica, Lacunas, LinhaFisica |
| [Cad_NFE.CommandButtonCopiarAutoEntradaFilial_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 840 | Find, LinhaFisica |
| [Cad_NFE.CommandButtonChaveR_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 1051 | LinhaFisica |
| [Cad_NFE.N_Nfe_Exit](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 1498 | CountIf |
| [Cad_NFE.ListBoxLista_Change](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 1642 | LinhaFisica, ListIndex |
| [Cad_NFE.ListBoxLista2_Change](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 1768 | LinhaFisica, ListIndex |
| [Cad_NFE.N_NfeFilial_Exit](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 2122 | CountIf |
| [Cad_NFE.CommandButtonCelCopiar_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 2544 | ActiveCell |
| [Cad_NFE.AtualizaPainelCelula](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE.frm) | 2548 | ActiveCell |
| [Cad_NFE_Saida.CommandButton_ExcluirFilial_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 44 | Find, ReferenciaFisica |
| [Cad_NFE_Saida.CommandButton_GravarFilial_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 93 | CountIf, Find, ReferenciaFisica, ActiveCell, Lacunas, LinhaFisica |
| [Cad_NFE_Saida.CommandButton_Excluir_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 478 | Find, ReferenciaFisica |
| [Cad_NFE_Saida.CommandButton_Gravar_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 565 | CountIf, Find, ReferenciaFisica, ActiveCell, Lacunas, LinhaFisica |
| [Cad_NFE_Saida.TextBox_N_Saida_Exit](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 1761 | CountIf |
| [Cad_NFE_Saida.TextBox_N_SaidaFilial_Exit](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 2121 | CountIf |
| [Cad_NFE_Saida.TextBoxFiltro_Change](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 2447 | ListIndex |
| [Cad_NFE_Saida.PreencheCampos](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 2453 | Lacunas |
| [Cad_NFE_Saida.PreencheCabecalho](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 2469 | Lacunas |
| [Cad_NFE_Saida.PreencheLista](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 2485 | ListIndex, Lacunas |
| [Cad_NFE_Saida.ListBoxLista_Change](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 2556 | ListIndex |
| [Cad_NFE_Saida.TextBoxFiltroFilial_Change](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 2617 | ListIndex |
| [Cad_NFE_Saida.PreencheCampos2](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 2623 | Lacunas |
| [Cad_NFE_Saida.PreencheCabecalho2](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 2639 | Lacunas |
| [Cad_NFE_Saida.PreencheLista2](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 2655 | ListIndex, Lacunas |
| [Cad_NFE_Saida.ListBoxListaFilial_Change](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/Cad_NFE_Saida.frm) | 2698 | ListIndex |
| [UserFormFiltro.CommandButtonFiltroCompra_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/UserFormFiltro.frm) | 22 | LinhaFisica, ReferenciaFisica |
| [UserFormFiltro.Command_Entrada_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/UserFormFiltro.frm) | 48 | LinhaFisica, ReferenciaFisica |
| [UserFormFiltro.CommandButtonFiltroAguardando_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/UserFormFiltro.frm) | 61 | LinhaFisica, ReferenciaFisica |
| [UserFormFiltro.CommandButtonRepasse_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/UserFormFiltro.frm) | 74 | LinhaFisica, ReferenciaFisica |
| [UserFormFiltro.CommandButtonVenda_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/UserFormFiltro.frm) | 87 | LinhaFisica, ReferenciaFisica |
| [UserFormFiltro.CommandButtonTransferência_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/UserFormFiltro.frm) | 100 | LinhaFisica, ReferenciaFisica |
| [UserFormFiltro.CommandButtonDevolução_Click](./_auditoria/fase5_20261101_202114Z/desenvolvimento/Form/UserFormFiltro.frm) | 113 | LinhaFisica, ReferenciaFisica |
| [MóduloPesquisa.ConfigurarListBoxPesquisa](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/MóduloPesquisa.bas) | 13 | Contrato de leitor/escritor confirmado por inspeção complementar |
| [MóduloPesquisa.PesquisarEntradasRapido](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/MóduloPesquisa.bas) | 36 | Contrato de leitor/escritor confirmado por inspeção complementar |
| [MóduloPesquisa.TransporParaListBox](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/MóduloPesquisa.bas) | 273 | Contrato de leitor/escritor confirmado por inspeção complementar |
| [modPreencherFormulario.PreencherDadosEntradas](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/modPreencherFormulario.bas) | 4 | Contrato de leitor/escritor confirmado por inspeção complementar |
| [modPreencherFormulario.PreencherDadosEntradasFilial](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/modPreencherFormulario.bas) | 53 | Contrato de leitor/escritor confirmado por inspeção complementar |
| [modIntegridadeOperacao.PrimeiraVagaGravacao](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/modIntegridadeOperacao.bas) | 4 | Contrato de leitor/escritor confirmado por inspeção complementar |
| [mdlCopiar.UF_NavegarCelula](./_auditoria/fase5_20261101_202114Z/desenvolvimento/MDL/mdlCopiar.bas) | 28 | Contrato de leitor/escritor confirmado por inspeção complementar |

## Cobertura das 17 planilhas

Contagem de 43.232 células com fórmula obtida por inspeção OOXML local do original, somente leitura, sem execução de macros nem extração de valores de estoque. Fórmulas compartilhadas/caches e sua correção funcional não foram reavaliadas integralmente. Metadados em [inventário de fórmulas](./_auditoria/fase6_20261101_224949Z/evidencias/formulas-17-abas-metadados.json). As sentinelas dos testes são sintéticas e não substituem todas as fórmulas operacionais.

| Planilha | Células com fórmula | Impacto / verificação requerida |
|---|---:|---|
| Menu Principal | 208 | Totais/painel devem contar somente projeções elegíveis; revisar SUMIF/SUM e faixas divergentes já existentes sem alterar regras nesta fase. |
| Entradas | 2344 | Origem matriz; dados A:L, auxiliares por linha, baixa e seleção. |
| Transferidos | 3 | Projeção/consulta de transferências; validar leitores e vínculos, sem presumir armazenamento novo. |
| Entradas Anual | 17971 | Histórico; P usa COUNTIF por placa em Todas Saidas e confunde reentrada com venda anterior. N/O dependem também de condição/formatação. |
| Demonstração | 154 | Referências físicas e indiretas; sentinelas demonstram #REF! e troca silenciosa. |
| Saidas | 100 | Saída matriz e fórmulas de margem/condição; vínculo com episódio de origem. |
| Todas Saidas | 16673 | Histórico de saída e fonte do COUNTIF anual por placa; precisa relacionar MovementID/EpisodeID. |
| Entradas Filial | 1093 | Origem filial/destino transferência; M/N/O e L chave; Delete parcial não cobre todos os auxiliares. |
| Entradas Anual Filial | 2974 | Histórico filial; P também usa placa; baixa/edição não podem procurar apenas NF-e. |
| Saidas Filial | 64 | Destino venda/devolução; proteção e fórmulas participam da recuperação. |
| Todas Saidas Filial | 1501 | Histórico filial; vínculo explícito à movimentação e entrada. |
| Carros PF | 130 | Contexto de outros veículos, fórmulas inclusive U; não reservar U:X como colunas livres do produto. |
| Clientes | 0 | Cadastro auxiliar; nenhum identificador novo foi inferido. Revisar chamadores de seleção no corte. |
| Veículos | 0 | Cadastro auxiliar de veículo; não é prova de identidade física única. |
| Doc. para fora | 0 | Documentação auxiliar; verificar referências dinâmicas e vinculação futura. |
| NCM e CEST | 0 | Referência fiscal e lista; lógica fiscal preservada. |
| Listas | 17 | Listas/validação; manter nomes, valores e referências de formulário. |

## Ordem segura de migração

1. Aprovar identidade, custódia, armazenamento/diário e tratamento de legados ambíguos.
2. Criar em modelo isolado o repositório/resolvedor por ID, envelope versionado, validações e bloqueio de pendências.
3. Adaptar pesquisa rápida, filtros legados, listas, pré-preenchimento e navegação; testar lacunas, baixados e reordenação.
4. Adaptar fórmulas, totais e históricos: consulta por identidade, sem COUNTIF por placa para inferir encerramento de episódio.
5. Adaptar cadastro, edição e exclusão por ID/versão; conferir todos os históricos antes de modificar.
6. Adaptar Venda/Transferência/Devolução no mesmo corte e somente então aplicar a baixa aprovada. Atualizar snapshots, backups e conciliação.
7. Executar eventos reais dos UserForms migrados com falhas em cada etapa, duas pastas e referências/formatações; verificar os cinco FAIL obrigatórios sem apagar a referência anterior.
8. Aprovar migração e corte reversível. Não habilitar gravadores antigos com o novo esquema.

Chamadas dinâmicas, Ribbon, OnAction, integrações e fórmulas/UDF precisam de revisão final no corte; nenhuma integração aparentemente legada foi removida. O mapa nominal é o conjunto confirmado nesta base, não uma garantia de inexistência de dependências externas não disponíveis.
