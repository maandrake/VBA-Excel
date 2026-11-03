# Revisão dirigida dos57 procedimentos — Entrega2

Análise estática do texto efetivamente embutido, com fonte/linha atual. Não equivale à execução dos57 procedimentos. [Inventário integral com corpos](../evidencias/revisao_57_leitores.json), [regressão funcional](../../../MATRIZ_REGRESSAO_ENTREGA2.md).

As gravações habilitadas usam ACE e revalidam IDs/versão. Rotinas legadas de apresentação ainda existem; sua presença não autoriza edição por posição. Os leitores ativos usam C1/C2; as rotinas de estoque por linha não devem ser usadas como nova autoridade. Helpers de clipboard/navegação e NCM pertencem a contextos auxiliares; preservados, sem migração indiscriminada.

| Componente/procedimento | Linha atual | Situação da revisão | Dependência e risco remanescente |
|---|---|---|---|
| clsOperacaoPlanilha.RegistrarDestino | 44 | Preservado; regressão incompleta | Snapshots/alocação legados não são mecanismo transacional dos gravadores ACE; FAIL históricos preservados. |
| clsOperacaoPlanilha.RestaurarCelula | 118 | Preservado; regressão incompleta | Snapshots/alocação legados não são mecanismo transacional dos gravadores ACE; FAIL históricos preservados. |
| modIntegridadeOperacao.ProximaLinhaGravacao | 19 | Preservado; regressão incompleta | Snapshots/alocação legados não são mecanismo transacional dos gravadores ACE; FAIL históricos preservados. |
| mdlCopiar.UF_CopiarCelula | 39 | Preservado; regressão incompleta | Contexto visual/posição permanece para apresentação/clipboard; não autoriza escrita ACE. |
| mdlCopiar.UF_AtualizaLabels | 62 | Preservado; regressão incompleta | Contexto visual/posição permanece para apresentação/clipboard; não autoriza escrita ACE. |
| mdlCopiar.MoveActiveCell | 111 | Preservado; regressão incompleta | Contexto visual/posição permanece para apresentação/clipboard; não autoriza escrita ACE. |
| MóduloPesquisa.AbrirRegistroDaLista | 29 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| MóduloPesquisa.UltimaLinhaUsada | 36 | Preservado; regressão incompleta | Contexto visual/posição permanece para apresentação/clipboard; não autoriza escrita ACE. |
| MóduloPesquisa.IrParaCelulaNFe | 79 | Bloqueio explícito | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NCM.ListBoxLista_Change | 69 | Preservado; regressão incompleta | Cadastro auxiliar preservado; não confundir seleção NCM com identidade de veículo. |
| Cad_NCM.TextBoxFiltro_Change | 157 | Preservado; regressão incompleta | Cadastro auxiliar preservado; não confundir seleção NCM com identidade de veículo. |
| Cad_NCM.PreencheCampos | 169 | Preservado; regressão incompleta | Rotina de apresentação legada preservada; conferir chamadas e elegibilidade/token antes de habilitar atalhos. |
| Cad_NCM.PreencheCabecalho | 185 | Preservado; regressão incompleta | Rotina de apresentação legada preservada; conferir chamadas e elegibilidade/token antes de habilitar atalhos. |
| Cad_NCM.PreencheLista | 201 | Preservado; regressão incompleta | Rotina de apresentação legada preservada; conferir chamadas e elegibilidade/token antes de habilitar atalhos. |
| Cad_NFE.CommandButton_Excluir_Click | 84 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE.CommandButton_ExcluirFilial_Click | 93 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE.CommandButton_Gravar_Click | 102 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE.CommandButton_GravarFilial_Click | 111 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE.CommandButtonAlterar_Click | 246 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE.CommandButtonCopiarAutoEntradaFilial_Click | 263 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE.CommandButtonChaveR_Click | 347 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE.N_Nfe_Exit | 533 | Preservado; regressão incompleta | Aviso legado CountIf sobre projeção física; não substitui validação ACE. Regressão de foco/duplicidade ainda necessária. |
| Cad_NFE.ListBoxLista_Change | 677 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE.ListBoxLista2_Change | 691 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE.N_NfeFilial_Exit | 933 | Preservado; regressão incompleta | Aviso legado CountIf sobre projeção física; não substitui validação ACE. Regressão de foco/duplicidade ainda necessária. |
| Cad_NFE.CommandButtonCelCopiar_Click | 1355 | Preservado; regressão incompleta | Contexto visual/posição permanece para apresentação/clipboard; não autoriza escrita ACE. |
| Cad_NFE.AtualizaPainelCelula | 1359 | Preservado; regressão incompleta | Contexto visual/posição permanece para apresentação/clipboard; não autoriza escrita ACE. |
| Cad_NFE_Saida.CommandButton_ExcluirFilial_Click | 33 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE_Saida.CommandButton_GravarFilial_Click | 37 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE_Saida.CommandButton_Excluir_Click | 144 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE_Saida.CommandButton_Gravar_Click | 175 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE_Saida.TextBox_N_Saida_Exit | 935 | Preservado; regressão incompleta | Aviso legado CountIf sobre projeção física; não substitui validação ACE. Regressão de foco/duplicidade ainda necessária. |
| Cad_NFE_Saida.TextBox_N_SaidaFilial_Exit | 1295 | Preservado; regressão incompleta | Aviso legado CountIf sobre projeção física; não substitui validação ACE. Regressão de foco/duplicidade ainda necessária. |
| Cad_NFE_Saida.TextBoxFiltro_Change | 1621 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE_Saida.PreencheCampos | 1625 | Preservado; regressão incompleta | Rotina de apresentação legada preservada; conferir chamadas e elegibilidade/token antes de habilitar atalhos. |
| Cad_NFE_Saida.PreencheCabecalho | 1641 | Preservado; regressão incompleta | Rotina de apresentação legada preservada; conferir chamadas e elegibilidade/token antes de habilitar atalhos. |
| Cad_NFE_Saida.PreencheLista | 1657 | Preservado; regressão incompleta | Rotina de apresentação legada preservada; conferir chamadas e elegibilidade/token antes de habilitar atalhos. |
| Cad_NFE_Saida.ListBoxLista_Change | 1728 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE_Saida.TextBoxFiltroFilial_Change | 1734 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| Cad_NFE_Saida.PreencheCampos2 | 1738 | Preservado; regressão incompleta | Rotina de apresentação legada preservada; conferir chamadas e elegibilidade/token antes de habilitar atalhos. |
| Cad_NFE_Saida.PreencheCabecalho2 | 1754 | Preservado; regressão incompleta | Rotina de apresentação legada preservada; conferir chamadas e elegibilidade/token antes de habilitar atalhos. |
| Cad_NFE_Saida.PreencheLista2 | 1770 | Preservado; regressão incompleta | Rotina de apresentação legada preservada; conferir chamadas e elegibilidade/token antes de habilitar atalhos. |
| Cad_NFE_Saida.ListBoxListaFilial_Change | 1813 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| UserFormFiltro.CommandButtonFiltroCompra_Click | 23 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| UserFormFiltro.Command_Entrada_Click | 46 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| UserFormFiltro.CommandButtonFiltroAguardando_Click | 56 | Bloqueio explícito | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| UserFormFiltro.CommandButtonRepasse_Click | 66 | Bloqueio explícito | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| UserFormFiltro.CommandButtonVenda_Click | 76 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| UserFormFiltro.CommandButtonTransferência_Click | 86 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| UserFormFiltro.CommandButtonDevolução_Click | 96 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| MóduloPesquisa.ConfigurarListBoxPesquisa | 13 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| MóduloPesquisa.PesquisarEntradasRapido | 20 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| MóduloPesquisa.TransporParaListBox | 49 | Preservado; regressão incompleta | Contexto visual/posição permanece para apresentação/clipboard; não autoriza escrita ACE. |
| modPreencherFormulario.PreencherDadosEntradas | 4 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| modPreencherFormulario.PreencherDadosEntradasFilial | 11 | Delega à camada integrada | Identidade/versão revalidadas pelo delegado; alcance funcional delimitado na matriz. |
| modIntegridadeOperacao.PrimeiraVagaGravacao | 4 | Preservado; regressão incompleta | Snapshots/alocação legados não são mecanismo transacional dos gravadores ACE; FAIL históricos preservados. |
| mdlCopiar.UF_NavegarCelula | 28 | Preservado; regressão incompleta | Contexto visual/posição permanece para apresentação/clipboard; não autoriza escrita ACE. |
