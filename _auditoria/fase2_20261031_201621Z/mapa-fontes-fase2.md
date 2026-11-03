# Mapa dos fontes de desenvolvimento — fase 2

34 fontes, 665 procedimentos exportados. Gerado por análise textual; não comprova compilação nem resolve chamadas dinâmicas. Acrescente o evento existente de Planilha15 ao incorporar os documentos: projeção de 666 procedimentos, sem módulos de teste. O XLSM base permanece com 50 componentes/648 procedimentos.

## clsOperacaoPlanilha

Fonte: [clsOperacaoPlanilha.cls](<./desenvolvimento/MDL/clsOperacaoPlanilha.cls>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| Class_Initialize | Private / Sub | 25 | — |
| Iniciar | Public / Sub | 29 | — |
| RegistrarDestino | Public / Sub | 43 | modIntegridadeOperacao.VerificarDestinoGravacao |
| MarcarBaixaIniciada | Public / Sub | 72 | — |
| MarcarSalvamentoIniciado | Public / Sub | 77 | — |
| Confirmar | Public / Sub | 81 | clsOperacaoPlanilha.Finalizar |
| Recuperar | Public / Function | 86 | clsOperacaoPlanilha.RestaurarCelula |
| RestaurarCelula | Private / Function | 112 | — |
| Finalizar | Public / Sub | 136 | — |
| Class_Terminate | Private / Sub | 156 | clsOperacaoPlanilha.Finalizar |

## mdlChamarFiltro

Fonte: [mdlChamarFiltro.bas](<./desenvolvimento/MDL/mdlChamarFiltro.bas>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ChamarFormFiltro | Public / Sub | 3 | — |

## mdlChamarForm

Fonte: [mdlChamarForm.bas](<./desenvolvimento/MDL/mdlChamarForm.bas>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ChamarForm | Public / Sub | 3 | — |

## mdlChamarFormDocFora

Fonte: [mdlChamarFormDocFora.bas](<./desenvolvimento/MDL/mdlChamarFormDocFora.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ChamarFormDocFora | Public / Sub | 3 | — |

## mdlChamarFormNCM

Fonte: [mdlChamarFormNCM.bas](<./desenvolvimento/MDL/mdlChamarFormNCM.bas>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ChamarFormNCM | Public / Sub | 3 | — |

## mdlChamarFormSaida

Fonte: [mdlChamarFormSaida.bas](<./desenvolvimento/MDL/mdlChamarFormSaida.bas>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ChamarFormSaida | Public / Sub | 3 | — |

## mdlChamarFormVeículos

Fonte: [mdlChamarFormVeículos.bas](<./desenvolvimento/MDL/mdlChamarFormVeículos.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ChamarFormVeículos | Public / Sub | 3 | — |

## mdlChamarTabCan

Fonte: [mdlChamarTabCan.bas](<./desenvolvimento/MDL/mdlChamarTabCan.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ChamarTabCan | Public / Sub | 3 | — |

## mdlChamarTabEnt

Fonte: [mdlChamarTabEnt.bas](<./desenvolvimento/MDL/mdlChamarTabEnt.bas>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ChamarTabEnt | Public / Sub | 3 | — |

## mdlChamarTabSai

Fonte: [mdlChamarTabSai.bas](<./desenvolvimento/MDL/mdlChamarTabSai.bas>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ChamarTabSai | Public / Sub | 3 | — |

## mdlCondCel

Fonte: [mdlCondCel.bas](<./desenvolvimento/MDL/mdlCondCel.bas>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ActiveCondition | Public / Function | 4 | mdlCondCel.GetStrippedValue |
| ColorIndexOfCF | Public / Function | 163 | mdlCondCel.ActiveCondition |
| ColorOfCF | Public / Function | 184 | mdlCondCel.ActiveCondition |
| GetStrippedValue | Public / Function | 204 | — |
| CountOfCF | Public / Function | 218 | mdlCondCel.ActiveCondition |
| SumByCFColorIndex | Public / Function | 234 | mdlCondCel.ColorIndexOfCF |

## mdlCopiar

Fonte: [mdlCopiar.bas](<./desenvolvimento/MDL/mdlCopiar.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| UF_NavegarCelula | Public / Sub | 28 | mdlCopiar.MoveActiveCell, mdlCopiar.UF_AtualizaLabels |
| UF_AtualizarPainel | Public / Sub | 34 | mdlCopiar.UF_AtualizaLabels |
| UF_CopiarCelula | Public / Sub | 39 | mdlCopiar.Clipboard_SetText, mdlCopiar.UF_AtualizaLabels, mdlCopiar.UF_SetStatus |
| UF_AtualizaLabels | Private / Sub | 62 | mdlCopiar.UF_SetStatus, mdlCopiar.ControleExiste |
| UF_SetStatus | Private / Sub | 94 | mdlCopiar.ControleExiste |
| ControleExiste | Private / Function | 101 | — |
| MoveActiveCell | Private / Sub | 111 | — |
| Clipboard_SetText | Private / Sub | 133 | — |

## mdlCriaBackup

Fonte: [mdlCriaBackup.bas](<./desenvolvimento/MDL/mdlCriaBackup.bas>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| SaveWorkbookBackup | Public / Sub | 4 | mdlCriaBackup.CriarBackupVersionado |
| SaveWorkbookBackupToFloppy | Public / Sub | 18 | mdlCriaBackup.CriarBackupVersionado |
| CriarBackupVersionado | Public / Function | 32 | mdlCriaBackup.VerificarArquivoBackup |
| VerificarArquivoBackup | Private / Sub | 71 | — |

## mdlCursor

Fonte: [mdlCursor.bas](<./desenvolvimento/MDL/mdlCursor.bas>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| MouseCursor | Public / Function | 27 | — |

## mdlMacroConexão

Fonte: [mdlMacroConexão.bas](<./desenvolvimento/MDL/mdlMacroConexão.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| MacroConexão | Public / Sub | 3 | — |

## mdlMacroConexãoClientes

Fonte: [mdlMacroConexãoClientes.bas](<./desenvolvimento/MDL/mdlMacroConexãoClientes.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ConexãoClientes | Public / Sub | 3 | — |

## mdlMacroDocFora

Fonte: [mdlMacroDocFora.bas](<./desenvolvimento/MDL/mdlMacroDocFora.bas>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| MacroDocFora | Public / Sub | 3 | — |

## mdlOcultarTítulo

Fonte: [mdlOcultarTítulo.bas](<./desenvolvimento/MDL/mdlOcultarTítulo.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| HideTitleBar | Public / Sub | 24 | — |

## mdlTelaCheia

Fonte: [mdlTelaCheia.bas](<./desenvolvimento/MDL/mdlTelaCheia.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| lsLigarTelaCheia | Public / Sub | 3 | — |
| lsDesligarTelaCheia | Public / Sub | 37 | — |

## modIntegridadeOperacao

Fonte: [modIntegridadeOperacao.bas](<./desenvolvimento/MDL/modIntegridadeOperacao.bas>). Codificação detectada: utf-8; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| PrimeiraVagaGravacao | Public / Function | 4 | — |
| ProximaLinhaGravacao | Public / Function | 19 | — |
| VerificarDestinoGravacao | Public / Sub | 27 | — |
| ValidarConversoesGravacao | Public / Sub | 37 | — |
| ValidarChaveTexto | Public / Sub | 46 | — |

## modPreencherFormulario

Fonte: [modPreencherFormulario.bas](<./desenvolvimento/MDL/modPreencherFormulario.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| PreencherDadosEntradas | Public / Sub | 4 | — |
| PreencherDadosEntradasFilial | Public / Sub | 53 | — |

## modSeleniumCommon

Fonte: [modSeleniumCommon.bas](<./desenvolvimento/MDL/modSeleniumCommon.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ConfigureChromedriverExecutable | Public / Sub | 18 | modSeleniumCommon.EnsureDriverPath |
| EnsureDriverPath | Public / Sub | 27 | — |
| WaitForReadyStateComplete | Public / Function | 42 | modSeleniumCommon.SegundosDecorridos, modSeleniumCommon.SleepMs |
| SleepMs | Public / Sub | 68 | modSeleniumCommon.SegundosDecorridos |
| TruncateForLog | Public / Function | 89 | — |
| WriteLog | Public / Sub | 105 | modSeleniumCommon.TruncateForLog |
| SegundosDecorridos | Public / Function | 144 | — |

## modSeleniumConsultaNFE

Fonte: [modSeleniumConsultaNFE.bas](<./desenvolvimento/MDL/modSeleniumConsultaNFE.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ConsultarDetranRetorno | Public / Function | 53 | modSeleniumConsultaNFE.SanitizePlacaNFE, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.BuildConsultaUrlNFE, modSeleniumConsultaNFE.GetDetranDriverNFE, modSeleniumConsultaNFE.SessaoDetranProvavelmenteAutenticadaNFE, modSeleniumConsultaNFE.WaitForUrlNFE, modSeleniumConsultaNFE.UrlAtualCorrespondeNFE, modSeleniumConsultaNFE.EstaPaginaDeLoginNFE, modSeleniumConsultaNFE.WaitDocumentReadyNFE, modSeleniumConsultaNFE.DetectarPlacaNaoCadastradaNFE, modSeleniumConsultaNFE.TentarConsultaBINPorPlaca, modSeleniumConsultaNFE.NzTexto, modSeleniumConsultaNFE.BuildCollectionFromDossieNFE, modSeleniumConsultaNFE.RasparDossieVeiculoNFE, modSeleniumConsultaNFE.EncerrarSessaoDriverNFE |
| GetDetranDriverNFE | Private / Function | 155 | modSeleniumConsultaNFE.TryReuseDetranDriverNFE, modSeleniumConsultaNFE.EncerrarSessaoDriverNFE, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.ConfigureChromedriverExecutableNFE, modSeleniumConsultaNFE.ConfigureAndStartChromeNFE |
| EncerrarSessaoDriverNFE | Private / Sub | 181 | — |
| EncerrarSessaoDetranNFE | Public / Sub | 190 | modSeleniumConsultaNFE.EncerrarSessaoDriverNFE |
| TryReuseDetranDriverNFE | Private / Function | 194 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.EncerrarSessaoDriverNFE |
| SanitizePlacaNFE | Private / Function | 213 | — |
| BuildConsultaUrlNFE | Private / Function | 233 | modSeleniumConsultaNFE.SanitizePlacaNFE |
| ConfigureChromedriverExecutableNFE | Private / Sub | 242 | modSeleniumConsultaNFE.EnsureDriverPathNFE |
| EnsureDriverPathNFE | Private / Sub | 255 | — |
| ConfigureAndStartChromeNFE | Private / Sub | 271 | — |
| SessaoDetranProvavelmenteAutenticadaNFE | Private / Function | 294 | — |
| EstaPaginaDeLoginNFE | Private / Function | 320 | — |
| WaitForUrlNFE | Private / Sub | 337 | modSeleniumConsultaNFE.NFE_ElapsedSeconds, modSeleniumConsultaNFE.SleepMsNFE |
| UrlAtualCorrespondeNFE | Private / Function | 357 | — |
| ObterJanelaChromeHandleNFE | Private / Function | 374 | modSeleniumConsultaNFE.EnumWindowsProcNFE |
| EnumWindowsProcNFE | Private / Function | 390 | — |
| WaitDocumentReadyNFE | Private / Function | 420 | modSeleniumConsultaNFE.WaitForReadyStateCompleteNFE |
| WaitForReadyStateCompleteNFE | Private / Function | 424 | modSeleniumConsultaNFE.NFE_ElapsedSeconds, modSeleniumConsultaNFE.SleepMsNFE |
| MostrarDriverWindowNFE | Private / Sub | 450 | modSeleniumConsultaNFE.NzTexto, modSeleniumConsultaNFE.ObterJanelaChromeHandleNFE |
| SleepMsNFE | Private / Sub | 482 | modSeleniumCommon.SleepMs |
| RasparDossieVeiculoNFE | Private / Function | 487 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.LimpaNBSP, modSeleniumConsultaNFE.GetValorPorRotuloNFE |
| GetValorPorRotuloNFE | Private / Function | 511 | — |
| BuildCollectionFromDossieNFE | Private / Function | 537 | modSeleniumConsultaNFE.NzTexto |
| GetValueNFE | Public / Function | 562 | modSeleniumConsultaNFE.NzTexto |
| TentarConsultaBINPorPlaca | Private / Function | 578 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.NavegarParaBINNFE, modSeleniumConsultaNFE.PreencherESubmeterPlacaBINNFE, modSeleniumConsultaNFE.WaitDocumentReadyNFE, modSeleniumConsultaNFE.SleepMsNFE, modSeleniumConsultaNFE.RasparDossieBINNFE |
| NavegarParaBINNFE | Private / Function | 597 | modSeleniumConsultaNFE.SwitchToDefaultContentNFE, modSeleniumConsultaNFE.WaitDocumentReadyNFE, modSeleniumConsultaNFE.SwitchToFrameByNameNFE, modSeleniumConsultaNFE.ClickByIdNFE, modSeleniumConsultaNFE.SleepMsNFE, modSeleniumConsultaNFE.LogNFE |
| PreencherESubmeterPlacaBINNFE | Private / Function | 639 | modSeleniumConsultaNFE.SwitchToDefaultContentNFE, modSeleniumConsultaNFE.TentarEncontrarCampoBINNFE, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.SwitchToFrameByNameNFE, modSeleniumConsultaNFE.ClickByNameNFE, modSeleniumConsultaNFE.ClickByIdNFE |
| TentarEncontrarCampoBINNFE | Private / Function | 677 | modSeleniumConsultaNFE.SwitchToDefaultContentNFE, modSeleniumConsultaNFE.SwitchToFrameByNameNFE, modSeleniumConsultaNFE.FindElementByNameNFE |
| RasparDossieBINNFE | Private / Function | 706 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.SwitchToDefaultContentNFE, modSeleniumConsultaNFE.SwitchToFrameByNameNFE, modSeleniumConsultaNFE.LimparPlacaBINNFE, modSeleniumConsultaNFE.LimpaNBSP, modSeleniumConsultaNFE.GetValorBINPorRotuloNFE |
| GetValorBINPorRotuloNFE | Private / Function | 757 | — |
| LimparPlacaBINNFE | Private / Function | 783 | — |
| SwitchToDefaultContentNFE | Private / Sub | 803 | — |
| SwitchToFrameByNameNFE | Private / Function | 809 | — |
| ClickByIdNFE | Private / Function | 828 | — |
| ClickByNameNFE | Private / Function | 844 | — |
| FindElementByNameNFE | Private / Function | 860 | — |
| DetectarPlacaNaoCadastradaNFE | Private / Function | 868 | — |
| LimpaNBSP | Private / Function | 889 | — |
| NFE_ElapsedSeconds | Private / Function | 893 | — |
| ConsultarDadosNFE | Public / Function | 903 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.ConsultarDetranRetorno, modSeleniumConsultaNFE.FormatDadosNFE |
| FormatDadosNFE | Public / Function | 922 | modSeleniumConsultaNFE.GetValueNFE, modSeleniumConsultaNFE.ExtrairAnoModelo, modSeleniumConsultaNFE.BuildResumoTexto, modSeleniumConsultaNFE.LimparMarcaModelo, modSeleniumConsultaNFE.LimparChassi, modSeleniumConsultaNFE.LimparCor |
| CopiarDadosDetranNFE | Public / Sub | 952 | modSeleniumConsultaNFE.ConsultarDadosNFE, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.CopiarParaClipboardNFE |
| BuildResumoTexto | Private / Function | 974 | modSeleniumConsultaNFE.NzTexto |
| LimparMarcaModelo | Private / Function | 1004 | — |
| LimparChassi | Private / Function | 1025 | — |
| ExtrairAnoModelo | Private / Function | 1046 | — |
| LimparCor | Private / Function | 1070 | — |
| NzTexto | Private / Function | 1085 | — |
| CommandButtonConsultaDetran | Public / Sub | 1093 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.CopiarDadosDetranNFE |
| CommandButton_Click | Public / Sub | 1107 | modSeleniumConsultaNFE.CommandButtonConsultaDetran |
| CommandButtonReexibirChrome | Public / Sub | 1111 | modSeleniumConsultaNFE.EncerrarSessaoDriverNFE, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.MostrarDriverWindowNFE |
| CopiarParaClipboardNFE | Private / Sub | 1144 | modSeleniumConsultaNFE.NzTexto, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.CopiarViaMsFormsDataObject, modSeleniumConsultaNFE.CopiarViaClipboardWinAPI |
| CopiarViaMsFormsDataObject | Private / Function | 1178 | — |
| CopiarViaClipboardWinAPI | Private / Function | 1221 | modSeleniumConsultaNFE.SleepMsNFE |
| LogNFE | Private / Sub | 1297 | modSeleniumCommon.WriteLog |

## modSeleniumRenavamConsulta

Fonte: [modSeleniumRenavamConsulta.bas](<./desenvolvimento/MDL/modSeleniumRenavamConsulta.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| AbrirConsultaRenavam | Public / Sub | 13 | modSeleniumRenavamConsulta.GetDetranDriverRenavam, modSeleniumRenavamConsulta.SessaoDetranProvavelmenteAutenticadaRenavam, modSeleniumCommon.WriteLog, modSeleniumRenavamConsulta.WaitForUrlRenavam, modSeleniumRenavamConsulta.UrlAtualCorrespondeRenavam, modSeleniumRenavamConsulta.EstaPaginaDeLoginRenavam, modSeleniumRenavamConsulta.NavegarMenuRenavam, modSeleniumRenavamConsulta.SelecionarUfRenavam, modSeleniumRenavamConsulta.PreencherPlacaRenavam, modSeleniumRenavamConsulta.EncerrarSessaoRenavam |
| CommandButtonDetranNetBIN_Handler | Public / Sub | 85 | modSeleniumRenavamConsulta.AbrirConsultaRenavam |
| EncerrarSessaoRenavam | Public / Sub | 98 | — |
| GetDetranDriverRenavam | Private / Function | 105 | modSeleniumRenavamConsulta.TryReuseDetranDriverRenavam, modSeleniumRenavamConsulta.EncerrarSessaoRenavam, modSeleniumCommon.WriteLog, modSeleniumCommon.ConfigureChromedriverExecutable, modSeleniumRenavamConsulta.ConfigureAndStartChrome |
| TryReuseDetranDriverRenavam | Private / Function | 131 | modSeleniumCommon.WriteLog, modSeleniumRenavamConsulta.EncerrarSessaoRenavam |
| SessaoDetranProvavelmenteAutenticadaRenavam | Private / Function | 150 | — |
| WaitForUrlRenavam | Private / Sub | 176 | modSeleniumCommon.SegundosDecorridos, modSeleniumCommon.SleepMs |
| UrlAtualCorrespondeRenavam | Private / Function | 196 | — |
| EstaPaginaDeLoginRenavam | Private / Function | 213 | — |
| NavegarMenuRenavam | Private / Function | 230 | modSeleniumRenavamConsulta.SwitchToDefaultContentSafe, modSeleniumRenavamConsulta.WaitDocumentReadyRenavam, modSeleniumRenavamConsulta.SwitchToFrameByNameSafe, modSeleniumRenavamConsulta.ClickByIdSafe, modSeleniumCommon.SleepMs, modSeleniumRenavamConsulta.AbrirConsultaOutrasUfs, modSeleniumCommon.WriteLog |
| AbrirConsultaOutrasUfs | Private / Function | 259 | modSeleniumRenavamConsulta.SwitchToFrameByNameSafe, modSeleniumCommon.WriteLog, modSeleniumRenavamConsulta.ClickByIdSafe, modSeleniumCommon.SleepMs, modSeleniumRenavamConsulta.SwitchToDefaultContentSafe, modSeleniumRenavamConsulta.EnsureConsultaFrameRenavam, modSeleniumCommon.SegundosDecorridos |
| WaitDocumentReadyRenavam | Private / Function | 326 | modSeleniumCommon.WaitForReadyStateComplete |
| ClickByIdSafe | Private / Function | 330 | — |
| SelecionarUfRenavam | Private / Function | 354 | modSeleniumRenavamConsulta.EnsureConsultaFrameRenavam |
| PreencherPlacaRenavam | Private / Function | 381 | modSeleniumRenavamConsulta.EnsureConsultaFrameRenavam, modSeleniumRenavamConsulta.FindElementByNameSafe |
| EnsureConsultaFrameRenavam | Private / Function | 405 | modSeleniumRenavamConsulta.SwitchToDefaultContentSafe, modSeleniumRenavamConsulta.SwitchToFrameByNameSafe, modSeleniumRenavamConsulta.FindElementByNameSafe, modSeleniumCommon.WriteLog, modSeleniumRenavamConsulta.FindElementByIdSafe, modSeleniumRenavamConsulta.FrameItemSafe |
| FrameItemSafe | Private / Function | 491 | — |
| SwitchToDefaultContentSafe | Private / Sub | 503 | — |
| FindElementByNameSafe | Private / Function | 509 | — |
| FindElementByIdSafe | Private / Function | 515 | — |
| FindElementByXPathSafe | Private / Function | 521 | — |
| SwitchToFrameByNameSafe | Private / Function | 527 | modSeleniumRenavamConsulta.FindElementByXPathSafe |
| ConfigureAndStartChrome | Private / Sub | 546 | — |

## modSelTest

Fonte: [modSelTest.bas](<./desenvolvimento/MDL/modSelTest.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| BuildStatusBadge | Private / Function | 21 | — |
| BuildStatusHeader | Private / Function | 29 | modSelTest.BuildStatusBadge |
| SelTest_Run | Public / Sub | 39 | modSelTest.ValidateSeleniumEnvironment, modSelTest.TryCreateDriver, modSelTest.BuildSeleniumComDiagnostic, modSeleniumCommon.ConfigureChromedriverExecutable, modSeleniumCommon.WaitForReadyStateComplete, modSelTest.ParseChromeMajorFromUA, modSelTest.GetChromeDriverMajor, modSelTest.UrlReachable, modSelTest.BuildStatusHeader, modSelTest.AppendSelfTestLog |
| SmokeTestConsultaDetran | Public / Sub | 194 | modSeleniumCommon.WriteLog, modSeleniumConsultaNFE.ConsultarDetranRetorno, modSelTest.BuildSmokeTestFieldSpecs, modSeleniumConsultaNFE.GetValueNFE, modSelTest.BuildStatusHeader, modSelTest.BuildStatusBadge |
| BuildSmokeTestFieldSpecs | Private / Function | 261 | modSelTest.AddFieldSpec |
| AddFieldSpec | Private / Sub | 305 | — |
| ValidateSeleniumEnvironment | Private / Function | 311 | modSelTest.GetChromeVersionString, modSelTest.GetChromeDriverVersionString, modSelTest.ExtractMajorVersion, modSelTest.CheckDotNetRequirement, modSelTest.CheckDotNet35Feature, modSelTest.BuildSeleniumInstallationSummary, modSelTest.TrimTrailingCrLf |
| RunAndCapture | Private / Function | 423 | modSeleniumCommon.SegundosDecorridos, modSeleniumCommon.SleepMs |
| ExtractMajorVersion | Private / Function | 454 | — |
| TrimTrailingCrLf | Private / Function | 472 | — |
| GetChromeVersionString | Private / Function | 479 | modSelTest.GetFileVersionSafe, modSelTest.TryRegRead |
| GetChromeDriverVersionString | Private / Function | 519 | modSelTest.GetFileVersionSafe, modSelTest.RunAndCapture |
| GetFileVersionSafe | Private / Function | 539 | — |
| TryRegRead | Private / Function | 553 | — |
| CheckDotNetRequirement | Private / Function | 563 | modSelTest.GetDotNetReleaseValue, modSelTest.MapDotNetReleaseToVersion |
| CheckDotNet35Feature | Private / Function | 588 | modSelTest.TryRegRead |
| GetDotNetReleaseValue | Private / Function | 645 | modSelTest.TryRegRead |
| MapDotNetReleaseToVersion | Private / Function | 673 | — |
| BuildSeleniumInstallationSummary | Private / Function | 700 | modSelTest.GetSeleniumComServerCommand, modSelTest.NormalizeCommandPath, modSelTest.DetectSeleniumInstall |
| DetectSeleniumInstall | Private / Function | 749 | modSelTest.GetSeleniumComServerCommand, modSelTest.NormalizeCommandPath, modSelTest.FileExists, modSelTest.GetParentFolderSafe, modSelTest.GetSeleniumCandidateFolders, modSelTest.FolderExists, modSelTest.FirstExistingFile, modSelTest.AppendPath |
| GetSeleniumComServerCommand | Private / Function | 814 | modSelTest.TryRegRead |
| NormalizeCommandPath | Private / Function | 838 | — |
| GetParentFolderSafe | Private / Function | 871 | — |
| GetSeleniumCandidateFolders | Private / Function | 881 | modSelTest.AddCandidateFolder, modSelTest.ExpandEnvPath |
| AddCandidateFolder | Private / Sub | 900 | — |
| ExpandEnvPath | Private / Function | 918 | — |
| FirstExistingFile | Private / Function | 936 | modSelTest.AppendPath, modSelTest.FileExists |
| AppendPath | Private / Function | 948 | — |
| FolderExists | Private / Function | 960 | — |
| FileExists | Private / Function | 970 | — |
| AppendSelfTestLog | Private / Sub | 977 | modSeleniumCommon.WriteLog, modSelTest.FolderExists |
| TryCreateDriver | Private / Function | 1011 | — |
| BuildSeleniumComDiagnostic | Private / Function | 1023 | modSelTest.TryRegRead, modSelTest.NormalizeCommandPath, modSelTest.DetectSeleniumInstall |
| UrlReachable | Private / Function | 1096 | modSeleniumCommon.WaitForReadyStateComplete |
| ParseChromeMajorFromUA | Private / Function | 1105 | — |
| GetChromeDriverMajor | Private / Function | 1118 | modSelTest.GetChromeDriverVersionString, modSelTest.ExtractMajorVersion |

## modValidacaoXML

Fonte: [modValidacaoXML.bas](<./desenvolvimento/MDL/modValidacaoXML.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| CarregarXMLNFe | Public / Function | 14 | modValidacaoXML.SelecionarArquivoXML, modValidacaoXML.CriarDocumentoXML, modValidacaoXML.TextoAtributoXML, modValidacaoXML.RemoverPrefixoNFe, modValidacaoXML.TextoNoXML, modValidacaoXML.SomenteDigitos |
| ValidarXMLNFe | Public / Function | 131 | modValidacaoXML.CarregarXMLNFe, modValidacaoXML.IdentificacaoXMLValida, modValidacaoXML.ValidarEstruturaEAutorizacao, modValidacaoXML.TextoNoXML, modValidacaoXML.AdicionarResultado, modValidacaoXML.DescricaoTipoNF, modValidacaoXML.ObterValorControle, modValidacaoXML.RemoverPrefixoNFe, modValidacaoXML.TextoAtributoXML, modValidacaoXML.NumerosNFeIguais, modValidacaoXML.ValidarDataEmissao, modValidacaoXML.LocalizarItemPorCodigo, modValidacaoXML.ListaCodigosProdutos, modValidacaoXML.AdicionarDependenciasNaoValidadas, modValidacaoXML.ValidarPlacaNoProduto, modValidacaoXML.ValidarOperacaoECFOP, modValidacaoXML.ValidarTransferenciaEntreEstabelecimentos, modValidacaoXML.ValidarBeneficioFiscal, modValidacaoXML.ValidarNumeroCampo, modValidacaoXML.ValidarICMSPorCST, modValidacaoXML.ValidarPISeCOFINSPorCST, modValidacaoXML.ValidarTotaisXML, modValidacaoXML.ExibirRelatorioVisual |
| ValidarEstruturaEAutorizacao | Private / Sub | 304 | modValidacaoXML.TextoNoXML, modValidacaoXML.AdicionarResultado, modValidacaoXML.DescricaoAmbiente, modValidacaoXML.IIfTexto, modValidacaoXML.TextoAtributoXML, modValidacaoXML.RemoverPrefixoNFe |
| ValidarOperacaoECFOP | Private / Sub | 377 | modValidacaoXML.NormalizarTexto, modValidacaoXML.AdicionarResultado, modValidacaoXML.AdicionarAviso |
| ValidarBeneficioFiscal | Private / Sub | 519 | modValidacaoXML.TextoNoXML, modValidacaoXML.AdicionarResultado, modValidacaoXML.AdicionarInformativo, modValidacaoXML.IIfTexto |
| ValidarTransferenciaEntreEstabelecimentos | Private / Sub | 638 | modValidacaoXML.NormalizarTexto, modValidacaoXML.SomenteDigitos, modValidacaoXML.TextoNoXML, modValidacaoXML.AdicionarResultado, modValidacaoXML.FormatarCNPJ |
| ValidarPlacaNoProduto | Private / Sub | 687 | modValidacaoXML.ObterValorControle, modValidacaoXML.TextoNoXML, modValidacaoXML.NormalizarAlfanumerico, modValidacaoXML.AdicionarResultado |
| ValidarICMSPorCST | Private / Sub | 717 | modValidacaoXML.TextoNoXML, modValidacaoXML.NormalizarCST, modValidacaoXML.AdicionarResultado, modValidacaoXML.ValidarNumeroCampo, modValidacaoXML.ObterValorControle, modValidacaoXML.TentarConverterNumeroXML, modValidacaoXML.ValidarNumeroCalculado, modValidacaoXML.AdicionarAviso, modValidacaoXML.AdicionarInformativo |
| ValidarPISeCOFINSPorCST | Private / Sub | 853 | modValidacaoXML.TextoNoXML, modValidacaoXML.ValidarPISPorCST, modValidacaoXML.ValidarCOFINSPorCST |
| ValidarPISPorCST | Private / Sub | 875 | modValidacaoXML.NormalizarCST, modValidacaoXML.AdicionarResultado, modValidacaoXML.TextoNoXML, modValidacaoXML.ValidarNumeroCampo, modValidacaoXML.DescricaoCSTSemIncidencia, modValidacaoXML.AdicionarInformativo |
| DescricaoCSTSemIncidencia | Private / Function | 935 | — |
| ValidarCOFINSPorCST | Private / Sub | 958 | modValidacaoXML.NormalizarCST, modValidacaoXML.AdicionarResultado, modValidacaoXML.TextoNoXML, modValidacaoXML.ValidarNumeroCampo, modValidacaoXML.ObterValorControle, modValidacaoXML.TentarConverterNumeroXML, modValidacaoXML.ValidarNumeroCalculado, modValidacaoXML.DescricaoCSTPISCOFINS, modValidacaoXML.AdicionarInformativo |
| ValidarTotaisXML | Private / Sub | 1042 | modValidacaoXML.SomarNosXML, modValidacaoXML.ValidarNumeroCalculado, modValidacaoXML.TextoNoXML |
| AdicionarDependenciasNaoValidadas | Private / Sub | 1112 | modValidacaoXML.AdicionarInformativo |
| ValidarNumeroCampo | Private / Sub | 1139 | modValidacaoXML.TentarConverterNumeroFormulario, modValidacaoXML.TentarConverterNumeroXML, modValidacaoXML.AdicionarResultado |
| ValidarNumeroCalculado | Private / Sub | 1186 | modValidacaoXML.TentarConverterNumeroXML, modValidacaoXML.AdicionarResultado, modValidacaoXML.FormatarNumeroRelatorio |
| SomarNosXML | Private / Function | 1226 | modValidacaoXML.TentarConverterNumeroXML, modValidacaoXML.AdicionarResultado |
| SelecionarArquivoXML | Private / Function | 1261 | — |
| CriarDocumentoXML | Private / Function | 1286 | — |
| TextoNoXML | Private / Function | 1304 | — |
| TextoAtributoXML | Private / Function | 1326 | — |
| LocalizarItemPorCodigo | Private / Function | 1348 | modValidacaoXML.TextoNoXML |
| ListaCodigosProdutos | Private / Function | 1380 | — |
| ObterValorControle | Private / Function | 1409 | — |
| TentarConverterNumeroXML | Private / Function | 1424 | — |
| TentarConverterNumeroFormulario | Private / Function | 1462 | — |
| ValidarDataEmissao | Private / Sub | 1538 | modValidacaoXML.TentarConverterDataFormulario, modValidacaoXML.ConverterDataISO, modValidacaoXML.AdicionarResultado |
| TentarConverterDataFormulario | Private / Function | 1588 | modValidacaoXML.CriarDataSegura |
| ConverterDataISO | Private / Function | 1651 | modValidacaoXML.CriarDataSegura |
| CriarDataSegura | Private / Function | 1678 | — |
| NumerosNFeIguais | Private / Function | 1706 | — |
| NormalizarCST | Private / Function | 1731 | — |
| DescricaoCSTPISCOFINS | Private / Function | 1745 | — |
| DescricaoTipoNF | Private / Function | 1763 | — |
| DescricaoAmbiente | Private / Function | 1780 | — |
| NormalizarTexto | Private / Function | 1800 | — |
| NormalizarAlfanumerico | Private / Function | 1837 | — |
| SomenteDigitos | Private / Function | 1859 | — |
| FormatarCNPJ | Private / Function | 1879 | modValidacaoXML.SomenteDigitos |
| RemoverPrefixoNFe | Private / Function | 1900 | — |
| FormatarNumeroRelatorio | Private / Function | 1912 | — |
| IIfTexto | Private / Function | 1918 | — |
| AdicionarResultado | Private / Sub | 1934 | — |
| AdicionarAviso | Private / Sub | 1969 | — |
| AdicionarInformativo | Private / Sub | 1993 | — |
| ExibirRelatorioVisual | Private / Sub | 2016 | — |
| IdentificacaoXMLValida | Private / Function | 2031 | modValidacaoXML.TextoAtributoXML, modValidacaoXML.RemoverPrefixoNFe, modValidacaoXML.SomenteDigitos, modValidacaoXML.TextoNoXML |

## MóduloOfficeRibbonEditor

Fonte: [MóduloOfficeRibbonEditor.bas](<./desenvolvimento/MDL/MóduloOfficeRibbonEditor.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ChamarFrmEntrada | Public / Sub | 3 | — |
| ChamarFrmSaida | Public / Sub | 8 | — |
| ChamarSalvar | Public / Sub | 13 | — |

## MóduloPesquisa

Fonte: [MóduloPesquisa.bas](<./desenvolvimento/MDL/MóduloPesquisa.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| ConfigurarListBoxPesquisa | Public / Sub | 13 | — |
| PesquisarEntradasRapido | Public / Sub | 36 | MóduloPesquisa.ConfigurarListBoxPesquisa, MóduloPesquisa.UltimaLinhaUsada, MóduloPesquisa.TransporParaListBox |
| AbrirRegistroDaLista | Public / Sub | 219 | — |
| UltimaLinhaUsada | Private / Function | 260 | — |
| TransporParaListBox | Private / Function | 273 | — |
| IrParaCelulaNFe | Public / Sub | 303 | — |

## MóduloRemoveFechar

Fonte: [MóduloRemoveFechar.bas](<./desenvolvimento/MDL/MóduloRemoveFechar.bas>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| RemoveBotao | Public / Sub | 15 | — |

## Cad_NCM

Fonte: [Cad_NCM.frm](<./desenvolvimento/Form/Cad_NCM.frm>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| FramePesquisa_Enter | Private / Sub | 25 | — |
| FramePesquisa_Exit | Private / Sub | 31 | — |
| ComboBoxCampos_Enter | Private / Sub | 47 | — |
| ComboBoxCampos_Exit | Private / Sub | 54 | — |
| TextBoxFiltro_Enter | Private / Sub | 61 | — |
| TextBoxFiltro_Exit | Private / Sub | 68 | — |
| ListBoxLista_Change | Private / Sub | 75 | — |
| FrameResultado_Enter | Private / Sub | 85 | — |
| FrameResultado_Exit | Private / Sub | 91 | — |
| TextBoxMotor_Enter | Private / Sub | 109 | — |
| TextBoxMotor_Exit | Private / Sub | 116 | — |
| TextBoxCEST_Enter | Private / Sub | 123 | — |
| TextBoxCEST_Exit | Private / Sub | 130 | — |
| TextBoxNCM_Enter | Private / Sub | 137 | — |
| TextBoxNCM_Exit | Private / Sub | 144 | — |
| ListBoxLista_Enter | Private / Sub | 151 | — |
| ListBoxLista_Exit | Private / Sub | 157 | — |
| TextBoxFiltro_Change | Private / Sub | 163 | Cad_NCM.PreencheLista |
| UserForm_Initialize | Private / Sub | 169 | Cad_NCM.PreencheCampos |
| PreencheCampos | Private / Sub | 175 | — |
| PreencheCabecalho | Private / Sub | 191 | — |
| PreencheLista | Private / Sub | 207 | Cad_NCM.PreencheCabecalho, Cad_NCM.Array2DTranspose |
| Array2DTranspose | Public / Function | 250 | — |
| CommandButtonFechar_Click | Private / Sub | 278 | — |
| CommandButtonFechar_MouseMove | Private / Sub | 283 | mdlCursor.MouseCursor |

## Cad_NFE

Fonte: [Cad_NFE.frm](<./desenvolvimento/Form/Cad_NFE.frm>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| UserForm_Initialize | Private / Sub | 21 | MóduloRemoveFechar.RemoveBotao, mdlCopiar.UF_AtualizarPainel, MóduloPesquisa.ConfigurarListBoxPesquisa |
| TextBoxPesquisa_Change | Private / Sub | 48 | MóduloPesquisa.PesquisarEntradasRapido |
| TextBoxPesquisaFilial_Change | Private / Sub | 60 | MóduloPesquisa.PesquisarEntradasRapido |
| CommandButtonCopiarEntradaPFilial_Click | Private / Sub | 72 | — |
| CommandButtonDetranNet_Click | Private / Sub | 95 | modSeleniumConsultaNFE.CommandButtonConsultaDetran |
| CommandButtonDetranNet_MouseDown | Private / Sub | 99 | — |
| CommandButtonDetranNet_MouseMove | Private / Sub | 104 | mdlCursor.MouseCursor |
| CommandButtonDetranNet_MouseUp | Private / Sub | 113 | — |
| CommandButtonDetranNetBIN_Click | Private / Sub | 118 | modSeleniumRenavamConsulta.CommandButtonDetranNetBIN_Handler |
| CommandButton_Excluir_Click | Private / Sub | 123 | — |
| CommandButton_Excluir_MouseMove | Private / Sub | 190 | mdlCursor.MouseCursor |
| CommandButton_ExcluirFilial_Click | Private / Sub | 195 | — |
| CommandButton_ExcluirFilial_MouseMove | Private / Sub | 268 | mdlCursor.MouseCursor |
| CommandButton_Gravar_Click | Private / Sub | 273 | modIntegridadeOperacao.ValidarConversoesGravacao, clsOperacaoPlanilha.Iniciar, modIntegridadeOperacao.PrimeiraVagaGravacao, clsOperacaoPlanilha.RegistrarDestino, clsOperacaoPlanilha.MarcarSalvamentoIniciado, clsOperacaoPlanilha.Confirmar, clsOperacaoPlanilha.Recuperar, clsOperacaoPlanilha.Finalizar |
| CommandButton_Gravar_MouseMove | Private / Sub | 405 | mdlCursor.MouseCursor |
| CommandButton_GravarFilial_Click | Private / Sub | 410 | modIntegridadeOperacao.ValidarConversoesGravacao, modIntegridadeOperacao.ValidarChaveTexto, clsOperacaoPlanilha.Iniciar, modIntegridadeOperacao.PrimeiraVagaGravacao, clsOperacaoPlanilha.RegistrarDestino, clsOperacaoPlanilha.Confirmar, clsOperacaoPlanilha.Recuperar, clsOperacaoPlanilha.Finalizar |
| CommandButton_GravarFilial_MouseMove | Private / Sub | 540 | mdlCursor.MouseCursor |
| CommandButton_Limpar_Click | Private / Sub | 545 | — |
| CommandButton_Limpar_MouseMove | Private / Sub | 573 | mdlCursor.MouseCursor |
| CommandButton_LimparFilial_Click | Private / Sub | 578 | — |
| CommandButton_LimparFilial_MouseMove | Private / Sub | 605 | mdlCursor.MouseCursor |
| MultiPage1_Change | Private / Sub | 610 | — |
| CommandButtonCadNFESaida_Click | Private / Sub | 625 | — |
| CommandButtonCadNFESaida_MouseMove | Private / Sub | 632 | mdlCursor.MouseCursor |
| CommandButtonCadNFESaida_MouseDown | Private / Sub | 638 | — |
| CommandButtonCadNFESaida_MouseUp | Private / Sub | 643 | — |
| CommandButtonFechar_Click | Private / Sub | 648 | modSeleniumConsultaNFE.EncerrarSessaoDetranNFE |
| CommandButtonFechar_MouseMove | Private / Sub | 654 | mdlCursor.MouseCursor |
| CommandButtonFechar_MouseDown | Private / Sub | 659 | — |
| CommandButtonFechar_MouseUp | Private / Sub | 664 | — |
| CommandButtonAlterar_Click | Private / Sub | 669 | mdlCursor.MouseCursor |
| CommandButtonCopiarEntrada_Click | Private / Sub | 800 | — |
| CommandButtonCopiarEntradaFilial_Click | Private / Sub | 819 | — |
| CommandButtonCopiarEntradaFilial_MouseMove | Private / Sub | 841 | mdlCursor.MouseCursor |
| CommandButtonCopiarAutoEntradaFilial_Click | Private / Sub | 846 | — |
| CommandButtonCopiarAutoEntradaFilial_MouseMove | Private / Sub | 973 | mdlCursor.MouseCursor |
| CommandButtonFiltroCompra_Click | Private / Sub | 978 | — |
| CommandButtonFiltroCompra_MouseMove | Private / Sub | 984 | mdlCursor.MouseCursor |
| CommandButtonAlterar_MouseMove | Private / Sub | 989 | mdlCursor.MouseCursor |
| CommandButtonLerXML_MouseMove | Private / Sub | 993 | mdlCursor.MouseCursor |
| CommandButtonFiltroCompraFilial_Click | Private / Sub | 998 | — |
| CommandButtonFiltroCompraFilial_MouseMove | Private / Sub | 1004 | mdlCursor.MouseCursor |
| CommandButtonNCM_Click | Private / Sub | 1009 | — |
| Command_EntradaFilial_Click | Private / Sub | 1015 | — |
| CommandButton_SaidasFilial_Click | Private / Sub | 1021 | — |
| Command_EntradaFilial_MouseMove | Private / Sub | 1027 | mdlCursor.MouseCursor |
| CommandButton_SaidasFilial_MouseMove | Private / Sub | 1032 | mdlCursor.MouseCursor |
| CommandButtonCopiarEntrada_MouseMove | Private / Sub | 1037 | mdlCursor.MouseCursor |
| CommandButtonNCM_MouseMove | Private / Sub | 1042 | mdlCursor.MouseCursor |
| CommandButtonNCMFilial_Click | Private / Sub | 1047 | — |
| CommandButtonNCMFilial_MouseMove | Private / Sub | 1052 | mdlCursor.MouseCursor |
| CommandButtonChaveR_Click | Private / Sub | 1057 | — |
| CommandButtonChaveR_MouseMove | Private / Sub | 1169 | mdlCursor.MouseCursor |
| FrameMenu_MouseMove | Private / Sub | 1174 | — |
| Command_Entrada_MouseMove | Private / Sub | 1181 | mdlCursor.MouseCursor |
| CommandButton_Saidas_MouseMove | Private / Sub | 1186 | mdlCursor.MouseCursor |
| Command_Entrada_Click | Private / Sub | 1191 | — |
| CommandButton_Saidas_Click | Private / Sub | 1197 | — |
| CommandButtonSalvarTudo_Click | Private / Sub | 1203 | mdlCriaBackup.SaveWorkbookBackup |
| CommandButtonSalvarTudo_MouseMove | Private / Sub | 1228 | mdlCursor.MouseCursor |
| CommandButtonSalvarTudo_MouseDown | Private / Sub | 1234 | — |
| CommandButtonSalvarTudo_MouseUp | Private / Sub | 1239 | — |
| Data_KeyPress | Private / Sub | 1244 | Cad_NFE.SóNúmeros |
| DataFilial_KeyPress | Private / Sub | 1260 | Cad_NFE.SóNúmeros |
| txt_Valor_AfterUpdate | Private / Sub | 1276 | — |
| UserForm_QueryClose | Private / Sub | 1282 | — |
| CommandButtonLerXML_Click | Private / Sub | 1295 | — |
| txt_ValorFilial_AfterUpdate | Private / Sub | 1431 | — |
| ComboBox_Operação_Enter | Private / Sub | 1437 | — |
| ComboBox_Operação_Exit | Private / Sub | 1445 | — |
| ComboBox_OutroEnvio_Enter | Private / Sub | 1452 | — |
| ComboBox_OutroEnvio_Exit | Private / Sub | 1460 | — |
| ComboBox_OutroEnvioFilial_Enter | Private / Sub | 1467 | — |
| ComboBox_OutroEnvioFilial_Exit | Private / Sub | 1475 | — |
| TextBoxChaveDeAcessoFilial_Enter | Private / Sub | 1482 | — |
| TextBoxChaveDeAcessoFilial_Exit | Private / Sub | 1489 | — |
| N_Nfe_Enter | Private / Sub | 1496 | — |
| N_Nfe_Exit | Private / Sub | 1504 | — |
| UN_Enter | Private / Sub | 1530 | — |
| UN_Exit | Private / Sub | 1538 | — |
| Frame_Dados_da_NFe_Enter | Private / Sub | 1545 | — |
| Frame_Dados_da_NFe_Exit | Private / Sub | 1551 | — |
| PLACA_Enter | Private / Sub | 1573 | — |
| PLACA_Exit | Private / Sub | 1581 | — |
| COD_Enter | Private / Sub | 1588 | — |
| COD_Exit | Private / Sub | 1596 | — |
| Combo_ANO_MOD_Enter | Private / Sub | 1603 | — |
| Combo_ANO_MOD_Exit | Private / Sub | 1611 | — |
| ComboBox_COR_Enter | Private / Sub | 1618 | — |
| ComboBox_COR_Exit | Private / Sub | 1626 | — |
| ComboBox_VEÍCULO_Enter | Private / Sub | 1633 | — |
| ComboBox_VEÍCULO_Exit | Private / Sub | 1641 | — |
| ListBoxLista_Change | Private / Sub | 1648 | modPreencherFormulario.PreencherDadosEntradas, modPreencherFormulario.PreencherDadosEntradasFilial |
| ListBoxLista_Enter | Private / Sub | 1764 | — |
| ListBoxLista_Exit | Private / Sub | 1769 | — |
| ListBoxLista2_Change | Private / Sub | 1774 | modPreencherFormulario.PreencherDadosEntradas, modPreencherFormulario.PreencherDadosEntradasFilial |
| Frame_Veículo_Enter | Private / Sub | 1890 | — |
| Frame_Veículo_Exit | Private / Sub | 1896 | — |
| Data_Enter | Private / Sub | 1922 | — |
| Data_Exit | Private / Sub | 1932 | — |
| Frame_Data_Enter | Private / Sub | 1939 | — |
| Frame_Data_Exit | Private / Sub | 1949 | — |
| txt_Valor_Enter | Private / Sub | 1959 | — |
| txt_Valor_Exit | Private / Sub | 1967 | — |
| Frame_Valores_Enter | Private / Sub | 1974 | — |
| Frame_Valores_Exit | Private / Sub | 1984 | — |
| TextBoxPesquisa_Enter | Private / Sub | 1994 | — |
| TextBoxPesquisa_Exit | Private / Sub | 2001 | — |
| TextBoxPesquisaFilial_Enter | Private / Sub | 2007 | — |
| TextBoxPesquisaFilial_Exit | Private / Sub | 2014 | — |
| ListBoxLista2_Enter | Private / Sub | 2020 | — |
| ListBoxLista2_Exit | Private / Sub | 2025 | — |
| Fame_PesquisaFilial_Enter | Private / Sub | 2030 | — |
| Fame_PesquisaFilial_Exit | Private / Sub | 2036 | — |
| Frame_Pesquisa_Enter | Private / Sub | 2042 | — |
| Frame_Pesquisa_Exit | Private / Sub | 2048 | — |
| Frame_Ações_Enter | Private / Sub | 2054 | — |
| Frame_Ações_Exit | Private / Sub | 2060 | — |
| Frame_AçõesFilial_Enter | Private / Sub | 2066 | — |
| Frame_AçõesFilial_Exit | Private / Sub | 2072 | — |
| FrameMenu_Enter | Private / Sub | 2078 | — |
| FrameMenu_Exit | Private / Sub | 2096 | — |
| ComboBox_OperaçãoFilial_Enter | Private / Sub | 2105 | — |
| ComboBox_OperaçãoFilial_Exit | Private / Sub | 2113 | — |
| N_NfeFilial_Enter | Private / Sub | 2120 | — |
| N_NfeFilial_Exit | Private / Sub | 2128 | — |
| UNFilial_Enter | Private / Sub | 2154 | — |
| UNFilial_Exit | Private / Sub | 2162 | — |
| Frame_Dados_da_NFeFilial_Enter | Private / Sub | 2169 | — |
| Frame_Dados_da_NFeFilial_Exit | Private / Sub | 2175 | — |
| PLACAFilial_Enter | Private / Sub | 2193 | — |
| PLACAFilial_Exit | Private / Sub | 2201 | — |
| CODFilial_Enter | Private / Sub | 2208 | — |
| CODFilial_Exit | Private / Sub | 2216 | — |
| Combo_ANO_MODFilial_Enter | Private / Sub | 2223 | — |
| Combo_ANO_MODFilial_Exit | Private / Sub | 2231 | — |
| ComboBox_CORFilial_Enter | Private / Sub | 2238 | — |
| ComboBox_CORFilial_Exit | Private / Sub | 2246 | — |
| ComboBox_VEÍCULOFilial_Enter | Private / Sub | 2253 | — |
| ComboBox_VEÍCULOFilial_Exit | Private / Sub | 2261 | — |
| Frame_VeículoFilial_Enter | Private / Sub | 2268 | — |
| Frame_VeículoFilial_Exit | Private / Sub | 2278 | — |
| DataFilial_Enter | Private / Sub | 2304 | — |
| DataFilial_Exit | Private / Sub | 2314 | — |
| Frame_DataFilial_Enter | Private / Sub | 2321 | — |
| Frame_DataFilial_Exit | Private / Sub | 2331 | — |
| txt_ValorFilial_Enter | Private / Sub | 2341 | — |
| txt_ValorFilial_Exit | Private / Sub | 2349 | — |
| Frame_ValoresFilial_Enter | Private / Sub | 2356 | — |
| Frame_ValoresFilial_Exit | Private / Sub | 2366 | — |
| Frame_PesquisaFilial_Enter | Private / Sub | 2376 | — |
| Frame_PesquisaFilial_Exit | Private / Sub | 2382 | — |
| PLACA_KeyPress | Private / Sub | 2391 | — |
| PLACA_AfterUpdate | Private / Sub | 2396 | — |
| SóNúmeros | Private / Function | 2401 | — |
| N_Nfe_KeyPress | Private / Sub | 2411 | Cad_NFE.SóNúmeros |
| UN_KeyPress | Private / Sub | 2416 | Cad_NFE.SóNúmeros |
| COD_KeyPress | Private / Sub | 2421 | Cad_NFE.SóNúmeros |
| txt_Valor_KeyPress | Private / Sub | 2426 | Cad_NFE.SóNúmeros |
| N_NfeFilial_KeyPress | Private / Sub | 2431 | Cad_NFE.SóNúmeros |
| UNFilial_KeyPress | Private / Sub | 2436 | Cad_NFE.SóNúmeros |
| CODFilial_KeyPress | Private / Sub | 2441 | Cad_NFE.SóNúmeros |
| txt_ValorFilial_KeyPress | Private / Sub | 2446 | Cad_NFE.SóNúmeros |
| CommandButtonCelEsquerda_MouseMove | Private / Sub | 2451 | mdlCursor.MouseCursor |
| CommandButtonCelEsquerda_MouseDown | Private / Sub | 2456 | — |
| CommandButtonCelEsquerda_MouseUp | Private / Sub | 2461 | — |
| CommandButtonCelDireita_MouseMove | Private / Sub | 2466 | mdlCursor.MouseCursor |
| CommandButtonCelDireita_MouseDown | Private / Sub | 2471 | — |
| CommandButtonCelDireita_MouseUp | Private / Sub | 2476 | — |
| CommandButtonCelAcima_MouseMove | Private / Sub | 2481 | mdlCursor.MouseCursor |
| CommandButtonCelAcima_MouseDown | Private / Sub | 2486 | — |
| CommandButtonCelAcima_MouseUp | Private / Sub | 2491 | — |
| CommandButtonCelAbaixo_MouseMove | Private / Sub | 2496 | mdlCursor.MouseCursor |
| CommandButtonCelAbaixo_MouseDown | Private / Sub | 2501 | — |
| CommandButtonCelAbaixo_MouseUp | Private / Sub | 2506 | — |
| CommandButtonCelCopiar_MouseMove | Private / Sub | 2511 | mdlCursor.MouseCursor |
| CommandButtonCelCopiar_MouseDown | Private / Sub | 2516 | — |
| CommandButtonCelCopiar_MouseUp | Private / Sub | 2521 | — |
| CommandButtonCelEsquerda_Click | Private / Sub | 2526 | mdlCopiar.UF_NavegarCelula |
| CommandButtonCelDireita_Click | Private / Sub | 2532 | mdlCopiar.UF_NavegarCelula |
| CommandButtonCelAcima_Click | Private / Sub | 2538 | mdlCopiar.UF_NavegarCelula |
| CommandButtonCelAbaixo_Click | Private / Sub | 2544 | mdlCopiar.UF_NavegarCelula |
| CommandButtonCelCopiar_Click | Private / Sub | 2550 | mdlCopiar.UF_CopiarCelula |
| AtualizaPainelCelula | Private / Sub | 2554 | Cad_NFE.ControleExiste |
| ControleExiste | Private / Function | 2562 | — |

## Cad_NFE_Saida

Fonte: [Cad_NFE_Saida.frm](<./desenvolvimento/Form/Cad_NFE_Saida.frm>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| UserForm_Initialize | Private / Sub | 21 | Cad_NFE_Saida.PreencheCampos, Cad_NFE_Saida.PreencheCampos2 |
| Command_EntradaFilial_Click | Private / Sub | 39 | — |
| Command_Saida_Click | Private / Sub | 45 | — |
| CommandButton_ExcluirFilial_Click | Private / Sub | 50 | — |
| CommandButton_GravarFilial_Click | Private / Sub | 99 | modIntegridadeOperacao.ValidarConversoesGravacao, modIntegridadeOperacao.VerificarDestinoGravacao, clsOperacaoPlanilha.Iniciar, modIntegridadeOperacao.PrimeiraVagaGravacao, clsOperacaoPlanilha.RegistrarDestino, modIntegridadeOperacao.ProximaLinhaGravacao, clsOperacaoPlanilha.MarcarBaixaIniciada, clsOperacaoPlanilha.MarcarSalvamentoIniciado, clsOperacaoPlanilha.Confirmar, clsOperacaoPlanilha.Recuperar, clsOperacaoPlanilha.Finalizar |
| CommandButton_LimparFilial_Click | Private / Sub | 362 | — |
| CommandButton_SaidasFilial_Click | Private / Sub | 393 | — |
| CommandButtonNCM_Click | Private / Sub | 399 | — |
| CommandButtonNCMFilial_Click | Private / Sub | 405 | — |
| CommandButtonCadNFE_Click | Private / Sub | 411 | — |
| CommandButtonFechar_Click | Private / Sub | 418 | — |
| CommandButtonSalvarTudo_Click | Private / Sub | 423 | mdlCriaBackup.SaveWorkbookBackup |
| Command_Entrada_Click | Private / Sub | 447 | — |
| CommandButton_Saidas_Click | Private / Sub | 453 | — |
| MultiPage1_Change | Private / Sub | 459 | — |
| ComboBox_PorFilial_AfterUpdate | Private / Sub | 468 | — |
| CommandButton_Excluir_Click | Private / Sub | 484 | — |
| CommandButtonCopiarSaida_Click | Private / Sub | 544 | — |
| CommandButtonCopiarSaida_MouseMove | Private / Sub | 561 | mdlCursor.MouseCursor |
| CommandButton_Fechar_Click | Private / Sub | 566 | — |
| CommandButton_Gravar_Click | Private / Sub | 571 | modValidacaoXML.CarregarXMLNFe, modValidacaoXML.ValidarXMLNFe, modIntegridadeOperacao.ValidarConversoesGravacao, modIntegridadeOperacao.VerificarDestinoGravacao, modIntegridadeOperacao.ValidarChaveTexto, clsOperacaoPlanilha.Iniciar, modIntegridadeOperacao.PrimeiraVagaGravacao, clsOperacaoPlanilha.RegistrarDestino, modIntegridadeOperacao.ProximaLinhaGravacao, clsOperacaoPlanilha.MarcarBaixaIniciada, clsOperacaoPlanilha.MarcarSalvamentoIniciado, clsOperacaoPlanilha.Confirmar, Cad_NFE_Saida.VisibilidadeOutroEnvio, clsOperacaoPlanilha.Recuperar, clsOperacaoPlanilha.Finalizar |
| CommandButton_Limpar_Click | Private / Sub | 1014 | Cad_NFE_Saida.VisibilidadeOutroEnvio |
| TextBox_Valor_EntradaFilial_AfterUpdate | Private / Sub | 1048 | — |
| TextBox_Valor_LucroFilial_AfterUpdate | Private / Sub | 1070 | — |
| TextBox_Valor_Saida_AfterUpdate | Private / Sub | 1092 | — |
| TextBox_Valor_SaidaFilial_AfterUpdate | Private / Sub | 1146 | — |
| TextBoxBaseICMSFilial_MouseMove | Private / Sub | 1168 | — |
| UserForm_QueryClose | Private / Sub | 1174 | — |
| CommandButtonSalvarTudo_MouseMove | Private / Sub | 1187 | mdlCursor.MouseCursor |
| CommandButtonCadNFE_MouseMove | Private / Sub | 1193 | mdlCursor.MouseCursor |
| CommandButtonFechar_MouseMove | Private / Sub | 1199 | mdlCursor.MouseCursor |
| CommandButton_Gravar_MouseMove | Private / Sub | 1204 | mdlCursor.MouseCursor |
| CommandButton_Limpar_MouseMove | Private / Sub | 1209 | mdlCursor.MouseCursor |
| CommandButton_Excluir_MouseMove | Private / Sub | 1214 | mdlCursor.MouseCursor |
| Command_Entrada_MouseMove | Private / Sub | 1219 | mdlCursor.MouseCursor |
| CommandButton_Saidas_MouseMove | Private / Sub | 1224 | mdlCursor.MouseCursor |
| CommandButtonNCM_MouseMove | Private / Sub | 1229 | mdlCursor.MouseCursor |
| CommandButton_GravarFilial_MouseMove | Private / Sub | 1234 | mdlCursor.MouseCursor |
| CommandButton_LimparFilial_MouseMove | Private / Sub | 1239 | mdlCursor.MouseCursor |
| CommandButton_ExcluirFilial_MouseMove | Private / Sub | 1244 | mdlCursor.MouseCursor |
| Command_EntradaFilial_MouseMove | Private / Sub | 1249 | mdlCursor.MouseCursor |
| CommandButton_SaidasFilial_MouseMove | Private / Sub | 1254 | mdlCursor.MouseCursor |
| CommandButtonNCMFilial_MouseMove | Private / Sub | 1259 | mdlCursor.MouseCursor |
| FrameMenu_MouseMove | Private / Sub | 1264 | — |
| FrameMenu_Enter | Private / Sub | 1270 | — |
| FrameMenu_Exit | Private / Sub | 1276 | — |
| Frame_Ações_Enter | Private / Sub | 1282 | — |
| Frame_Ações_Exit | Private / Sub | 1292 | — |
| Frame_Dados_da_NFe_Enter | Private / Sub | 1298 | — |
| Frame_Dados_da_NFe_Exit | Private / Sub | 1304 | — |
| Frame_Veículo_Enter | Private / Sub | 1310 | — |
| Frame_Veículo_Exit | Private / Sub | 1324 | — |
| Frame_Valores_Enter | Private / Sub | 1330 | — |
| Frame_Valores_Exit | Private / Sub | 1340 | — |
| Frame_Data_Enter | Private / Sub | 1346 | — |
| Frame_Data_Exit | Private / Sub | 1360 | — |
| Frame_Pesquisa_Enter | Private / Sub | 1366 | — |
| Frame_Pesquisa_Exit | Private / Sub | 1372 | — |
| Frame_AçõesFilial_Enter | Private / Sub | 1378 | — |
| Frame_AçõesFilial_Exit | Private / Sub | 1388 | — |
| Frame_PesquisaFilial_Enter | Private / Sub | 1394 | — |
| Frame_PesquisaFilial_Exit | Private / Sub | 1400 | — |
| Frame_Dados_da_NFeFilial_Enter | Private / Sub | 1406 | — |
| Frame_Dados_da_NFeFilial_Exit | Private / Sub | 1412 | — |
| Frame_VeículoFilial_Enter | Private / Sub | 1418 | — |
| Frame_VeículoFilial_Exit | Private / Sub | 1428 | — |
| Frame_ValoresFilial_Enter | Private / Sub | 1434 | — |
| Frame_ValoresFilial_Exit | Private / Sub | 1444 | — |
| Frame_ImpostoFilial_Enter | Private / Sub | 1450 | — |
| Frame_ImpostoFilial_Exit | Private / Sub | 1460 | — |
| Frame_DataFilial_Enter | Private / Sub | 1466 | — |
| Frame_DataFilial_Exit | Private / Sub | 1480 | — |
| Fame_PesquisaFilial_Enter | Private / Sub | 1486 | — |
| Fame_PesquisaFilial_Exit | Private / Sub | 1492 | — |
| TextBox_ESTOQUE_Enter | Private / Sub | 1498 | — |
| TextBox_ESTOQUE_Exit | Private / Sub | 1505 | — |
| ComboBox_Operação_Change | Private / Sub | 1512 | Cad_NFE_Saida.VisibilidadeOutroEnvio, Cad_NFE_Saida.ZerarCampo |
| VisibilidadeOutroEnvio | Public / Sub | 1731 | — |
| ZerarCampo | Private / Sub | 1742 | — |
| ComboBox_Operação_Enter | Private / Sub | 1746 | — |
| ComboBox_Operação_Exit | Private / Sub | 1753 | — |
| TextBox_N_Saida_Enter | Private / Sub | 1760 | — |
| TextBox_N_Saida_Exit | Private / Sub | 1767 | — |
| TextBox_N_Entrada_Enter | Private / Sub | 1793 | — |
| TextBox_N_Entrada_Exit | Private / Sub | 1800 | — |
| ComboBox_OutroEnvio_Enter | Private / Sub | 1807 | — |
| ComboBox_OutroEnvio_Exit | Private / Sub | 1814 | — |
| TextBoxChaveDeAcesso_Enter | Private / Sub | 1821 | — |
| TextBoxChaveDeAcesso_Exit | Private / Sub | 1828 | — |
| ComboBox_OutroEnvioFilial_Enter | Private / Sub | 1835 | — |
| ComboBox_OutroEnvioFilial_Exit | Private / Sub | 1842 | — |
| TextBoxChaveDeAcessoFilial_Enter | Private / Sub | 1849 | — |
| TextBoxChaveDeAcessoFilial_Exit | Private / Sub | 1856 | — |
| COD_Enter | Private / Sub | 1863 | — |
| COD_Exit | Private / Sub | 1870 | — |
| ComboBox_VEÍCULO_Enter | Private / Sub | 1877 | — |
| ComboBox_VEÍCULO_Exit | Private / Sub | 1884 | — |
| TextBox_PLACA_Enter | Private / Sub | 1891 | — |
| TextBox_PLACA_Exit | Private / Sub | 1898 | — |
| ComboBox_Ano_Enter | Private / Sub | 1905 | — |
| ComboBox_Ano_Exit | Private / Sub | 1912 | — |
| ComboBox_COR_Enter | Private / Sub | 1919 | — |
| ComboBox_COR_Exit | Private / Sub | 1926 | — |
| TextBox_Valor_Entrada_Enter | Private / Sub | 1933 | — |
| TextBox_Valor_Entrada_Exit | Private / Sub | 1940 | — |
| TextBox_Valor_Saida_Enter | Private / Sub | 1949 | — |
| TextBox_Valor_Saida_Exit | Private / Sub | 1956 | — |
| TextBox_Valor_Lucro_Enter | Private / Sub | 1963 | — |
| TextBox_Valor_Lucro_Exit | Private / Sub | 1970 | — |
| TextBox_Data_Enter | Private / Sub | 1994 | — |
| TextBox_Data_Exit | Private / Sub | 2001 | — |
| ComboBox_Por_Enter | Private / Sub | 2008 | — |
| ComboBox_Por_Exit | Private / Sub | 2015 | — |
| TextBoxBaseICMS_Enter | Private / Sub | 2022 | — |
| TextBoxBaseICMS_Exit | Private / Sub | 2029 | — |
| TextBox_ICMS_Enter | Private / Sub | 2036 | — |
| TextBox_ICMS_Exit | Private / Sub | 2043 | — |
| TextBoxBasePISeCOFINS_Enter | Private / Sub | 2050 | — |
| TextBoxBasePISeCOFINS_Exit | Private / Sub | 2057 | — |
| TextBox_PIS_Enter | Private / Sub | 2064 | — |
| TextBox_PIS_Exit | Private / Sub | 2071 | — |
| TextBox_COFINS_Enter | Private / Sub | 2078 | — |
| TextBox_COFINS_Exit | Private / Sub | 2085 | — |
| TextBox_ESTOQUEFilial_Enter | Private / Sub | 2092 | — |
| TextBox_ESTOQUEFilial_Exit | Private / Sub | 2099 | — |
| ComboBox_OperaçãoFilial_Enter | Private / Sub | 2106 | — |
| ComboBox_OperaçãoFilial_Exit | Private / Sub | 2113 | — |
| TextBox_N_SaidaFilial_Enter | Private / Sub | 2120 | — |
| TextBox_N_SaidaFilial_Exit | Private / Sub | 2127 | — |
| TextBox_N_EntradaFilial_Enter | Private / Sub | 2153 | — |
| TextBox_N_EntradaFilial_Exit | Private / Sub | 2160 | — |
| CODFilial_Enter | Private / Sub | 2167 | — |
| CODFilial_Exit | Private / Sub | 2174 | — |
| ComboBox_VEÍCULOFilial_Enter | Private / Sub | 2181 | — |
| ComboBox_VEÍCULOFilial_Exit | Private / Sub | 2188 | — |
| TextBox_PLACAFilial_Enter | Private / Sub | 2195 | — |
| TextBox_PLACAFilial_Exit | Private / Sub | 2202 | — |
| ComboBox_AnoFilial_Enter | Private / Sub | 2209 | — |
| ComboBox_AnoFilial_Exit | Private / Sub | 2216 | — |
| ComboBox_CORFilial_Enter | Private / Sub | 2223 | — |
| ComboBox_CORFilial_Exit | Private / Sub | 2230 | — |
| TextBox_Valor_EntradaFilial_Enter | Private / Sub | 2237 | — |
| TextBox_Valor_EntradaFilial_Exit | Private / Sub | 2244 | — |
| TextBox_Valor_SaidaFilial_Enter | Private / Sub | 2251 | — |
| TextBox_Valor_SaidaFilial_Exit | Private / Sub | 2258 | — |
| TextBox_Valor_LucroFilial_Enter | Private / Sub | 2265 | — |
| TextBox_Valor_LucroFilial_Exit | Private / Sub | 2272 | — |
| TextBox_DataFilial_Enter | Private / Sub | 2279 | — |
| TextBox_DataFilial_Exit | Private / Sub | 2286 | — |
| ComboBox_PorFilial_Enter | Private / Sub | 2293 | — |
| ComboBox_PorFilial_Exit | Private / Sub | 2300 | — |
| TextBoxBaseICMSFilial_Enter | Private / Sub | 2307 | — |
| TextBoxBaseICMSFilial_Exit | Private / Sub | 2314 | — |
| TextBox_ICMSFilial_Enter | Private / Sub | 2321 | — |
| TextBox_ICMSFilial_Exit | Private / Sub | 2328 | — |
| TextBoxBasePISeCOFINSFilial_Enter | Private / Sub | 2335 | — |
| TextBoxBasePISeCOFINSFilial_Exit | Private / Sub | 2342 | — |
| TextBox_PISFilial_Enter | Private / Sub | 2349 | — |
| TextBox_PISFilial_Exit | Private / Sub | 2356 | — |
| TextBox_COFINSFilial_Enter | Private / Sub | 2363 | — |
| TextBox_COFINSFilial_Exit | Private / Sub | 2370 | — |
| ComboBoxCampos_Enter | Private / Sub | 2376 | — |
| ComboBoxCampos_Exit | Private / Sub | 2383 | — |
| ComboBoxCamposFilial_Enter | Private / Sub | 2390 | — |
| ComboBoxCamposFilial_Exit | Private / Sub | 2397 | — |
| TextBoxFiltroFilial_Enter | Private / Sub | 2404 | — |
| TextBoxFiltroFilial_Exit | Private / Sub | 2411 | — |
| TextBoxFiltro_Enter | Private / Sub | 2418 | — |
| TextBoxFiltro_Exit | Private / Sub | 2425 | — |
| ListBoxLista_Enter | Private / Sub | 2432 | — |
| ListBoxLista_Exit | Private / Sub | 2437 | — |
| ListBoxListaFilial_Enter | Private / Sub | 2442 | — |
| ListBoxListaFilial_Exit | Private / Sub | 2447 | — |
| TextBoxFiltro_Change | Private / Sub | 2453 | Cad_NFE_Saida.PreencheLista |
| PreencheCampos | Private / Sub | 2459 | — |
| PreencheCabecalho | Private / Sub | 2475 | — |
| PreencheLista | Private / Sub | 2491 | Cad_NFE_Saida.PreencheCabecalho, Cad_NFE_Saida.Array2DTranspose |
| Array2DTranspose | Public / Function | 2534 | — |
| ListBoxLista_Change | Private / Sub | 2562 | — |
| TextBoxFiltroFilial_Change | Private / Sub | 2623 | Cad_NFE_Saida.PreencheLista2 |
| PreencheCampos2 | Private / Sub | 2629 | — |
| PreencheCabecalho2 | Private / Sub | 2645 | — |
| PreencheLista2 | Private / Sub | 2661 | Cad_NFE_Saida.PreencheCabecalho2, Cad_NFE_Saida.Array2DTranspose |
| ListBoxListaFilial_Change | Private / Sub | 2704 | — |
| SóNúmeros | Private / Function | 2764 | — |
| TextBox_ESTOQUE_KeyPress | Private / Sub | 2774 | Cad_NFE_Saida.SóNúmeros |
| TextBox_N_Entrada_KeyPress | Private / Sub | 2779 | Cad_NFE_Saida.SóNúmeros |
| TextBox_N_Saida_KeyPress | Private / Sub | 2784 | Cad_NFE_Saida.SóNúmeros |
| COD_KeyPress | Private / Sub | 2789 | Cad_NFE_Saida.SóNúmeros |
| TextBox_Valor_Entrada_KeyPress | Private / Sub | 2794 | Cad_NFE_Saida.SóNúmeros |
| TextBox_Valor_Saida_KeyPress | Private / Sub | 2799 | Cad_NFE_Saida.SóNúmeros |
| TextBox_ESTOQUEFilial_KeyPress | Private / Sub | 2804 | Cad_NFE_Saida.SóNúmeros |
| TextBox_N_SaidaFilial_KeyPress | Private / Sub | 2809 | Cad_NFE_Saida.SóNúmeros |
| TextBox_N_EntradaFilial_KeyPress | Private / Sub | 2814 | Cad_NFE_Saida.SóNúmeros |
| CODFilial_KeyPress | Private / Sub | 2819 | Cad_NFE_Saida.SóNúmeros |
| TextBox_Valor_EntradaFilial_KeyPress | Private / Sub | 2824 | Cad_NFE_Saida.SóNúmeros |
| TextBox_Valor_SaidaFilial_KeyPress | Private / Sub | 2829 | Cad_NFE_Saida.SóNúmeros |

## frmRelatorioXML

Fonte: [frmRelatorioXML.frm](<./desenvolvimento/Form/frmRelatorioXML.frm>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| CarregarRelatorio | Public / Sub | 18 | — |
| CommandButton_Salvar_Click | Private / Sub | 93 | frmRelatorioXML.TextoSeguroListBox |
| CommandButton_Salvar_MouseMove | Private / Sub | 199 | mdlCursor.MouseCursor |
| TextoSeguroListBox | Private / Function | 204 | — |
| cmdFechar_Click | Private / Sub | 214 | — |
| cmdFechar_MouseMove | Private / Sub | 219 | mdlCursor.MouseCursor |
| cmdFechar_MouseDown | Private / Sub | 224 | — |
| cmdFechar_MouseUp | Private / Sub | 229 | — |

## UserFormFiltro

Fonte: [UserFormFiltro.frm](<./desenvolvimento/Form/UserFormFiltro.frm>). Codificação detectada: windows-1252; Option Explicit: true. Codificações ASCII são compatíveis com UTF-8/Windows-1252; o round-trip de bytes foi conferido conforme a base.

| Procedimento | Escopo / tipo | Linha inicial | Chamadas aproximadas |
|---|---|---|---|
| UserForm_Initialize | Private / Sub | 18 | mdlOcultarTítulo.HideTitleBar |
| CommandButtonFechar_Click | Private / Sub | 23 | — |
| CommandButtonFiltroCompra_Click | Private / Sub | 28 | — |
| CommandButtonFiltroCompra_MouseMove | Private / Sub | 35 | mdlCursor.MouseCursor |
| FrameMenu_MouseMove | Private / Sub | 41 | — |
| Command_Entrada_Click | Private / Sub | 54 | — |
| Command_Entrada_MouseMove | Private / Sub | 61 | mdlCursor.MouseCursor |
| CommandButtonFiltroAguardando_Click | Private / Sub | 67 | — |
| CommandButtonFiltroAguardando_MouseMove | Private / Sub | 74 | mdlCursor.MouseCursor |
| CommandButtonRepasse_Click | Private / Sub | 80 | — |
| CommandButtonRepasse_MouseMove | Private / Sub | 87 | mdlCursor.MouseCursor |
| CommandButtonVenda_Click | Private / Sub | 93 | — |
| CommandButtonVenda_MouseMove | Private / Sub | 100 | mdlCursor.MouseCursor |
| CommandButtonTransferência_Click | Private / Sub | 106 | — |
| CommandButtonTransferência_MouseMove | Private / Sub | 113 | mdlCursor.MouseCursor |
| CommandButtonDevolução_Click | Private / Sub | 119 | — |
| CommandButtonDevolução_MouseMove | Private / Sub | 126 | mdlCursor.MouseCursor |
| CommandButtonMatriz_Click | Private / Sub | 132 | — |
| CommandButtonMatriz_MouseMove | Private / Sub | 137 | mdlCursor.MouseCursor |
| CommandButtonFilial_Click | Private / Sub | 143 | — |
| CommandButtonFilial_MouseMove | Private / Sub | 148 | mdlCursor.MouseCursor |

