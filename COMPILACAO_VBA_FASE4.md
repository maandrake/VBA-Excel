# Resultado atual — compilação VBA real em 01/11/2026

**PASS de importação e compilação na fixture integrada; aplicação NÃO HOMOLOGADA.** O roteiro histórico abaixo foi executado pela interface nativa. Não reimporte componentes na cópia já integrada: isso pode criar duplicidades. O XLSM entregue está em [Fixture_Auditoria_Fase4_Integrada_Testada.xlsm](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Entrega_Fase4_20261101/Fixture_Auditoria_Fase4_Integrada_Testada.xlsm>), SHA-256 dcbd72e46e8b182a9810254bcce460ea2cb17260b4bcb47b622ddc3d4f9879a0.

Foram incorporados 34 componentes do sistema e três de testes, com cinco pares FRM/FRX e 18 documentos existentes. Contagem real: 55 (30 BAS, duas classes, cinco UserForms, 18 documentos); 17 abas/CodeNames conferidos. [Inventário atualizado](./_auditoria/fase4_20261031_230457Z/inventario-integracao.json), [corpos incorporados](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/fontes-integradas-entrega.json). O VBA salvo foi inspecionado localmente; esse passo complementa a compilação efetivamente feita pelo Excel.

## Evidência de compilação

No VBE, Depurar > Compilar VBAProject encontrou erro de sintaxe na suíte modTestesFase2. C4-01 corrigiu a chamada ExigirF2 com Call e parênteses. O comando foi repetido após as alterações incrementais C4-02/C4-03/C4-04; terminou sem erros e ficou desabilitado. A última conferência após salvar mostrou Compilar VBAProject desabilitado. Suítes reais retornaram PASS com 9, 377 e 11 verificações; nenhuma ausência de erro do VS Code foi usada como compilação. [Execução](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/resultado-final.json), [matriz](./MATRIZ_REGRESSAO_VBA_FASE4.md).

Excel 16.0 build 20527 x64. O acesso COM VBProject ainda retornou 0x800A03EC na sondagem exclusiva; a interface manual permaneceu permitida. Sem alteração de Trust Center, AccessVBOM, ExecutionPolicy, Registro, locais confiáveis, segurança de macros ou instalações. MSForms e MSXML6 foram executados nos casos da matriz; outras dependências de execução tardia/ActiveX não foram certificadas pela compilação. Não instalar ou registrar dependências automaticamente.

## Repetir apenas os testes autorizados

1. Preserve o XLSM e o pacote finais com SHA-256. Trabalhe numa nova cópia de nome iniciado por Fixture_Auditoria_ dentro de TEMP, fora do OneDrive. Confira Menu Principal!Z1 = AUDITORIA_VBA_FICTICIO. A cópia contém sentinelas fictícias e não deve receber dados reais.
2. Abra a cópia no Excel e Alt+F11. Se houver restrição de segurança, registre a mensagem; não altere proteções. Selecione o projeto pelo nome exato, não outro livro.
3. Depurar > Compilar VBAProject. Se habilitado, execute. Registre erro, linha/módulo e mensagem integral. Se já estiver desabilitado após compilação, registre esse estado; sozinho ele não prova toda a importação, use o inventário/corpos também.
4. Ctrl+G abre a Janela Imediata. Digite uma linha por vez: **? Auditoria_ExecutarTestesSeguros()**; depois **? AuditoriaFase2_Executar()**; depois **? AuditoriaFase4_IntegridadeComplementar()**. Saídas de referência: 9, 377 e 11 verificações PASS. As funções modificam somente células sentinela da fixture e podem deixar dados de teste. Registre a saída integral, não apenas ausência de mensagens.
5. **? AuditoriaFase3_CicloVida()** ainda não tem resultado aprovado. O VBE está em Interromper em módulo de classe e intercepta a exceção esperada vbObjectError + 2100 antes do tratador. A mudança temporária para Interromper em erros não tratados foi solicitada e não foi aplicada sem resposta. Não tratar o breakpoint como defeito de compilação ou PASS; não modificar automaticamente essa preferência. Se autorizado, registrar a configuração anterior, executar e restaurar a preferência ao final.
6. Não execute os botões de Venda/Transferência/Devolução/Excluir/Alterar sem massa fictícia e critérios D01/D02. Não chame Selenium/consultas fiscais/DETRAN. O roteiro de integração abaixo serve à recomposição em fixture nova, não à reimportação sobre a entrega existente.

C4-02 normaliza caminho curto/longo por FSO ShortPath nas três guardas da suíte; preserva o limite TEMP, prefixo e marcador. Nenhuma guarda foi removida. O comentário histórico PREPARADA, NAO EXECUTADA em um módulo de teste descreve seu preparo anterior; os resultados atuais estão documentados nesta execução, sem refatoração apenas para trocar o comentário.

## Restauração

O estado anterior dos relatórios está em [backup dos documentos](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/documentos-antes-consolidacao/). Checkpoints do XLSM e fontes anteriores estão em [evidências](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/), com manifestos SHA-256. Restaure somente por cópia para nome novo dentro de TEMP. Não substituir o original, nem considerar um backup conciliação de eventos ocorridos depois dele. O [relatório de integração](./RELATORIO_INTEGRACAO_FASE4.md) detalha os limites A08/A09.

---

# Roteiro histórico preservado — reconstrução a partir da fixture vazia

As contagens/status de bloqueio abaixo descrevem tentativas anteriores e foram superados pela execução acima. Use as instruções de importação somente em fixture nova, conferindo fontes finais/FRMFRX/CodeNames.

# Compilação e execução assistida — Fase 4

Estado em 20261031_230457Z: **AGUARDANDO AÇÃO NO EDITOR VBA**. O bloqueio anterior 0x800A03EC não foi sondado novamente. Importação manual não depende de conceder acesso programático ao projeto, mas pode ser impedida pelas proteções/referências do ambiente. Não alterar políticas, Trust Center, AccessVBOM, Registro ou instalar dependências.

## 1. Abrir a cópia correta

Abra no Excel [Fixture_Auditoria_Fase4.xlsm](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Fixture_Auditoria_Fase4.xlsm>). Use Arquivo > Abrir > Procurar e cole o caminho completo. A cópia é idêntica ao esqueleto fictício anterior, SHA-256 inicial `ff4f6b37d529e7404ae31bc65ceb1aeb3ba8f54aeb007b71c1e47c9bb9b9585d`. Está fora do OneDrive; não abra o XLSM de desenvolvimento arquivado com dados reais.

Confira na aba Menu Principal a célula Z1 = `AUDITORIA_VBA_FICTICIO`. O prefixo do nome deve continuar `Fixture_Auditoria_` e o arquivo deve continuar dentro de TEMP. Se aparecer restrição de edição/macros ou projeto protegido, registre a mensagem e pare a etapa; não desabilite proteções nem altere configurações.

## 2. Selecionar projeto e documentos

Pressione **Alt+F11**, depois **Ctrl+R** (Explorador de Projetos). Selecione exclusivamente VBAProject (Fixture_Auditoria_Fase4.xlsm). Use **F4** para Propriedades. As 17 abas devem ter exatamente os nomes abaixo; a propriedade `(Name)` é o CodeName, diferente da propriedade `Name` da aba.

| Aba | CodeName final | Corpo a colar no módulo existente |
|---|---|---|
| Menu Principal | Planilha12 | [Planilha12.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha12.txt>) |
| Entradas | Planilha1 | [Planilha1.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha1.txt>) |
| Transferidos | Planilha13 | [Planilha13.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha13.txt>) |
| Entradas Anual | Planilha3 | [Planilha3.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha3.txt>) |
| Demonstração | Planilha4 | [Planilha4.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha4.txt>) |
| Saidas | Planilha2 | [Planilha2.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha2.txt>) |
| Todas Saidas | Planilha17 | [Planilha17.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha17.txt>) |
| Entradas Filial | Planilha15 | [Planilha15.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha15.txt>) |
| Entradas Anual Filial | Planilha10 | [Planilha10.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha10.txt>) |
| Saidas Filial | Planilha16 | [Planilha16.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha16.txt>) |
| Todas Saidas Filial | Planilha18 | [Planilha18.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha18.txt>) |
| Carros PF | Planilha7 | [Planilha7.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha7.txt>) |
| Clientes | Planilha6 | [Planilha6.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha6.txt>) |
| Veículos | Planilha9 | [Planilha9.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha9.txt>) |
| Doc. para fora | Planilha11 | [Planilha11.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha11.txt>) |
| NCM e CEST | Planilha5 | [Planilha5.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha5.txt>) |
| Listas | Planilha8 | [Planilha8.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/Planilha8.txt>) |

Se houver colisão entre CodeNames, renomeie primeiro TODOS os 17 componentes novos para nomes provisórios exclusivos (por exemplo TmpF4_01 até TmpF4_17), só nesta fixture, e então aplique os nomes finais. Confira a correspondência pela aba, não pela posição anterior do componente. Em EstaPastaDeTrabalho confira `(Name) = EstaPastaDeTrabalho`.

Abra cada TXT em editor de texto, copie seu conteúdo e substitua o corpo do módulo de documento correspondente. Para a pasta use [EstaPastaDeTrabalho.txt](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Documentos/EstaPastaDeTrabalho.txt>). Não importe TXT/CLS como substituto de planilha/ThisWorkbook. Não cole cabeçalhos VERSION/Attribute nos módulos. Os TXT já contêm somente o corpo. Preserve CommandButton1_Click de Planilha15; não crie um botão novo.

## 3. Importar os componentes, sem executá-los

Pacote único: [Importar_No_VBE/](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE>). No VBE use **Arquivo > Importar arquivo** (Ctrl+M), repetindo por componente. Preserve os nomes do atributo VB_Name; não aceite módulos duplicados com sufixo 1. Se a fixture já contiver uma importação parcial, registre o estado e não importe outra versão por cima sem preservar/exportar o existente.

Sequência: **1)** classe de integridade; **2)** helper de integridade; **3)** outros 27 BAS; **4)** cinco FRM; **5)** classe fake; **6)** dois BAS de teste. A ordem evita dúvidas de dependência; não execute macros durante a importação.

| Etapa | Componente/arquivo a importar |
|---|---|
| 1 | [clsOperacaoPlanilha.cls](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/clsOperacaoPlanilha.cls>) |
| 2 | [modIntegridadeOperacao.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/modIntegridadeOperacao.bas>) |
| 3 | [mdlChamarFiltro.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlChamarFiltro.bas>) |
| 3 | [mdlChamarTabCan.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlChamarTabCan.bas>) |
| 3 | [mdlChamarTabEnt.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlChamarTabEnt.bas>) |
| 3 | [mdlChamarTabSai.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlChamarTabSai.bas>) |
| 3 | [mdlCondCel.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlCondCel.bas>) |
| 3 | [mdlCopiar.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlCopiar.bas>) |
| 3 | [mdlCriaBackup.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlCriaBackup.bas>) |
| 3 | [mdlCursor.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlCursor.bas>) |
| 3 | [mdlMacroConexão.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlMacroConexão.bas>) |
| 3 | [mdlMacroConexãoClientes.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlMacroConexãoClientes.bas>) |
| 3 | [mdlMacroDocFora.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlMacroDocFora.bas>) |
| 3 | [mdlOcultarTítulo.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlOcultarTítulo.bas>) |
| 3 | [mdlTelaCheia.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlTelaCheia.bas>) |
| 3 | [modSeleniumCommon.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/modSeleniumCommon.bas>) |
| 3 | [modSeleniumConsultaNFE.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/modSeleniumConsultaNFE.bas>) |
| 3 | [modSeleniumRenavamConsulta.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/modSeleniumRenavamConsulta.bas>) |
| 3 | [modSelTest.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/modSelTest.bas>) |
| 3 | [MóduloOfficeRibbonEditor.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/MóduloOfficeRibbonEditor.bas>) |
| 3 | [MóduloPesquisa.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/MóduloPesquisa.bas>) |
| 3 | [MóduloRemoveFechar.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/MóduloRemoveFechar.bas>) |
| 3 | [modValidacaoXML.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/modValidacaoXML.bas>) |
| 4 | [Cad_NCM.frm](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Form/Cad_NCM.frm>) |
| 4 | [Cad_NFE.frm](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Form/Cad_NFE.frm>) |
| 4 | [Cad_NFE_Saida.frm](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Form/Cad_NFE_Saida.frm>) |
| 4 | [frmRelatorioXML.frm](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Form/frmRelatorioXML.frm>) |
| 4 | [mdlChamarForm.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlChamarForm.bas>) |
| 4 | [mdlChamarFormDocFora.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlChamarFormDocFora.bas>) |
| 4 | [mdlChamarFormNCM.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlChamarFormNCM.bas>) |
| 4 | [mdlChamarFormSaida.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlChamarFormSaida.bas>) |
| 4 | [mdlChamarFormVeículos.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/mdlChamarFormVeículos.bas>) |
| 4 | [modPreencherFormulario.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/MDL/modPreencherFormulario.bas>) |
| 4 | [UserFormFiltro.frm](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Form/UserFormFiltro.frm>) |
| 5 | [AuditoriaListaFake.cls](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Testes/AuditoriaListaFake.cls>) |
| 6 | [modTestesAuditoria.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Testes/modTestesAuditoria.bas>) |
| 6 | [modTestesFase2.bas](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Importar_No_VBE/Testes/modTestesFase2.bas>) |

Os cinco FRM têm seus FRX de mesmo nome na mesma pasta. Importe **somente o FRM**; o VBE lê o FRX automaticamente. Não copie apenas o código do formulário, não importe FRX separadamente e não renomeie controles. Se houver erro de carga, preserve o arquivo LOG que o VBE criar e informe a mensagem. Os arquivos de staging são CP1252/CRLF; não os resalve em UTF-8.

Expectativa após importar tudo: **30 módulos padrão + 2 classes + 5 UserForms + 18 documentos = 55 componentes**. Conte os itens reais no Explorador; não informe 55 apenas por estar neste roteiro. O [inventário esperado](<C:/Users/MARCOSALEXANDREPERDO/OneDrive - MARCOS DESPACHANTE/Marcos Despachante/Software/VBA-Excel/_auditoria/fase4_20261031_230457Z/inventario-integracao.json>) contém todos os nomes e hashes. A contagem real ainda não existe.

## 4. Conferir referências e designers

No projeto correto, **Ferramentas > Referências**: registre a lista marcada e qualquer `AUSENTE/MISSING`. VBA/Excel/OLE Automation são as referências do host. Confira a disponibilidade das bibliotecas declaradas abaixo; marque somente a biblioteca instalada correspondente se estiver disponível e necessária à reprodução do projeto. Não selecione DLL/OCX por tentativa, não substitua por outra versão sem registrar, não remova uma referência para esconder erro. Se uma dependência não existir/for incompatível, pare a compilação e devolva o diagnóstico.

- OLE Automation (metadado original 2.0).
- Microsoft Office 16.0 Object Library (metadado original 2.0).
- Microsoft Forms 2.0 Object Library (metadado original 2.0).
- Microsoft Internet Controls (metadado original 1.1).
- Microsoft Windows Common Controls 6.0 (SP6) (metadado original 2.2).
- Microsoft Search CoClasses Type Library (metadado original 1.0).
- Microsoft Search Interface Type Library (metadado original 1.0).
- Microsoft XML, v6.0 (metadado original 6.0).
- Selenium Type Library (metadado original 2.0).

O caminho antigo de uma DLL não prova ausência: [diagnóstico da Fase 3](<C:/Users/MARCOSALEXANDREPERDO/OneDrive - MARCOS DESPACHANTE/Marcos Despachante/Software/VBA-Excel/_auditoria/fase3_20261031_213925Z/bibliotecas-registradas.json>). Forms pode aparecer ao importar o primeiro UserForm; Excel instalado é x64. Registro de Selenium/OCX não comprova compatibilidade. Não executar Selenium/DETRAN/refresh/autenticação para conferir referências.

Abra cada UserForm **no designer**, sem F5/Show. Confira `(Name)`, controles, propriedades, eventos e erros ActiveX. Não criar Frame_ImpostoFilial: a correção preservou os handlers e comentou quatro acessos ao nome ausente. Nomes definidos/RowSources/ControlSources/fórmulas operacionais e Ribbon não são reconstruídos por este esqueleto. Não testar os gravadores ainda; a aprovação dos auxiliares não libera esses eventos.

## 5. Compilar e guardar a evidência

Use **Depurar > Compilar VBAProject** no projeto da fixture. Se houver erro, registre mensagem integral, módulo, procedimento, linha destacada e texto da linha; não corrija por adivinhação. Informe também referência marcada como ausente, log de FRM e etapa da importação. Vou preparar a menor correção na cópia após receber o erro, e a compilação deverá ser repetida.

Se o comando estiver cinza, isso isoladamente não comprova compilação: confirme que não está em modo execução/interrupção nem selecionando outro projeto. Se continuar cinza, informe esse estado sem classificar PASS. Quando o comando efetivamente concluir sem mensagem, registre a ação e o resultado; salve **somente a fixture** (Ctrl+S com o projeto correto). Não use salvar todos.

## 6. Executar apenas as três funções existentes

Após compilação sem erro e conferência de segurança, **Ctrl+G** abre a Janela Imediata. Selecione um módulo da fixture antes de cada comando. Execute cada linha separadamente, com Enter, e copie a saída integral:

```vb
? ThisWorkbook.FullName
? ThisWorkbook.Worksheets.Count
? ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2
? Auditoria_ExecutarTestesSeguros()
? AuditoriaFase2_Executar()
? AuditoriaFase3_CicloVida()
```

As três primeiras linhas devem confirmar caminho da Fase 4, 17 e marcador. Se divergirem, não execute as funções. Os testes limpam/escrevem apenas intervalos pequenos da fixture: Entradas A1:K10, Entradas Filial A1:L10, A4:D4 em oito abas dos fluxos e Saidas F4 (formatação condicional). Esses dados não devem servir como estado inicial dos testes dos gravadores. Uma saída PASS de função não é PASS de formulário, XML, backup ou Ribbon. A função de ciclo de vida possui quatro asserções; a outra imprime sua contagem real. Não substituir esses resultados pelos COM da Fase 3.

Em erro, copie número/descrição e a linha destacada (se disponível). Não use End/Reset repetidamente para prosseguir após falha sem registrar, pois isso pode impedir a restauração dos estados Application. Depois de registrar, informe os quatro estados: `? Application.EnableEvents`, `? Application.ScreenUpdating`, `? Application.DisplayAlerts`, `? Application.Calculation`. Não altere valores manualmente supondo o estado anterior.

## 7. Comprovar revisão e devolver o resultado

No VBE exporte os **37 componentes importados**, usando Arquivo > Exportar arquivo em cada um, para [Exportados_VBE/](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Exportados_VBE>). Exportar FRM deve gerar FRX ao lado. Exporte os 18 documentos para [Documentos_Exportados_VBE/](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Documentos_Exportados_VBE>) como evidência, sem reimportá-los como classes. Não sobrescreva o pacote de referência. Vou comparar atributos, corpos, controles/recursos e bytes normalizados quando o VBE modificar cabeçalhos; hash binário diferente de exportação não prova mudança lógica.

Confira na classe os campos mIniciada e captura/restauração Errors(xlNumberAsText).Ignore. Nos quatro eventos gravadores, devem existir contexto e helper da revisão. Só conferir trechos não substitui a comparação completa dos exportados. Salve, feche e reabra somente a fixture e repita a conferência de projeto/referências/compilação antes de registrar persistência. Se não puder executar algum passo, indique BLOQUEADO ou NÃO EXECUTADO.

**Devolva em uma resposta:** caminho do XLSM salvo; contagem real por tipo; ação/resultado da compilação ou mensagem integral com localização; referências ausentes/erros FRM; saída das três funções; caminho dos exportados. Não precisa aprovar identidade, baixa, cancelamento ou segurança para relatar esses resultados.

## Restauração

O backup inicial da fixture está em [backup da fixture](<C:/Users/MARCOSALEXANDREPERDO/OneDrive - MARCOS DESPACHANTE/Marcos Despachante/Software/VBA-Excel/_auditoria/fase4_20261031_230457Z/backup/Fixture_Auditoria_Fase3.xlsm>) e tem o hash inicial acima. Os [114 arquivos de backup](<C:/Users/MARCOSALEXANDREPERDO/OneDrive - MARCOS DESPACHANTE/Marcos Despachante/Software/VBA-Excel/_auditoria/fase4_20261031_230457Z/backup>) estão no manifesto. Não sobrescreva uma tentativa com erro: preserve/exporte-a, feche somente essa cópia, restaure o backup sob **outro nome com prefixo Fixture_Auditoria_** dentro desta pasta TEMP e use o mesmo pacote conferido. A restauração é de arquivo inicial sem integração; não recupera uma operação de estoque.

**Resultado atual:** zero componentes revisados incorporados comprovados, compilação/execução VBA BLOQUEADAS pela intervenção pendente; sistema NÃO HOMOLOGADO. Não há nova evidência que altere os quatro cenários de conciliação e a recuperação incompleta da Fase 3.
