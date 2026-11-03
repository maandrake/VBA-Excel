> Atualização efetiva de 01/11/2026: a fixture da Fase 4 foi integrada e compilada no Excel real. Três suítes concluídas: 9 + 377 + 11 verificações PASS; recuperação ampla protegida continua FAIL e A08/A09 críticos. Sistema NÃO HOMOLOGADO. O [relatório atual](./RELATORIO_INTEGRACAO_FASE4.md) e a [matriz](./MATRIZ_REGRESSAO_VBA_FASE4.md) substituem os estados de importação bloqueada descritos historicamente abaixo.

# Continuidade — Fase 4 (20261031_230457Z)

Estado atual: [RELATORIO_INTEGRACAO_FASE4](./RELATORIO_INTEGRACAO_FASE4.md), [roteiro manual assistido](./COMPILACAO_VBA_FASE4.md), [matriz funcional](./MATRIZ_REGRESSAO_VBA_FASE4.md) e [proposta de identidade](./PLANO_IDENTIDADE_REGISTROS.md). Fontes e suítes copiados sem alteração; hashes conferidos e 114 backups criados. Nenhuma nova auditoria geral/teste equivalente nem tentativa VBProject. Integração, compilação e funções VBA aguardam intervenção manual; zero revisados incorporados comprovados e contagem total VBE não aferida. A08/A09 críticos; sistema NÃO HOMOLOGADO. Conteúdo anterior preservado abaixo.

---

# Atualização do mapa — Fase 3 (20261031_213925Z)

A nova [cópia de desenvolvimento](./_auditoria/fase3_20261031_213925Z/desenvolvimento/) mantém 34 fontes/665 procedimentos exportados. Outros 33 fontes são idênticos à Fase 2; o grafo anterior permanece referência histórica. O [grafo atual](./_auditoria/fase3_20261031_213925Z/evidencias/estatica-final/analise-estatica.json) foi recontado, sem duplicidade pública ou assinatura modificada.

Somente [clsOperacaoPlanilha.cls](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) mudou: mIniciada controla uso único, Recuperar recusa contexto inativo e o sexto item de snapshot preserva Errors(xlNumberAsText).Ignore. Os quatro consumidores dos formulários permanecem byte a byte iguais à Fase 2.

| Procedimento da classe atual | Escopo | Linha |
|---|---|---|
| [Class_Initialize](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | Private | 26 |
| [Iniciar](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | Public | 30 |
| [RegistrarDestino](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | Public | 45 |
| [MarcarBaixaIniciada](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | Public | 75 |
| [MarcarSalvamentoIniciado](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | Public | 80 |
| [Confirmar](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | Public | 84 |
| [Recuperar](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | Public | 89 |
| [RestaurarCelula](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | Private | 119 |
| [Finalizar](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | Public | 144 |
| [Class_Terminate](./_auditoria/fase3_20261031_213925Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) | Private | 164 |

A suíte existente em [modTestesFase2.bas](./_auditoria/fase3_20261031_213925Z/scripts/modTestesFase2.bas) acrescenta AuditoriaFase3_CicloVida/InicioRepetidoRecusadoF3 e a asserção do metadado. São testes preparados, fora do inventário do produto; não executados. Não houve renomeação de eventos ou criação de controles.

O [pacote manual](./_auditoria/fase3_20261031_213925Z/pacote-importacao.json) tem 34 componentes do sistema, três de teste e 18 documentos previstos: 55 esperados, **zero importados**. Não confundir a projeção de 52 componentes/666 procedimentos sem testes com o XLSM ainda intacto de 50/648. [Referências registradas](./_auditoria/fase3_20261031_213925Z/bibliotecas-registradas.json) não equivalem a IsBroken/compatibilidade resolvidos pelo host.

As dependências anuais sem identidade e o projeto A08 constam do [relatório](./RELATORIO_HOMOLOGACAO_FASE3.md). Nenhum campo, coluna ou política nova foi criado. Mapa anterior abaixo preservado como histórico.

---

# Mapa do projeto — atualização da fase 2

Execução 20261031_201621Z. O mapa histórico abaixo descreve a base e a primeira revisão; linhas atuais estão no [mapa completo de desenvolvimento](./_auditoria/fase2_20261031_201621Z/mapa-fontes-fase2.md) e no [grafo JSON](./_auditoria/fase2_20261031_201621Z/evidencias/estatica-final/analise-estatica.json).

## Inventários distintos

| Inventário | Componentes / fontes | Procedimentos | Situação |
|---|---|---|---|
| XLSM original e cópia binária de desenvolvimento | 50 componentes, 17 planilhas | 648 | Binário intacto, revisão não incorporada. |
| Fontes exportados da fase 1 | 32 fontes | 647 | Preservados integralmente na execução anterior. |
| Fontes exportados da fase 2 | 34 fontes | 665 | 647 assinaturas existentes preservadas e 18 procedimentos novos. |
| Projeção depois de importar fontes e documentos | 52 componentes, sem testes | 666 | Não realizada; soma o evento de Planilha15. |
| Preparador com três componentes de teste | 55 componentes esperados | Não atestado | Importação e compilação bloqueadas; fixture de integração incompleta. |

Há 17 fontes existentes alterados nesta fase e dois novos. Os cinco pares de formulários/designer/FRX permanecem íntegros; nenhum controle foi renomeado ou criado. Não houve remoção de eventos, UDFs, módulos públicos ou integrações.

## Arquivos com alterações da fase 2

| Fonte | Procedimentos atuais | Tipo de mudança |
|---|---|---|
| [mdlChamarForm.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlChamarForm.bas>) | 1 | Proprietário ou diagnóstico legado |
| [mdlChamarFormSaida.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlChamarFormSaida.bas>) | 1 | Proprietário ou diagnóstico legado |
| [mdlChamarTabCan.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlChamarTabCan.bas>) | 1 | Proprietário ou diagnóstico legado |
| [mdlChamarTabEnt.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlChamarTabEnt.bas>) | 1 | Proprietário ou diagnóstico legado |
| [mdlChamarTabSai.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlChamarTabSai.bas>) | 1 | Proprietário ou diagnóstico legado |
| [mdlCondCel.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlCondCel.bas>) | 6 | Condição/limites |
| [mdlCriaBackup.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlCriaBackup.bas>) | 4 | Backup versionado |
| [mdlMacroConexão.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlMacroConexão.bas>) | 1 | Proprietário ou diagnóstico legado |
| [mdlMacroConexãoClientes.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlMacroConexãoClientes.bas>) | 1 | Proprietário ou diagnóstico legado |
| [mdlMacroDocFora.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlMacroDocFora.bas>) | 1 | Proprietário ou diagnóstico legado |
| [modSeleniumCommon.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modSeleniumCommon.bas>) | 7 | Espera/timeout |
| [modSeleniumConsultaNFE.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modSeleniumConsultaNFE.bas>) | 55 | Espera/timeout |
| [modSelTest.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modSelTest.bas>) | 36 | Espera/timeout |
| [modValidacaoXML.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modValidacaoXML.bas>) | 47 | Estrutura e números XML |
| [Cad_NFE.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE.frm>) | 185 | Proprietário e mitigação de gravação |
| [Cad_NFE_Saida.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE_Saida.frm>) | 197 | Proprietário e mitigação de gravação |
| [UserFormFiltro.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/UserFormFiltro.frm>) | 21 | Proprietário ou diagnóstico legado |
| [clsOperacaoPlanilha.cls](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls>) | 10 | Novo auxiliar |
| [modIntegridadeOperacao.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modIntegridadeOperacao.bas>) | 5 | Novo auxiliar |

## Novas dependências técnicas

Os quatro eventos de gravação de Cad_NFE/Cad_NFE_Saida chamam [modIntegridadeOperacao](./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modIntegridadeOperacao.bas) para pré-conversão e destino, e instanciam [clsOperacaoPlanilha](./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls). A classe registra snapshots, compensa em ordem inversa antes de Delete/Save e restaura os estados de Application; não possui diário persistente.

modValidacaoXML chama IdentificacaoXMLValida antes de consumir um DOM recebido diretamente; SomarNosXML reporta entradas inválidas por AdicionarResultado. Funções fiscais existentes continuam idênticas. SleepMsNFE chama modSeleniumCommon.SleepMs, que usa a declaração Sleep compatível com VBA7/legada. O backup preserva os dois nomes públicos antigos e chama CriarBackupVersionado/VerificarArquivoBackup.

## Sequências de persistência e lacunas

O [mapa de múltiplas abas](./_auditoria/fase2_20261031_201621Z/operacoes-multiplas.json) cobre referências literais: oito procedimentos, dos quais seis alteram dados e dois navegam. Quatro gravadores receberam contexto de recuperação limitado; Alterar e ExcluirFilial continuam sem mapeamento/recuperação integral. Dependências indiretas, ActiveSheet intencional, controles, Ribbon, OnAction, chamadas dinâmicas e fórmulas exigem a conferência manual; o grafo não é prova de ausência de chamadas.

| Evento / operação | Sequência mantida |
|---|---|
| Cad_NFE.Gravar | Entradas A:K → Entradas Anual A:J → interface → Save |
| Cad_NFE.GravarFilial | Entradas Filial A:L → Entradas Anual Filial A:J → interface; sem Save nesse evento |
| Saída Venda | Saidas A:J/L:P → Todas Saidas A:J/L:P → baixa Entradas A:N → Save |
| Saída Transferência | Saidas → Entradas Filial A:L → Entradas Anual Filial A:J → Todas Saidas → baixa Entradas → Save |
| Saída Demonstração | Saidas → Demonstração A:J → Todas Saidas → baixa Entradas → Save |
| Saída Devolução filial | Saidas Filial → Todas Saidas Filial → Entradas A:K → Entradas Anual A:J → baixa Entradas Filial A:M → Save |
| Saída Cancelado | Saidas → Todas Saidas → Save; sem XML e sem baixa nesse ramo |
| Cad_NFE.Alterar | Limpa origem, grava primeira vaga, limpa seleção anterior em Entradas Anual e grava primeira vaga; vínculo anual indefinido |
| Cad_NFE.ExcluirFilial | Delete A:M em Entradas Filial → busca/exclui A:P no histórico anual; sem reversão estrutural |

## Sete pontos Delete atuais

| Fonte / procedimento | Linha | Comando preservado |
|---|---|---|
| [CommandButton_Excluir_Click](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE.frm>) | 146 | Selection.Delete Shift:=xlUp |
| [CommandButton_ExcluirFilial_Click](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE.frm>) | 217 | c.Resize(1, 13).Delete Shift:=xlUp |
| [CommandButton_ExcluirFilial_Click](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE.frm>) | 223 | cAnual.Resize(1, 16).Delete Shift:=xlUp |
| [CommandButton_ExcluirFilial_Click](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE_Saida.frm>) | 66 | Selection.EntireRow.Delete |
| [CommandButton_GravarFilial_Click](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE_Saida.frm>) | 270 | c.Resize(1, 13).Delete Shift:=xlUp |
| [CommandButton_Excluir_Click](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE_Saida.frm>) | 500 | Selection.EntireRow.Delete |
| [CommandButton_Gravar_Click](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE_Saida.frm>) | 896 | c.Resize(1, 14).Delete Shift:=xlUp |

O [mapeamento de fórmulas A08](./_auditoria/fase2_20261031_201621Z/formulas-a08.json) distingue campos e auxiliares, inclusive fórmulas compartilhadas. A troca dos Delete permanece condicionada à decisão/testes descritos no [plano](./PLANO_CORRECAO_FASE2.md).

## Planilhas, documentos e referências

Os nomes e CodeNames continuam sendo os 17 do [esquema da fixture](./_auditoria/fase2_20261031_201621Z/scripts/fixture-esquema.json); os documentos somam 18 componentes incluindo EstaPastaDeTrabalho. O único procedimento de planilha identificado é o evento já existente de Planilha15. Não foi criado botão substituto.

As referências declaradas da base continuam documentadas no mapa histórico. Cache EXD não equivale a biblioteca resolvida/instalada. MSXML 6/RegExp foram exercitados por COM; Forms, OCX e Selenium não tiveram importação comprovada. O [guia manual](./_auditoria/fase2_20261031_201621Z/COMPILACAO_MANUAL.md) explica os passos e o que falta reconstruir na fixture.

Não restam referências literais Sheets/Worksheets sem qualificação nos fontes; três ActiveWorkbook foram conservados nas conexões de proprietário indefinido. A navegação contextual ActiveSheet/ActiveCell permaneceu.

---

# Registro histórico — primeira auditoria

# Mapa do projeto VBA — 31/10/2026

## Inventário

A base original contém 39 arquivos: 27 módulos padrão, 5 formulários e seus 5 recursos FRX, uma pasta XLSM e um README. Não há classes independentes, Word DOCM/DOTM nem Access MDB/ACCDB disponíveis. Os componentes auxiliares de teste estão fora desse inventário original.

Pasta principal: [Estoque Fera Veículos 2026.xlsm](<./Excel/Estoque Fera Veículos 2026.xlsm>). O pacote contém 17 planilhas, 50 componentes VBA e 648 procedimentos distintos no ramo VBA7/Win64: 647 exportados e um evento de documento. Os 32 fontes exportados coincidem com os embutidos em termos de código; veja [comparacao-fontes-embutidos.json](<./_auditoria/20261031_174936Z/comparacao-fontes-embutidos.json>).

[manifesto-original.json](<./_auditoria/20261031_174936Z/manifesto-original.json>) contém todos os arquivos, tamanhos e hashes. [planilhas.json](<./_auditoria/20261031_174936Z/planilhas.json>) contém distribuição de fórmulas por coluna sem expor valores das células. A extração local de fontes está em [embutido_original/](<./_auditoria/20261031_174936Z/embutido_original/>).

## Planilhas e módulos de documento

| Planilha real | CodeName | Células com fórmula |
|---|---|---:|
| Menu Principal | Planilha12 | 208 |
| Entradas | Planilha1 | 2344 |
| Transferidos | Planilha13 | 3 |
| Entradas Anual | Planilha3 | 17971 |
| Demonstração | Planilha4 | 154 |
| Saidas | Planilha2 | 100 |
| Todas Saidas | Planilha17 | 16673 |
| Entradas Filial | Planilha15 | 1093 |
| Entradas Anual Filial | Planilha10 | 2974 |
| Saidas Filial | Planilha16 | 64 |
| Todas Saidas Filial | Planilha18 | 1501 |
| Carros PF | Planilha7 | 130 |
| Clientes | Planilha6 | 0 |
| Veículos | Planilha9 | 0 |
| Doc. para fora | Planilha11 | 0 |
| NCM e CEST | Planilha5 | 0 |
| Listas | Planilha8 | 17 |

EstaPastaDeTrabalho possui apenas atributos e Option Explicit. Dezesseis módulos de planilha também não têm procedimentos. Planilha15 (Entradas Filial) contém CommandButton1_Click, com mensagem de funcionalidade indisponível; o comentário embutido indica evento legado sem controles. O evento foi preservado. Esses 18 módulos não estavam exportados no workspace e foram lidos diretamente do projeto embutido. Não devem ser importados como classes normais.

## Dependências e pontos de entrada

- Ribbon: ChamarFrmEntrada, ChamarFrmSaida e ChamarSalvar foram confirmados no customUI14.xml. Abrem Cad_NFE/Cad_NFE_Saida e salvam a pasta. Nomes e assinaturas foram preservados.
- Cad_NFE: pesquisa e cadastro em Entradas/Entradas Filial; históricos Entradas Anual/Entradas Anual Filial; integração com Cad_NFE_Saida, NCM, filtros, leitura XML e módulos Selenium.
- Cad_NFE_Saida: Venda, Transferência, Devolução e Cancelado; Saidas, Todas Saidas, abas filiais, entradas e Demonstração. Contém cálculos e movimentações, portanto exige regressão funcional antes de refatoração ampla.
- modValidacaoXML: MSXML, código do item, número, chave, data, CFOP, benefício, CST e totais; apresenta frmRelatorioXML. O módulo permaneceu byte a byte intacto.
- MóduloPesquisa: leitura em arrays de Entradas/Entradas Filial, 14 colunas de resultado, origem na coluna 0, linha na 12, chave na 13. Pesquisa e navegação têm contratos diferentes dos campos de armazenamento.
- modPreencherFormulario: transfere colunas do resultado para controles reais, sem alterar o mapeamento nesta revisão.
- Cad_NCM: pesquisa a aba NCM e CEST; recursos e controles foram conferidos no FRX. Três filtros semelhantes distribuem-se entre NCM e saídas.
- UserFormFiltro: filtros sobre ActiveSheet e intervalos fixos; preservar o comportamento contextual até testar pasta/aba ativas.
- modSeleniumCommon, modSeleniumConsultaNFE, modSeleniumRenavamConsulta e modSelTest: Selenium/Chrome, sessão de Detran/Renavam, APIs Windows e diagnóstico. Não executados nesta auditoria.
- mdlCopiar, mdlCursor, mdlOcultarTítulo, MóduloRemoveFechar e mdlTelaCheia: navegação, clipboard e interface. A navegação contextual pode usar ActiveCell intencionalmente.
- mdlCondCel: avaliação de formatação condicional e UDFs; confira fórmulas e chamadas dinâmicas antes de remoção ou alteração.
- mdlCriaBackup e macros de conexão: destinos e integrações legadas; pendências A12/A18 no relatório.

## Referências declaradas e componentes COM

[referencias-declaradas.json](<./_auditoria/20261031_174936Z/referencias-declaradas.json>) registra as referências extraídas dos metadados do projeto. **Não foi possível verificar IsBroken no Office**, devido ao bloqueio VBOM. Caminhos declarados fora do workspace não foram abertos para investigação.

| Referência declarada | Versão |
|---|---|
| OLE Automation | 2.0 |
| Microsoft Office 16.0 Object Library | 2.0 |
| Microsoft Forms 2.0 Object Library | 2.0 |
| Microsoft Forms 2.0 Object Library | 2.0 |
| Microsoft Internet Controls | 1.1 |
| Microsoft Internet Controls | 1.1 |
| Microsoft Windows Common Controls 6.0 (SP6) | 2.2 |
| Microsoft Windows Common Controls 6.0 (SP6) | 2.2 |
| Microsoft Search CoClasses Type Library | 1.0 |
| Microsoft Search Interface Type Library | 1.0 |
| Microsoft XML, v6.0 | 6.0 |
| Selenium Type Library | 2.0 |

Os metadados incluem entradas de controles/extensores ligadas aos caches temporários MSForms.exd, SHDocVw.exd e MSComctlLib.exd, além de Selenium64.tlb. Registros nulos de controle foram excluídos da tabela; a lista não equivale a References/IsBroken resolvidos pelo VBIDE. Elas exigem verificação na máquina de destino; não se afirma que estejam quebradas nesta instalação. O arquivo declara dependências específicas de Office/Windows e Selenium, portanto compatibilidade de 32/64 bits não foi certificada.

ProgIDs literais encontrados: Forms.DataObject; MSForms.DataObject; MSXML2.DOMDocument.3.0; MSXML2.DOMDocument.6.0; Scripting.Dictionary; Scripting.FileSystemObject; VBScript.RegExp; WScript.Shell. O ChromeDriver também é criado a partir de constante pública. A disponibilidade do MSXML 6.0 e do Excel 16.0 foi comprovada por testes COM; Selenium e as demais integrações não foram inicializados. Não há connections.xml no pacote; as três macros de atualização citam nomes de conexões que não estão nesta pasta.

## Procedimentos e chamadas estáticas da cópia revisada

A lista abaixo inclui todos os **647 procedimentos exportados**, mantendo o número original. Linhas são da cópia revisada. As chamadas foram identificadas por tokens conhecidos após retirar comentários/strings; o grafo é aproximado e não resolve todos os membros COM, chamadas implícitas de eventos, OnAction, fórmulas ou CallByName. Ausência de chamada no grafo não autoriza remover código.

O conjunto completo e reproduzível está em [analise-estatica.json](<./_auditoria/20261031_174936Z/evidencias/revisado/analise-estatica.json>).

### mdlChamarFiltro

Fonte: [mdlChamarFiltro.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlChamarFiltro.bas>). Codificação: utf-8. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ChamarFormFiltro | Public | 3 | — |

### mdlChamarForm

Fonte: [mdlChamarForm.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlChamarForm.bas>). Codificação: utf-8. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ChamarForm | Public | 3 | — |

### mdlChamarFormDocFora

Fonte: [mdlChamarFormDocFora.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlChamarFormDocFora.bas>). Codificação: windows-1252. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ChamarFormDocFora | Public | 3 | — |

### mdlChamarFormNCM

Fonte: [mdlChamarFormNCM.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlChamarFormNCM.bas>). Codificação: utf-8. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ChamarFormNCM | Public | 3 | — |

### mdlChamarFormSaida

Fonte: [mdlChamarFormSaida.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlChamarFormSaida.bas>). Codificação: utf-8. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ChamarFormSaida | Public | 3 | — |

### mdlChamarFormVeículos

Fonte: [mdlChamarFormVeículos.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlChamarFormVeículos.bas>). Codificação: windows-1252. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ChamarFormVeículos | Public | 3 | — |

### mdlChamarTabCan

Fonte: [mdlChamarTabCan.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlChamarTabCan.bas>). Codificação: windows-1252. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ChamarTabCan | Public | 3 | — |

### mdlChamarTabEnt

Fonte: [mdlChamarTabEnt.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlChamarTabEnt.bas>). Codificação: utf-8. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ChamarTabEnt | Public | 3 | — |

### mdlChamarTabSai

Fonte: [mdlChamarTabSai.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlChamarTabSai.bas>). Codificação: utf-8. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ChamarTabSai | Public | 3 | — |

### mdlCondCel

Fonte: [mdlCondCel.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlCondCel.bas>). Codificação: utf-8. Procedimentos: 6.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ActiveCondition | Public | 4 | mdlCondCel.GetStrippedValue |
| ColorIndexOfCF | Public | 164 | mdlCondCel.ActiveCondition |
| ColorOfCF | Public | 185 | mdlCondCel.ActiveCondition |
| GetStrippedValue | Public | 205 | — |
| CountOfCF | Public | 218 | mdlCondCel.ActiveCondition |
| SumByCFColorIndex | Public | 234 | mdlCondCel.ColorIndexOfCF |

### mdlCopiar

Fonte: [mdlCopiar.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlCopiar.bas>). Codificação: windows-1252. Procedimentos: 8.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| UF_NavegarCelula | Public | 28 | mdlCopiar.MoveActiveCell, mdlCopiar.UF_AtualizaLabels |
| UF_AtualizarPainel | Public | 34 | mdlCopiar.UF_AtualizaLabels |
| UF_CopiarCelula | Public | 39 | mdlCopiar.Clipboard_SetText, mdlCopiar.UF_AtualizaLabels, mdlCopiar.UF_SetStatus |
| UF_AtualizaLabels | Private | 62 | mdlCopiar.UF_SetStatus, mdlCopiar.ControleExiste |
| UF_SetStatus | Private | 94 | mdlCopiar.ControleExiste |
| ControleExiste | Private | 101 | — |
| MoveActiveCell | Private | 111 | — |
| Clipboard_SetText | Private | 133 | — |

### mdlCriaBackup

Fonte: [mdlCriaBackup.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlCriaBackup.bas>). Codificação: windows-1252. Procedimentos: 2.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| SaveWorkbookBackup | Public | 4 | — |
| SaveWorkbookBackupToFloppy | Public | 51 | — |

### mdlCursor

Fonte: [mdlCursor.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlCursor.bas>). Codificação: utf-8. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| MouseCursor | Public | 27 | — |

### mdlMacroConexão

Fonte: [mdlMacroConexão.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlMacroConexão.bas>). Codificação: windows-1252. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| MacroConexão | Public | 3 | — |

### mdlMacroConexãoClientes

Fonte: [mdlMacroConexãoClientes.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlMacroConexãoClientes.bas>). Codificação: windows-1252. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ConexãoClientes | Public | 3 | — |

### mdlMacroDocFora

Fonte: [mdlMacroDocFora.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlMacroDocFora.bas>). Codificação: utf-8. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| MacroDocFora | Public | 3 | — |

### mdlOcultarTítulo

Fonte: [mdlOcultarTítulo.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlOcultarTítulo.bas>). Codificação: windows-1252. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| HideTitleBar | Public | 24 | — |

### mdlTelaCheia

Fonte: [mdlTelaCheia.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/mdlTelaCheia.bas>). Codificação: windows-1252. Procedimentos: 2.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| lsLigarTelaCheia | Public | 3 | — |
| lsDesligarTelaCheia | Public | 37 | — |

### modPreencherFormulario

Fonte: [modPreencherFormulario.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/modPreencherFormulario.bas>). Codificação: windows-1252. Procedimentos: 2.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| PreencherDadosEntradas | Public | 4 | — |
| PreencherDadosEntradasFilial | Public | 53 | — |

### modSeleniumCommon

Fonte: [modSeleniumCommon.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/modSeleniumCommon.bas>). Codificação: windows-1252. Procedimentos: 7.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ConfigureChromedriverExecutable | Public | 12 | modSeleniumCommon.EnsureDriverPath |
| EnsureDriverPath | Public | 21 | — |
| WaitForReadyStateComplete | Public | 36 | modSeleniumCommon.SegundosDecorridos, modSeleniumCommon.SleepMs |
| SleepMs | Public | 62 | modSeleniumCommon.SegundosDecorridos |
| TruncateForLog | Public | 70 | — |
| WriteLog | Public | 86 | modSeleniumCommon.TruncateForLog |
| SegundosDecorridos | Public | 125 | — |

### modSeleniumConsultaNFE

Fonte: [modSeleniumConsultaNFE.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/modSeleniumConsultaNFE.bas>). Codificação: windows-1252. Procedimentos: 55.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ConsultarDetranRetorno | Public | 53 | modSeleniumConsultaNFE.SanitizePlacaNFE, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.BuildConsultaUrlNFE, modSeleniumConsultaNFE.GetDetranDriverNFE, modSeleniumConsultaNFE.SessaoDetranProvavelmenteAutenticadaNFE, modSeleniumConsultaNFE.WaitForUrlNFE, modSeleniumConsultaNFE.UrlAtualCorrespondeNFE, modSeleniumConsultaNFE.EstaPaginaDeLoginNFE, modSeleniumConsultaNFE.WaitDocumentReadyNFE, modSeleniumConsultaNFE.DetectarPlacaNaoCadastradaNFE, modSeleniumConsultaNFE.TentarConsultaBINPorPlaca, modSeleniumConsultaNFE.NzTexto, modSeleniumConsultaNFE.BuildCollectionFromDossieNFE, modSeleniumConsultaNFE.RasparDossieVeiculoNFE, modSeleniumConsultaNFE.EncerrarSessaoDriverNFE |
| GetDetranDriverNFE | Private | 155 | modSeleniumConsultaNFE.TryReuseDetranDriverNFE, modSeleniumConsultaNFE.EncerrarSessaoDriverNFE, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.ConfigureChromedriverExecutableNFE, modSeleniumConsultaNFE.ConfigureAndStartChromeNFE |
| EncerrarSessaoDriverNFE | Private | 181 | — |
| EncerrarSessaoDetranNFE | Public | 190 | modSeleniumConsultaNFE.EncerrarSessaoDriverNFE |
| TryReuseDetranDriverNFE | Private | 194 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.EncerrarSessaoDriverNFE |
| SanitizePlacaNFE | Private | 213 | — |
| BuildConsultaUrlNFE | Private | 233 | modSeleniumConsultaNFE.SanitizePlacaNFE |
| ConfigureChromedriverExecutableNFE | Private | 242 | modSeleniumConsultaNFE.EnsureDriverPathNFE |
| EnsureDriverPathNFE | Private | 255 | — |
| ConfigureAndStartChromeNFE | Private | 271 | — |
| SessaoDetranProvavelmenteAutenticadaNFE | Private | 294 | — |
| EstaPaginaDeLoginNFE | Private | 320 | — |
| WaitForUrlNFE | Private | 337 | modSeleniumConsultaNFE.NFE_ElapsedSeconds, modSeleniumConsultaNFE.SleepMsNFE |
| UrlAtualCorrespondeNFE | Private | 357 | — |
| ObterJanelaChromeHandleNFE | Private | 374 | modSeleniumConsultaNFE.EnumWindowsProcNFE |
| EnumWindowsProcNFE | Private | 390 | — |
| WaitDocumentReadyNFE | Private | 420 | modSeleniumConsultaNFE.WaitForReadyStateCompleteNFE |
| WaitForReadyStateCompleteNFE | Private | 424 | modSeleniumConsultaNFE.NFE_ElapsedSeconds, modSeleniumConsultaNFE.SleepMsNFE |
| MostrarDriverWindowNFE | Private | 450 | modSeleniumConsultaNFE.NzTexto, modSeleniumConsultaNFE.ObterJanelaChromeHandleNFE |
| SleepMsNFE | Private | 482 | modSeleniumConsultaNFE.NFE_ElapsedSeconds |
| RasparDossieVeiculoNFE | Private | 491 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.LimpaNBSP, modSeleniumConsultaNFE.GetValorPorRotuloNFE |
| GetValorPorRotuloNFE | Private | 515 | — |
| BuildCollectionFromDossieNFE | Private | 541 | modSeleniumConsultaNFE.NzTexto |
| GetValueNFE | Public | 566 | modSeleniumConsultaNFE.NzTexto |
| TentarConsultaBINPorPlaca | Private | 582 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.NavegarParaBINNFE, modSeleniumConsultaNFE.PreencherESubmeterPlacaBINNFE, modSeleniumConsultaNFE.WaitDocumentReadyNFE, modSeleniumConsultaNFE.SleepMsNFE, modSeleniumConsultaNFE.RasparDossieBINNFE |
| NavegarParaBINNFE | Private | 601 | modSeleniumConsultaNFE.SwitchToDefaultContentNFE, modSeleniumConsultaNFE.WaitDocumentReadyNFE, modSeleniumConsultaNFE.SwitchToFrameByNameNFE, modSeleniumConsultaNFE.ClickByIdNFE, modSeleniumConsultaNFE.SleepMsNFE, modSeleniumConsultaNFE.LogNFE |
| PreencherESubmeterPlacaBINNFE | Private | 643 | modSeleniumConsultaNFE.SwitchToDefaultContentNFE, modSeleniumConsultaNFE.TentarEncontrarCampoBINNFE, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.SwitchToFrameByNameNFE, modSeleniumConsultaNFE.ClickByNameNFE, modSeleniumConsultaNFE.ClickByIdNFE |
| TentarEncontrarCampoBINNFE | Private | 681 | modSeleniumConsultaNFE.SwitchToDefaultContentNFE, modSeleniumConsultaNFE.SwitchToFrameByNameNFE, modSeleniumConsultaNFE.FindElementByNameNFE |
| RasparDossieBINNFE | Private | 710 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.SwitchToDefaultContentNFE, modSeleniumConsultaNFE.SwitchToFrameByNameNFE, modSeleniumConsultaNFE.LimparPlacaBINNFE, modSeleniumConsultaNFE.LimpaNBSP, modSeleniumConsultaNFE.GetValorBINPorRotuloNFE |
| GetValorBINPorRotuloNFE | Private | 761 | — |
| LimparPlacaBINNFE | Private | 787 | — |
| SwitchToDefaultContentNFE | Private | 807 | — |
| SwitchToFrameByNameNFE | Private | 813 | — |
| ClickByIdNFE | Private | 832 | — |
| ClickByNameNFE | Private | 848 | — |
| FindElementByNameNFE | Private | 864 | — |
| DetectarPlacaNaoCadastradaNFE | Private | 872 | — |
| LimpaNBSP | Private | 893 | — |
| NFE_ElapsedSeconds | Private | 897 | — |
| ConsultarDadosNFE | Public | 907 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.ConsultarDetranRetorno, modSeleniumConsultaNFE.FormatDadosNFE |
| FormatDadosNFE | Public | 926 | modSeleniumConsultaNFE.GetValueNFE, modSeleniumConsultaNFE.ExtrairAnoModelo, modSeleniumConsultaNFE.BuildResumoTexto, modSeleniumConsultaNFE.LimparMarcaModelo, modSeleniumConsultaNFE.LimparChassi, modSeleniumConsultaNFE.LimparCor |
| CopiarDadosDetranNFE | Public | 956 | modSeleniumConsultaNFE.ConsultarDadosNFE, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.CopiarParaClipboardNFE |
| BuildResumoTexto | Private | 978 | modSeleniumConsultaNFE.NzTexto |
| LimparMarcaModelo | Private | 1008 | — |
| LimparChassi | Private | 1029 | — |
| ExtrairAnoModelo | Private | 1050 | — |
| LimparCor | Private | 1074 | — |
| NzTexto | Private | 1089 | — |
| CommandButtonConsultaDetran | Public | 1097 | modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.CopiarDadosDetranNFE |
| CommandButton_Click | Public | 1111 | modSeleniumConsultaNFE.CommandButtonConsultaDetran |
| CommandButtonReexibirChrome | Public | 1115 | modSeleniumConsultaNFE.EncerrarSessaoDriverNFE, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.MostrarDriverWindowNFE |
| CopiarParaClipboardNFE | Private | 1148 | modSeleniumConsultaNFE.NzTexto, modSeleniumConsultaNFE.LogNFE, modSeleniumConsultaNFE.CopiarViaMsFormsDataObject, modSeleniumConsultaNFE.CopiarViaClipboardWinAPI |
| CopiarViaMsFormsDataObject | Private | 1182 | — |
| CopiarViaClipboardWinAPI | Private | 1225 | modSeleniumConsultaNFE.SleepMsNFE |
| LogNFE | Private | 1301 | modSeleniumCommon.WriteLog |

### modSeleniumRenavamConsulta

Fonte: [modSeleniumRenavamConsulta.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/modSeleniumRenavamConsulta.bas>). Codificação: windows-1252. Procedimentos: 23.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| AbrirConsultaRenavam | Public | 13 | modSeleniumRenavamConsulta.GetDetranDriverRenavam, modSeleniumRenavamConsulta.SessaoDetranProvavelmenteAutenticadaRenavam, modSeleniumCommon.WriteLog, modSeleniumRenavamConsulta.WaitForUrlRenavam, modSeleniumRenavamConsulta.UrlAtualCorrespondeRenavam, modSeleniumRenavamConsulta.EstaPaginaDeLoginRenavam, modSeleniumRenavamConsulta.NavegarMenuRenavam, modSeleniumRenavamConsulta.SelecionarUfRenavam, modSeleniumRenavamConsulta.PreencherPlacaRenavam, modSeleniumRenavamConsulta.EncerrarSessaoRenavam |
| CommandButtonDetranNetBIN_Handler | Public | 85 | modSeleniumRenavamConsulta.AbrirConsultaRenavam |
| EncerrarSessaoRenavam | Public | 98 | — |
| GetDetranDriverRenavam | Private | 105 | modSeleniumRenavamConsulta.TryReuseDetranDriverRenavam, modSeleniumRenavamConsulta.EncerrarSessaoRenavam, modSeleniumCommon.WriteLog, modSeleniumCommon.ConfigureChromedriverExecutable, modSeleniumRenavamConsulta.ConfigureAndStartChrome |
| TryReuseDetranDriverRenavam | Private | 131 | modSeleniumCommon.WriteLog, modSeleniumRenavamConsulta.EncerrarSessaoRenavam |
| SessaoDetranProvavelmenteAutenticadaRenavam | Private | 150 | — |
| WaitForUrlRenavam | Private | 176 | modSeleniumCommon.SegundosDecorridos, modSeleniumCommon.SleepMs |
| UrlAtualCorrespondeRenavam | Private | 196 | — |
| EstaPaginaDeLoginRenavam | Private | 213 | — |
| NavegarMenuRenavam | Private | 230 | modSeleniumRenavamConsulta.SwitchToDefaultContentSafe, modSeleniumRenavamConsulta.WaitDocumentReadyRenavam, modSeleniumRenavamConsulta.SwitchToFrameByNameSafe, modSeleniumRenavamConsulta.ClickByIdSafe, modSeleniumCommon.SleepMs, modSeleniumRenavamConsulta.AbrirConsultaOutrasUfs, modSeleniumCommon.WriteLog |
| AbrirConsultaOutrasUfs | Private | 259 | modSeleniumRenavamConsulta.SwitchToFrameByNameSafe, modSeleniumCommon.WriteLog, modSeleniumRenavamConsulta.ClickByIdSafe, modSeleniumCommon.SleepMs, modSeleniumRenavamConsulta.SwitchToDefaultContentSafe, modSeleniumRenavamConsulta.EnsureConsultaFrameRenavam, modSeleniumCommon.SegundosDecorridos |
| WaitDocumentReadyRenavam | Private | 326 | modSeleniumCommon.WaitForReadyStateComplete |
| ClickByIdSafe | Private | 330 | — |
| SelecionarUfRenavam | Private | 354 | modSeleniumRenavamConsulta.EnsureConsultaFrameRenavam |
| PreencherPlacaRenavam | Private | 381 | modSeleniumRenavamConsulta.EnsureConsultaFrameRenavam, modSeleniumRenavamConsulta.FindElementByNameSafe |
| EnsureConsultaFrameRenavam | Private | 405 | modSeleniumRenavamConsulta.SwitchToDefaultContentSafe, modSeleniumRenavamConsulta.SwitchToFrameByNameSafe, modSeleniumRenavamConsulta.FindElementByNameSafe, modSeleniumCommon.WriteLog, modSeleniumRenavamConsulta.FindElementByIdSafe, modSeleniumRenavamConsulta.FrameItemSafe |
| FrameItemSafe | Private | 491 | — |
| SwitchToDefaultContentSafe | Private | 503 | — |
| FindElementByNameSafe | Private | 509 | — |
| FindElementByIdSafe | Private | 515 | — |
| FindElementByXPathSafe | Private | 521 | — |
| SwitchToFrameByNameSafe | Private | 527 | modSeleniumRenavamConsulta.FindElementByXPathSafe |
| ConfigureAndStartChrome | Private | 546 | — |

### modSelTest

Fonte: [modSelTest.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/modSelTest.bas>). Codificação: windows-1252. Procedimentos: 36.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| BuildStatusBadge | Private | 21 | — |
| BuildStatusHeader | Private | 29 | modSelTest.BuildStatusBadge |
| SelTest_Run | Public | 39 | modSelTest.ValidateSeleniumEnvironment, modSelTest.TryCreateDriver, modSelTest.BuildSeleniumComDiagnostic, modSeleniumCommon.ConfigureChromedriverExecutable, modSeleniumCommon.WaitForReadyStateComplete, modSelTest.ParseChromeMajorFromUA, modSelTest.GetChromeDriverMajor, modSelTest.UrlReachable, modSelTest.BuildStatusHeader, modSelTest.AppendSelfTestLog |
| SmokeTestConsultaDetran | Public | 194 | modSeleniumCommon.WriteLog, modSeleniumConsultaNFE.ConsultarDetranRetorno, modSelTest.BuildSmokeTestFieldSpecs, modSeleniumConsultaNFE.GetValueNFE, modSelTest.BuildStatusHeader, modSelTest.BuildStatusBadge |
| BuildSmokeTestFieldSpecs | Private | 261 | modSelTest.AddFieldSpec |
| AddFieldSpec | Private | 305 | — |
| ValidateSeleniumEnvironment | Private | 311 | modSelTest.GetChromeVersionString, modSelTest.GetChromeDriverVersionString, modSelTest.ExtractMajorVersion, modSelTest.CheckDotNetRequirement, modSelTest.CheckDotNet35Feature, modSelTest.BuildSeleniumInstallationSummary, modSelTest.TrimTrailingCrLf |
| RunAndCapture | Private | 423 | — |
| ExtractMajorVersion | Private | 444 | — |
| TrimTrailingCrLf | Private | 462 | — |
| GetChromeVersionString | Private | 469 | modSelTest.GetFileVersionSafe, modSelTest.TryRegRead |
| GetChromeDriverVersionString | Private | 509 | modSelTest.GetFileVersionSafe, modSelTest.RunAndCapture |
| GetFileVersionSafe | Private | 529 | — |
| TryRegRead | Private | 543 | — |
| CheckDotNetRequirement | Private | 553 | modSelTest.GetDotNetReleaseValue, modSelTest.MapDotNetReleaseToVersion |
| CheckDotNet35Feature | Private | 578 | modSelTest.TryRegRead |
| GetDotNetReleaseValue | Private | 635 | modSelTest.TryRegRead |
| MapDotNetReleaseToVersion | Private | 663 | — |
| BuildSeleniumInstallationSummary | Private | 690 | modSelTest.GetSeleniumComServerCommand, modSelTest.NormalizeCommandPath, modSelTest.DetectSeleniumInstall |
| DetectSeleniumInstall | Private | 739 | modSelTest.GetSeleniumComServerCommand, modSelTest.NormalizeCommandPath, modSelTest.FileExists, modSelTest.GetParentFolderSafe, modSelTest.GetSeleniumCandidateFolders, modSelTest.FolderExists, modSelTest.FirstExistingFile, modSelTest.AppendPath |
| GetSeleniumComServerCommand | Private | 804 | modSelTest.TryRegRead |
| NormalizeCommandPath | Private | 828 | — |
| GetParentFolderSafe | Private | 861 | — |
| GetSeleniumCandidateFolders | Private | 871 | modSelTest.AddCandidateFolder, modSelTest.ExpandEnvPath |
| AddCandidateFolder | Private | 890 | — |
| ExpandEnvPath | Private | 908 | — |
| FirstExistingFile | Private | 926 | modSelTest.AppendPath, modSelTest.FileExists |
| AppendPath | Private | 938 | — |
| FolderExists | Private | 950 | — |
| FileExists | Private | 960 | — |
| AppendSelfTestLog | Private | 967 | modSeleniumCommon.WriteLog, modSelTest.FolderExists |
| TryCreateDriver | Private | 1001 | — |
| BuildSeleniumComDiagnostic | Private | 1013 | modSelTest.TryRegRead, modSelTest.NormalizeCommandPath, modSelTest.DetectSeleniumInstall |
| UrlReachable | Private | 1086 | modSeleniumCommon.WaitForReadyStateComplete |
| ParseChromeMajorFromUA | Private | 1095 | — |
| GetChromeDriverMajor | Private | 1108 | modSelTest.GetChromeDriverVersionString, modSelTest.ExtractMajorVersion |

### modValidacaoXML

Fonte: [modValidacaoXML.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/modValidacaoXML.bas>). Codificação: windows-1252. Procedimentos: 46.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| CarregarXMLNFe | Public | 14 | modValidacaoXML.SelecionarArquivoXML, modValidacaoXML.CriarDocumentoXML, modValidacaoXML.TextoAtributoXML, modValidacaoXML.RemoverPrefixoNFe, modValidacaoXML.TextoNoXML, modValidacaoXML.SomenteDigitos |
| ValidarXMLNFe | Public | 131 | modValidacaoXML.CarregarXMLNFe, modValidacaoXML.ValidarEstruturaEAutorizacao, modValidacaoXML.TextoNoXML, modValidacaoXML.AdicionarResultado, modValidacaoXML.DescricaoTipoNF, modValidacaoXML.ObterValorControle, modValidacaoXML.RemoverPrefixoNFe, modValidacaoXML.TextoAtributoXML, modValidacaoXML.NumerosNFeIguais, modValidacaoXML.ValidarDataEmissao, modValidacaoXML.LocalizarItemPorCodigo, modValidacaoXML.ListaCodigosProdutos, modValidacaoXML.AdicionarDependenciasNaoValidadas, modValidacaoXML.ValidarPlacaNoProduto, modValidacaoXML.ValidarOperacaoECFOP, modValidacaoXML.ValidarTransferenciaEntreEstabelecimentos, modValidacaoXML.ValidarBeneficioFiscal, modValidacaoXML.ValidarNumeroCampo, modValidacaoXML.ValidarICMSPorCST, modValidacaoXML.ValidarPISeCOFINSPorCST, modValidacaoXML.ValidarTotaisXML, modValidacaoXML.ExibirRelatorioVisual |
| ValidarEstruturaEAutorizacao | Private | 299 | modValidacaoXML.TextoNoXML, modValidacaoXML.AdicionarResultado, modValidacaoXML.DescricaoAmbiente, modValidacaoXML.IIfTexto, modValidacaoXML.TextoAtributoXML, modValidacaoXML.RemoverPrefixoNFe |
| ValidarOperacaoECFOP | Private | 372 | modValidacaoXML.NormalizarTexto, modValidacaoXML.AdicionarResultado, modValidacaoXML.AdicionarAviso |
| ValidarBeneficioFiscal | Private | 514 | modValidacaoXML.TextoNoXML, modValidacaoXML.AdicionarResultado, modValidacaoXML.AdicionarInformativo, modValidacaoXML.IIfTexto |
| ValidarTransferenciaEntreEstabelecimentos | Private | 633 | modValidacaoXML.NormalizarTexto, modValidacaoXML.SomenteDigitos, modValidacaoXML.TextoNoXML, modValidacaoXML.AdicionarResultado, modValidacaoXML.FormatarCNPJ |
| ValidarPlacaNoProduto | Private | 682 | modValidacaoXML.ObterValorControle, modValidacaoXML.TextoNoXML, modValidacaoXML.NormalizarAlfanumerico, modValidacaoXML.AdicionarResultado |
| ValidarICMSPorCST | Private | 712 | modValidacaoXML.TextoNoXML, modValidacaoXML.NormalizarCST, modValidacaoXML.AdicionarResultado, modValidacaoXML.ValidarNumeroCampo, modValidacaoXML.ObterValorControle, modValidacaoXML.TentarConverterNumeroXML, modValidacaoXML.ValidarNumeroCalculado, modValidacaoXML.AdicionarAviso, modValidacaoXML.AdicionarInformativo |
| ValidarPISeCOFINSPorCST | Private | 848 | modValidacaoXML.TextoNoXML, modValidacaoXML.ValidarPISPorCST, modValidacaoXML.ValidarCOFINSPorCST |
| ValidarPISPorCST | Private | 870 | modValidacaoXML.NormalizarCST, modValidacaoXML.AdicionarResultado, modValidacaoXML.TextoNoXML, modValidacaoXML.ValidarNumeroCampo, modValidacaoXML.DescricaoCSTSemIncidencia, modValidacaoXML.AdicionarInformativo |
| DescricaoCSTSemIncidencia | Private | 930 | — |
| ValidarCOFINSPorCST | Private | 953 | modValidacaoXML.NormalizarCST, modValidacaoXML.AdicionarResultado, modValidacaoXML.TextoNoXML, modValidacaoXML.ValidarNumeroCampo, modValidacaoXML.ObterValorControle, modValidacaoXML.TentarConverterNumeroXML, modValidacaoXML.ValidarNumeroCalculado, modValidacaoXML.DescricaoCSTPISCOFINS, modValidacaoXML.AdicionarInformativo |
| ValidarTotaisXML | Private | 1037 | modValidacaoXML.SomarNosXML, modValidacaoXML.ValidarNumeroCalculado, modValidacaoXML.TextoNoXML |
| AdicionarDependenciasNaoValidadas | Private | 1107 | modValidacaoXML.AdicionarInformativo |
| ValidarNumeroCampo | Private | 1134 | modValidacaoXML.TentarConverterNumeroFormulario, modValidacaoXML.TentarConverterNumeroXML, modValidacaoXML.AdicionarResultado |
| ValidarNumeroCalculado | Private | 1181 | modValidacaoXML.TentarConverterNumeroXML, modValidacaoXML.AdicionarResultado, modValidacaoXML.FormatarNumeroRelatorio |
| SomarNosXML | Private | 1221 | modValidacaoXML.TentarConverterNumeroXML |
| SelecionarArquivoXML | Private | 1251 | — |
| CriarDocumentoXML | Private | 1276 | — |
| TextoNoXML | Private | 1294 | — |
| TextoAtributoXML | Private | 1316 | — |
| LocalizarItemPorCodigo | Private | 1338 | modValidacaoXML.TextoNoXML |
| ListaCodigosProdutos | Private | 1370 | — |
| ObterValorControle | Private | 1399 | — |
| TentarConverterNumeroXML | Private | 1414 | — |
| TentarConverterNumeroFormulario | Private | 1448 | — |
| ValidarDataEmissao | Private | 1524 | modValidacaoXML.TentarConverterDataFormulario, modValidacaoXML.ConverterDataISO, modValidacaoXML.AdicionarResultado |
| TentarConverterDataFormulario | Private | 1574 | modValidacaoXML.CriarDataSegura |
| ConverterDataISO | Private | 1637 | modValidacaoXML.CriarDataSegura |
| CriarDataSegura | Private | 1664 | — |
| NumerosNFeIguais | Private | 1692 | — |
| NormalizarCST | Private | 1717 | — |
| DescricaoCSTPISCOFINS | Private | 1731 | — |
| DescricaoTipoNF | Private | 1749 | — |
| DescricaoAmbiente | Private | 1766 | — |
| NormalizarTexto | Private | 1786 | — |
| NormalizarAlfanumerico | Private | 1823 | — |
| SomenteDigitos | Private | 1845 | — |
| FormatarCNPJ | Private | 1865 | modValidacaoXML.SomenteDigitos |
| RemoverPrefixoNFe | Private | 1886 | — |
| FormatarNumeroRelatorio | Private | 1898 | — |
| IIfTexto | Private | 1904 | — |
| AdicionarResultado | Private | 1920 | — |
| AdicionarAviso | Private | 1955 | — |
| AdicionarInformativo | Private | 1979 | — |
| ExibirRelatorioVisual | Private | 2002 | — |

### MóduloOfficeRibbonEditor

Fonte: [MóduloOfficeRibbonEditor.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/MóduloOfficeRibbonEditor.bas>). Codificação: windows-1252. Procedimentos: 3.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ChamarFrmEntrada | Public | 3 | — |
| ChamarFrmSaida | Public | 8 | — |
| ChamarSalvar | Public | 13 | — |

### MóduloPesquisa

Fonte: [MóduloPesquisa.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/MóduloPesquisa.bas>). Codificação: windows-1252. Procedimentos: 6.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| ConfigurarListBoxPesquisa | Public | 13 | — |
| PesquisarEntradasRapido | Public | 36 | MóduloPesquisa.ConfigurarListBoxPesquisa, MóduloPesquisa.UltimaLinhaUsada, MóduloPesquisa.TransporParaListBox |
| AbrirRegistroDaLista | Public | 219 | — |
| UltimaLinhaUsada | Private | 260 | — |
| TransporParaListBox | Private | 273 | — |
| IrParaCelulaNFe | Public | 303 | — |

### MóduloRemoveFechar

Fonte: [MóduloRemoveFechar.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/MóduloRemoveFechar.bas>). Codificação: windows-1252. Procedimentos: 1.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| RemoveBotao | Public | 15 | — |

### Cad_NCM

Fonte: [Cad_NCM.frm](<./_auditoria/20261031_174936Z/desenvolvimento/Form/Cad_NCM.frm>). Codificação: windows-1252. Procedimentos: 25.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| FramePesquisa_Enter | Private | 25 | — |
| FramePesquisa_Exit | Private | 31 | — |
| ComboBoxCampos_Enter | Private | 47 | — |
| ComboBoxCampos_Exit | Private | 54 | — |
| TextBoxFiltro_Enter | Private | 61 | — |
| TextBoxFiltro_Exit | Private | 68 | — |
| ListBoxLista_Change | Private | 75 | — |
| FrameResultado_Enter | Private | 85 | — |
| FrameResultado_Exit | Private | 91 | — |
| TextBoxMotor_Enter | Private | 109 | — |
| TextBoxMotor_Exit | Private | 116 | — |
| TextBoxCEST_Enter | Private | 123 | — |
| TextBoxCEST_Exit | Private | 130 | — |
| TextBoxNCM_Enter | Private | 137 | — |
| TextBoxNCM_Exit | Private | 144 | — |
| ListBoxLista_Enter | Private | 151 | — |
| ListBoxLista_Exit | Private | 157 | — |
| TextBoxFiltro_Change | Private | 163 | Cad_NCM.PreencheLista |
| UserForm_Initialize | Private | 169 | Cad_NCM.PreencheCampos |
| PreencheCampos | Private | 175 | — |
| PreencheCabecalho | Private | 191 | — |
| PreencheLista | Private | 207 | Cad_NCM.PreencheCabecalho, Cad_NCM.Array2DTranspose |
| Array2DTranspose | Public | 250 | — |
| CommandButtonFechar_Click | Private | 278 | — |
| CommandButtonFechar_MouseMove | Private | 283 | mdlCursor.MouseCursor |

### Cad_NFE

Fonte: [Cad_NFE.frm](<./_auditoria/20261031_174936Z/desenvolvimento/Form/Cad_NFE.frm>). Codificação: windows-1252. Procedimentos: 185.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| UserForm_Initialize | Private | 21 | MóduloRemoveFechar.RemoveBotao, mdlCopiar.UF_AtualizarPainel, MóduloPesquisa.ConfigurarListBoxPesquisa |
| TextBoxPesquisa_Change | Private | 48 | MóduloPesquisa.PesquisarEntradasRapido |
| TextBoxPesquisaFilial_Change | Private | 60 | MóduloPesquisa.PesquisarEntradasRapido |
| CommandButtonCopiarEntradaPFilial_Click | Private | 72 | — |
| CommandButtonDetranNet_Click | Private | 95 | modSeleniumConsultaNFE.CommandButtonConsultaDetran |
| CommandButtonDetranNet_MouseDown | Private | 99 | — |
| CommandButtonDetranNet_MouseMove | Private | 104 | mdlCursor.MouseCursor |
| CommandButtonDetranNet_MouseUp | Private | 113 | — |
| CommandButtonDetranNetBIN_Click | Private | 118 | modSeleniumRenavamConsulta.CommandButtonDetranNetBIN_Handler |
| CommandButton_Excluir_Click | Private | 123 | — |
| CommandButton_Excluir_MouseMove | Private | 190 | mdlCursor.MouseCursor |
| CommandButton_ExcluirFilial_Click | Private | 195 | — |
| CommandButton_ExcluirFilial_MouseMove | Private | 268 | mdlCursor.MouseCursor |
| CommandButton_Gravar_Click | Private | 273 | — |
| CommandButton_Gravar_MouseMove | Private | 385 | mdlCursor.MouseCursor |
| CommandButton_GravarFilial_Click | Private | 390 | — |
| CommandButton_GravarFilial_MouseMove | Private | 492 | mdlCursor.MouseCursor |
| CommandButton_Limpar_Click | Private | 497 | — |
| CommandButton_Limpar_MouseMove | Private | 525 | mdlCursor.MouseCursor |
| CommandButton_LimparFilial_Click | Private | 530 | — |
| CommandButton_LimparFilial_MouseMove | Private | 557 | mdlCursor.MouseCursor |
| MultiPage1_Change | Private | 562 | — |
| CommandButtonCadNFESaida_Click | Private | 577 | — |
| CommandButtonCadNFESaida_MouseMove | Private | 584 | mdlCursor.MouseCursor |
| CommandButtonCadNFESaida_MouseDown | Private | 590 | — |
| CommandButtonCadNFESaida_MouseUp | Private | 595 | — |
| CommandButtonFechar_Click | Private | 600 | modSeleniumConsultaNFE.EncerrarSessaoDetranNFE |
| CommandButtonFechar_MouseMove | Private | 606 | mdlCursor.MouseCursor |
| CommandButtonFechar_MouseDown | Private | 611 | — |
| CommandButtonFechar_MouseUp | Private | 616 | — |
| CommandButtonAlterar_Click | Private | 621 | mdlCursor.MouseCursor |
| CommandButtonCopiarEntrada_Click | Private | 752 | — |
| CommandButtonCopiarEntradaFilial_Click | Private | 771 | — |
| CommandButtonCopiarEntradaFilial_MouseMove | Private | 793 | mdlCursor.MouseCursor |
| CommandButtonCopiarAutoEntradaFilial_Click | Private | 798 | — |
| CommandButtonCopiarAutoEntradaFilial_MouseMove | Private | 925 | mdlCursor.MouseCursor |
| CommandButtonFiltroCompra_Click | Private | 930 | — |
| CommandButtonFiltroCompra_MouseMove | Private | 936 | mdlCursor.MouseCursor |
| CommandButtonAlterar_MouseMove | Private | 941 | mdlCursor.MouseCursor |
| CommandButtonLerXML_MouseMove | Private | 945 | mdlCursor.MouseCursor |
| CommandButtonFiltroCompraFilial_Click | Private | 950 | — |
| CommandButtonFiltroCompraFilial_MouseMove | Private | 956 | mdlCursor.MouseCursor |
| CommandButtonNCM_Click | Private | 961 | — |
| Command_EntradaFilial_Click | Private | 967 | — |
| CommandButton_SaidasFilial_Click | Private | 973 | — |
| Command_EntradaFilial_MouseMove | Private | 979 | mdlCursor.MouseCursor |
| CommandButton_SaidasFilial_MouseMove | Private | 984 | mdlCursor.MouseCursor |
| CommandButtonCopiarEntrada_MouseMove | Private | 989 | mdlCursor.MouseCursor |
| CommandButtonNCM_MouseMove | Private | 994 | mdlCursor.MouseCursor |
| CommandButtonNCMFilial_Click | Private | 999 | — |
| CommandButtonNCMFilial_MouseMove | Private | 1004 | mdlCursor.MouseCursor |
| CommandButtonChaveR_Click | Private | 1009 | — |
| CommandButtonChaveR_MouseMove | Private | 1121 | mdlCursor.MouseCursor |
| FrameMenu_MouseMove | Private | 1126 | — |
| Command_Entrada_MouseMove | Private | 1133 | mdlCursor.MouseCursor |
| CommandButton_Saidas_MouseMove | Private | 1138 | mdlCursor.MouseCursor |
| Command_Entrada_Click | Private | 1143 | — |
| CommandButton_Saidas_Click | Private | 1149 | — |
| CommandButtonSalvarTudo_Click | Private | 1155 | mdlCriaBackup.SaveWorkbookBackup |
| CommandButtonSalvarTudo_MouseMove | Private | 1180 | mdlCursor.MouseCursor |
| CommandButtonSalvarTudo_MouseDown | Private | 1186 | — |
| CommandButtonSalvarTudo_MouseUp | Private | 1191 | — |
| Data_KeyPress | Private | 1196 | Cad_NFE.SóNúmeros |
| DataFilial_KeyPress | Private | 1212 | Cad_NFE.SóNúmeros |
| txt_Valor_AfterUpdate | Private | 1228 | — |
| UserForm_QueryClose | Private | 1234 | — |
| CommandButtonLerXML_Click | Private | 1247 | — |
| txt_ValorFilial_AfterUpdate | Private | 1383 | — |
| ComboBox_Operação_Enter | Private | 1389 | — |
| ComboBox_Operação_Exit | Private | 1397 | — |
| ComboBox_OutroEnvio_Enter | Private | 1404 | — |
| ComboBox_OutroEnvio_Exit | Private | 1412 | — |
| ComboBox_OutroEnvioFilial_Enter | Private | 1419 | — |
| ComboBox_OutroEnvioFilial_Exit | Private | 1427 | — |
| TextBoxChaveDeAcessoFilial_Enter | Private | 1434 | — |
| TextBoxChaveDeAcessoFilial_Exit | Private | 1441 | — |
| N_Nfe_Enter | Private | 1448 | — |
| N_Nfe_Exit | Private | 1456 | — |
| UN_Enter | Private | 1482 | — |
| UN_Exit | Private | 1490 | — |
| Frame_Dados_da_NFe_Enter | Private | 1497 | — |
| Frame_Dados_da_NFe_Exit | Private | 1503 | — |
| PLACA_Enter | Private | 1525 | — |
| PLACA_Exit | Private | 1533 | — |
| COD_Enter | Private | 1540 | — |
| COD_Exit | Private | 1548 | — |
| Combo_ANO_MOD_Enter | Private | 1555 | — |
| Combo_ANO_MOD_Exit | Private | 1563 | — |
| ComboBox_COR_Enter | Private | 1570 | — |
| ComboBox_COR_Exit | Private | 1578 | — |
| ComboBox_VEÍCULO_Enter | Private | 1585 | — |
| ComboBox_VEÍCULO_Exit | Private | 1593 | — |
| ListBoxLista_Change | Private | 1600 | modPreencherFormulario.PreencherDadosEntradas, modPreencherFormulario.PreencherDadosEntradasFilial |
| ListBoxLista_Enter | Private | 1716 | — |
| ListBoxLista_Exit | Private | 1721 | — |
| ListBoxLista2_Change | Private | 1726 | modPreencherFormulario.PreencherDadosEntradas, modPreencherFormulario.PreencherDadosEntradasFilial |
| Frame_Veículo_Enter | Private | 1842 | — |
| Frame_Veículo_Exit | Private | 1848 | — |
| Data_Enter | Private | 1874 | — |
| Data_Exit | Private | 1884 | — |
| Frame_Data_Enter | Private | 1891 | — |
| Frame_Data_Exit | Private | 1901 | — |
| txt_Valor_Enter | Private | 1911 | — |
| txt_Valor_Exit | Private | 1919 | — |
| Frame_Valores_Enter | Private | 1926 | — |
| Frame_Valores_Exit | Private | 1936 | — |
| TextBoxPesquisa_Enter | Private | 1946 | — |
| TextBoxPesquisa_Exit | Private | 1953 | — |
| TextBoxPesquisaFilial_Enter | Private | 1959 | — |
| TextBoxPesquisaFilial_Exit | Private | 1966 | — |
| ListBoxLista2_Enter | Private | 1972 | — |
| ListBoxLista2_Exit | Private | 1977 | — |
| Fame_PesquisaFilial_Enter | Private | 1982 | — |
| Fame_PesquisaFilial_Exit | Private | 1988 | — |
| Frame_Pesquisa_Enter | Private | 1994 | — |
| Frame_Pesquisa_Exit | Private | 2000 | — |
| Frame_Ações_Enter | Private | 2006 | — |
| Frame_Ações_Exit | Private | 2012 | — |
| Frame_AçõesFilial_Enter | Private | 2018 | — |
| Frame_AçõesFilial_Exit | Private | 2024 | — |
| FrameMenu_Enter | Private | 2030 | — |
| FrameMenu_Exit | Private | 2048 | — |
| ComboBox_OperaçãoFilial_Enter | Private | 2057 | — |
| ComboBox_OperaçãoFilial_Exit | Private | 2065 | — |
| N_NfeFilial_Enter | Private | 2072 | — |
| N_NfeFilial_Exit | Private | 2080 | — |
| UNFilial_Enter | Private | 2106 | — |
| UNFilial_Exit | Private | 2114 | — |
| Frame_Dados_da_NFeFilial_Enter | Private | 2121 | — |
| Frame_Dados_da_NFeFilial_Exit | Private | 2127 | — |
| PLACAFilial_Enter | Private | 2145 | — |
| PLACAFilial_Exit | Private | 2153 | — |
| CODFilial_Enter | Private | 2160 | — |
| CODFilial_Exit | Private | 2168 | — |
| Combo_ANO_MODFilial_Enter | Private | 2175 | — |
| Combo_ANO_MODFilial_Exit | Private | 2183 | — |
| ComboBox_CORFilial_Enter | Private | 2190 | — |
| ComboBox_CORFilial_Exit | Private | 2198 | — |
| ComboBox_VEÍCULOFilial_Enter | Private | 2205 | — |
| ComboBox_VEÍCULOFilial_Exit | Private | 2213 | — |
| Frame_VeículoFilial_Enter | Private | 2220 | — |
| Frame_VeículoFilial_Exit | Private | 2230 | — |
| DataFilial_Enter | Private | 2256 | — |
| DataFilial_Exit | Private | 2266 | — |
| Frame_DataFilial_Enter | Private | 2273 | — |
| Frame_DataFilial_Exit | Private | 2283 | — |
| txt_ValorFilial_Enter | Private | 2293 | — |
| txt_ValorFilial_Exit | Private | 2301 | — |
| Frame_ValoresFilial_Enter | Private | 2308 | — |
| Frame_ValoresFilial_Exit | Private | 2318 | — |
| Frame_PesquisaFilial_Enter | Private | 2328 | — |
| Frame_PesquisaFilial_Exit | Private | 2334 | — |
| PLACA_KeyPress | Private | 2343 | — |
| PLACA_AfterUpdate | Private | 2348 | — |
| SóNúmeros | Private | 2353 | — |
| N_Nfe_KeyPress | Private | 2363 | Cad_NFE.SóNúmeros |
| UN_KeyPress | Private | 2368 | Cad_NFE.SóNúmeros |
| COD_KeyPress | Private | 2373 | Cad_NFE.SóNúmeros |
| txt_Valor_KeyPress | Private | 2378 | Cad_NFE.SóNúmeros |
| N_NfeFilial_KeyPress | Private | 2383 | Cad_NFE.SóNúmeros |
| UNFilial_KeyPress | Private | 2388 | Cad_NFE.SóNúmeros |
| CODFilial_KeyPress | Private | 2393 | Cad_NFE.SóNúmeros |
| txt_ValorFilial_KeyPress | Private | 2398 | Cad_NFE.SóNúmeros |
| CommandButtonCelEsquerda_MouseMove | Private | 2403 | mdlCursor.MouseCursor |
| CommandButtonCelEsquerda_MouseDown | Private | 2408 | — |
| CommandButtonCelEsquerda_MouseUp | Private | 2413 | — |
| CommandButtonCelDireita_MouseMove | Private | 2418 | mdlCursor.MouseCursor |
| CommandButtonCelDireita_MouseDown | Private | 2423 | — |
| CommandButtonCelDireita_MouseUp | Private | 2428 | — |
| CommandButtonCelAcima_MouseMove | Private | 2433 | mdlCursor.MouseCursor |
| CommandButtonCelAcima_MouseDown | Private | 2438 | — |
| CommandButtonCelAcima_MouseUp | Private | 2443 | — |
| CommandButtonCelAbaixo_MouseMove | Private | 2448 | mdlCursor.MouseCursor |
| CommandButtonCelAbaixo_MouseDown | Private | 2453 | — |
| CommandButtonCelAbaixo_MouseUp | Private | 2458 | — |
| CommandButtonCelCopiar_MouseMove | Private | 2463 | mdlCursor.MouseCursor |
| CommandButtonCelCopiar_MouseDown | Private | 2468 | — |
| CommandButtonCelCopiar_MouseUp | Private | 2473 | — |
| CommandButtonCelEsquerda_Click | Private | 2478 | mdlCopiar.UF_NavegarCelula |
| CommandButtonCelDireita_Click | Private | 2484 | mdlCopiar.UF_NavegarCelula |
| CommandButtonCelAcima_Click | Private | 2490 | mdlCopiar.UF_NavegarCelula |
| CommandButtonCelAbaixo_Click | Private | 2496 | mdlCopiar.UF_NavegarCelula |
| CommandButtonCelCopiar_Click | Private | 2502 | mdlCopiar.UF_CopiarCelula |
| AtualizaPainelCelula | Private | 2506 | Cad_NFE.ControleExiste |
| ControleExiste | Private | 2514 | — |

### Cad_NFE_Saida

Fonte: [Cad_NFE_Saida.frm](<./_auditoria/20261031_174936Z/desenvolvimento/Form/Cad_NFE_Saida.frm>). Codificação: windows-1252. Procedimentos: 197.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| UserForm_Initialize | Private | 21 | Cad_NFE_Saida.PreencheCampos, Cad_NFE_Saida.PreencheCampos2 |
| Command_EntradaFilial_Click | Private | 39 | — |
| Command_Saida_Click | Private | 45 | — |
| CommandButton_ExcluirFilial_Click | Private | 50 | — |
| CommandButton_GravarFilial_Click | Private | 99 | — |
| CommandButton_LimparFilial_Click | Private | 321 | — |
| CommandButton_SaidasFilial_Click | Private | 352 | — |
| CommandButtonNCM_Click | Private | 358 | — |
| CommandButtonNCMFilial_Click | Private | 364 | — |
| CommandButtonCadNFE_Click | Private | 370 | — |
| CommandButtonFechar_Click | Private | 377 | — |
| CommandButtonSalvarTudo_Click | Private | 382 | mdlCriaBackup.SaveWorkbookBackup |
| Command_Entrada_Click | Private | 406 | — |
| CommandButton_Saidas_Click | Private | 412 | — |
| MultiPage1_Change | Private | 418 | — |
| ComboBox_PorFilial_AfterUpdate | Private | 427 | — |
| CommandButton_Excluir_Click | Private | 443 | — |
| CommandButtonCopiarSaida_Click | Private | 503 | — |
| CommandButtonCopiarSaida_MouseMove | Private | 520 | mdlCursor.MouseCursor |
| CommandButton_Fechar_Click | Private | 525 | — |
| CommandButton_Gravar_Click | Private | 530 | modValidacaoXML.CarregarXMLNFe, modValidacaoXML.ValidarXMLNFe, Cad_NFE_Saida.VisibilidadeOutroEnvio |
| CommandButton_Limpar_Click | Private | 935 | Cad_NFE_Saida.VisibilidadeOutroEnvio |
| TextBox_Valor_EntradaFilial_AfterUpdate | Private | 969 | — |
| TextBox_Valor_LucroFilial_AfterUpdate | Private | 991 | — |
| TextBox_Valor_Saida_AfterUpdate | Private | 1013 | — |
| TextBox_Valor_SaidaFilial_AfterUpdate | Private | 1067 | — |
| TextBoxBaseICMSFilial_MouseMove | Private | 1089 | — |
| UserForm_QueryClose | Private | 1095 | — |
| CommandButtonSalvarTudo_MouseMove | Private | 1108 | mdlCursor.MouseCursor |
| CommandButtonCadNFE_MouseMove | Private | 1114 | mdlCursor.MouseCursor |
| CommandButtonFechar_MouseMove | Private | 1120 | mdlCursor.MouseCursor |
| CommandButton_Gravar_MouseMove | Private | 1125 | mdlCursor.MouseCursor |
| CommandButton_Limpar_MouseMove | Private | 1130 | mdlCursor.MouseCursor |
| CommandButton_Excluir_MouseMove | Private | 1135 | mdlCursor.MouseCursor |
| Command_Entrada_MouseMove | Private | 1140 | mdlCursor.MouseCursor |
| CommandButton_Saidas_MouseMove | Private | 1145 | mdlCursor.MouseCursor |
| CommandButtonNCM_MouseMove | Private | 1150 | mdlCursor.MouseCursor |
| CommandButton_GravarFilial_MouseMove | Private | 1155 | mdlCursor.MouseCursor |
| CommandButton_LimparFilial_MouseMove | Private | 1160 | mdlCursor.MouseCursor |
| CommandButton_ExcluirFilial_MouseMove | Private | 1165 | mdlCursor.MouseCursor |
| Command_EntradaFilial_MouseMove | Private | 1170 | mdlCursor.MouseCursor |
| CommandButton_SaidasFilial_MouseMove | Private | 1175 | mdlCursor.MouseCursor |
| CommandButtonNCMFilial_MouseMove | Private | 1180 | mdlCursor.MouseCursor |
| FrameMenu_MouseMove | Private | 1185 | — |
| FrameMenu_Enter | Private | 1191 | — |
| FrameMenu_Exit | Private | 1197 | — |
| Frame_Ações_Enter | Private | 1203 | — |
| Frame_Ações_Exit | Private | 1213 | — |
| Frame_Dados_da_NFe_Enter | Private | 1219 | — |
| Frame_Dados_da_NFe_Exit | Private | 1225 | — |
| Frame_Veículo_Enter | Private | 1231 | — |
| Frame_Veículo_Exit | Private | 1245 | — |
| Frame_Valores_Enter | Private | 1251 | — |
| Frame_Valores_Exit | Private | 1261 | — |
| Frame_Data_Enter | Private | 1267 | — |
| Frame_Data_Exit | Private | 1281 | — |
| Frame_Pesquisa_Enter | Private | 1287 | — |
| Frame_Pesquisa_Exit | Private | 1293 | — |
| Frame_AçõesFilial_Enter | Private | 1299 | — |
| Frame_AçõesFilial_Exit | Private | 1309 | — |
| Frame_PesquisaFilial_Enter | Private | 1315 | — |
| Frame_PesquisaFilial_Exit | Private | 1321 | — |
| Frame_Dados_da_NFeFilial_Enter | Private | 1327 | — |
| Frame_Dados_da_NFeFilial_Exit | Private | 1333 | — |
| Frame_VeículoFilial_Enter | Private | 1339 | — |
| Frame_VeículoFilial_Exit | Private | 1349 | — |
| Frame_ValoresFilial_Enter | Private | 1355 | — |
| Frame_ValoresFilial_Exit | Private | 1365 | — |
| Frame_ImpostoFilial_Enter | Private | 1371 | — |
| Frame_ImpostoFilial_Exit | Private | 1381 | — |
| Frame_DataFilial_Enter | Private | 1387 | — |
| Frame_DataFilial_Exit | Private | 1401 | — |
| Fame_PesquisaFilial_Enter | Private | 1407 | — |
| Fame_PesquisaFilial_Exit | Private | 1413 | — |
| TextBox_ESTOQUE_Enter | Private | 1419 | — |
| TextBox_ESTOQUE_Exit | Private | 1426 | — |
| ComboBox_Operação_Change | Private | 1433 | Cad_NFE_Saida.VisibilidadeOutroEnvio, Cad_NFE_Saida.ZerarCampo |
| VisibilidadeOutroEnvio | Public | 1652 | — |
| ZerarCampo | Private | 1663 | — |
| ComboBox_Operação_Enter | Private | 1667 | — |
| ComboBox_Operação_Exit | Private | 1674 | — |
| TextBox_N_Saida_Enter | Private | 1681 | — |
| TextBox_N_Saida_Exit | Private | 1688 | — |
| TextBox_N_Entrada_Enter | Private | 1714 | — |
| TextBox_N_Entrada_Exit | Private | 1721 | — |
| ComboBox_OutroEnvio_Enter | Private | 1728 | — |
| ComboBox_OutroEnvio_Exit | Private | 1735 | — |
| TextBoxChaveDeAcesso_Enter | Private | 1742 | — |
| TextBoxChaveDeAcesso_Exit | Private | 1749 | — |
| ComboBox_OutroEnvioFilial_Enter | Private | 1756 | — |
| ComboBox_OutroEnvioFilial_Exit | Private | 1763 | — |
| TextBoxChaveDeAcessoFilial_Enter | Private | 1770 | — |
| TextBoxChaveDeAcessoFilial_Exit | Private | 1777 | — |
| COD_Enter | Private | 1784 | — |
| COD_Exit | Private | 1791 | — |
| ComboBox_VEÍCULO_Enter | Private | 1798 | — |
| ComboBox_VEÍCULO_Exit | Private | 1805 | — |
| TextBox_PLACA_Enter | Private | 1812 | — |
| TextBox_PLACA_Exit | Private | 1819 | — |
| ComboBox_Ano_Enter | Private | 1826 | — |
| ComboBox_Ano_Exit | Private | 1833 | — |
| ComboBox_COR_Enter | Private | 1840 | — |
| ComboBox_COR_Exit | Private | 1847 | — |
| TextBox_Valor_Entrada_Enter | Private | 1854 | — |
| TextBox_Valor_Entrada_Exit | Private | 1861 | — |
| TextBox_Valor_Saida_Enter | Private | 1870 | — |
| TextBox_Valor_Saida_Exit | Private | 1877 | — |
| TextBox_Valor_Lucro_Enter | Private | 1884 | — |
| TextBox_Valor_Lucro_Exit | Private | 1891 | — |
| TextBox_Data_Enter | Private | 1915 | — |
| TextBox_Data_Exit | Private | 1922 | — |
| ComboBox_Por_Enter | Private | 1929 | — |
| ComboBox_Por_Exit | Private | 1936 | — |
| TextBoxBaseICMS_Enter | Private | 1943 | — |
| TextBoxBaseICMS_Exit | Private | 1950 | — |
| TextBox_ICMS_Enter | Private | 1957 | — |
| TextBox_ICMS_Exit | Private | 1964 | — |
| TextBoxBasePISeCOFINS_Enter | Private | 1971 | — |
| TextBoxBasePISeCOFINS_Exit | Private | 1978 | — |
| TextBox_PIS_Enter | Private | 1985 | — |
| TextBox_PIS_Exit | Private | 1992 | — |
| TextBox_COFINS_Enter | Private | 1999 | — |
| TextBox_COFINS_Exit | Private | 2006 | — |
| TextBox_ESTOQUEFilial_Enter | Private | 2013 | — |
| TextBox_ESTOQUEFilial_Exit | Private | 2020 | — |
| ComboBox_OperaçãoFilial_Enter | Private | 2027 | — |
| ComboBox_OperaçãoFilial_Exit | Private | 2034 | — |
| TextBox_N_SaidaFilial_Enter | Private | 2041 | — |
| TextBox_N_SaidaFilial_Exit | Private | 2048 | — |
| TextBox_N_EntradaFilial_Enter | Private | 2074 | — |
| TextBox_N_EntradaFilial_Exit | Private | 2081 | — |
| CODFilial_Enter | Private | 2088 | — |
| CODFilial_Exit | Private | 2095 | — |
| ComboBox_VEÍCULOFilial_Enter | Private | 2102 | — |
| ComboBox_VEÍCULOFilial_Exit | Private | 2109 | — |
| TextBox_PLACAFilial_Enter | Private | 2116 | — |
| TextBox_PLACAFilial_Exit | Private | 2123 | — |
| ComboBox_AnoFilial_Enter | Private | 2130 | — |
| ComboBox_AnoFilial_Exit | Private | 2137 | — |
| ComboBox_CORFilial_Enter | Private | 2144 | — |
| ComboBox_CORFilial_Exit | Private | 2151 | — |
| TextBox_Valor_EntradaFilial_Enter | Private | 2158 | — |
| TextBox_Valor_EntradaFilial_Exit | Private | 2165 | — |
| TextBox_Valor_SaidaFilial_Enter | Private | 2172 | — |
| TextBox_Valor_SaidaFilial_Exit | Private | 2179 | — |
| TextBox_Valor_LucroFilial_Enter | Private | 2186 | — |
| TextBox_Valor_LucroFilial_Exit | Private | 2193 | — |
| TextBox_DataFilial_Enter | Private | 2200 | — |
| TextBox_DataFilial_Exit | Private | 2207 | — |
| ComboBox_PorFilial_Enter | Private | 2214 | — |
| ComboBox_PorFilial_Exit | Private | 2221 | — |
| TextBoxBaseICMSFilial_Enter | Private | 2228 | — |
| TextBoxBaseICMSFilial_Exit | Private | 2235 | — |
| TextBox_ICMSFilial_Enter | Private | 2242 | — |
| TextBox_ICMSFilial_Exit | Private | 2249 | — |
| TextBoxBasePISeCOFINSFilial_Enter | Private | 2256 | — |
| TextBoxBasePISeCOFINSFilial_Exit | Private | 2263 | — |
| TextBox_PISFilial_Enter | Private | 2270 | — |
| TextBox_PISFilial_Exit | Private | 2277 | — |
| TextBox_COFINSFilial_Enter | Private | 2284 | — |
| TextBox_COFINSFilial_Exit | Private | 2291 | — |
| ComboBoxCampos_Enter | Private | 2297 | — |
| ComboBoxCampos_Exit | Private | 2304 | — |
| ComboBoxCamposFilial_Enter | Private | 2311 | — |
| ComboBoxCamposFilial_Exit | Private | 2318 | — |
| TextBoxFiltroFilial_Enter | Private | 2325 | — |
| TextBoxFiltroFilial_Exit | Private | 2332 | — |
| TextBoxFiltro_Enter | Private | 2339 | — |
| TextBoxFiltro_Exit | Private | 2346 | — |
| ListBoxLista_Enter | Private | 2353 | — |
| ListBoxLista_Exit | Private | 2358 | — |
| ListBoxListaFilial_Enter | Private | 2363 | — |
| ListBoxListaFilial_Exit | Private | 2368 | — |
| TextBoxFiltro_Change | Private | 2374 | Cad_NFE_Saida.PreencheLista |
| PreencheCampos | Private | 2380 | — |
| PreencheCabecalho | Private | 2396 | — |
| PreencheLista | Private | 2412 | Cad_NFE_Saida.PreencheCabecalho, Cad_NFE_Saida.Array2DTranspose |
| Array2DTranspose | Public | 2455 | — |
| ListBoxLista_Change | Private | 2483 | — |
| TextBoxFiltroFilial_Change | Private | 2544 | Cad_NFE_Saida.PreencheLista2 |
| PreencheCampos2 | Private | 2550 | — |
| PreencheCabecalho2 | Private | 2566 | — |
| PreencheLista2 | Private | 2582 | Cad_NFE_Saida.PreencheCabecalho2, Cad_NFE_Saida.Array2DTranspose |
| ListBoxListaFilial_Change | Private | 2625 | — |
| SóNúmeros | Private | 2685 | — |
| TextBox_ESTOQUE_KeyPress | Private | 2695 | Cad_NFE_Saida.SóNúmeros |
| TextBox_N_Entrada_KeyPress | Private | 2700 | Cad_NFE_Saida.SóNúmeros |
| TextBox_N_Saida_KeyPress | Private | 2705 | Cad_NFE_Saida.SóNúmeros |
| COD_KeyPress | Private | 2710 | Cad_NFE_Saida.SóNúmeros |
| TextBox_Valor_Entrada_KeyPress | Private | 2715 | Cad_NFE_Saida.SóNúmeros |
| TextBox_Valor_Saida_KeyPress | Private | 2720 | Cad_NFE_Saida.SóNúmeros |
| TextBox_ESTOQUEFilial_KeyPress | Private | 2725 | Cad_NFE_Saida.SóNúmeros |
| TextBox_N_SaidaFilial_KeyPress | Private | 2730 | Cad_NFE_Saida.SóNúmeros |
| TextBox_N_EntradaFilial_KeyPress | Private | 2735 | Cad_NFE_Saida.SóNúmeros |
| CODFilial_KeyPress | Private | 2740 | Cad_NFE_Saida.SóNúmeros |
| TextBox_Valor_EntradaFilial_KeyPress | Private | 2745 | Cad_NFE_Saida.SóNúmeros |
| TextBox_Valor_SaidaFilial_KeyPress | Private | 2750 | Cad_NFE_Saida.SóNúmeros |

### frmRelatorioXML

Fonte: [frmRelatorioXML.frm](<./_auditoria/20261031_174936Z/desenvolvimento/Form/frmRelatorioXML.frm>). Codificação: windows-1252. Procedimentos: 8.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| CarregarRelatorio | Public | 18 | — |
| CommandButton_Salvar_Click | Private | 93 | frmRelatorioXML.TextoSeguroListBox |
| CommandButton_Salvar_MouseMove | Private | 199 | mdlCursor.MouseCursor |
| TextoSeguroListBox | Private | 204 | — |
| cmdFechar_Click | Private | 214 | — |
| cmdFechar_MouseMove | Private | 219 | mdlCursor.MouseCursor |
| cmdFechar_MouseDown | Private | 224 | — |
| cmdFechar_MouseUp | Private | 229 | — |

### UserFormFiltro

Fonte: [UserFormFiltro.frm](<./_auditoria/20261031_174936Z/desenvolvimento/Form/UserFormFiltro.frm>). Codificação: windows-1252. Procedimentos: 21.

| Procedimento | Escopo | Linha | Chamadas estáticas candidatas |
|---|---|---:|---|
| UserForm_Initialize | Private | 18 | mdlOcultarTítulo.HideTitleBar |
| CommandButtonFechar_Click | Private | 23 | — |
| CommandButtonFiltroCompra_Click | Private | 28 | — |
| CommandButtonFiltroCompra_MouseMove | Private | 35 | mdlCursor.MouseCursor |
| FrameMenu_MouseMove | Private | 41 | — |
| Command_Entrada_Click | Private | 54 | — |
| Command_Entrada_MouseMove | Private | 61 | mdlCursor.MouseCursor |
| CommandButtonFiltroAguardando_Click | Private | 67 | — |
| CommandButtonFiltroAguardando_MouseMove | Private | 74 | mdlCursor.MouseCursor |
| CommandButtonRepasse_Click | Private | 80 | — |
| CommandButtonRepasse_MouseMove | Private | 87 | mdlCursor.MouseCursor |
| CommandButtonVenda_Click | Private | 93 | — |
| CommandButtonVenda_MouseMove | Private | 100 | mdlCursor.MouseCursor |
| CommandButtonTransferência_Click | Private | 106 | — |
| CommandButtonTransferência_MouseMove | Private | 113 | mdlCursor.MouseCursor |
| CommandButtonDevolução_Click | Private | 119 | — |
| CommandButtonDevolução_MouseMove | Private | 126 | mdlCursor.MouseCursor |
| CommandButtonMatriz_Click | Private | 132 | — |
| CommandButtonMatriz_MouseMove | Private | 137 | mdlCursor.MouseCursor |
| CommandButtonFilial_Click | Private | 143 | — |
| CommandButtonFilial_MouseMove | Private | 148 | mdlCursor.MouseCursor |

## Outros mecanismos e limites

As cinco definições FRM apontam para seus respectivos FRX, preservados por hash. Os nomes usados nas correções foram confirmados nos streams de controles dos recursos. Frame_ImpostoFilial não foi encontrado; não houve criação de substituto. Não foi feita validação visual de abertura dos UserForms no Excel nesta execução.

A macro ChamarTabCan aponta para Cancelado, Estorno e Devolução, ausente no arquivo. ChamarFormDocFora e ChamarFormVeículos são pontos públicos legados que exibem indisponibilidade e foram mantidos. O Ribbon foi identificado no pacote; outras associações de botões, eventos e UDFs exigem inspeção no Office antes de remoção.

Chamadas explícitas: 18 instruções Call foram inspecionadas; todos os nomes simples possuem definição local ou declaração de API. Membros qualificados e chamadas dinâmicas continuam sujeitos à validação do host. Evidência: [chamadas-explicitas.json](<./_auditoria/20261031_174936Z/chamadas-explicitas.json>).


## Integração efetiva da Fase 4 — 01/11/2026

A fixture salva contém 55 componentes reais: 30 módulos padrão, duas classes, cinco UserForms, 17 documentos de planilha e o documento EstaPastaDeTrabalho. Foram importados 34 componentes do sistema e três da suíte; todos os 37 corpos correspondem ao pacote final sob normalização documentada, além dos 18 corpos de documento. [Inventário por componente](./_auditoria/fase4_20261031_230457Z/inventario-integracao.json), [comparação salva](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/fontes-integradas-entrega.json).

| Aba | CodeName real | Correspondência ao esquema |
|---|---|---|
| Menu Principal | Planilha12 | PASS |
| Entradas | Planilha1 | PASS |
| Transferidos | Planilha13 | PASS |
| Entradas Anual | Planilha3 | PASS |
| Demonstração | Planilha4 | PASS |
| Saidas | Planilha2 | PASS |
| Todas Saidas | Planilha17 | PASS |
| Entradas Filial | Planilha15 | PASS |
| Entradas Anual Filial | Planilha10 | PASS |
| Saidas Filial | Planilha16 | PASS |
| Todas Saidas Filial | Planilha18 | PASS |
| Carros PF | Planilha7 | PASS |
| Clientes | Planilha6 | PASS |
| Veículos | Planilha9 | PASS |
| Doc. para fora | Planilha11 | PASS |
| NCM e CEST | Planilha5 | PASS |
| Listas | Planilha8 | PASS |

Documento da pasta: EstaPastaDeTrabalho. Planilha5 recebeu o corpo Option Explicit faltante; Planilha15 preserva o evento existente sem criar controle novo. Cinco formulários importados: Cad_NCM, Cad_NFE, Cad_NFE_Saida, frmRelatorioXML e UserFormFiltro. As duas classes são clsOperacaoPlanilha e AuditoriaListaFake. A função de teste AuditoriaFase4_IntegridadeComplementar foi adicionada ao modTestesFase2 existente, sem criar módulo de produção. A única alteração nova no sistema é a retirada do ClearContents redundante em RestaurarCelula; contratos de eventos, controles e regras fiscais preservados. [Fontes finais](./_auditoria/fase4_20261031_230457Z/desenvolvimento/).

As relações de baixa e vínculo anual continuam propostas no [plano de identidade](./PLANO_IDENTIDADE_REGISTROS.md); nenhuma identidade persistida nova foi implantada. A compilação real não prova execução das integrações Selenium/Access/Word ou todos os controles legados.


## Continuação — Fase 5 (20261101_202114Z)

[Relatório dos testes integrados reais](./RELATORIO_TESTES_INTEGRADOS_FASE5.md), [matriz e evidências](./MATRIZ_REGRESSAO_FASE5.md), [mapa A08/A09](./MAPA_INTEGRIDADE_ESTOQUE.md), [recuperação](./PLANO_RECUPERACAO_FALHAS.md). Base sintética compilada no VBE com 56 componentes/17 planilhas. Eventos de negócio: 20 PASS, cinco FAIL, um BLOQUEADO para injeção de Save; 13 casos XML PASS, suítes auxiliares reais 377+11 asserções PASS e classe com duas pastas PASS. Ciclo de vida BLOQUEADO; recuperação incompleta anterior permanece FAIL. Venda/Transferência/Devolução não preservaram integridade após a baixa. Originais e fontes herdados intactos; nenhum dos sete Delete ou regra fiscal modificado. D01/D02 pendentes e sistema NÃO HOMOLOGADO. Resultados anteriores preservados; consultar matriz para modalidades e testes não repetidos.


## Extensão experimental da Fase 6

O produto conserva seus componentes. O mestre experimental adiciona modPrototiposFase6 (57 componentes totais); fixtures criam cinco abas sintéticas, totalizando 22 abas/62 componentes, verificados em arquivo salvo. Não existe reserva aprovada de U:X ou novas abas no produto. [Mapa nominal de 57 procedimentos candidatos/contratos, 17 abas e fórmulas](./MATRIZ_IMPACTO_BAIXA_FASE6.md); [contrato de identidade e migração](./ARQUITETURA_IDENTIDADE_FASE6.md). O mapa distingue pesquisa rápida que atravessa lacunas dos filtros contíguos e dos históricos sem identidade inequívoca.


## Continuidade — Fase 7 (20261102_025115Z)

Mapa dirigido reconfirmado de 57 procedimentos, contratos de leitor, fórmulas/históricos e 17 abas. Mestre experimental F7 com 58 componentes: 57 herdados sem alteração funcional canônica e um módulo de teste; nenhuma nova estrutura operacional no principal. [Relatório](./RELATORIO_TECNICO_FASE7.md), [matriz](./MATRIZ_REGRESSAO_FASE7.md), [contrato](./CONTRATO_IDENTIDADE_FASE7.md), [mapa](./MAPA_MIGRACAO_LEITORES_FASE7.md) e [decisões](./PACOTE_DECISOES_FASE7.md). Checkpoints e evidências em [execução F7](./_auditoria/fase7_20261102_025115Z/). Portão2 não aprovado; NÃO HOMOLOGADO.


## Fase 8 — 02/11/2026: preparação experimental do Portão 2

Execução [_auditoria/fase8_20261102_121422Z](./_auditoria/fase8_20261102_121422Z/). Base anterior:2.106 hashes conferidos sem divergência na abertura e35 backups novos verificados. Originais e sete Delete operacionais preservados. Mestre experimental final:60 componentes/17 abas, compilação real no Excel16.0/20527/64bits; fontes herdados sem alteração de literais/corpos. Portão2 não executado; sistema NÃO HOMOLOGADO.

Quatro defeitos F7 reproduzidos e corrigidos somente nos protótipos: data/selo persistente com política explícita, anual por episódio, consulta por ID após sort/lacuna/reuso e auxiliares com valores discriminantes. Guarda adicional comprovada: recusar selo ERRO coincidente; três FAIL preservados e três PASS após correção.70 asserções únicas aprovadas na seleção de resultados experimentais e4 volumes medidos (100/1000/5000/10000), com limites por revisão na matriz.26 integrados repetidos após guarda final;15 guardas F7 reexecutados em sessão fria. Não houve teste novo de recuperação integral ou aprovação de gravadores reais.

[Relatório técnico](./RELATORIO_TECNICO_FASE8.md), [matriz](./MATRIZ_REGRESSAO_FASE8.md), [correções](./CORRECOES_PROTOTIPOS_FASE8.md), [armazenamento](./PROPOSTA_ARMAZENAMENTO_FASE8.md), [persistência](./CONTRATO_PERSISTENCIA_FASE8.md), [M0–M9](./PLANO_IMPLANTACAO_FASE8.md), [decisões Portão2](./PACOTE_APROVACAO_PORTAO2.md). Cinco FAIL oficiais F5 e FAIL integrais F6 continuam pendentes. Nenhuma promoção, migração real, alteração fiscal, instalação ou mudança de segurança.


## Fase 9 — 20261102_142753Z — persistência experimental

Portão experimental somente; Portão 2 NÃO aprovado e sistema NÃO HOMOLOGADO. Entrega F8 R2 e 2.268 hashes conferidos; 43 backups novos verificados. ACE 16/ADODB-ADOX x64 do Office instalado executou transações reais em bancos sintéticos no TEMP. Mestre F9 compilado no VBE: 63 componentes/17 abas, com os 60 herdados preservados e os sete Delete intactos. Nenhuma integração no principal, migração real, regra fiscal alterada, instalação, segurança modificada, commit/push ou promoção.

65 asserções VBA: 64 PASS e 1 FAIL de Ignore após Sort, preservado com prova corretiva separada. Quatro benchmarks VBA PASS e três testes COM entre processos PASS, sem somar sondagens/hashes. Ciclo de sete movimentos/recibos, rollback pré-commit, bloqueio de pendência, escritor/fencing e reconstrução em sessão fria exercitados. Erro real de SaveAs e cancelamento BeforeSave preservaram a autoridade. Falha física de commit, energia/crash/disco cheio, rollback incompleto e snapshot sob escrita concorrente NÃO EXECUTADOS. Os sete FAIL oficiais F5/F6 permanecem nas classificações anteriores.

[Relatório](./RELATORIO_TECNICO_FASE9.md), [matriz](./MATRIZ_REGRESSAO_FASE9.md), [armazenamento](./COMPARATIVO_ARMAZENAMENTO_FASE9.md), [esquema](./ESQUEMA_TRANSACIONAL_FASE9.md), [durabilidade](./TESTES_DURABILIDADE_FASE9.md), [concorrência](./TESTES_CONCORRENCIA_FASE9.md), [reconstrução](./PROTOCOLO_RECONSTRUCAO_EXCEL_FASE9.md), [M0–M4](./PLANO_MIGRACAO_M0_M4_FASE9.md), [decisões](./DECISOES_PENDENTES_FASE9.md).


## Adendo Fase10 — componentes experimentais

A [fixture F10](./_auditoria/fase10_20261102_155920Z/entrega/Fixture_Integridade_Fase10_EXPERIMENTAL_NAO_HOMOLOGADA.xlsm) tem66 componentes/17 planilhas, acrescentando três módulos à referência63 da F9: modPersistenciaFase10 (arquivo externo modLaboratorioFase10.bas), modCenariosFase10 e modFronteirasFase10. Somam87 procedimentos novos de laboratório; não substituem os57 leitores ou gravadores do produto.

Dependências: ADO/ADOX/DAO ACE16, CryptoAPI local e MSForms; Audit10 chama AuditBatch10 e PreparedAudit10; BuildFull10/VerifyFull10 usam SnapshotFull10; ConfirmAck10 reabre o arquivo e registra confirmação na autoridade. [Catálogo e relações](./ESQUEMA_CONSOLIDADO_FASE10.md), [inventário de procedimentos](./_auditoria/fase10_20261102_155920Z/evidencias/procedimentos_f10.json). A reserialização interna de designers não foi apresentada como igualdade binária dos FRX.


## Candidata Entrega 1 — mapa incremental

[Fontes extraídos](./_desenvolvimento/entrega1_20261102T182823Z/fontes/) / [inventário](./_desenvolvimento/entrega1_20261102T182823Z/evidencias/inventario_integracao.json). 69 componentes: 66 herdados mais modCandidatoIntegrado, clsTesteCandidato e modTestesEntrega1. As 17 planilhas e seus CodeNames foram preservados; assinaturas dos cinco formulários comparadas. Streams de designer reserializados pelo Excel não são byte a byte iguais; regressão completa de controles permanece pendente.

Cad_NFE/seleção → MóduloPesquisa e modPreencherFormulario → modCandidatoIntegrado → Prepare10/Apply10 → ACE esquema 5 → SheetProjections → CandidatoReconstruir/CampoCandidato → oito abas de estoque/históricos. modCenariosFase10 fornece auditoria/ACK. modValidacaoXML mantém regras e recebe somente provedor local de arquivo para teste. EstaPastaDeTrabalho carrega cache na abertura.

SheetProjections acrescenta ViewID por destino sobre ProjectionID/EpisodeID/MovementID; CandidateSettings vincula BaseID, WorkbookID e AuthorityPath. Não há segunda autoridade em células. Leitores incompatíveis e alteração/exclusão manual não foram habilitados. Consulte [entrega](./ENTREGA_1_CANDIDATO_INTEGRADO.md) para alcance e [defeitos](./REGISTRO_DEFEITOS_ABERTOS.md) para migração restante.

## 02/11/2026 — continuação da Entrega 1: reabertura real

Checkpoint verificado em ./_desenvolvimento/entrega1_20261102T182823Z/continuacao_reabertura_20261102T205915Z/. O usuário habilitou macros manualmente. E1_SessaoReaberta revelou FAIL por diferença entre aliases MARCOS~1/MARCOSALEXANDREPERDO do mesmo arquivo. Corrigido somente modCandidatoIntegrado.ConnC1 mediante C1_MesmoArquivoAutoridade (existência + ShortPath); BaseID e vínculo persistido preservados. O Editor VBA recompilou sem erro. Três verificações imediatas retornaram True/False/False (mesmo arquivo/cópia distinta/ausente). Nova execução na sessão aberta PASS e nova instância Excel /x PID 50212 confirmou CACHE_AUDITADO automático e E1_SessaoReaberta PASS, sem repetir movimentos.

Log acumulado 47 execuções de asserção: 40 PASS/7 FAIL; falhas anteriores preservadas. As três consultas imediatas e a compilação são evidências separadas, não somadas como cenários de UserForm. Banco SHA-256 inalterado (3198e07699f644f64c325a9fb517470b64950aa5a262cc354629ed61f1503dc3), revisão 8, oito movimentos/recibos, três ativos, zero Pending. XLSM corrigido SHA-256 5b2a0d55fcea077dd034ace4c4892bbc9fdd46c69760d7d318d2162f54771749. Contagem binária 69 componentes; corpos dos cinco forms iguais ao checkpoint (VB_Base reserializado). 39 originais e backup F10 novamente conferidos, sem divergência. E1-04 encerrado no escopo de reabertura normal; E1-03, conciliação, fluxo D e homologação permanecem pendentes. Nenhuma mudança de segurança, produção, regra fiscal, commit ou publicação.


## Continuidade — Entrega2 (02/11/2026)

Implementação dirigida na candidata sintética, sem nova auditoria geral. [Relatório](./ENTREGA_2_OPERACOES_CORRIGIDAS.md), [matriz de resultados reais](./MATRIZ_REGRESSAO_ENTREGA2.md), [fontes/inventário](./_desenvolvimento/entrega2_20261102T215146Z/fontes/), [revisão dos57 leitores](./_desenvolvimento/entrega2_20261102T215146Z/fontes/REVISAO_LEITORES_ENTREGA2.md) e [manifestos/evidências](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/).

69componentes/17abas/5forms;436assinaturas existentes dos formulários preservadas. Retificação auditável, reentrada por controles e classificação de exclusão integradas; oito componentes modificados funcionalmente. Revisão16/esquema6,16movimentos/recibos e zeroPending no mestre final. Forenses de falhas separados. ReferênciaR0:10PASS; continuação61execuções(54PASS/7FAIL preservados),55IDs. Compilação real final PASS, reabertura interativa E2 BLOQUEADA no aviso de macros. Análise estática/consultas ADO não substituem eventos. Sistema NÃO HOMOLOGADO; [pendências](./REGISTRO_DEFEITOS_ABERTOS.md) e [manual](./MANUAL_INSTALACAO_RESTAURACAO.md) prevalecem para uso da candidata. Originais e Entrega1 intactos.
