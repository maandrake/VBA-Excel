# Plano de correção — fase 2

Execução: 20261031_201621Z. Continuação da auditoria anterior, sem promoção para produção.

## Base confirmada

39 originais e 39 backups conferidos por SHA-256; 50 componentes, 17 planilhas e 648 procedimentos recontados no XLSM. Os cinco fontes modificados reproduzem os 16 blocos registrados de A01–A07; não foram detectadas alterações posteriores. Foram preservados os 155 arquivos da execução anterior e criado novo backup de 40 arquivos de desenvolvimento e seis documentos. Evidência: [manifesto inicial](./_auditoria/fase2_20261031_201621Z/manifesto-inicial-fase2.json) e [verificação da fase 1](./_auditoria/fase2_20261031_201621Z/verificacao-fase1.json).

## A08 — decisão requerida sobre a baixa

Os sete pontos Delete foram reproduzidos por operações COM em tabelas fictícias com sentinelas, apoiadas no mapeamento das fórmulas reais. Exclusões parciais produziram referências quebradas e desalinhamento; EntireRow.Delete também removeu uma sentinela lateral independente. Não se propõe substituir todos os intervalos por linha inteira.

| Modelo testado | Benefício observado | Limitação observada |
|---|---|---|
| Compactar campos, sem Delete | Fórmulas de linha e chave preservadas no protótipo | Uma referência direta a A4 passou a apontar para outro registro, sem erro visível. Exige migração das dependências por chave antes de integrar. |
| Limpar apenas campos, sem mover linhas | Fórmulas e referências das demais linhas preservadas | Deixa lacuna; filtros atuais param cedo. Requer adaptar pesquisas e definir comportamento de reutilização das linhas e referências ao registro removido. |

Proposta para aprovação: desenvolver a **baixa sem deslocar linhas físicas**, limpando somente os campos de dados efetivamente gravados e preservando fórmulas, junto com pesquisas que ignorem lacunas. Antes de integrar, testar reutilização da primeira vaga e referências ao registro baixado; se isso puder associar outra nota a uma referência antiga, a integração deve permanecer bloqueada até definir vínculo por chave. Não criar campo/status fiscal ou controle novo para contornar essa decisão. A política de manter/excluir o histórico anual também precisa permanecer igual à operação original ou receber aprovação explícita.

Evidências: [sete exclusões](./_auditoria/fase2_20261031_201621Z/testes-a08-com.json), [protótipos finais](./_auditoria/fase2_20261031_201621Z/testes-a08-prototipos-v2-com.json), [fórmulas reais](./_auditoria/fase2_20261031_201621Z/formulas-a08.json). As primeiras preparações de protótipo falharam no marshaling COM e foram conservadas como evidência; não são PASS de integridade.

**A08: proposta testada parcialmente por operações equivalentes; implementação depende de decisão. Os sete Delete continuam inalterados nesta etapa.**

## A09 — integridade operacional

Pré-validar conversões, origem, destinos, proteção e capacidade antes de qualquer escrita. Capturar os destinos planejados por célula (valor, fórmula, tipo e formato) e compensar falhas controladas **apenas antes de iniciar Delete**. Depois de baixa estrutural ou confirmação/salvamento, não prometer restauração por esse mecanismo; apontar necessidade de conciliação. Preservar sequência e regras de Venda, Transferência, Devolução e Cancelado, sem criar unicidade histórica nova. A edição anual que depende da célula ativa exige esclarecer a identificação do registro histórico.

O mecanismo implementado é limitado a compensação em memória, sem ACID ou recuperação garantida depois de encerramento inesperado do Excel. Testes devem injetar falhas em cada etapa controlável e comparar valores, fórmulas, formatos, ausência de duplicação, nova tentativa e estado de Application.


## Situação de entrega

Os 20 grupos originais permanecem classificados em 2 críticos, 7 altos, 7 médios e 4 baixos. Há sete correções anteriores preservadas, cinco grupos com novos ajustes técnicos em fonte (A10/A14/A15/A18/A19), cinco com mitigação parcial (A09/A11/A12/A16/A17) e três dependentes de decisão/proposta (A08/A13/A20). Nenhum grupo recebeu aprovação funcional VBA nesta fase.

A cópia tem 34 fontes e 665 procedimentos exportados; os 647 contratos existentes foram preservados. Depois de incorporar fontes e documentos, a projeção seria de 52 componentes/666 procedimentos, sem módulos de teste. O XLSM disponível continua com 50/648 e não incorpora a revisão.

Compilação real: **BLOQUEADO**, erro Office 0x800A03EC, zero módulos importados. [Roteiro manual](./_auditoria/fase2_20261031_201621Z/COMPILACAO_MANUAL.md). [Testes de recuperação](./_auditoria/fase2_20261031_201621Z/testes-recuperacao-reprodutivel-v2.json): 70 PASS por COM equivalente, antes de Delete/Save, sem executar o VBA. A08 e A09 continuam críticos, sem homologação.

## Tratamento por identificador

### A01 — Símbolo celula não declarado na transferência

- **Gravidade / situação:** Alta; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [Cad_NFE_Saida.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE_Saida.frm>); CommandButton_Gravar_Click.
- **Causa raiz:** A única ocorrência executável usa celula.Errors, sem declaração. Option Explicit está ativo.
- **Solução e comportamento esperado:** Destino corrigido para .Cells(linhaDemo, 12).Errors; correção anterior preservada.
- **Risco de regressão e limites:** Correção local; evento VBA bloqueado.
- **Testes necessários/observados:** Transferência e conferência da célula da chave.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/testes-estaticos-fase2.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A02 — Frame_ImpostoFilial ausente no formulário

- **Gravidade / situação:** Alta; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [Cad_NFE_Saida.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE_Saida.frm>); Eventos Frame_ImpostoFilial e controles de impostos.
- **Causa raiz:** Quatro instruções acessam esse nome; ele não aparece nos streams de controles do FRX. Não há Controls.Add nem chamada dinâmica desse nome nos fontes.
- **Solução e comportamento esperado:** Quatro acessos ao controle ausente continuam comentados; handlers e designer preservados.
- **Risco de regressão e limites:** Comportamento visual ainda depende de regressão real.
- **Testes necessários/observados:** Abrir formulário, percorrer eventos e conferir aparência e impostos.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/verificacao-fase1.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A03 — Chave NF-e sem garantia explícita de texto

- **Gravidade / situação:** Alta; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [Cad_NFE.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE.frm>); GravarFilial em Cad_NFE e transferência em Cad_NFE_Saida.
- **Causa raiz:** Cadastro filial e transferência em Cad_NFE_Saida gravam L sem definir texto localmente. Teste COM em formato padrão converteu a chave fictícia para Double.
- **Solução e comportamento esperado:** Texto antes de String preservado; pré-validação aceita vazio legado e exige 44 dígitos quando preenchido.
- **Risco de regressão e limites:** Obrigatoriedade fiscal não alterada; evento VBA bloqueado.
- **Testes necessários/observados:** Chaves com zeros, inválidas, salvar/reabrir; equivalente COM passou.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/regressao-excel-com.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A04 — Leitura de ListBox sem linha selecionada

- **Gravidade / situação:** Média; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [Cad_NCM.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NCM.frm>); ListBoxLista_Click e seleção dos três filtros.
- **Causa raiz:** NCM lê List(-1,...). Nas saídas, ListCount > 0 não garante ListIndex >= 0.
- **Solução e comportamento esperado:** Guardas ListIndex < 0 preservadas antes de acessar listas ou usar a coluna.
- **Risco de regressão e limites:** Inspeção/JS aprovados; cliques dos formulários não executados.
- **Testes necessários/observados:** Sem seleção, lista vazia, múltiplos resultados e seleção válida.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/testes-estaticos-fase2.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A05 — Redimensionamento repetido por coluna em três filtros

- **Gravidade / situação:** Baixa; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [Cad_NCM.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NCM.frm>); PreencheLista dos filtros NCM e saídas.
- **Causa raiz:** PreencheLista do NCM e PreencheLista/PreencheLista2 de saídas executam ReDim Preserve para cada célula da linha encontrada.
- **Solução e comportamento esperado:** Uma realocação por registro e colunas em cache, conforme revisão anterior.
- **Risco de regressão e limites:** Sem benchmark VBA; crescimento por registro pode ser quadrático.
- **Testes necessários/observados:** 35 cenários JS anteriores reproduzidos; comparar listas no VBA.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/regressao-fase1/testes-estaticos.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A06 — Ribbon pode salvar outra pasta ativa

- **Gravidade / situação:** Alta; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [MóduloOfficeRibbonEditor.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/MóduloOfficeRibbonEditor.bas>); ChamarSalvar.
- **Causa raiz:** ChamarSalvar usa ActiveWorkbook.Save apesar de pertencer ao Ribbon desta pasta.
- **Solução e comportamento esperado:** ThisWorkbook.Save preservado; equivalente COM salvou A com B ativa e B continuou sem salvar.
- **Risco de regressão e limites:** Callback/Ribbon VBA ainda bloqueados.
- **Testes necessários/observados:** Executar callback real com duas pastas fictícias; conferir disco.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/testes-backup-contexto-v4-com.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A07 — Módulo sem Option Explicit

- **Gravidade / situação:** Baixa; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [modPreencherFormulario.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modPreencherFormulario.bas>); Declarações de modPreencherFormulario.
- **Causa raiz:** Único fonte exportado sem declaração explícita obrigatória.
- **Solução e comportamento esperado:** Option Explicit preservado; diretiva em 34/34 fontes.
- **Risco de regressão e limites:** Diretiva não comprova ausência de erros de declaração.
- **Testes necessários/observados:** Compilar módulo e projeto para conferir declarações e tipos.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/testes-estaticos-fase2.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A08 — Exclusões parciais podem quebrar fórmulas e vínculos

- **Gravidade / situação:** Crítica; PENDENTE DE DECISAO.
- **Fonte revisado / procedimento:** [Cad_NFE.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE.frm>); Sete Delete, conforme tabela no relatório.
- **Causa raiz:** Delete Shift:=xlUp em A:M da filial deixa N:O fora do deslocamento. Há 363 fórmulas em N e 364 em O na pasta original. Reprodução fictícia em Excel produziu =#REF! em N2 e sentinela desalinhada em O2. Há sete pontos Delete, incluindo duas exclusões de linha inteira.
- **Solução e comportamento esperado:** Proposta de baixa sem deslocar linhas e filtros tolerantes a lacunas; não integrada, autorização solicitada sem resposta registrada.
- **Risco de regressão e limites:** CRÍTICO: exclusões atuais quebram referências; troca de baixa exige aprovação.
- **Testes necessários/observados:** Sete reproduções COM, compactação versus limpeza, lacunas, reutilização e vínculos/histórico.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/testes-a08-com.json>).
- **Critério de aprovação:** Aprovar modelo de baixa; demonstrar ausência de referências quebradas/trocadas, lacunas invisíveis e perda de histórico.

### A09 — Gravações em várias abas sem transação ou reversão

- **Gravidade / situação:** Crítica; MITIGACAO PARCIAL EM FONTE.
- **Fonte revisado / procedimento:** [Cad_NFE_Saida.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE_Saida.frm>); Quatro Gravar/GravarFilial; Alterar e ExcluirFilial.
- **Causa raiz:** GravarFilial escreve em Saidas Filial, Todas Saidas Filial, Entradas e Entradas Anual antes de localizar/apagar a entrada. CDbl/CDate podem falhar após gravações parciais. Cad_NFE também converte depois de gravar campos.
- **Solução e comportamento esperado:** Pré-conversão, origem/destinos e snapshots nos quatro gravadores. Compensação em memória antes de Delete/Save; fronteiras bloqueiam restauração automática. Edição anual e exclusão em duas abas pendentes.
- **Risco de regressão e limites:** CRÍTICO: sem ACID; Delete/Save/crash e interface não são restaurados. Conciliar antes de repetir.
- **Testes necessários/observados:** 70 verificações COM equivalentes; suíte VBA preparada para falhas por etapa, estados e nova tentativa.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/testes-recuperacao-reprodutivel-v2.json>).
- **Critério de aprovação:** Demonstrar recuperação no VBA em todas as fronteiras ou aprovar procedimento de conciliação; resolver edição anual e exclusões.

### A10 — Operações dependem da pasta ativa

- **Gravidade / situação:** Alta; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [Cad_NFE.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE.frm>); Formulários e módulos de abertura/navegação.
- **Causa raiz:** 40 linhas com Sheets/Worksheets sem qualificação e oito referências ActiveWorkbook remanescentes. Alguns usos podem ser intencionais, mas há referências em exclusões e salvamentos.
- **Solução e comportamento esperado:** Referências literais Sheets/Worksheets qualificadas pelo proprietário conhecido; três ActiveWorkbook de conexões e navegação contextual preservados.
- **Risco de regressão e limites:** Não houve troca global de ActiveCell/ActiveSheet nem validação dos eventos de navegação.
- **Testes necessários/observados:** Duas pastas abertas; conferir pesquisas, gravações e salvamento. Equivalente de salvamento passou.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/evidencias/estatica-final/analise-estatica.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A11 — Navegação aponta para aba inexistente

- **Gravidade / situação:** Média; MITIGACAO PARCIAL EM FONTE.
- **Fonte revisado / procedimento:** [mdlChamarTabCan.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlChamarTabCan.bas>); ChamarTabCan.
- **Causa raiz:** Cancelado, Estorno e Devolução não existe entre as 17 planilhas. O código não identifica uma substituição inequívoca.
- **Solução e comportamento esperado:** Diagnóstico de erro com nome legado mantido; nenhuma aba substituta criada.
- **Risco de regressão e limites:** Funcionalidade continua indisponível; tratar erro não resolve dependência.
- **Testes necessários/observados:** Responsável confirmar vínculo do botão/Ribbon e destino correto.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/evidencias/estatica-final/analise-estatica.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A12 — Três macros de atualização referem conexões ausentes

- **Gravidade / situação:** Média; MITIGACAO PARCIAL EM FONTE.
- **Fonte revisado / procedimento:** [mdlMacroConexão.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlMacroConexão.bas>); MacroConexão, ConexãoClientes, MacroDocFora.
- **Causa raiz:** O pacote não contém xl/connections.xml. Os módulos de conexão usam Consulta - Tutores, Consulta - Estudantes e Consulta - Documentos para fora.
- **Solução e comportamento esperado:** Captura de ActiveWorkbook e diagnóstico de conexão ausente; nomes e assinaturas mantidos.
- **Risco de regressão e limites:** Dependência externa indefinida; não foram criados bancos/tabelas/conexões.
- **Testes necessários/observados:** Confirmar proprietário e conexões antes de atualização; nenhuma consulta real realizada.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/evidencias/estatica-final/analise-estatica.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A13 — Cancelamento dispensa leitura e validação do XML

- **Gravidade / situação:** Alta; PENDENTE DE DECISAO.
- **Fonte revisado / procedimento:** [Cad_NFE_Saida.frm](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE_Saida.frm>); CommandButton_Gravar_Click, operacaoCancelada.
- **Causa raiz:** operacaoCancelada evita CarregarXMLNFe e ValidarXMLNFe; os demais ramos mantêm gravação. O validador já avisa que XML normal não comprova cancelamento.
- **Solução e comportamento esperado:** Regra preservada: pula XML, grava Saidas/Todas Saidas e não baixa Entradas nesse ramo. Propor separação entre registro interno e comprovação de evento, mediante decisão fiscal.
- **Risco de regressão e limites:** Modelo COM de duas escritas não valida evento fiscal nem autorização.
- **Testes necessários/observados:** Conferir dois históricos, permanência da origem e duplicidade com responsável e dados fictícios.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/operacoes-multiplas.json>).
- **Critério de aprovação:** Responsável definir evidência e efeitos de Cancelado; executar regressão sem mudar regra fiscal implicitamente.

### A14 — Documento XML fornecido diretamente evita verificações do carregador

- **Gravidade / situação:** Média; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [modValidacaoXML.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modValidacaoXML.bas>); ValidarXMLNFe; IdentificacaoXMLValida.
- **Causa raiz:** O parâmetro documentoXML pode entrar sem a exigência de exatamente um infNFe aplicada em CarregarXMLNFe. XPaths globais podem misturar grupos de documentos compostos. O fluxo normal que usa o carregador já limita a cardinalidade.
- **Solução e comportamento esperado:** Documento direto exige um infNFe, Id NFe+44 dígitos, nNF positivo com até nove dígitos e ausência de DTD. Carregador e regras fiscais preservados.
- **Risco de regressão e limites:** Não faz XSD, assinatura digital ou autenticidade; validador VBA bloqueado.
- **Testes necessários/observados:** DOM válido, zero/dois infNFe, Id/nNF inválidos, DTD/malformação; conferir relatório VBA.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/testes-xml-condicoes-com.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A15 — Somatório ignora nós numéricos inválidos

- **Gravidade / situação:** Média; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [modValidacaoXML.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modValidacaoXML.bas>); SomarNosXML; TentarConverterNumeroXML.
- **Causa raiz:** SomarNosXML adiciona somente os valores convertíveis e não sinaliza explicitamente os demais. Um item não selecionado pode ficar fora da soma sem erro próprio.
- **Solução e comportamento esperado:** Nó presente inválido adiciona erro ao relatório; overflow de Val retorna False.
- **Risco de regressão e limites:** Não muda tributos/tolerâncias; modelo COM/.NET não executa Val nem relatório VBA.
- **Testes necessários/observados:** Válidos, inválidos, vazios, overflow e soma de múltiplos itens; confirmar recusa no VBA.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/testes-xml-condicoes-com.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A16 — Subprocesso sem limite de espera

- **Gravidade / situação:** Média; MITIGACAO PARCIAL EM FONTE.
- **Fonte revisado / procedimento:** [modSelTest.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modSelTest.bas>); RunAndCapture.
- **Causa raiz:** RunAndCapture espera execObj.status sair de zero, usando DoEvents sem timeout; StdOut só é drenado depois.
- **Solução e comportamento esperado:** Prazo de 10 s e encerramento somente do WScript.Exec criado, pausa e diagnóstico.
- **Risco de regressão e limites:** Parcial: StdOut só é drenado ao fim; StdErr não é drenado. Pipe cheio pode terminar por timeout.
- **Testes necessários/observados:** Subprocesso fictício lento encerrado por COM; testar VBA com saída grande e erro.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/testes-backup-contexto-v4-com.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A17 — Espera ocupa CPU e permite reentrada por DoEvents

- **Gravidade / situação:** Baixa; MITIGACAO PARCIAL EM FONTE.
- **Fonte revisado / procedimento:** [modSeleniumCommon.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modSeleniumCommon.bas>); SleepMs; SleepMsNFE.
- **Causa raiz:** SleepMs e variantes usam espera ativa. Os cálculos já compensam meia-noite, mas a espera continua ocupando o thread.
- **Solução e comportamento esperado:** Sleep kernel32 em intervalos curtos, declarações VBA7/legada; variante NFE delega ao helper.
- **Risco de regressão e limites:** Parcial: DoEvents/reentrada permanecem; sem runtime x86 ou benchmark.
- **Testes necessários/observados:** VBA: zero, negativo, duração curta, meia-noite, cancelamento e benchmark de CPU.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/evidencias/estatica-final/analise-estatica.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A18 — Backup legado salva a pasta ativa e reutiliza destino

- **Gravidade / situação:** Alta; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [mdlCriaBackup.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlCriaBackup.bas>); SaveWorkbookBackup; SaveWorkbookBackupToFloppy; CriarBackupVersionado.
- **Causa raiz:** As rotinas chamam .Save antes de SaveCopyAs; o primeiro destino é .bak fixo e o segundo D:\nome. Não há versionamento de cada execução nem verificação de conteúdo.
- **Solução e comportamento esperado:** ThisWorkbook, SaveCopyAs sem Save, pasta timestamp/contador exclusiva e cabeçalho/tamanho. Seletor substitui destino fixo D:.
- **Risco de regressão e limites:** Verificação VBA básica sem hash interno; rotina bloqueada. Destino/retenção operacional exigem política.
- **Testes necessários/observados:** Duas versões, outra pasta ativa, SHA-256 externo, restauração, destino inválido e colisão.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/testes-backup-contexto-v4-com.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A19 — Condição xlNotBetween calculada incorretamente

- **Gravidade / situação:** Média; CORRECAO TECNICA EM FONTE; VBA BLOQUEADO.
- **Fonte revisado / procedimento:** [mdlCondCel.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/mdlCondCel.bas>); ActiveCondition; GetStrippedValue.
- **Causa raiz:** A expressão Not(valor <= mínimo) And (valor >= máximo) não representa valores fora do intervalo: falha abaixo do mínimo e inclui o máximo. A avaliação de fórmulas também depende do contexto ativo.
- **Solução e comportamento esperado:** Fora do intervalo usa menor que mínimo OU maior que máximo; leitura de =5, =10 e literais corrigida. CDbl usa limite limpo; Evaluate usa planilha proprietária.
- **Risco de regressão e limites:** Referências em limites, fórmulas relativas, tipos modernos e StopIfTrue não foram generalizados/validados.
- **Testes necessários/observados:** 4/5/7/10/11 contra DisplayFormat nativo e sete leituras passaram no equivalente; executar UDF real.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/testes-backup-contexto-v4-com.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

### A20 — Logs e caminhos operacionais fora da pasta do projeto

- **Gravidade / situação:** Baixa; PENDENTE DE DECISAO.
- **Fonte revisado / procedimento:** [modSeleniumCommon.bas](<./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modSeleniumCommon.bas>); WriteLog; LogNFE; constantes do driver.
- **Causa raiz:** WriteLog usa C:\Temp; o driver tem caminho fixo em Program Files. Logs de consultas podem conter identificadores pessoais.
- **Solução e comportamento esperado:** Proposta: destino configurável autorizado, códigos de eventos, mascaramento e retenção pelo responsável. Nenhum log/consulta real executado.
- **Risco de regressão e limites:** Sem alteração: caminhos fixos e exposição potencial de identificadores permanecem.
- **Testes necessários/observados:** Depois de definir política, testar sanitização fictícia e retenção sem dados reais.
- **Evidência:** [registro](<./_auditoria/fase2_20261031_201621Z/evidencias/estatica-final/analise-estatica.json>).
- **Critério de aprovação:** Importar, compilar e executar os testes descritos na fixture; comparar com base e registrar limitações antes de homologar.

## Reversibilidade e decisões

O [manifesto inicial](./_auditoria/fase2_20261031_201621Z/manifesto-inicial-fase2.json) identifica 46 arquivos no novo backup, conferidos antes das alterações. As [alterações por bloco](./_auditoria/fase2_20261031_201621Z/alteracoes-fase2.json) permitem revisar e reverter os ajustes. Restaurar somente na cópia de desenvolvimento, preservando os pares FRM/FRX.

A08 aguarda a decisão solicitada sobre desenvolver baixa sem deslocar linhas e pesquisas tolerantes a lacunas. Isso não aprova integrar uma estratégia que reutilize linha e troque identidade de referência. A09 exige solução/decisão adicional depois de Delete/Save, fechamento inesperado, edição anual e exclusões relacionadas. A13 requer decisão fiscal separada.

A18: recomenda-se manifesto SHA-256 externo para cada versão fechada, restauração periódica e política aprovada de destino/retenção. Uma opção a avaliar é manter sete versões diárias, quatro semanais e doze mensais; nenhuma retenção/limpeza foi implementada ou executada.

Sem commit, push, publicação ou promoção. Homologação impedida pelas pendências críticas e pela regressão VBA bloqueada.
