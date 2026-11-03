# Plano de correção — fase 2

Execução: 20261031_201621Z. Continuação da auditoria anterior, sem promoção para produção.

## Base confirmada

39 originais e 39 backups conferidos por SHA-256; 50 componentes, 17 planilhas e 648 procedimentos recontados no XLSM. Os cinco fontes modificados reproduzem os 16 blocos registrados de A01–A07; não foram detectadas alterações posteriores. Foram preservados os 155 arquivos da execução anterior e criado novo backup de 40 arquivos de desenvolvimento e seis documentos. Evidência: [manifesto inicial](./_auditoria/fase2_20261031_201621Z/manifesto-inicial-fase2.json) e [verificação da fase 1](./_auditoria/fase2_20261031_201621Z/verificacao-fase1.json).

## A08 — decisão requerida sobre a baixa

Os sete Delete foram reproduzidos por COM em arquivos novos fictícios. Exclusões parciais produziram referências quebradas e desalinhamento; EntireRow.Delete também removeu uma sentinela lateral independente. Não se propõe substituir todos os intervalos por linha inteira.

| Modelo testado | Benefício observado | Limitação observada |
|---|---|---|
| Compactar campos, sem Delete | Fórmulas de linha e chave preservadas no protótipo | Uma referência direta a A4 passou a apontar para outro registro, sem erro visível. Exige migração das dependências por chave antes de integrar. |
| Limpar apenas campos, sem mover linhas | Fórmulas e referências das demais linhas preservadas | Deixa lacuna; filtros atuais param cedo. Requer adaptar pesquisas e definir comportamento de reutilização das linhas e referências ao registro removido. |

Proposta para aprovação: desenvolver a **baixa sem deslocar linhas físicas**, limpando somente os campos de dados efetivamente gravados e preservando fórmulas, junto com pesquisas que ignorem lacunas. Antes de integrar, testar reutilização da primeira vaga e referências ao registro baixado; se isso puder associar outra nota a uma referência antiga, a integração deve permanecer bloqueada até definir vínculo por chave. Não criar campo/status fiscal ou controle novo para contornar essa decisão. A política de manter/excluir o histórico anual também precisa permanecer igual à operação original ou receber aprovação explícita.

Evidências: [sete exclusões](./_auditoria/fase2_20261031_201621Z/testes-a08-com.json), [protótipos finais](./_auditoria/fase2_20261031_201621Z/testes-a08-prototipos-v2-com.json), [fórmulas reais](./_auditoria/fase2_20261031_201621Z/formulas-a08.json). As primeiras preparações de protótipo falharam no marshaling COM e foram conservadas como evidência; não são PASS de integridade.

**A08: proposta testada parcialmente por operações equivalentes; implementação depende de decisão. Os sete Delete continuam inalterados nesta etapa.**

## A09 — integridade operacional

Pré-validar conversões, origem, destinos, proteção e capacidade antes de qualquer escrita. Capturar os destinos planejados por célula (valor, fórmula, tipo e formato) e compensar falhas controladas **apenas antes de iniciar Delete**. Depois de baixa estrutural ou confirmação/salvamento, não prometer restauração por esse mecanismo; apontar necessidade de conciliação. Preservar sequência e regras de Venda, Transferência, Devolução e Cancelado, sem criar unicidade histórica nova. A edição anual que depende da célula ativa exige esclarecer a identificação do registro histórico.

O mecanismo será limitado a compensação em memória, sem ACID ou recuperação garantida depois de encerramento inesperado do Excel. Testes devem injetar falhas em cada etapa controlável e comparar valores, fórmulas, formatos, ausência de duplicação, nova tentativa e estado de Application.

## Matriz A01–A20

| ID | Gravidade | Problema / causa | Arquivo original | Solução e situação inicial | Risco / teste e critério |
|---|---|---|---|---|---|
| A01 | Alta | Símbolo celula não declarado na transferência | [Cad_NFE_Saida.frm](./Form/Cad_NFE_Saida.frm) | Correção anterior confirmada em fonte; regressão VBA pendente. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A02 | Alta | Frame_ImpostoFilial ausente no formulário | [Cad_NFE_Saida.frm](./Form/Cad_NFE_Saida.frm) | Correção anterior confirmada em fonte; regressão VBA pendente. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A03 | Alta | Chave NF-e sem garantia explícita de texto | [Cad_NFE.frm](./Form/Cad_NFE.frm) | Correção anterior confirmada em fonte; regressão VBA pendente. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A04 | Média | Leitura de ListBox sem linha selecionada | [Cad_NCM.frm](./Form/Cad_NCM.frm) | Correção anterior confirmada em fonte; regressão VBA pendente. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A05 | Baixa | Redimensionamento repetido por coluna em três filtros | [Cad_NCM.frm](./Form/Cad_NCM.frm) | Correção anterior confirmada em fonte; regressão VBA pendente. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A06 | Alta | Ribbon pode salvar outra pasta ativa | [MóduloOfficeRibbonEditor.bas](./MDL/MóduloOfficeRibbonEditor.bas) | Correção anterior confirmada em fonte; regressão VBA pendente. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A07 | Baixa | Módulo sem Option Explicit | [modPreencherFormulario.bas](./MDL/modPreencherFormulario.bas) | Correção anterior confirmada em fonte; regressão VBA pendente. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A08 | Crítica | Exclusões parciais podem quebrar fórmulas e vínculos | [Cad_NFE.frm](./Form/Cad_NFE.frm) | Aguardando decisão da baixa. | Fórmulas, dependências, lacunas, reutilização e aprovação do responsável. |
| A09 | Crítica | Gravações em várias abas sem transação ou reversão | [Cad_NFE_Saida.frm](./Form/Cad_NFE_Saida.frm) | Pré-validação e compensação limitada em desenvolvimento. | Falhas por etapa; recuperação antes de Delete; limitações explícitas. |
| A10 | Alta | Operações dependem da pasta ativa | [Cad_NFE.frm](./Form/Cad_NFE.frm) | Análise e correção técnica incremental, quando reproduzível. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A11 | Média | Navegação aponta para aba inexistente | [mdlChamarTabCan.bas](./MDL/mdlChamarTabCan.bas) | Análise e correção técnica incremental, quando reproduzível. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A12 | Média | Três macros de atualização referem conexões ausentes | [mdlMacroConexão.bas](./MDL/mdlMacroConexão.bas) | Análise e correção técnica incremental, quando reproduzível. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A13 | Alta | Cancelamento dispensa leitura e validação do XML | [Cad_NFE_Saida.frm](./Form/Cad_NFE_Saida.frm) | Análise e correção técnica incremental, quando reproduzível. | Preservar cancelamento; decisão fiscal separada. |
| A14 | Média | Documento XML fornecido diretamente evita verificações do carregador | [modValidacaoXML.bas](./MDL/modValidacaoXML.bas) | Análise e correção técnica incremental, quando reproduzível. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A15 | Média | Somatório ignora nós numéricos inválidos | [modValidacaoXML.bas](./MDL/modValidacaoXML.bas) | Análise e correção técnica incremental, quando reproduzível. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A16 | Média | Subprocesso sem limite de espera | [modSelTest.bas](./MDL/modSelTest.bas) | Análise e correção técnica incremental, quando reproduzível. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A17 | Baixa | Espera ocupa CPU e permite reentrada por DoEvents | [modSeleniumCommon.bas](./MDL/modSeleniumCommon.bas) | Análise e correção técnica incremental, quando reproduzível. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A18 | Alta | Backup legado salva a pasta ativa e reutiliza destino | [mdlCriaBackup.bas](./MDL/mdlCriaBackup.bas) | Análise e correção técnica incremental, quando reproduzível. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A19 | Média | Condição xlNotBetween calculada incorretamente | [mdlCondCel.bas](./MDL/mdlCondCel.bas) | Análise e correção técnica incremental, quando reproduzível. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |
| A20 | Baixa | Logs e caminhos operacionais fora da pasta do projeto | [modSeleniumCommon.bas](./MDL/modSeleniumCommon.bas) | Análise e correção técnica incremental, quando reproduzível. | Comparação de contratos, testes focados e compilação/execução VBA pendentes. |

Cada alteração terá diff textual, manifesto de hashes e evidência própria. A contagem de problemas fechados será atualizada ao final, sem confundir fonte corrigido, teste COM equivalente e macro executada.

## Compilação

Excel 16.0 build 20527, x64, confirmado pela imagem PE do executável instalado. O Office recusou VBProject com mensagem original “O acesso de programação ao projeto do Visual Basic não é confiável”, HRESULT 0x800A03EC. Nenhum módulo importado e nenhuma macro compilada/executada. Evidência: [probe](./_auditoria/fase2_20261031_201621Z/compilacao-probe.json). Preparar roteiro manual e fixture de dependências sem contornar a proteção.
