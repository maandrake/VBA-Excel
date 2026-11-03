# Matriz de testes — Fase 3

Execução 20261031_213925Z. PASS = executado com resultado esperado na modalidade indicada; FAIL = divergência observada no critério indicado; BLOQUEADO = impedimento técnico/permissão; NÃO EXECUTADO = caso ainda não realizado. Aprovar um oráculo de reprodução de defeito não aprova a integridade da aplicação.

## Verificações desta fase

| Frente | Modalidade | Resultado observado | Evidência |
|---|---|---|---|
| Integridade/contratos | Análise estática | 12 PASS, zero FAIL; 39 originais, execuções anteriores, 69 backups, 665 contratos e preservação dos fontes | [JSON](./_auditoria/fase3_20261031_213925Z/verificacao-estatica-fase3.json) |
| Bibliotecas | Inspeção local somente leitura | Caminhos e registros consultados; IsBroken/ativação não certificados | [JSON](./_auditoria/fase3_20261031_213925Z/bibliotecas-registradas.json) |
| Importar/compilar | Compilação real VBA | BLOQUEADO, 0x800A03EC; zero importações/compilações | [Erro original](./_auditoria/fase3_20261031_213925Z/office-fase3.json) |
| Executar auxiliares/eventos | Execução funcional VBA | BLOQUEADO pela integração pendente; zero macros executadas | [Roteiro manual](./_auditoria/fase3_20261031_213925Z/COMPILACAO_MANUAL.md) |
| A08 | Operações equivalentes COM / mapa em memória | Três estratégias, quatro experiências; defeitos A/B reproduzidos, C apenas protótipo condicionado | [JSON](./_auditoria/fase3_20261031_213925Z/a08-estrategias-com.json) |
| A09 metadado anterior | Operação equivalente COM | FAIL de preservação para Ignore=True; valor/chave continuaram preservados | [Antes](./_auditoria/fase3_20261031_213925Z/a09-metadado-ignore-com.json) |
| A09 metadado corrigido | Operação equivalente COM | Quatro casos PASS (vazio, fórmula, texto, chave), sem executar classe VBA | [Depois](./_auditoria/fase3_20261031_213925Z/a09-metadado-corrigido-com.json) |
| XML | MSXML/RegExp COM + simulação .NET | 12 PASS no parser/modelo; soma/erro de relatório VBA não executados | [JSON](./_auditoria/fase3_20261031_213925Z/xml-fase3-com.json) |
| Backup | Operação equivalente COM | Cinco cenários PASS, incluindo versões/restauração e erros esperados | [JSON](./_auditoria/fase3_20261031_213925Z/backup-fase3-com.json) |

## A09 — campos antes/depois e fronteiras

Os estados de fronteira são modelados em PowerShell; métodos de célula, Delete e Save/SaveAs são COM reais em fixtures. A classe VBA não foi chamada. Cada snapshot compara 200 células de cinco abas (Value2, Formula, HasFormula e NumberFormat) e os quatro estados Application. Metadado Ignore tem ensaio próprio. A comparação não é uma imagem integral de formatos/objetos do workbook.

| Caso executado | Campos recompostos? | Estado Application | Resultado de integridade no escopo | Oráculo do ensaio |
|---|---|---|---|---|
| AntesPrimeira | Sim | Preservado | PASS no equivalente | PASS da observação esperada |
| Durante1 | Sim | Preservado | PASS no equivalente | PASS da observação esperada |
| Durante2 | Sim | Preservado | PASS no equivalente | PASS da observação esperada |
| Durante3 | Sim | Preservado | PASS no equivalente | PASS da observação esperada |
| Durante4 | Sim | Preservado | PASS no equivalente | PASS da observação esperada |
| Entre1e2 | Sim | Preservado | PASS no equivalente | PASS da observação esperada |
| Entre2e3 | Sim | Preservado | PASS no equivalente | PASS da observação esperada |
| Entre3e4 | Sim | Preservado | PASS no equivalente | PASS da observação esperada |
| AntesDelete | Sim | Preservado | PASS no equivalente | PASS da observação esperada |
| DeleteRecusado | Não | Preservado | FAIL de recomposição automática — conciliação necessária | PASS da observação esperada |
| AntesSaveAposDelete | Não | Preservado | FAIL de recomposição automática — conciliação necessária | PASS da observação esperada |
| AntesSaveSemDelete | Sim | Preservado | PASS no equivalente | PASS da observação esperada |
| FalhaSaveAsSemDelete | Não | Preservado | FAIL de recomposição automática — conciliação necessária | PASS da observação esperada |
| ErroDepoisSave | Não | Preservado | FAIL de recomposição automática — conciliação necessária | PASS da observação esperada |
| FalhaNaCompensacao | Não | Preservado | FAIL — recuperação incompleta | PASS da observação esperada |

Em Durante2, uma nova tentativa depois de compensar manteve CountIf = 1 nos quatro destinos do modelo. Todas as experiências mantiveram a outra pasta fictícia intocada. DeleteRecusado usa proteção adicionada somente à fixture depois de escrever; não representa executar uma rotina sobre planilha real protegida. SaveAs para pasta inexistente não equivale ao Save do evento. ErroDepoisSave alterou o arquivo em disco, conforme hash. FalhaNaCompensacao registrou 40 falhas de célula e continuou as outras tentativas.

[Dados completos](./_auditoria/fase3_20261031_213925Z/a09-fronteiras-com.json). As propriedades nulas na apresentação resumida do console foram um problema de Select-Object/hashtable; o JSON foi conferido diretamente e não foi reclassificado.

## Regressão integrada ainda necessária

Todas as linhas abaixo se referem à **execução funcional VBA**, ainda BLOQUEADA. Os resultados COM acima não são transferidos a esses casos.

| Operação | Casos para executar na fixture compilada | Conferência |
|---|---|---|
| Entrada matriz | Válida, N_Nfe duplicada/vazia, txt_Valor/Data inválidos | Entradas A:K e Entradas Anual A:J, sem escrita parcial |
| Entrada filial | Válida, duplicada, chave vazia/inválida/com 44 dígitos e zeros | A:L, anual A:J, M:O, Ignore e chave String |
| Venda | Válida/duplicada, XML divergente, erro antes/dentro/entre escritas | Saidas/Todas Saidas, origem, histórico e fórmulas |
| Transferência | Quatro destinos, chave e dados, fórmulas M:O, baixa | Correspondência NF-e/veículo e recuperação antes das fronteiras |
| Devolução | Origem existente/inexistente, quatro gravações e erro parcial | Origem filial, entradas matriz/anual, histórico e fórmulas |
| Cancelado | Comportamento atual, duplicidade, dois históricos, permanência da entrada | Não exigir nova regra fiscal por este teste |
| Pesquisa rápida | Matriz/filial, sem resultado, múltiplos, lacunas e seleção | Origem/linha/chave e navegação para registro correto |
| ListBox/filtros | Lista vazia, ListIndex=-1, seleção válida, coluna inválida | UserForms reais; filtro legado separado da pesquisa rápida |
| XML/relatório | Completo válido, malformado, zero/dois infNFe, Id/nNF inválidos, itens inválidos/overflow/somas | ValidarXMLNFe, AdicionarResultado e frmRelatorioXML reais |
| Backup | Proprietário, duas pastas, duas versões, restauração, destino inválido e escrita bloqueada | Três rotinas VBA, sem Save inesperado nem retenção automática |
| Ribbon | Outra pasta ativa | Callback salva somente a fixture proprietária; conferir disco |
| Condição/Forms | Limites, tipos usados, abertura/designer/ActiveX | Não generalizar comparação nativa anterior para UDF |
| Recuperação | Proteção/mesclagem/matriz/spill/outro workbook, erro de restauração | Classe real, estado Application e metadado Ignore |
| Edição/exclusão anual | Vínculo inexistente/múltiplo/único aprovado; falha entre duas abas | Não escolher registro por primeira ocorrência/seleção anterior |
| Save da macro | Erro antes/durante/depois do ThisWorkbook.Save | Evidência real; SaveAs equivalente não satisfaz este caso |

A suíte [modTestesFase2.bas](./_auditoria/fase3_20261031_213925Z/scripts/modTestesFase2.bas) foi estendida, sem duplicar uma nova infraestrutura VBA. AuditoriaFase3_CicloVida e a asserção Ignore estão PREPARADAS, NÃO IMPORTADAS/COMPILADAS/EXECUTADAS. A suíte antiga de pesquisa e sua classe fake permanecem disponíveis.

Selenium/Detran/bancos/logs reais e crash forçado do Excel: **NÃO EXECUTADO**. Não são necessários para autorizar operações fictícias locais, e não foram usados como atalhos. Não houve testes sobre documentos/estoques reais.

Critério de homologação continua não satisfeito: importação/compilação e eventos essenciais pendentes, A08/A09 críticos e decisões de identidade/fiscalidade abertas.
