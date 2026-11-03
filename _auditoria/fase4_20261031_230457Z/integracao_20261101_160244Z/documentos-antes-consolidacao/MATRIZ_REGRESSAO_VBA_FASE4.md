# Matriz de regressão VBA — Fase 4

Execução 20261031_230457Z. Esta matriz registra o estado efetivo, sem repetir os ensaios COM da Fase 3. **Zero funções VBA executadas nesta fase.** PASS exige execução aprovada na modalidade indicada; FAIL exige execução com divergência; BLOQUEADO indica dependência impeditiva; NÃO EXECUTADO indica trabalho ainda não realizado. Ausência de execução não é PASS nem FAIL.

## Executado nesta fase

| Verificação | Modalidade | Resultado | Evidência |
|---|---|---|---|
| Originais, fontes/backup da Fase 3, pacote, documentos e fixture | Integridade SHA-256 local | PASS: 211 comparações, zero divergências | [Manifesto](./_auditoria/fase4_20261031_230457Z/manifesto-inicial-fase4.json) |
| Cópias para Fase 4 | Cópia de arquivos com conferência SHA-256 | PASS: 114 backups; 39 arquivos de fonte/FRX, três componentes de teste e pacote isolado preservados | [Manifesto](./_auditoria/fase4_20261031_230457Z/manifesto-inicial-fase4.json) |
| Código das suítes e escopo de escrita | Inspeção dirigida | Lidos os três componentes existentes; não alterados nem executados | [Roteiro](./COMPILACAO_VBA_FASE4.md) |
| Importação/compilação | Compilação real VBA | BLOQUEADO: intervenção manual pendente; nenhuma nova sondagem VBProject | [Inventário](./_auditoria/fase4_20261031_230457Z/inventario-integracao.json) |

Esses PASS de arquivo não representam validação funcional. O bloqueio original permanece como evidência histórica; o estado de hoje no VBE não foi presumido nem sondado novamente.

## Primeiro lote — executar antes dos formulários

| ID | Função existente / objetivo | Resultado real Fase 4 | Limites |
|---|---|---|---|
| V4-01 | Auditoria_ExecutarTestesSeguros: pesquisa vazia, múltipla, erro de célula, origem, linha, chave, sem resultado e TruncateForLog | BLOQUEADO — aguardando importação/compilação | Usa AuditoriaListaFake; não certifica MSForms, navegação visual nem grava logs |
| V4-02 | AuditoriaFase2_Executar: classe real, snapshots, falha artificial após escrita de destino, restauração de conteúdo/fórmula/chave/Ignore e Application, repetição e condição | BLOQUEADO — aguardando importação/compilação | Escreve A4:D4 em oito abas; não chama eventos gravadores. Recusa de baixa/Save usa flags, sem Delete/Save real |
| V4-03 | AuditoriaFase3_CicloVida: contexto não iniciado/finalizado e recusa de reutilização | BLOQUEADO — aguardando importação/compilação | Quatro asserções preparadas; nenhuma observada em runtime |

Comparação futura: pesquisa possui nove asserções preparadas; a suíte da classe imprime a contagem efetivamente atingida; a função de ciclo de vida imprime seu resultado. Não inserir números projetados como execução. Guardas exigem TEMP, prefixo Fixture_Auditoria_ e marcador fictício.

## Cobertura A09 a completar no VBA

| Caso | Suíte disponível / próximo mecanismo | Estado efetivo |
|---|---|---|
| Inicialização, contexto ativo/inativo e reutilização | V4-02/V4-03 | BLOQUEADO |
| Fórmula, vazio, texto começando por =, chave e Ignore | V4-02; comparar também tipo String/Len=44/NumberFormat no registro de evidência | BLOQUEADO |
| NumberFormat de cada tipo | A classe captura/restaura; a suíte atual não afirma formato de todos os tipos separadamente. Conferência adicional necessária somente depois de integrar | BLOQUEADO |
| Ordem inversa e destinos sobrepostos | Inspeção anterior confirma loop inverso; prova direta de ordem ainda não coberta pela suíte | BLOQUEADO; instrumentação futura após compilar |
| Pré-conversões inválidas, proteção, mesclagem, matriz/spill e outro workbook | Testes diretos dirigidos dos auxiliares após primeiro lote | BLOQUEADO |
| Falha individual na restauração, demais células e diagnóstico | Fixture protegida/instrumentação controlada, preservando estado anterior | BLOQUEADO; não reclassificar o resultado COM incompleto |
| Duas pastas abertas | Conferir proprietário/destinos e estados de ambas no VBA | BLOQUEADO |
| Gravadores e falhas antes de Delete/Save | Interface/instrumentação exclusiva da fixture, sem renomear eventos | BLOQUEADO: compilação e configuração operacional pendentes |
| Alterar/anual, ExcluirFilial, baixa de transferência/devolução | Relação de identidade aprovada e conciliação | BLOQUEADO também por D01/D02 |
| Depois de Delete/Save/crash | Sem compensação automática aprovada | NÃO EXECUTADO; quatro conciliações e uma recuperação incompleta históricas permanecem |

## Regressão integrada posterior

| Operação | Casos mínimos / resultado a conferir | Modalidade e estado atual |
|---|---|---|
| Entrada matriz/filial | Válida, duplicada, campos/data/valor inválidos, chave 44 como texto, fórmulas e históricos | Execução funcional VBA: BLOQUEADO |
| Venda | Duas gravações, origem/histórico, duplicidade e falhas antes da baixa | Execução funcional VBA: BLOQUEADO |
| Transferência | Saidas, Entradas Filial/Anual Filial, Todas Saidas; M:O, chave e veículo | Execução funcional VBA: BLOQUEADO; baixa depende D01/D02 |
| Devolução | Origem existente/inexistente, quatro gravações, histórico/fórmulas | Execução funcional VBA: BLOQUEADO; baixa depende D01/D02 |
| Cancelado | Duas gravações e permanência da entrada, preservando regra atual | Execução funcional VBA: BLOQUEADO; comprovação fiscal D03 separada |
| Pesquisa/ListBox | Sem resultado, múltiplos, seleção válida/-1, lacuna, navegação e seleção antiga | Execução funcional VBA: BLOQUEADO; fake não substitui controle real |
| XML | Estrutura/Id/nNF/chave/itens/número/overflow/somatórios e frmRelatorioXML real | Execução funcional VBA: BLOQUEADO; fixture XML completa/configuração pendentes |
| Backup | Três rotinas existentes, versões, outra pasta ativa, falha e restauração | Execução funcional VBA: BLOQUEADO; sem política de retenção automática |
| Ribbon | Callback proprietário com outra pasta fictícia ativa; vínculo Ribbon separadamente | Execução funcional VBA: BLOQUEADO |

Selenium, DETRAN, refresh/conexões externas e operações reais: **NÃO EXECUTADO**, fora do escopo permitido para testes desta fase. Nenhuma suíte equivalente nova foi criada. Os casos COM anteriores permanecem em [MATRIZ_TESTES_FASE3](./MATRIZ_TESTES_FASE3.md), sem serem contados novamente.

Para cada execução futura registrar antes/depois, resultado literal, módulo/linha em erro, estados Application, arquivo salvo e revisão exportada. Um relato do operador será identificado como evidência manual fornecida, e não como execução automática observada pelo agente. Critérios de integração e homologação continuam não satisfeitos.


## 01/11/2026 — tentativa assistida pela interface (20261101_143216Z)

O usuário autorizou controle do computador e uso contínuo do Excel em fixtures. Foram preservados nove arquivos/backup da fixture, com hashes conferidos, antes de qualquer edição. [Manifesto](./_auditoria/fase4_20261031_230457Z/interface_20261101_143216Z/manifesto-inicial.json). A tentativa usou a interface Windows e não repetiu o acesso programático ao VBProject.

O controle localizou Excel e o diálogo Abrir. Houve expiração/reconexão, recusa do índice de edição em cache e recusa de clique por divergência entre a janela principal e modal. Após atualização/ativação e tentativa por teclado, o foco permaneceu na Caixa de Pesquisa; o campo Nome do arquivo não recebeu texto. Nenhum VBA foi importado, compilado ou executado. A fixture continua com SHA-256 inicial; isso não comprova compilação. [Erros exatos](./_auditoria/fase4_20261031_230457Z/interface_20261101_143216Z/resultado-interface.json).

**BLOQUEADO por controle da janela modal**, separado do bloqueio VBProject histórico. Foi solicitado somente abrir a fixture já preparada e deixá-la em primeiro plano, para retomar o VBE; não se solicitou nova aprovação nem mudança de segurança. A instância Excel não foi encerrada e nenhuma instância do usuário foi fechada. Continuam NÃO INTEGRADO/NÃO HOMOLOGADO e A08/A09 críticos.


## 01/11/2026 — fixture e VBE observados (20261101_150908Z)

Após a confirmação do usuário, a interface mostrou a fixture com as 17 abas esperadas e o VBE com **18 componentes reais, todos documentos**. Zero módulos padrão, classes ou UserForms revisados presentes. Os 17 CodeNames ainda são os nomes sequenciais da fixture, diferentes do mapeamento exigido; não foram ajustados. A contagem esperada de 55 não foi usada como resultado de importação. [Inventário observado](./_auditoria/fase4_20261031_230457Z/interface_20261101_150908Z/resultado-vbe.json).

Ctrl+M abriu Importar arquivo, e o campo Nome recebeu o caminho conferido de clsOperacaoPlanilha.cls. A primeira entrada tinha escape de barras incorreto; foi corrigida com separadores / antes de qualquer confirmação. O clique em Abrir foi recusado pela ferramenta: `point (1560, 1068) is over EXCEL.EXE "&Abrir", not target window EXCEL.EXE "Microsoft Visual Basic for Applications - Fixture_Auditoria_Fase4.xlsm"`. Enter/Alt+A mantiveram o diálogo aberto. Houve também `coordinate input geometry is unavailable` e timeout de SetValue; a atualização posterior confirmou o caminho correto, mas não a importação.

A consulta de caminho na Janela Imediata foi tentada, porém não houve saída verificável; não é PASS de VBA. Foram observados o mesmo bloqueio de geometria/ativação tanto em sky quanto no mecanismo cua_repl. Não se alterou segurança, não houve novo acesso VBProject e não se executaram consultas externas. **Importação revisada: 0; compilação: BLOQUEADO; suítes VBA: 0 execuções; sistema NÃO HOMOLOGADO.**

O arquivo da fixture em disco continua com o SHA-256 inicial `ff4f6b37d529e7404ae31bc65ceb1aeb3ba8f54aeb007b71c1e47c9bb9b9585d`. Seis arquivos foram respaldados e conferidos antes desta atualização, em [backup](./_auditoria/fase4_20261031_230457Z/interface_20261101_150908Z/backup/), com [manifesto](./_auditoria/fase4_20261031_230457Z/interface_20261101_150908Z/manifesto-backup.json). A interface ficou no diálogo de importação, sem exclusões, salvamento ou fechamento de instâncias do usuário. Para continuar, a confirmação/importação pelo VBE exige ação manual enquanto o controle de janelas não operar corretamente; seguir o [roteiro único](./COMPILACAO_VBA_FASE4.md).
