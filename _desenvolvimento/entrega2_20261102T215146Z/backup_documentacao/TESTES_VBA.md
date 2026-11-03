> Atualização efetiva de 01/11/2026: a fixture da Fase 4 foi integrada e compilada no Excel real. Três suítes concluídas: 9 + 377 + 11 verificações PASS; recuperação ampla protegida continua FAIL e A08/A09 críticos. Sistema NÃO HOMOLOGADO. O [relatório atual](./RELATORIO_INTEGRACAO_FASE4.md) e a [matriz](./MATRIZ_REGRESSAO_VBA_FASE4.md) substituem os estados de importação bloqueada descritos historicamente abaixo.

# Continuidade — Fase 4 (20261031_230457Z)

Estado atual: [RELATORIO_INTEGRACAO_FASE4](./RELATORIO_INTEGRACAO_FASE4.md), [roteiro manual assistido](./COMPILACAO_VBA_FASE4.md), [matriz funcional](./MATRIZ_REGRESSAO_VBA_FASE4.md) e [proposta de identidade](./PLANO_IDENTIDADE_REGISTROS.md). Fontes e suítes copiados sem alteração; hashes conferidos e 114 backups criados. Nenhuma nova auditoria geral/teste equivalente nem tentativa VBProject. Integração, compilação e funções VBA aguardam intervenção manual; zero revisados incorporados comprovados e contagem total VBE não aferida. A08/A09 críticos; sistema NÃO HOMOLOGADO. Conteúdo anterior preservado abaixo.

---

# Atualização de testes — Fase 3 (20261031_213925Z)

Estado atual na [MATRIZ_TESTES_FASE3](./MATRIZ_TESTES_FASE3.md). **Zero importações, zero compilações, zero macros VBA executadas**: bloqueio 0x800A03EC. Suítes existentes reaproveitadas e estendidas, sem converter PASS COM/modelo em PASS VBA.

Evidências novas: 12 contratos estáticos; três estratégias A08 em quatro experiências; 15 experiências de fronteira A09 (dez recomposições, quatro dependentes de conciliação e uma recuperação incompleta); perda de metadado demonstrada e quatro regressões equivalentes corrigidas; 12 casos XML de parser/modelo e cinco cenários de backup. Não somar células, asserções internas ou testes antigos repetidos como cobertura funcional.

O [relatório](./RELATORIO_HOMOLOGACAO_FASE3.md) registra limites e riscos; o [roteiro manual](./_auditoria/fase3_20261031_213925Z/COMPILACAO_MANUAL.md) contém os passos que dependem do operador. A comparação de 200 células é restrita a valor/fórmula/HasFormula/NumberFormat, com metadado Ignore testado separadamente; não é imagem integral de workbook.

Precisão sobre A08: a pesquisa rápida de entradas percorre lacunas; a parada em primeiro vazio corresponde aos filtros legados de saídas/NCM. Validade estrutural XML não prova autorização/autenticidade. Registros anteriores abaixo preservados como históricos.

---

# Testes VBA — atualização da fase 2

Execução 20261031_201621Z. **Compilação, importação e execução VBA: BLOQUEADAS.** Esta seção é o estado atual; o registro anterior foi preservado abaixo. Testes COM equivalentes, modelos JS/.NET e análise textual não comprovam comportamento dos eventos dos UserForms.

## Resultados observados

| Suíte / modalidade | Resultado | Escopo e evidência |
|---|---|---|
| Integridade e contratos, estática/JS fase 2 | **12 PASS, 0 FAIL** na execução final | 39 originais/backups, 155 arquivos anteriores, 46 backups novos, FRX/designers, 647 assinaturas, codificação/CRLF, sete Delete, nove rotinas fiscais e pré-validação nos quatro gravadores. [JSON](./_auditoria/fase2_20261031_201621Z/testes-estaticos-fase2.json). |
| Suíte anterior, estática/JS | **10 PASS, 0 FAIL**, incluindo 35 cenários de filtros | Executada contra a revisão da fase 1, preservada; não simula novos eventos. [JSON](./_auditoria/fase2_20261031_201621Z/regressao-fase1/testes-estaticos.json). |
| Suíte anterior, Excel/MSXML COM | **6 + 6 PASS** | Chave como texto, persistência de fórmula/formato, reprodução de exclusão; namespace, malformação/DTD e cardinalidade. [Excel](./_auditoria/fase2_20261031_201621Z/regressao-excel-com.json), [XML](./_auditoria/fase2_20261031_201621Z/regressao-xml-com.json). |
| A08, Excel COM diagnóstico | **7 PASS de reprodução do defeito** | Comandos equivalentes em tabelas fictícias, referências/sentinelas; não são PASS de integridade. [JSON](./_auditoria/fase2_20261031_201621Z/testes-a08-com.json). |
| A08, dois protótipos COM | **2 PASS de observação esperada** | Compactação troca identidade de referência; limpeza cria lacuna que esconde resultados no filtro atual. Nenhuma estratégia aprovada/integrada. [JSON final](./_auditoria/fase2_20261031_201621Z/testes-a08-prototipos-v2-com.json). |
| A09, equivalente COM inicial | **38 PASS** | 16 falhas por etapa + 16 novas tentativas + seis sucessos antes de baixa/Save. [JSON](./_auditoria/fase2_20261031_201621Z/testes-a09-com.json). |
| A09, equivalente COM reproduzível final | **70 PASS, 0 FAIL** | 32 falhas entre/dentro das escritas + 32 novas tentativas + seis sucessos. Compara células, tipos/fórmulas/formatos e estados antes/depois. Não executa a classe VBA. [JSON](./_auditoria/fase2_20261031_201621Z/testes-recuperacao-reprodutivel-v2.json), [script](./_auditoria/fase2_20261031_201621Z/scripts/Test-Fase2Recuperacao.ps1). |
| Estrutura XML e números, MSXML/RegExp COM + modelo .NET | **13 PASS** | Um/zero/dois infNFe, Id/nNF inválidos; decimal, zero, negativo, vírgula, texto, vazio e overflow; item inválido no somatório. .NET modela conversão; não executa Val ou relatório VBA. [JSON](./_auditoria/fase2_20261031_201621Z/testes-xml-condicoes-com.json). |
| A19, comparação com DisplayFormat nativo | **5 PASS** | Limites 5 e 10, valores 4/5/7/10/11. Detectou que Formula1/2 retornam =5/=10. É comparação do critério, não UDF VBA. [JSON](./_auditoria/fase2_20261031_201621Z/testes-xml-condicoes-com.json). |
| Backup, duas pastas, limites e timeout, COM final | **20 PASS, 0 FAIL** | Sete verificações de contexto/backup, sete de leitura de limites, cinco de condição nativa e uma de subprocesso próprio. [JSON](./_auditoria/fase2_20261031_201621Z/testes-backup-contexto-v4-com.json), [script](./_auditoria/fase2_20261031_201621Z/scripts/Test-Fase2BackupContextoV4.ps1). |
| Sintaxe PowerShell | **Sem erros de parser na execução final** | Analisa PS1 sem executá-los nem mudar ExecutionPolicy. [JSON](./_auditoria/fase2_20261031_201621Z/sintaxe-powershell-fase2.json). |
| Problems do VS Code | Sem erros exibidos | Não há evidência de compilação VBA nesse resultado. [Registro](./_auditoria/fase2_20261031_201621Z/problems-fase2.json). |
| Importação / compilação VBA | **BLOQUEADO** | Office 16.0 build 20527 x64; zero componentes importados, erro 0x800A03EC. [Erro original](./_auditoria/fase2_20261031_201621Z/compilacao-probe.json). |
| Suítes VBA e eventos de UserForms | **BLOQUEADO; zero execuções** | Suítes antigas revisadas; nova suíte preparada. Nenhum PASS de macro, Ribbon ou fluxo fiscal. |

Os totais das suítes repetidas não devem ser somados como cenários independentes. A suíte de 70 aprofunda a de 38; as comparações de formatação também se repetem. Não há índice de cobertura funcional VBA medido.

## Falhas e ajustes da infraestrutura preservados

- A08: as primeiras preparações de protótipos falharam por marshaling de System.Object[,] e System.Double para String; os JSON iniciais foram preservados. O ensaio final usou cópia de intervalo nativa e retornou as duas observações descritas.
- Backup: a primeira tentativa parou ao tentar ler hash de arquivo aberto sem compartilhamento. V2 parou ao abrir arquivo homônimo na mesma instância. V3 executou 20 verificações, com **18 PASS e 2 FAIL** por comparar formato suposto 0.00, enquanto Excel devolveu #.000. A inspeção real isolou os dois FAIL no formato: chave, conteúdo e fórmula estavam corretos. V4 passou ao comparar o formato **efetivamente observado antes e depois**, usando instância própria separada. [Tentativa inicial](./_auditoria/fase2_20261031_201621Z/testes-backup-contexto-com.json), [V2](./_auditoria/fase2_20261031_201621Z/testes-backup-contexto-v2-com.json), [V3](./_auditoria/fase2_20261031_201621Z/testes-backup-contexto-v3-com.json), [diagnóstico](./_auditoria/fase2_20261031_201621Z/diagnostico-backup-v3.json). Não se alterou o resultado histórico para PASS.
- Recuperação reproduzível: preparação inicial executou zero casos porque NumberFormat = General foi recusado pelo COM local. V2 usa formato capturado do Excel, com 70 PASS. [Erro inicial](./_auditoria/fase2_20261031_201621Z/testes-recuperacao-reprodutivel-inicial.json).
- Estática inicial: **10 PASS e 2 FAIL de asserção**. O teste de codificação confundia ASCII válido em Windows-1252 com troca de codificação, e o teste de celula.Errors incluía ocorrências legítimas de outros procedimentos. Foram corrigidos por round-trip de bytes e escopo no gravador, sem apagar os FAIL. [Inicial](./_auditoria/fase2_20261031_201621Z/testes-estaticos-fase2-inicial.json).
- Parser PowerShell: os erros individuais eram zero, mas o primeiro agregador somou a propriedade Count dos hashtables e mostrou 33. O JSON inicial foi preservado e o agregador corrigido para quantidade de mensagens. [Inicial](./_auditoria/fase2_20261031_201621Z/sintaxe-powershell-fase2-inicial.json). A execução final do parser não detectou erros.
- Nenhum desses ajustes foi contornado por mudança de segurança. Arquivos PS1 foram preparados para manutenção; a política de execução não foi alterada.

## Matriz por operação

PASS significa teste executado com resultado esperado na modalidade indicada; FAIL significa divergência executada; BLOQUEADO significa impedimento técnico/permissão; NÃO EXECUTADO significa caso ainda não realizado. Um equivalente COM aprovado não transfere PASS para a macro correspondente.

| Operação / cenário | Verificação disponível | Evento/integração VBA | Pendência |
|---|---|---|---|
| Venda válida | PASS do mecanismo equivalente com duas escritas e seis sucessos por fluxo | BLOQUEADO | Dados reais do evento, totais, tributos e baixa/fórmulas completos não exercitados. |
| Venda duplicada | PASS de nova tentativa após compensação com CountIf = 1 por destino | BLOQUEADO | Não testa o ramo real de duplicidade, nem repetição após Save/Delete. |
| Venda inválida/vazia | Presença estática de pré-conversão e guardas | BLOQUEADO | Exceções CDbl/CDate e campos obrigatórios pelo evento: NÃO EXECUTADO. |
| Venda com falha parcial | PASS COM antes de baixa/Save, inclusive dentro da linha | BLOQUEADO | Falhas durante/depois de Delete/Save: NÃO EXECUTADO; mecanismo recusa prometer reversão. |
| Transferência, dados/chave/históricos | PASS COM de quatro destinos/recuperação; chave 44 preservada na reabertura independente | BLOQUEADO | Fórmulas M:O, estado das abas e evento de transferência integrados ainda não validados. |
| Devolução, gravações/exclusão | PASS COM dos quatro destinos antes da baixa; A08 reproduziu defeito de A:M | BLOQUEADO | Recuperação estrutural após exclusão e histórico completo pendentes. |
| Cancelado, comportamento atual | Inspeção confirmou duas gravações e ausência de baixa nesse ramo; PASS do modelo de duas escritas | BLOQUEADO | Validade fiscal e eventual exigência de evento dependem do responsável. |
| Pesquisa sem resultado/múltiplos/válido | PASS do modelo JS anterior e contratos estáticos; suíte fake preparada | BLOQUEADO | Pesquisa real e navegação do formulário não executadas. |
| ListBox negativo/vazio/selecionado | Guardas confirmadas em fonte | BLOQUEADO | Clique, ListIndex, ColumnCount e apresentação MSForms reais não executados. |
| XML válido/malformado/múltiplos infNFe | PASS MSXML/RegExp e modelo estrutural | BLOQUEADO | ValidarXMLNFe, interface e relatório fiscal não executados; sem XSD/assinatura. |
| XML com número inválido/vazio/overflow | PASS do critério COM/.NET e inspeção do erro no somatório | BLOQUEADO | Execução de Val, AdicionarResultado e recusa efetiva pelo VBA pendentes. |
| Backup proprietário/versionamento/restauração | PASS COM com duas versões, hashes, chave, fórmula e formato; base em disco não salva | BLOQUEADO | Macro real, destino inválido, concorrência e política de retenção ainda não executados. |
| Ribbon com duas pastas | PASS do salvamento COM pelo proprietário | BLOQUEADO | Callback/Vínculo Ribbon da fixture ainda não executados. |
| Edição anual e exclusão entre duas abas | Inspeção de risco em Alterar/ExcluirFilial | BLOQUEADO | Mapeamento anual e recuperação estrutural ainda indefinidos. |
| Sleep/reentrada/compatibilidade x86 | Declaração e chamadas conferidas em fonte | BLOQUEADO | Duração, CPU, meia-noite e reentrada no VBA: NÃO EXECUTADO. |
| Subprocesso lento | PASS COM de término apenas do processo criado, em cerca de 2,4 s com custo de encerramento | BLOQUEADO | Rotina VBA de 10 s, StdOut/StdErr grandes e filhos de processo não testados. |
| Selenium/Detran/bancos externos | Análise estática apenas | NÃO EXECUTADO | Nenhuma autenticação, consulta ou refresh externo autorizado/executado nesta fase. |
| Word | Nenhum teste desta fase | NÃO EXECUTADO | Não há documento Word nem funcionalidade Word identificada para regressão neste projeto. |

## Reexecução e compilação

O [roteiro manual](./_auditoria/fase2_20261031_201621Z/COMPILACAO_MANUAL.md) explica importação, CodeNames, bibliotecas, limitações da fixture e registro real de compilação. O [esquema](./_auditoria/fase2_20261031_201621Z/scripts/fixture-esquema.json) possui os 17 nomes/CodeNames e código dos documentos, sem dados de estoque.

[modTestesAuditoria.bas](./_auditoria/fase2_20261031_201621Z/scripts/modTestesAuditoria.bas) e [AuditoriaListaFake.cls](./_auditoria/fase2_20261031_201621Z/scripts/AuditoriaListaFake.cls) foram revistos; não executados. [modTestesFase2.bas](./_auditoria/fase2_20261031_201621Z/scripts/modTestesFase2.bas) foi preparado para testar a classe diretamente, fronteiras Delete/Save, texto literal, chave e ActiveCondition; não foi importado, compilado ou executado. O preparador agora recusa execução automática de testes. A fixture completa de eventos continua pendente de RowSources, referências resolvidas, nomes definidos, fórmulas e objetos/Ribbon reais.

Os scripts COM de regressão, recuperação e backup foram executados por comandos inline com a mesma lógica salva. [Reproduzir-A08.ps1](./_auditoria/fase2_20261031_201621Z/scripts/Reproduzir-A08.ps1) é um roteiro salvo posteriormente para reprodução futura; esse arquivo específico não foi executado. Os sete casos A08 executados têm evidência nos JSON.

As instâncias Office usadas foram criadas pelos próprios testes e fechadas por suas referências COM. Não houve encerramento global, modificação de segurança, macro financeira real, consulta de rede, instalação ou promoção.

---

# Registro histórico — primeira auditoria

# Testes VBA e evidências — 31/10/2026

## Estado da validação

**Compilação real: BLOQUEADA. Rotinas VBA revisadas: NÃO EXECUTADAS.** O Office recusou o acesso programático ao projeto. Nenhuma política, AccessVBOM ou proteção de macro foi alterada.

A base permaneceu intacta. As operações Office usaram pastas novas com dados fictícios em C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Auditoria_20261031_174936Z, fora do OneDrive. Os arquivos temporários são evidências, não dependências do produto.

## Excel 16.0: execução real por COM

Evidência: [testes_excel_com.json](<./_auditoria/20261031_174936Z/testes_excel_com.json>). Foram executadas operações equivalentes do modelo de objetos, **não os eventos VBA dos UserForms**.

| Verificação | Resultado observado |
|---|---|
| Chave fictícia de 44 dígitos no formato padrão | Convertida para System.Double; representação original perdida. Reprodução do risco confirmada. |
| Chave com zeros após formato texto | System.String, 44 caracteres; PASS. |
| Chave longa variada após formato texto | System.String, 44 caracteres; PASS. |
| Salvar/reabrir XLSX fictício | Chave idêntica; PASS. |
| Salvar/reabrir fórmula e formato | Fórmula =1+2, resultado 3 e NumberFormat preservados; PASS. |
| Excluir A2:M2 com fórmula em N e sentinelas em O | A3 deslocada; O2 manteve a sentinela anterior; N2 tornou-se =#REF!. Reprodução do risco confirmada. |

O arquivo [regressao_excel.xlsx](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Auditoria_20261031_174936Z/regressao_excel.xlsx>) foi salvo/reaberto. A exclusão ocorreu em outra pasta nova, fechada sem salvar. A reprodução demonstra o mecanismo de risco; não afirma que todo Delete produza o mesmo resultado.

Uma preparação inicial usando NumberFormat General foi recusada pela instalação em português. A repetição usou o formato padrão da célula (Geral) e concluiu as seis verificações. Não foi uma falha de VBA, pois as rotinas não foram chamadas.

## MSXML 6.0: execução real por COM

Evidência: [testes_xml_com.json](<./_auditoria/20261031_174936Z/testes_xml_com.json>). Somente literais XML sintéticos, sem importação fiscal, com resolveExternals=False e ProhibitDTD=True.

- **PASS:** XML sintetico bem formado.
- **PASS:** XPath com namespace.
- **PASS:** Chave no atributo preservada como texto.
- **PASS:** XML malformado rejeitado.
- **PASS:** DTD interna rejeitada.
- **PASS:** Dois infNFe identificados.

São testes do parser e da estratégia XPath. Não atestam XSD, assinatura, autorização fiscal ou a função ValidarXMLNFe completa.

## Análise estática e simulação

Evidência: [testes-estaticos.json](<./_auditoria/20261031_174936Z/testes-estaticos.json>). **10 verificações aprovadas, zero falhas**:

- 39 originais e backups preservados por SHA-256.
- Somente os cinco fontes previstos foram modificados na copia.
- Cinco FRX e pasta XLSM da copia continuam identicos.
- Modulos, assinaturas e eventos preservados; ramos VBA7 de 32 e 64 bits analisados.
- Regras XML intactas byte a byte e calculos tributarios dos formularios preservados.
- Guardas ListIndex antes de ler listas NCM, saidas matriz e filial.
- Referencia celula e referencias invalidas de Frame removidas do codigo executavel.
- Callback Ribbon salva a propria pasta.
- Modelo JS da alteracao de alocacao: listas equivalentes em 35 cenarios.
- Sem duplicidade publica e mesmo total de procedimentos.

O modelo JavaScript comparou as matrizes antes/depois em 35 cenários: zero a 500 registros, termo vazio, caixa diferente, ausente e zeros. Isso verifica o movimento da realocação nesse modelo; não executa VBA nem mede tempo no Excel.

As análises completas estão em [análise original](<./_auditoria/20261031_174936Z/evidencias/original/analise-estatica.json>) e [análise revisada](<./_auditoria/20261031_174936Z/evidencias/revisado/analise-estatica.json>). Os ramos VBA7 de 32/64 bits foram examinados textualmente; isso não prova ABI/compatibilidade em execução. Problems não mostrou erros nos cinco fontes alterados; esse painel não comprova compilação VBA.

## Bloqueios e tentativas

| Etapa | Situação |
|---|---|
| Arquivo PowerShell | O .ps1 foi recusado pela política atual. Não houve bypass. Comandos diretos autorizados foram executados no terminal sem mudar políticas. |
| Excel COM na sessão restrita | Falhou com 0x80070520: [office_probe.json](<./_auditoria/20261031_174936Z/office_probe.json>). Fora da restrição, os testes fictícios funcionaram. |
| VBProject/VBE | Acesso recusado: [office_excel_detail.json](<./_auditoria/20261031_174936Z/office_excel_detail.json>). A propriedade direta pode retornar null; Reflection preservou a exceção real. |
| Preparação inicial de fixture VBA | Nenhum módulo importado nem fixture revisada salva. Estado intermediário de importação não comprova sucesso. |
| Word | Automação pendente sem documento/bookmark confirmado. A instância oculta WINWORD PID 38888, criada pelo teste às 14:57:19, foi identificada e encerrada; a sessão de teste foi cancelada. Nenhum PASS Word. |
| Compilar e rodar VBA | NÃO EXECUTADO. Nenhuma macro fiscal, de exclusão, gravação ou Selenium foi chamada. |

## Infraestrutura

- [auditar-vba.mjs](<./_auditoria/20261031_174936Z/scripts/auditar-vba.mjs>): inventário, procedimentos, repetição e grafo aproximado; Node sem pacotes externos.
- [testar-auditoria.mjs](<./_auditoria/20261031_174936Z/scripts/testar-auditoria.mjs>): integridade e regressão estática/JS. Execute com Node, fornecendo opcionalmente a pasta desta execução de auditoria.
- [Test-ExcelCom.ps1](<./_auditoria/20261031_174936Z/scripts/Test-ExcelCom.ps1>): recria seis casos Excel em diretório temporário novo. O roteiro equivalente foi executado diretamente no terminal; o arquivo requer que a política local permita scripts.
- [Test-XMLCom.ps1](<./_auditoria/20261031_174936Z/scripts/Test-XMLCom.ps1>): parser com dados sintéticos; roteiro equivalente executado no terminal.
- [Preparar-TestesVBA.ps1](<./_auditoria/20261031_174936Z/scripts/Preparar-TestesVBA.ps1>): prepara pasta nova com fontes e abas fictícias, sem abrir o XLSM original. Interrompe se VBProject estiver indisponível, sem alterar segurança. Por padrão não executa macros; a opção RunTests exige prévia revisão e compilação da fixture.
- [modTestesAuditoria.bas](<./_auditoria/20261031_174936Z/scripts/modTestesAuditoria.bas>) e [AuditoriaListaFake.cls](<./_auditoria/20261031_174936Z/scripts/AuditoriaListaFake.cls>): suíte **preparada, não executada**, com nove asserções sobre pesquisa, erro de célula, origem, linha, chave e log. A classe é um dublê de teste, não um controle de produção substituto.
- [Test-OfficeCapabilities.ps1](<./_auditoria/20261031_174936Z/scripts/Test-OfficeCapabilities.ps1>): sondagem exploratória. Word é opcional e tem limitação de inicialização nesta máquina; não considere a regressão aprovada por sua presença.

## Roteiro manual seguro

1. Crie um XLSM novo em TEMP, fora do OneDrive, com as 17 abas do [MAPA_PROJETO.md](<./MAPA_PROJETO.md>). Não execute macros sobre uma cópia contendo estoque/notas reais. Use somente registros fictícios.
2. Importe no editor VBA os BAS e os cinco pares FRM/FRX da [cópia revisada](<./_auditoria/20261031_174936Z/desenvolvimento/>). Confira designer, nomes e referências. Não importe módulos de planilha como classes comuns; use os módulos de documento correspondentes quando necessário.
3. Use Depurar > Compilar VBAProject. Registre módulo, procedimento, linha e mensagem. Esse passo não foi executado aqui. Se houver exigência de mudança de segurança, obtenha autorização específica antes.
4. Preencha Listas e NCM e CEST com dados fictícios compatíveis com os RowSources reais. Abra cada formulário; teste Clear, ListIndex=-1, seleção válida, pesquisa vazia/sem resultado e abas matriz/filial. Confirme o vínculo dos eventos.
5. Na fixture preparada pelo script, importe os dois componentes de teste e execute Auditoria_ExecutarTestesSeguros. A suíte exige nome, TEMP e marcador de dados fictícios; nunca a rode no arquivo real.
6. Venda: teste número válido, vazio, duplicado, moeda/data inválidas e XML divergente. Confira gravações e baixa da entrada. A ausência de transação continua pendente.
7. Transferência: compare Entradas Filial A:L, chave de texto, Entradas Anual Filial e fórmulas M:O. Insira falhas artificiais após cada gravação e verifique recuperação.
8. Devolução filial: confira Saidas Filial, Todas Saidas Filial, Entradas e Entradas Anual; teste entrada inexistente, duplicados e falha após a primeira gravação.
9. Cancelamento: confirme a dispensa de XML com o responsável; confira a permanência da entrada conforme a regra atual. Não trate XML normal como prova de evento sem regra definida.
10. Exclusão: use fórmulas sentinela e valores diferentes por linha em todas as colunas auxiliares. Valide os sete Delete, inclusive N:O e históricos, contra a política aprovada.
11. XML: teste namespace, arquivo vazio/malformado, DTD, dois infNFe, chave inválida, vários itens, números inválidos e data impossível. Compare relatórios campo a campo, preservando CFOP, CST, cBenef e tributos.
12. Mantenha duas pastas fictícias abertas para testar Salvar. Compare o estado Saved e os arquivos gravados. Teste recuperação dos estados globais com valores prévios diferentes.
13. Salve, feche e reabra a fixture. Compare valores, fórmulas, formatos, referências, Ribbon, designer e hashes dos fontes exportados antes de qualquer promoção autorizada.

Consultas externas, logins, operações fiscais reais, integrações Access e automação Word do produto não foram validadas nesta execução.

## Sintaxe dos scripts auxiliares

Os quatro arquivos PowerShell foram analisados pelo parser instalado, sem executá-los: zero erros sintáticos. Evidência: [sintaxe_powershell.json](<./_auditoria/20261031_174936Z/sintaxe_powershell.json>). Esse resultado não é compilação VBA nem aprovação funcional da suíte preparada.


## 01/11/2026 — tentativa assistida pela interface (20261101_143216Z)

O usuário autorizou controle do computador e uso contínuo do Excel em fixtures. Foram preservados nove arquivos/backup da fixture, com hashes conferidos, antes de qualquer edição. [Manifesto](./_auditoria/fase4_20261031_230457Z/interface_20261101_143216Z/manifesto-inicial.json). A tentativa usou a interface Windows e não repetiu o acesso programático ao VBProject.

O controle localizou Excel e o diálogo Abrir. Houve expiração/reconexão, recusa do índice de edição em cache e recusa de clique por divergência entre a janela principal e modal. Após atualização/ativação e tentativa por teclado, o foco permaneceu na Caixa de Pesquisa; o campo Nome do arquivo não recebeu texto. Nenhum VBA foi importado, compilado ou executado. A fixture continua com SHA-256 inicial; isso não comprova compilação. [Erros exatos](./_auditoria/fase4_20261031_230457Z/interface_20261101_143216Z/resultado-interface.json).

**BLOQUEADO por controle da janela modal**, separado do bloqueio VBProject histórico. Foi solicitado somente abrir a fixture já preparada e deixá-la em primeiro plano, para retomar o VBE; não se solicitou nova aprovação nem mudança de segurança. A instância Excel não foi encerrada e nenhuma instância do usuário foi fechada. Continuam NÃO INTEGRADO/NÃO HOMOLOGADO e A08/A09 críticos.


## Fase 4 — execuções reais de 01/11/2026

Compilação VBA PASS por Depurar > Compilar, após erro de sintaxe C4-01 corrigido. Auditoria_ExecutarTestesSeguros: 9 PASS; AuditoriaFase2_Executar: 377 PASS; AuditoriaFase4_IntegridadeComplementar: 11 PASS. Reexecução de 377 com outra pasta ativa e após C4-04 não soma novos cenários. O ciclo AuditoriaFase3_CicloVida foi iniciado, mas o depurador interceptou erro esperado: BLOQUEADO, sem retorno PASS. Não se alterou a preferência sem resposta à solicitação.

FAIL real: proteção ampla impede compensação integral (quatro falhas). C4-04 retirou ClearContents destrutivo anterior ao formato; mantém dados alterados e exige conciliação. PASS complementar de falha individual isolada em outra aba não reclassifica o FAIL. Macro CriarBackupVersionado executada duas vezes e callback ChamarSalvar com outra pasta ativa; backups reabertos por COM, sem confundir modalidades. Pesquisa matriz/MSForms com seleção válida, -1 e sem resultado: PASS. XML: sem infNFe, múltiplos, chave curta e nNF ABC rejeitados; LoadXML malformado retorna False/-1072896685. Não certifica XML completo, tributos ou evento fiscal.

[Casos, estados e limites atuais](./MATRIZ_REGRESSAO_VBA_FASE4.md), [saídas e integridade](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/). Os JSON de UI transcrevem resultados efetivamente observados, não constituem logs automáticos da suíte. Nenhum evento gravador completo Venda/Transferência/Devolução/Cancelamento foi executado nesta retomada. Esses cenários continuam NÃO EXECUTADOS, e funções dependentes de identidade permanecem BLOQUEADAS por D01/D02.


Complemento da mesma execução Fase 4: seis casos dos botões reais de entrada executados (cinco matriz: válida/duplicada/valor inválido/data inválida/número vazio; um filial válido), todos PASS nos critérios documentados. Matriz sem células alteradas nas recusas; filial preserva String 44/@, fórmulas sentinela M/N e referência anual. A:J corresponde aos dois históricos. A entrega fictícia final foi salva com SHA-256 dcbd72e46e8b182a9810254bcce460ea2cb17260b4bcb47b622ddc3d4f9879a0. Fontes dos gravadores não modificados; não equivale a testar Venda/Transferência/Devolução/Cancelamento nem falhas entre suas gravações. [Evidências](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/entradas-vba-real.json).


## Continuação — Fase 5 (20261101_202114Z)

[Relatório dos testes integrados reais](./RELATORIO_TESTES_INTEGRADOS_FASE5.md), [matriz e evidências](./MATRIZ_REGRESSAO_FASE5.md), [mapa A08/A09](./MAPA_INTEGRIDADE_ESTOQUE.md), [recuperação](./PLANO_RECUPERACAO_FALHAS.md). Base sintética compilada no VBE com 56 componentes/17 planilhas. Eventos de negócio: 20 PASS, cinco FAIL, um BLOQUEADO para injeção de Save; 13 casos XML PASS, suítes auxiliares reais 377+11 asserções PASS e classe com duas pastas PASS. Ciclo de vida BLOQUEADO; recuperação incompleta anterior permanece FAIL. Venda/Transferência/Devolução não preservaram integridade após a baixa. Originais e fontes herdados intactos; nenhum dos sete Delete ou regra fiscal modificado. D01/D02 pendentes e sistema NÃO HOMOLOGADO. Resultados anteriores preservados; consultar matriz para modalidades e testes não repetidos.


## Execução experimental Fase 6 — 20261101_224949Z

Compilação real e execução no Excel/VBE, sem VBProject/COM: 43 cenários únicos, 28 PASS e 15 FAIL; dois checks de guarda PASS separados; 86 tentativas preservadas. São testes de motor VBA experimental, não reexecução dos eventos privados do produto. Pesquisa rápida existente chamada em C_LEITOR_LEGADO.

Cenários com falta de assertiva específica foram registrados NÃO EXECUTADO e repetidos após correção do teste; bloqueio de reabertura por memória insuficiente/0x8000FFFF foi preservado. SaveCopyAs em destino inexistente não foi apresentado como Save legado. Crash, multiwriter e migração real NÃO EXECUTADOS. [Matriz com resultados e modalidades](./MATRIZ_REGRESSAO_FASE6.md), [evidências estruturadas](./_auditoria/fase6_20261101_224949Z/evidencias/resultados-consolidados.json) e [limites de recuperação](./ARQUITETURA_RECUPERACAO_FASE6.md).


## Continuidade — Fase 7 (20261102_025115Z)

VBA experimental real compilado: 18 verificações de adaptadores PASS; 12 cenários frios/15 checks PASS. Matriz consolidada 38 verificações únicas: 34 PASS/quatro FAIL, com um PASS numérico não discriminante explicitamente limitado; 40 execuções, sem somar repetição. Eventos reais migrados NÃO EXECUTADOS; recuperação integral F6 continua FAIL. [Relatório](./RELATORIO_TECNICO_FASE7.md), [matriz](./MATRIZ_REGRESSAO_FASE7.md), [contrato](./CONTRATO_IDENTIDADE_FASE7.md), [mapa](./MAPA_MIGRACAO_LEITORES_FASE7.md) e [decisões](./PACOTE_DECISOES_FASE7.md). Checkpoints e evidências em [execução F7](./_auditoria/fase7_20261102_025115Z/). Portão2 não aprovado; NÃO HOMOLOGADO.


## Fase 8 — 02/11/2026: preparação experimental do Portão 2

Execução [_auditoria/fase8_20261102_121422Z](./_auditoria/fase8_20261102_121422Z/). Base anterior:2.106 hashes conferidos sem divergência na abertura e35 backups novos verificados. Originais e sete Delete operacionais preservados. Mestre experimental final:60 componentes/17 abas, compilação real no Excel16.0/20527/64bits; fontes herdados sem alteração de literais/corpos. Portão2 não executado; sistema NÃO HOMOLOGADO.

Quatro defeitos F7 reproduzidos e corrigidos somente nos protótipos: data/selo persistente com política explícita, anual por episódio, consulta por ID após sort/lacuna/reuso e auxiliares com valores discriminantes. Guarda adicional comprovada: recusar selo ERRO coincidente; três FAIL preservados e três PASS após correção.70 asserções únicas aprovadas na seleção de resultados experimentais e4 volumes medidos (100/1000/5000/10000), com limites por revisão na matriz.26 integrados repetidos após guarda final;15 guardas F7 reexecutados em sessão fria. Não houve teste novo de recuperação integral ou aprovação de gravadores reais.

[Relatório técnico](./RELATORIO_TECNICO_FASE8.md), [matriz](./MATRIZ_REGRESSAO_FASE8.md), [correções](./CORRECOES_PROTOTIPOS_FASE8.md), [armazenamento](./PROPOSTA_ARMAZENAMENTO_FASE8.md), [persistência](./CONTRATO_PERSISTENCIA_FASE8.md), [M0–M9](./PLANO_IMPLANTACAO_FASE8.md), [decisões Portão2](./PACOTE_APROVACAO_PORTAO2.md). Cinco FAIL oficiais F5 e FAIL integrais F6 continuam pendentes. Nenhuma promoção, migração real, alteração fiscal, instalação ou mudança de segurança.


## Fase 9 — 20261102_142753Z — persistência experimental

Portão experimental somente; Portão 2 NÃO aprovado e sistema NÃO HOMOLOGADO. Entrega F8 R2 e 2.268 hashes conferidos; 43 backups novos verificados. ACE 16/ADODB-ADOX x64 do Office instalado executou transações reais em bancos sintéticos no TEMP. Mestre F9 compilado no VBE: 63 componentes/17 abas, com os 60 herdados preservados e os sete Delete intactos. Nenhuma integração no principal, migração real, regra fiscal alterada, instalação, segurança modificada, commit/push ou promoção.

65 asserções VBA: 64 PASS e 1 FAIL de Ignore após Sort, preservado com prova corretiva separada. Quatro benchmarks VBA PASS e três testes COM entre processos PASS, sem somar sondagens/hashes. Ciclo de sete movimentos/recibos, rollback pré-commit, bloqueio de pendência, escritor/fencing e reconstrução em sessão fria exercitados. Erro real de SaveAs e cancelamento BeforeSave preservaram a autoridade. Falha física de commit, energia/crash/disco cheio, rollback incompleto e snapshot sob escrita concorrente NÃO EXECUTADOS. Os sete FAIL oficiais F5/F6 permanecem nas classificações anteriores.

[Relatório](./RELATORIO_TECNICO_FASE9.md), [matriz](./MATRIZ_REGRESSAO_FASE9.md), [armazenamento](./COMPARATIVO_ARMAZENAMENTO_FASE9.md), [esquema](./ESQUEMA_TRANSACIONAL_FASE9.md), [durabilidade](./TESTES_DURABILIDADE_FASE9.md), [concorrência](./TESTES_CONCORRENCIA_FASE9.md), [reconstrução](./PROTOCOLO_RECONSTRUCAO_EXCEL_FASE9.md), [M0–M4](./PLANO_MIGRACAO_M0_M4_FASE9.md), [decisões](./DECISOES_PENDENTES_FASE9.md).


## Adendo Fase10 — execução comprovada

Compilação real e execução VBA na instância sintética PID52068:121 asserções únicas PASS,12 amostras benchmark PASS. Testes abrangeram esquema novo/migrado, ciclo com reentrada, transações, adulteração, intenção incompleta, custódia, projeções, ACK, Sort/Ignore, Save cancelado, backup e restauração. Quatro experiências de concorrência COM e seis verificações frias COM PASS têm modalidade separada.

A instância fria PID41728 foi bloqueada pelo Office antes de produzir frio_f10.tsv; seis verificações VBA continuam BLOQUEADAS. Sem alteração de segurança. Não foram exercitadas energia/crash/commit físico/disco cheio/backup quente. [Matriz completa](./MATRIZ_REGRESSAO_FASE10.md), [benchmark e versões](./BENCHMARK_FASE10.md), [limites de recuperação](./TESTES_RECUPERACAO_FASE10.md). Os sete FAIL históricos do produto não mudaram de classificação.


## Entrega 1 — execução integrada real

[44 asserções registradas](./MATRIZ_REGRESSAO_ENTREGA.md), 38 PASS/6 FAIL preservados; revisão R3 com 33 IDs únicos e uma repetição corretiva. Modalidades: eventos reais de UserForms, auxiliares VBA reais, compilação no Editor VBA, inspeção binária/SHA e consulta ADO direta identificadas separadamente. Nenhum resultado COM equivalente substituiu um evento.

PASS em cenários sintéticos de entrada, venda, transferência, devolução, destino duplicado, pesquisa, versão obsoleta, reentrada por identidade explícita, Save cancelado real e reconstrução. Falhas durante transação: fatos revertidos e intenção PENDENTE_CONCILIACAO; recusa de repetição não é recuperação integral. Backup reaberto/auditado não equivale a restauração operacional completa.

Compilação final PASS ([menu](./_desenvolvimento/entrega1_20261102T182823Z/evidencias/compilacao-final-menu.txt)); reabertura interativa BLOQUEADA no aviso de macros. Script de preparação BLOQUEADO por ExecutionPolicy, sem bypass. Os sete FAIL oficiais anteriores continuam preservados. [Registro completo e limitações](./REGISTRO_DEFEITOS_ABERTOS.md).

## 02/11/2026 — continuação da Entrega 1: reabertura real

Checkpoint verificado em ./_desenvolvimento/entrega1_20261102T182823Z/continuacao_reabertura_20261102T205915Z/. O usuário habilitou macros manualmente. E1_SessaoReaberta revelou FAIL por diferença entre aliases MARCOS~1/MARCOSALEXANDREPERDO do mesmo arquivo. Corrigido somente modCandidatoIntegrado.ConnC1 mediante C1_MesmoArquivoAutoridade (existência + ShortPath); BaseID e vínculo persistido preservados. O Editor VBA recompilou sem erro. Três verificações imediatas retornaram True/False/False (mesmo arquivo/cópia distinta/ausente). Nova execução na sessão aberta PASS e nova instância Excel /x PID 50212 confirmou CACHE_AUDITADO automático e E1_SessaoReaberta PASS, sem repetir movimentos.

Log acumulado 47 execuções de asserção: 40 PASS/7 FAIL; falhas anteriores preservadas. As três consultas imediatas e a compilação são evidências separadas, não somadas como cenários de UserForm. Banco SHA-256 inalterado (3198e07699f644f64c325a9fb517470b64950aa5a262cc354629ed61f1503dc3), revisão 8, oito movimentos/recibos, três ativos, zero Pending. XLSM corrigido SHA-256 5b2a0d55fcea077dd034ace4c4892bbc9fdd46c69760d7d318d2162f54771749. Contagem binária 69 componentes; corpos dos cinco forms iguais ao checkpoint (VB_Base reserializado). 39 originais e backup F10 novamente conferidos, sem divergência. E1-04 encerrado no escopo de reabertura normal; E1-03, conciliação, fluxo D e homologação permanecem pendentes. Nenhuma mudança de segurança, produção, regra fiscal, commit ou publicação.
