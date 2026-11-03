# Continuidade — Fase 4 (20261031_230457Z)

Estado atual: [RELATORIO_INTEGRACAO_FASE4](./RELATORIO_INTEGRACAO_FASE4.md), [roteiro manual assistido](./COMPILACAO_VBA_FASE4.md), [matriz funcional](./MATRIZ_REGRESSAO_VBA_FASE4.md) e [proposta de identidade](./PLANO_IDENTIDADE_REGISTROS.md). Fontes e suítes copiados sem alteração; hashes conferidos e 114 backups criados. Nenhuma nova auditoria geral/teste equivalente nem tentativa VBProject. Integração, compilação e funções VBA aguardam intervenção manual; zero revisados incorporados comprovados e contagem total VBE não aferida. A08/A09 críticos; sistema NÃO HOMOLOGADO. Conteúdo anterior preservado abaixo.

---

# Atualização — Fase 3 (20261031_213925Z)

**Sistema NÃO HOMOLOGADO.** A importação automatizada continua bloqueada por 0x800A03EC; não houve compilação nem execução VBA. [Relatório de homologação da Fase 3](./RELATORIO_HOMOLOGACAO_FASE3.md), [plano](./PLANO_CORRECAO_FASE3.md), [matriz](./MATRIZ_TESTES_FASE3.md) e [decisões pendentes](./PENDENCIAS_DECISAO.md) descrevem o estado atual. Os registros anteriores abaixo são históricos.

Os 39 originais, 155 arquivos da Fase 1, 146 da Fase 2 e 46 backups anteriores foram conferidos. O novo [backup de 69 arquivos](./_auditoria/fase3_20261031_213925Z/backup/) e o [manifesto](./_auditoria/fase3_20261031_213925Z/manifesto-inicial-fase3.json) permitem reversão.

Somente a classe de compensação recebeu dois ajustes técnicos: uso único/recusa inativa e preservação do metadado Ignore do erro de número como texto. Outros 33 fontes, FRM/FRX, fiscalidade e sete Delete preservados. Contratos estáticos: 12 PASS. Novos testes COM demonstraram dez recomposições antes das fronteiras, quatro cenários exigindo conciliação e um de recuperação incompleta; não aprovam os eventos VBA.

A08/A09 continuam críticos. A08 foi comparada em três estratégias; recomenda-se C somente depois de identidade, persistência e migração aprovadas. Foram encontrados valores repetidos nos históricos; nenhum identificador foi presumido. A13/A20 aguardam decisões. O XLSM da cópia não contém a revisão; esqueleto e pacote manual estão fora do OneDrive. Sem promoção, commit, publicação ou mudança de segurança.

---

# Relatório de auditoria — atualização da fase 2

Execução **20261031_201621Z**, em 31/10/2026, continuando o registro histórico abaixo. **Entrega em fontes de desenvolvimento; compilação e homologação pendentes.** Esta seção substitui os estados atuais do relatório anterior, conservado ao final.

## Resultado e preservação

Os 39 arquivos originais e seus backups anteriores passaram na comparação SHA-256. Os 155 arquivos da execução anterior permanecem intactos. As cinco revisões de A01–A07 reproduzem exatamente os 16 blocos registrados, sem mudanças posteriores detectadas. Novo backup: [46 arquivos](./_auditoria/fase2_20261031_201621Z/backup/) — 40 de desenvolvimento e seis documentos — com hashes conferidos antes e depois das alterações. [Manifesto](./_auditoria/fase2_20261031_201621Z/manifesto-inicial-fase2.json), [verificação anterior](./_auditoria/fase2_20261031_201621Z/verificacao-fase1.json), [verificação final](./_auditoria/fase2_20261031_201621Z/testes-estaticos-fase2.json).

A base XLSM foi recontada: **17 abas, 50 componentes, 648 procedimentos**. A nova cópia contém **34 fontes exportados e 665 procedimentos**; os 647 contratos de procedimentos exportados existentes, os atributos de módulo, cinco designers/FRX e codificações compatíveis com a base foram preservados. Incorporar os novos fontes com os documentos resultaria em 52 componentes/666 procedimentos, sem testes; isso é uma projeção. O XLSM da cópia continua idêntico à base e **não incorpora as correções**.

O diagnóstico mantém os **20 grupos: 2 críticos, 7 altos, 7 médios e 4 baixos**. Sete correções anteriores preservadas; cinco grupos com ajustes técnicos novos (A10/A14/A15/A18/A19); cinco com mitigação parcial (A09/A11/A12/A16/A17); três dependentes de decisão/proposta (A08/A13/A20). Esses números descrevem tratamento em fonte, sem aprovação funcional VBA.

## Compilação e dependências

Excel 16.0, build 20527, arquitetura x64 confirmada pelo PE 0x8664. Uma pasta nova foi criada por COM, sem abrir a base em Office. VBProject foi recusado: **“O acesso de programação ao projeto do Visual Basic não é confiável”**, HRESULT **0x800A03EC**. Zero módulos importados; nenhuma compilação ou macro executada. [Erro original](./_auditoria/fase2_20261031_201621Z/compilacao-probe.json).

Nenhuma proteção foi alterada. O [preparador revisado](./_auditoria/fase2_20261031_201621Z/scripts/Preparar-TestesVBA.ps1) agora contempla CLS, FRM/FRX, staging de codificação, CodeNames e documentos com esquema extraído, além de registrar referências. Ele não executa macros, e seu caminho de importação permanece não validado por causa do bloqueio. [Roteiro manual](./_auditoria/fase2_20261031_201621Z/COMPILACAO_MANUAL.md).

MSXML 6 e VBScript.RegExp funcionaram nos testes equivalentes. As declarações de Forms, Common Controls/OCX, Internet Controls, Search e Selenium não comprovam registro ou compatibilidade x64; sua importação e referências resolvidas continuam bloqueadas. Não houve instalação nem remoção de dependências. RowSources, nomes definidos, objetos e Ribbon da fixture precisam de conferência manual antes dos eventos dos formulários.

## A08 — sete exclusões investigadas, correção não integrada

| Ponto | Fonte / procedimento | Registro e intervalo | Dependências e risco |
|---|---|---|---|
| D01 | Cad_NFE, Excluir | Entradas, A:N | Campos A:K e auxiliares L:N; referências diretas externas podem virar #REF!. |
| D02 | Cad_NFE, ExcluirFilial | Entradas Filial, A:M | Campos A:L, auxiliares M:O; N:O não acompanham a exclusão. |
| D03 | Cad_NFE, ExcluirFilial | Entradas Anual Filial, A:P | Campos A:J e fórmulas K:P; segunda exclusão pode falhar depois da primeira. |
| D04 | Cad_NFE_Saida, ExcluirFilial | Saidas Filial, linha inteira | Campos A:J/L:P, auxiliares K/Q e totais; apaga também conteúdo lateral da linha. |
| D05 | Cad_NFE_Saida, GravarFilial | Entradas Filial, A:M | Baixa depois de quatro gravações; mesmos riscos de N:O. |
| D06 | Cad_NFE_Saida, Excluir | Saidas, linha inteira | Fórmulas K/Q, totais e conteúdo lateral; histórico Todas Saidas não é apagado por esse evento. |
| D07 | Cad_NFE_Saida, Gravar | Entradas, A:N | Baixa após gravações de saída/transferência; dependências diretas podem quebrar. |

Localizações exatas constam do [mapa final](./MAPA_PROJETO.md) e do [plano](./PLANO_CORRECAO_FASE2.md). O [mapeamento de fórmulas](./_auditoria/fase2_20261031_201621Z/formulas-a08.json) foi extraído do pacote original, sem executar dados reais. Entradas Filial possui 363 fórmulas em N e 364 em O. Há também referências por placa e históricos, UDFs, totais e fórmulas compartilhadas; a fixture de reprodução usa sentinelas, sem simular a totalidade da pasta original.

Sete comportamentos foram reproduzidos por COM: referências quebradas e sentinelas fora do deslocamento nas exclusões parciais; remoção de sentinela lateral na exclusão de linha inteira. PASS no JSON significa reprodução do defeito, não aprovação de integridade.

Dois protótipos explicaram por que não cabe uma troca automática: compactar campos sem Delete preservou fórmulas, mas trocou silenciosamente a identidade de uma referência direta; limpar somente campos manteve outras linhas, mas os filtros atuais pararam na lacuna. Reutilizar a vaga pode criar outro vínculo incorreto. Proposta e aprovação solicitada no [plano](./PLANO_CORRECAO_FASE2.md); resposta ainda não registrada. **Os sete Delete permanecem literalmente iguais.**

## A09 — mitigação parcial com limites explícitos

Os quatro eventos Gravar/GravarFilial de [Cad_NFE](./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE.frm) e [Cad_NFE_Saida](./_auditoria/fase2_20261031_201621Z/desenvolvimento/Form/Cad_NFE_Saida.frm) agora antecipam as mesmas conversões CDbl/CDate, verificam destinos e registram snapshots antes das escritas. A devolução da filial localiza a origem antes de gravar. Proteção, células mescladas, matriz/derramamento e pasta somente leitura são recusadas pelo mecanismo nos pontos previstos.

[clsOperacaoPlanilha](./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/clsOperacaoPlanilha.cls) captura valores, fórmulas e formatos por célula e tenta compensá-los em ordem inversa. Texto começando por = e chave de 44 dígitos são restaurados como texto. Estado anterior de eventos/tela/alertas/cálculo é restaurado independentemente; erros desses setters são registrados na Janela Imediata. [modIntegridadeOperacao](./_auditoria/fase2_20261031_201621Z/desenvolvimento/MDL/modIntegridadeOperacao.bas) centraliza pré-conversão e escolha técnica de destino, mantendo a regra original de vaga.

A classe **recusa recuperação automática depois de iniciar Delete, Save ou confirmar**. Não restaura interface, nomes, referências estruturais, objetos, filtros ou disco já salvo. Não possui diário persistente; encerramento inesperado impede compensação. Esses limites são também mensagens do código, não garantia de rollback.

Sequências preservadas:

- Entrada: Entradas → Entradas Anual → interface → Save.
- Entrada filial: Entradas Filial → Entradas Anual Filial → interface; evento original não salva nesse ponto.
- Venda: Saidas → Todas Saidas → baixa Entradas → Save. Demonstração acrescenta sua aba na posição original.
- Transferência: Saidas → Entradas Filial → Entradas Anual Filial → Todas Saidas → baixa Entradas → Save.
- Devolução filial: Saidas Filial → Todas Saidas Filial → Entradas → Entradas Anual → baixa Entradas Filial → Save.
- Cancelado: Saidas → Todas Saidas → Save, sem carregar/validar XML e sem baixa de origem nesse ramo.

O [mapa de operações](./_auditoria/fase2_20261031_201621Z/operacoes-multiplas.json) identificou oito procedimentos com múltiplas abas literais, dois apenas de navegação. Além dos quatro gravadores, **CommandButtonAlterar_Click** limpa a seleção anterior de Entradas Anual sem identificar inequivocamente o registro; **ExcluirFilial** exclui em duas abas sem compensação estrutural. Essas duas rotinas não receberam um mapeamento inventado. São pendências de A09. Duplicidade histórica depois de falha de baixa/salvamento e nova tentativa também continua sem garantia.

O teste equivalente inicial passou em 38 verificações. A suíte reproduzível posterior passou em **70**: 32 falhas entre/dentro das escritas, 32 novas tentativas e seis sucessos antes das fronteiras de baixa/salvamento. Evidência inclui estados antes/depois. **Nenhum desses resultados é execução dos eventos VBA. A09 continua crítico.**

## Outras alterações e limites

A10: não restam referências literais Sheets/Worksheets sem proprietário nos fontes; os três ActiveWorkbook preservados pertencem às conexões de proprietário indefinido. A11/A12 receberam diagnóstico, sem inventar abas/conexões. A13 mantém Cancelado e requer decisão fiscal.

A14/A15 adicionam identificação estrutural do DOM direto e erro explícito para números inválidos presentes no somatório, inclusive tratamento de overflow. Nove rotinas fiscais especializadas foram comparadas e ficaram idênticas. Não foram alterados CFOP, CST, cBenef, ICMS, PIS, COFINS ou tolerâncias. Não há validação de autenticidade, XSD ou assinatura comprovada por esses testes.

A16 limita espera de subprocesso próprio a dez segundos; captura de saída continua incompleta e pode terminar por timeout com pipe cheio. A17 reduz espera ativa usando Sleep em intervalos curtos; DoEvents/reentrada e benchmark continuam pendentes. A18 versiona backup do proprietário sem salvar antes a base e verifica tamanho/cabeçalho; SHA-256 e restauração foram testados externamente, sem afirmar hash interno. A19 corrige o critério fora do intervalo e limites retornados como =5/=10; fórmulas relativas, limites por referência e tipos modernos continuam exigindo testes. A20 conserva caminhos/logs e registra proposta de destino, privacidade e retenção, sem consultas externas.

Não foram removidos eventos semelhantes nem integrações legadas. Não houve refatoração extensa de mapeamento fiscal/estoque ou substituição indiscriminada de Select, Activate e On Error Resume Next. Não há ganho de tempo VBA medido. Novos auxiliares separam pré-validação, recuperação e backup para revisão e testes, com 17 fontes existentes alterados e dois novos.

## Evidências, pendências e entrega

[TESTES_VBA](./TESTES_VBA.md) contém modalidades, resultados, falhas de preparação e matriz por operação. [Plano](./PLANO_CORRECAO_FASE2.md) registra os vinte grupos, risco e critérios individuais. [Histórico](./HISTORICO_ALTERACOES.md) e [alterações por bloco](./_auditoria/fase2_20261031_201621Z/alteracoes-fase2.json) permitem rastrear/reverter.

A nova [área de desenvolvimento](./_auditoria/fase2_20261031_201621Z/desenvolvimento/) não está pronta para homologação. A08/A09 continuam críticos, compilação/execução VBA bloqueadas, A13 exige decisão e integrações legadas ainda não foram resolvidas. Não houve promoção, commit, push, publicação, alteração de segurança, execução de consultas reais ou movimentação de estoque/dados fiscais reais.

---

# Registro histórico — primeira auditoria

# Relatório de auditoria VBA — 31/10/2026

## Diagnóstico

A base exige correções antes da liberação: há referências que impedem compilação por inspeção, gravações sem recuperação e exclusões capazes de desalinhamento e quebra de fórmulas. Foram corrigidos **7 grupos de problemas em 5 fontes de desenvolvimento**; **13 grupos permanecem pendentes**. Não houve promoção para produção.

Foram inventariados **39 arquivos originais**: 27 BAS, 5 FRM, 5 FRX, 1 XLSM e 1 README. O Excel contém **50 componentes VBA**: os 32 exportados e 18 módulos de documento (17 planilhas e EstaPastaDeTrabalho). São **648 procedimentos no projeto embutido**, dos quais **647 nos fontes exportados**. Não foram encontrados documentos Word, bancos Access nem classes independentes na base original.

Os 32 códigos exportados coincidem com os fontes extraídos do XLSM, desconsiderando atributos e apresentação da exportação. Isso não valida o p-code compilado. A extração foi local, por OOXML/CFB e descompressão VBA, sem executar macros.

| Gravidade | Grupos |
|---|---:|
| Crítica | 2 |
| Alta | 7 |
| Média | 7 |
| Baixa | 4 |
| **Total** | **20** |

A contagem agrupa causas, não ocorrências. As 31 famílias de corpos repetidos encontradas são candidatos a manutenção, não 31 defeitos adicionais. Não foram encontradas duplicidades públicas entre módulos padrão no ramo VBA7/Win64 analisado.

## Originais, backup e revisão

- Base: [Excel/](<./Excel/>), [MDL/](<./MDL/>) e [Form/](<./Form/>) — intactos; 39 hashes conferidos antes e depois.
- Backup: [backup/](<./_auditoria/20261031_174936Z/backup/>).
- Manifesto SHA-256: [manifesto-original.json](<./_auditoria/20261031_174936Z/manifesto-original.json>).
- Fontes revisados: [desenvolvimento/](<./_auditoria/20261031_174936Z/desenvolvimento/>).
- Texto anterior/posterior de cada ajuste: [alteracoes.json](<./_auditoria/20261031_174936Z/alteracoes.json>).
- Testes Office: C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Auditoria_20261031_174936Z, fora do OneDrive, somente com dados fictícios. Nenhum Excel original foi aberto no Office pelos testes.

**O XLSM da cópia de desenvolvimento permanece idêntico ao original e NÃO contém as correções dos fontes.** A importação foi bloqueada pelo Office. Não distribua esse binário como versão corrigida. Os fontes devem ser importados, compilados e testados em fixture antes da promoção.

## Achados e tratamento

As localizações são linhas da **base original**. O [MAPA_PROJETO.md](<./MAPA_PROJETO.md>) contém procedimentos e linhas revisadas.

### A01 — Símbolo celula não declarado na transferência

**Gravidade:** Alta. **Situação:** Corrigido nos fontes.

**Localização:** [Cad_NFE_Saida.frm](<./Form/Cad_NFE_Saida.frm>), linha 750.

**Evidência:** A única ocorrência executável usa celula.Errors, sem declaração. Option Explicit está ativo.

**Tratamento:** A referência agora aponta para .Cells(linhaDemo, 12), a célula da chave já gravada no mesmo With. Compilação real pendente.

### A02 — Frame_ImpostoFilial ausente no formulário

**Gravidade:** Alta. **Situação:** Corrigido nos fontes.

**Localização:** [Cad_NFE_Saida.frm](<./Form/Cad_NFE_Saida.frm>), linha 1369.

**Evidência:** Quatro instruções acessam esse nome; ele não aparece nos streams de controles do FRX. Não há Controls.Add nem chamada dinâmica desse nome nos fontes.

**Tratamento:** As quatro instruções inválidas foram comentadas. Eventos Enter/Exit e alterações dos controles existentes foram preservados; nenhum substituto foi criado.

### A03 — Chave NF-e sem garantia explícita de texto

**Gravidade:** Alta. **Situação:** Corrigido nos fontes.

**Localização:** [Cad_NFE.frm](<./Form/Cad_NFE.frm>), linha 435.

**Evidência:** Cadastro filial e transferência em Cad_NFE_Saida gravam L sem definir texto localmente. Teste COM em formato padrão converteu a chave fictícia para Double.

**Tratamento:** NumberFormat = "@" antes da atribuição como String em ambos os pontos. Teste COM de operação equivalente preservou 44 dígitos e zeros após reabertura; eventos VBA não executados.

### A04 — Leitura de ListBox sem linha selecionada

**Gravidade:** Média. **Situação:** Corrigido nos fontes.

**Localização:** [Cad_NCM.frm](<./Form/Cad_NCM.frm>), linha 76.

**Evidência:** NCM lê List(-1,...). Nas saídas, ListCount > 0 não garante ListIndex >= 0.

**Tratamento:** Guarda antes da leitura nas três listas. Guardas também impedem filtro sem coluna selecionada. Nomes dos controles conferidos no FRX.

### A05 — Redimensionamento repetido por coluna em três filtros

**Gravidade:** Baixa. **Situação:** Corrigido nos fontes.

**Localização:** [Cad_NCM.frm](<./Form/Cad_NCM.frm>), linha 229.

**Evidência:** PreencheLista do NCM e PreencheLista/PreencheLista2 de saídas executam ReDim Preserve para cada célula da linha encontrada.

**Tratamento:** Uma expansão por registro e cache da quantidade de colunas; dimensões, cabeçalho, ordem e término na primeira célula vazia preservados. Modelo JS: 35 cenários equivalentes; tempo no VBA não medido.

### A06 — Ribbon pode salvar outra pasta ativa

**Gravidade:** Alta. **Situação:** Corrigido nos fontes.

**Localização:** [MóduloOfficeRibbonEditor.bas](<./MDL/MóduloOfficeRibbonEditor.bas>), linha 17.

**Evidência:** ChamarSalvar usa ActiveWorkbook.Save apesar de pertencer ao Ribbon desta pasta.

**Tratamento:** Alterado somente esse callback para ThisWorkbook.Save. Assinatura preservada; teste com duas pastas ainda pendente.

### A07 — Módulo sem Option Explicit

**Gravidade:** Baixa. **Situação:** Corrigido nos fontes.

**Localização:** [modPreencherFormulario.bas](<./MDL/modPreencherFormulario.bas>), linha 1.

**Evidência:** Único fonte exportado sem declaração explícita obrigatória.

**Tratamento:** Option Explicit adicionado; parâmetros, tipos e mapeamentos existentes preservados. Agora 32/32 fontes têm a diretiva.

### A08 — Exclusões parciais podem quebrar fórmulas e vínculos

**Gravidade:** Crítica. **Situação:** Pendente.

**Localização:** [Cad_NFE.frm](<./Form/Cad_NFE.frm>), linha 217.

**Evidência:** Delete Shift:=xlUp em A:M da filial deixa N:O fora do deslocamento. Há 363 fórmulas em N e 364 em O na pasta original. Reprodução fictícia em Excel produziu =#REF! em N2 e sentinela desalinhada em O2. Há sete pontos Delete, incluindo duas exclusões de linha inteira.

**Tratamento:** Não ampliar nem substituir exclusões automaticamente. Definir política de baixa por registro, preservar fórmulas e validar todas as abas com estoque fictício antes de corrigir.

### A09 — Gravações em várias abas sem transação ou reversão

**Gravidade:** Crítica. **Situação:** Pendente.

**Localização:** [Cad_NFE_Saida.frm](<./Form/Cad_NFE_Saida.frm>), linha 99.

**Evidência:** GravarFilial escreve em Saidas Filial, Todas Saidas Filial, Entradas e Entradas Anual antes de localizar/apagar a entrada. CDbl/CDate podem falhar após gravações parciais. Cad_NFE também converte depois de gravar campos.

**Tratamento:** Pré-validar dados e destinos e desenhar recuperação com snapshot por registro. Requer testes de falha em cada etapa; não foi reescrito sem regressão funcional completa.

### A10 — Operações dependem da pasta ativa

**Gravidade:** Alta. **Situação:** Pendente.

**Localização:** [Cad_NFE.frm](<./Form/Cad_NFE.frm>), linha 132.

**Evidência:** 40 linhas com Sheets/Worksheets sem qualificação e oito referências ActiveWorkbook remanescentes. Alguns usos podem ser intencionais, mas há referências em exclusões e salvamentos.

**Tratamento:** Priorizar qualificação por proprietário nos fluxos destrutivos e testar duas pastas abertas. Não substituir globalmente usos de navegação intencional da célula ativa.

### A11 — Navegação aponta para aba inexistente

**Gravidade:** Média. **Situação:** Pendente.

**Localização:** [mdlChamarTabCan.bas](<./MDL/mdlChamarTabCan.bas>), linha 5.

**Evidência:** Cancelado, Estorno e Devolução não existe entre as 17 planilhas. O código não identifica uma substituição inequívoca.

**Tratamento:** Investigar o vínculo legado com o usuário; nenhum nome de aba foi inventado e a macro pública não foi removida.

### A12 — Três macros de atualização referem conexões ausentes

**Gravidade:** Média. **Situação:** Pendente.

**Localização:** [mdlMacroConexão.bas](<./MDL/mdlMacroConexão.bas>), linha 9.

**Evidência:** O pacote não contém xl/connections.xml. Os módulos de conexão usam Consulta - Tutores, Consulta - Estudantes e Consulta - Documentos para fora.

**Tratamento:** Confirmar se são integrações descontinuadas ou dependentes de outra pasta. Não atualizar conexões externas nem criar bancos/tabelas fictícios para contornar a ausência.

### A13 — Cancelamento dispensa leitura e validação do XML

**Gravidade:** Alta. **Situação:** Pendente.

**Localização:** [Cad_NFE_Saida.frm](<./Form/Cad_NFE_Saida.frm>), linha 560.

**Evidência:** operacaoCancelada evita CarregarXMLNFe e ValidarXMLNFe; os demais ramos mantêm gravação. O validador já avisa que XML normal não comprova cancelamento.

**Tratamento:** Regra existente preservada. Definir a evidência de evento de cancelamento e critérios de estoque com aprovação específica; a auditoria não atesta a correção fiscal dessa decisão.

### A14 — Documento XML fornecido diretamente evita verificações do carregador

**Gravidade:** Média. **Situação:** Pendente.

**Localização:** [modValidacaoXML.bas](<./MDL/modValidacaoXML.bas>), linha 155.

**Evidência:** O parâmetro documentoXML pode entrar sem a exigência de exatamente um infNFe aplicada em CarregarXMLNFe. XPaths globais podem misturar grupos de documentos compostos. O fluxo normal que usa o carregador já limita a cardinalidade.

**Tratamento:** Adicionar testes de documento composto e validação estrutural comum aos dois caminhos antes de modificar o contrato. Não houve validação XSD/assinatura digital nesta auditoria.

### A15 — Somatório ignora nós numéricos inválidos

**Gravidade:** Média. **Situação:** Pendente.

**Localização:** [modValidacaoXML.bas](<./MDL/modValidacaoXML.bas>), linha 1221.

**Evidência:** SomarNosXML adiciona somente os valores convertíveis e não sinaliza explicitamente os demais. Um item não selecionado pode ficar fora da soma sem erro próprio.

**Tratamento:** Criar fixture de múltiplos itens com número inválido e conferir o relatório. Ajustar a propagação de erros sem mudar tolerâncias nem tributos, depois de validar a regressão.

### A16 — Subprocesso sem limite de espera

**Gravidade:** Média. **Situação:** Pendente.

**Localização:** [modSelTest.bas](<./MDL/modSelTest.bas>), linha 423.

**Evidência:** RunAndCapture espera execObj.status sair de zero, usando DoEvents sem timeout; StdOut só é drenado depois.

**Tratamento:** Implementar prazo e captura progressiva de saída/erro, encerrando apenas o processo criado pelo teste. Não executar ferramentas externas reais nesta rodada.

### A17 — Espera ocupa CPU e permite reentrada por DoEvents

**Gravidade:** Baixa. **Situação:** Pendente.

**Localização:** [modSeleniumCommon.bas](<./MDL/modSeleniumCommon.bas>), linha 62.

**Evidência:** SleepMs e variantes usam espera ativa. Os cálculos já compensam meia-noite, mas a espera continua ocupando o thread.

**Tratamento:** Considerar Sleep da API com intervalos curtos e proteção de reentrada, preservando compatibilidade e testando cancelamento/timeouts. Sem benchmark real nesta rodada.

### A18 — Backup legado salva a pasta ativa e reutiliza destino

**Gravidade:** Alta. **Situação:** Pendente.

**Localização:** [mdlCriaBackup.bas](<./MDL/mdlCriaBackup.bas>), linha 5.

**Evidência:** As rotinas chamam .Save antes de SaveCopyAs; o primeiro destino é .bak fixo e o segundo D:\nome. Não há versionamento de cada execução nem verificação de conteúdo.

**Tratamento:** O backup desta auditoria é separado, datado e verificado. Modernizar a rotina do produto após definir destino autorizado e regra para documento ainda não salvo.

### A19 — Condição xlNotBetween calculada incorretamente

**Gravidade:** Média. **Situação:** Pendente.

**Localização:** [mdlCondCel.bas](<./MDL/mdlCondCel.bas>), linha 122.

**Evidência:** A expressão Not(valor <= mínimo) And (valor >= máximo) não representa valores fora do intervalo: falha abaixo do mínimo e inclui o máximo. A avaliação de fórmulas também depende do contexto ativo.

**Tratamento:** Confirmar chamadas por fórmulas e botões e testar mínimo-1, mínimo, meio, máximo e máximo+1 contra a formatação do Excel antes de corrigir o utilitário. Nenhum UDF foi removido.

### A20 — Logs e caminhos operacionais fora da pasta do projeto

**Gravidade:** Baixa. **Situação:** Pendente.

**Localização:** [modSeleniumCommon.bas](<./MDL/modSeleniumCommon.bas>), linha 86.

**Evidência:** WriteLog usa C:\Temp; o driver tem caminho fixo em Program Files. Logs de consultas podem conter identificadores pessoais.

**Tratamento:** Definir destino e retenção locais autorizados e avaliar sanitização. As consultas de Detran, rede, clipboard e gravação desses logs não foram executadas.

## Organização e desempenho

As declarações dos três filtros alterados foram organizadas; a quantidade de colunas é obtida uma vez e a matriz cresce uma vez por registro. Com cinco colunas, as expansões no modelo passaram de 5R para R, mantendo resultados em 35 cenários. O tempo no VBA não foi medido; a realocação por registro ainda pode ser quadrática. Uma futura leitura em arrays exige testar a parada na primeira célula vazia, preservada nesta revisão.

O módulo de pesquisa já usa leitura em blocos para arrays. Evitou-se reformatar indiscriminadamente, alterar assinaturas, remover eventos semelhantes ou trocar a regra de primeira célula vazia por última linha sem regressão real. Há 106 linhas com .Select/.Activate e 58 com On Error Resume Next na revisão; são indícios contextuais, não erros automáticos.

## Testes e limites

- **6 verificações reais no Excel 16.0 por COM**, com resultados esperados; duas reproduzem riscos. Chaves de texto, fórmula e formato persistiram após salvar/reabrir arquivo fictício.
- **6 verificações reais no MSXML 6.0 por COM**, com XML sintético: namespace, chave, rejeição de malformação/DTD e identificação de dois infNFe. Não equivalem à execução de ValidarXMLNFe nem à validação fiscal.
- **10 verificações estáticas/JS**, todas aprovadas, incluindo 35 cenários de equivalência do filtro.
- **Compilação real e execução VBA: não realizadas.** O Office retornou “O acesso de programação ao projeto do Visual Basic não é confiável”. A importação não ocorreu. Reflection revelou a exceção que a propriedade direta retornava como null.
- **Word:** a automação ficou pendente antes de produzir evidência de documento. A instância oculta do teste foi identificada e encerrada. Nenhum PASS funcional Word; não há Word no projeto para regressão.
- **Selenium, Detran, bancos e movimentações fiscais/estoque:** não executados. Não houve consulta externa, autenticação ou operação fiscal real.

O [TESTES_VBA.md](<./TESTES_VBA.md>) diferencia modalidades, bloqueios, evidências e roteiro manual. Não há garantia de compilação, compatibilidade integral de Office 32/64 bits nem regressão funcional completa.

## Continuidade

1. Priorizar os dois achados críticos em fixture de estoque fictício, com fórmulas sentinela em todas as abas.
2. Concluir importação e compilação no Excel. A alternativa manual é o editor VBA com as proteções atuais. Qualquer mudança de segurança exige autorização específica; não há autorização implícita para contornar o bloqueio.
3. Validar Venda, Transferência, Devolução e Cancelamento com o responsável, sem alterar tributos durante os testes.
4. Rodar a suíte preparada e os testes de falha/recuperação; comparar com a base antes de incorporar os fontes em novo XLSM de desenvolvimento.
5. Investigar integrações legadas antes de remover código. Versionar fontes, scripts e relatórios; não publicar backups ou novos binários com dados.
