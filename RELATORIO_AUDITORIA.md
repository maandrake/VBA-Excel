> Atualização efetiva de 01/11/2026: a fixture da Fase 4 foi integrada e compilada no Excel real. Três suítes concluídas: 9 + 377 + 11 verificações PASS; recuperação ampla protegida continua FAIL e A08/A09 críticos. Sistema NÃO HOMOLOGADO. O [relatório atual](./RELATORIO_INTEGRACAO_FASE4.md) e a [matriz](./MATRIZ_REGRESSAO_VBA_FASE4.md) substituem os estados de importação bloqueada descritos historicamente abaixo.

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


## Adendo Fase 4 — 01/11/2026: evidência real

Importação e compilação reais concluídas na fixture (55 componentes, 17 CodeNames). Correções C4-01–C4-03 em testes/integração e C4-04 em RestaurarCelula: remover ClearContents anterior à restauração do formato evita limpar conteúdo quando formatação falha. Proteção ampla ainda causa quatro falhas de compensação; não classificar A09 resolvida. Sete Delete permanecem literais, A08 depende de política/identidade aprovada. D03/D04 permanecem pendentes; nenhuma regra fiscal/log real alterado.

As suítes 9/377/11 PASS testam rotinas reais nos limites descritos, sem provar os botões de gravação. Backup/callback Ribbon/pesquisa matriz/ListBox e rejeições XML tiveram execução VBA direta; reabertura de backups foi COM e está separada. Ciclo de vida bloqueado pelo depurador; XML completo, gravadores e demais ramos ainda não executados. [Detalhamento e correções](./RELATORIO_INTEGRACAO_FASE4.md), [matriz](./MATRIZ_REGRESSAO_VBA_FASE4.md). Nenhum novo inventário geral de 20 grupos foi fabricado; há dois críticos abertos, com redução técnica parcial de A09.


## Continuação — Fase 5 (20261101_202114Z)

[Relatório dos testes integrados reais](./RELATORIO_TESTES_INTEGRADOS_FASE5.md), [matriz e evidências](./MATRIZ_REGRESSAO_FASE5.md), [mapa A08/A09](./MAPA_INTEGRIDADE_ESTOQUE.md), [recuperação](./PLANO_RECUPERACAO_FALHAS.md). Base sintética compilada no VBE com 56 componentes/17 planilhas. Eventos de negócio: 20 PASS, cinco FAIL, um BLOQUEADO para injeção de Save; 13 casos XML PASS, suítes auxiliares reais 377+11 asserções PASS e classe com duas pastas PASS. Ciclo de vida BLOQUEADO; recuperação incompleta anterior permanece FAIL. Venda/Transferência/Devolução não preservaram integridade após a baixa. Originais e fontes herdados intactos; nenhum dos sete Delete ou regra fiscal modificado. D01/D02 pendentes e sistema NÃO HOMOLOGADO. Resultados anteriores preservados; consultar matriz para modalidades e testes não repetidos.


## Continuação Fase 6 — 20261101_224949Z

A análise estrutural foi limitada às causas e dependências A08/A09. Os cinco FAIL F5 continuam mantidos. Protótipos VBA compilados compararam Delete, limpeza e baixa lógica com IDs declarados: 43 cenários únicos, 28 PASS / 15 FAIL, além de dois checks de guarda. C é recomendação condicionada; referências físicas, pesquisa/contagem/fórmula por placa e recuperação protegida ainda apresentam FAIL. Não houve correção desses pontos no produto. [Protótipos](./PROTOTIPOS_INTEGRIDADE_FASE6.md), [matriz](./MATRIZ_REGRESSAO_FASE6.md) e [decisões](./DECISOES_ARQUITETURA_FASE6.md).


## Continuidade — Fase 7 (20261102_025115Z)

A08/A09 permanecem críticos. Contrato de identidade e migração conjunta preparados em Portão1; nenhuma política/gravador principal alterado. Quatro FAIL experimentais/legados preservados e contenção fria separada de recuperação integral. [Relatório](./RELATORIO_TECNICO_FASE7.md), [matriz](./MATRIZ_REGRESSAO_FASE7.md), [contrato](./CONTRATO_IDENTIDADE_FASE7.md), [mapa](./MAPA_MIGRACAO_LEITORES_FASE7.md) e [decisões](./PACOTE_DECISOES_FASE7.md). Checkpoints e evidências em [execução F7](./_auditoria/fase7_20261102_025115Z/). Portão2 não aprovado; NÃO HOMOLOGADO.


## Fase 8 — 02/11/2026: preparação experimental do Portão 2

Execução [_auditoria/fase8_20261102_121422Z](./_auditoria/fase8_20261102_121422Z/). Base anterior:2.106 hashes conferidos sem divergência na abertura e35 backups novos verificados. Originais e sete Delete operacionais preservados. Mestre experimental final:60 componentes/17 abas, compilação real no Excel16.0/20527/64bits; fontes herdados sem alteração de literais/corpos. Portão2 não executado; sistema NÃO HOMOLOGADO.

Quatro defeitos F7 reproduzidos e corrigidos somente nos protótipos: data/selo persistente com política explícita, anual por episódio, consulta por ID após sort/lacuna/reuso e auxiliares com valores discriminantes. Guarda adicional comprovada: recusar selo ERRO coincidente; três FAIL preservados e três PASS após correção.70 asserções únicas aprovadas na seleção de resultados experimentais e4 volumes medidos (100/1000/5000/10000), com limites por revisão na matriz.26 integrados repetidos após guarda final;15 guardas F7 reexecutados em sessão fria. Não houve teste novo de recuperação integral ou aprovação de gravadores reais.

[Relatório técnico](./RELATORIO_TECNICO_FASE8.md), [matriz](./MATRIZ_REGRESSAO_FASE8.md), [correções](./CORRECOES_PROTOTIPOS_FASE8.md), [armazenamento](./PROPOSTA_ARMAZENAMENTO_FASE8.md), [persistência](./CONTRATO_PERSISTENCIA_FASE8.md), [M0–M9](./PLANO_IMPLANTACAO_FASE8.md), [decisões Portão2](./PACOTE_APROVACAO_PORTAO2.md). Cinco FAIL oficiais F5 e FAIL integrais F6 continuam pendentes. Nenhuma promoção, migração real, alteração fiscal, instalação ou mudança de segurança.


## Fase 9 — 20261102_142753Z — persistência experimental

Portão experimental somente; Portão 2 NÃO aprovado e sistema NÃO HOMOLOGADO. Entrega F8 R2 e 2.268 hashes conferidos; 43 backups novos verificados. ACE 16/ADODB-ADOX x64 do Office instalado executou transações reais em bancos sintéticos no TEMP. Mestre F9 compilado no VBE: 63 componentes/17 abas, com os 60 herdados preservados e os sete Delete intactos. Nenhuma integração no principal, migração real, regra fiscal alterada, instalação, segurança modificada, commit/push ou promoção.

65 asserções VBA: 64 PASS e 1 FAIL de Ignore após Sort, preservado com prova corretiva separada. Quatro benchmarks VBA PASS e três testes COM entre processos PASS, sem somar sondagens/hashes. Ciclo de sete movimentos/recibos, rollback pré-commit, bloqueio de pendência, escritor/fencing e reconstrução em sessão fria exercitados. Erro real de SaveAs e cancelamento BeforeSave preservaram a autoridade. Falha física de commit, energia/crash/disco cheio, rollback incompleto e snapshot sob escrita concorrente NÃO EXECUTADOS. Os sete FAIL oficiais F5/F6 permanecem nas classificações anteriores.

[Relatório](./RELATORIO_TECNICO_FASE9.md), [matriz](./MATRIZ_REGRESSAO_FASE9.md), [armazenamento](./COMPARATIVO_ARMAZENAMENTO_FASE9.md), [esquema](./ESQUEMA_TRANSACIONAL_FASE9.md), [durabilidade](./TESTES_DURABILIDADE_FASE9.md), [concorrência](./TESTES_CONCORRENCIA_FASE9.md), [reconstrução](./PROTOCOLO_RECONSTRUCAO_EXCEL_FASE9.md), [M0–M4](./PLANO_MIGRACAO_M0_M4_FASE9.md), [decisões](./DECISOES_PENDENTES_FASE9.md).


## Adendo Fase10 — laboratório transacional (02/11/2026)

A08/A09 permanecem críticos no produto. A Fase10 entregou candidato sintético schema4/66 componentes,121 asserções VBA únicas PASS,12 medições benchmark PASS e10 experiências/verificações COM PASS. Bloqueio de macros impediu seis verificações frias VBA; teste COM não as substitui. Novas guardas de intenção/custódia e reconstrução de metadados foram comprovadas somente no laboratório. Não foram alterados os sete Delete nem os gravadores reais. [Resultados, falhas preservadas e riscos](./RELATORIO_TECNICO_FASE10.md).


## Atualização — Entrega 1 sintética, 02/11/2026

[Relatório de implementação](./ENTREGA_1_CANDIDATO_INTEGRADO.md) e [matriz](./MATRIZ_REGRESSAO_ENTREGA.md). Não é uma nova auditoria nem promoção. Na candidata, os fluxos Entrada/Venda/Transferência/Devolução usam ACE 16/esquema 5 e baixa lógica por identidade. Os sete Delete foram removidos apenas dessa cópia. Eventos reais demonstraram sucesso e contenção de falhas, com pendências de conciliação ainda sem encerramento operacional. No original, A08/A09 e os sete FAIL históricos não mudam de classificação.

Compilação real PASS; reabertura interativa BLOQUEADA no aviso de macros; implantação/rebind e restauração completa pendentes. Original XLSM SHA-256 2ad1c635da68b51a68d5d2c0d3cd638f1ff8064942b299699f2594d7a9ca8850, e 39 arquivos originais conferidos. Candidata NÃO HOMOLOGADA; [defeitos abertos](./REGISTRO_DEFEITOS_ABERTOS.md).

## 02/11/2026 — continuação da Entrega 1: reabertura real

Checkpoint verificado em ./_desenvolvimento/entrega1_20261102T182823Z/continuacao_reabertura_20261102T205915Z/. O usuário habilitou macros manualmente. E1_SessaoReaberta revelou FAIL por diferença entre aliases MARCOS~1/MARCOSALEXANDREPERDO do mesmo arquivo. Corrigido somente modCandidatoIntegrado.ConnC1 mediante C1_MesmoArquivoAutoridade (existência + ShortPath); BaseID e vínculo persistido preservados. O Editor VBA recompilou sem erro. Três verificações imediatas retornaram True/False/False (mesmo arquivo/cópia distinta/ausente). Nova execução na sessão aberta PASS e nova instância Excel /x PID 50212 confirmou CACHE_AUDITADO automático e E1_SessaoReaberta PASS, sem repetir movimentos.

Log acumulado 47 execuções de asserção: 40 PASS/7 FAIL; falhas anteriores preservadas. As três consultas imediatas e a compilação são evidências separadas, não somadas como cenários de UserForm. Banco SHA-256 inalterado (3198e07699f644f64c325a9fb517470b64950aa5a262cc354629ed61f1503dc3), revisão 8, oito movimentos/recibos, três ativos, zero Pending. XLSM corrigido SHA-256 5b2a0d55fcea077dd034ace4c4892bbc9fdd46c69760d7d318d2162f54771749. Contagem binária 69 componentes; corpos dos cinco forms iguais ao checkpoint (VB_Base reserializado). 39 originais e backup F10 novamente conferidos, sem divergência. E1-04 encerrado no escopo de reabertura normal; E1-03, conciliação, fluxo D e homologação permanecem pendentes. Nenhuma mudança de segurança, produção, regra fiscal, commit ou publicação.


## Continuidade — Entrega2 (02/11/2026)

Implementação dirigida na candidata sintética, sem nova auditoria geral. [Relatório](./ENTREGA_2_OPERACOES_CORRIGIDAS.md), [matriz de resultados reais](./MATRIZ_REGRESSAO_ENTREGA2.md), [fontes/inventário](./_desenvolvimento/entrega2_20261102T215146Z/fontes/), [revisão dos57 leitores](./_desenvolvimento/entrega2_20261102T215146Z/fontes/REVISAO_LEITORES_ENTREGA2.md) e [manifestos/evidências](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/).

69componentes/17abas/5forms;436assinaturas existentes dos formulários preservadas. Retificação auditável, reentrada por controles e classificação de exclusão integradas; oito componentes modificados funcionalmente. Revisão16/esquema6,16movimentos/recibos e zeroPending no mestre final. Forenses de falhas separados. ReferênciaR0:10PASS; continuação61execuções(54PASS/7FAIL preservados),55IDs. Compilação real final PASS, reabertura interativa E2 BLOQUEADA no aviso de macros. Análise estática/consultas ADO não substituem eventos. Sistema NÃO HOMOLOGADO; [pendências](./REGISTRO_DEFEITOS_ABERTOS.md) e [manual](./MANUAL_INSTALACAO_RESTAURACAO.md) prevalecem para uso da candidata. Originais e Entrega1 intactos.


### Continuação da reabertura E2

O bloqueio anterior foi superado após habilitação manual. Execução real VBA auxiliar E1_SESSAO_REABERTA_E2_R3 PASS, inicialização automáticaCACHE_AUDITADO, revisão16/projeçãoVERIFIED/16recibos/0Pending, hashes do par inalterados. [Evidência](./_desenvolvimento/entrega2_20261102T215146Z/continuacao_reabertura_20261103T010409Z/evidencias_reabertura.json). Log acumulado62execuções/55PASS/7FAIL; não são62cenários únicos. Nenhuma fonte ou operação de estoque foi alterada nesta continuação. Demais limites e NÃO HOMOLOGADA mantidos.


## Manutenção do repositório em 03/11/2026

O gitlink autorreferente foi removido em `c560de1`. Os cinco arquivos do backup adicional da Fase 4 foram realocados com SHA-256 conferido pelo [manifesto histórico](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/manifesto-backup-adicional.json). O [mapa de caminhos](./MAPA_PROJETO.md#backups-de-manutenção) permite localizar os arquivos pelos registros antigos.
