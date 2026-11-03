# Histórico de alterações — 31/10/2026

## Execução 20261031_174936Z

Base Git observada: 2ae7792, com histórico anterior 67c9c11 e 94bfbaf. A pasta já era um repositório Git. Não foram feitos commits, push, publicação nem instalação de programas. Os fontes e o XLSM originais permaneceram intactos. O status inicial foi obtido enquanto a pasta de auditoria já estava sendo criada; a integridade da base é demonstrada pelos hashes, não por inferência desse status.

### Preparação

- Inventariados 39 arquivos; copiados para [backup/](<./_auditoria/20261031_174936Z/backup/>) e [desenvolvimento/](<./_auditoria/20261031_174936Z/desenvolvimento/>).
- Conferidos hashes SHA-256 dos 39 backups antes de alterar qualquer fonte.
- Extraídos localmente os 50 componentes embutidos e comparados os 32 fontes exportados com o XLSM.
- Criada pasta de testes fora do OneDrive; nenhum dado de estoque/cliente/XML real foi levado ao Office para execução de macros.

### Fontes alterados somente na cópia de desenvolvimento

| Fonte | Alterações |
|---|---|
| [Cad_NCM.frm](<./_auditoria/20261031_174936Z/desenvolvimento/Form/Cad_NCM.frm>) | Guarda ListIndex, guarda de coluna, declarações organizadas, cache de quantidade de colunas e ReDim por registro. |
| [Cad_NFE.frm](<./_auditoria/20261031_174936Z/desenvolvimento/Form/Cad_NFE.frm>) | Formato texto e atribuição String na chave de acesso filial. |
| [Cad_NFE_Saida.frm](<./_auditoria/20261031_174936Z/desenvolvimento/Form/Cad_NFE_Saida.frm>) | Célula da chave qualificada; quatro referências inválidas a Frame_ImpostoFilial comentadas; eventos preservados; chave como texto na transferência; guardas das duas listas e otimização dos dois filtros. |
| [modPreencherFormulario.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/modPreencherFormulario.bas>) | Option Explicit, sem alterar mapeamentos ou assinaturas. |
| [MóduloOfficeRibbonEditor.bas](<./_auditoria/20261031_174936Z/desenvolvimento/MDL/MóduloOfficeRibbonEditor.bas>) | ChamarSalvar usa ThisWorkbook.Save. |

São 7 grupos de problemas corrigidos em 5 arquivos, registrados em 16 blocos de substituição de [alteracoes.json](<./_auditoria/20261031_174936Z/alteracoes.json>). Os cinco FRX, o módulo XML, os cálculos tributários dos formulários e os 647 procedimentos exportados foram preservados. Não houve remoção ou renomeação de evento, procedimento ou controle.

### Verificação efetiva

- 10 verificações de regressão estática/JS aprovadas, incluindo 35 cenários do modelo de alocação.
- 6 casos reais COM no Excel 16.0 aprovados conforme expectativas, incluindo reproduções de perda de chave e de desalinhamento/#REF! por exclusão parcial.
- 6 casos reais COM no MSXML 6.0 aprovados com conteúdo sintético.
- Quatro scripts PowerShell passaram na análise sintática do parser instalado; isso não comprova seu comportamento em execução.
- Acesso programático ao projeto VBA bloqueado pelo Office; importação, compilação e suíte VBA não executadas. O XLSM da cópia permanece original.
- Sondagem Word pendente encerrada somente na instância oculta criada pelo teste; nenhum teste Word declarado aprovado.
- Nenhuma alteração de política de execução, Trust Center, AccessVBOM, antivírus, credenciais ou regra fiscal.

### Arquivos adicionados

- [AGENTS.md](<./AGENTS.md>), [RELATORIO_AUDITORIA.md](<./RELATORIO_AUDITORIA.md>), [MAPA_PROJETO.md](<./MAPA_PROJETO.md>), [TESTES_VBA.md](<./TESTES_VBA.md>) e este histórico.
- [.gitignore](<./.gitignore>): exclui backups, fontes extraídos duplicados e nova cópia do binário com dados do versionamento.
- [scripts/](<./_auditoria/20261031_174936Z/scripts/>): analisador, verificações estáticas, testes COM e suíte VBA preparada.
- [LEIA_ANTES_DE_USAR.md](<./_auditoria/20261031_174936Z/desenvolvimento/LEIA_ANTES_DE_USAR.md>): aviso de importação/compilação pendente.
- Manifestos, mapa de fórmulas, referências declaradas e resultados estruturados de testes na [pasta desta execução](<./_auditoria/20261031_174936Z/>).

### Pendências

13 grupos permanecem abertos, incluindo 2 críticos. O [RELATORIO_AUDITORIA.md](<./RELATORIO_AUDITORIA.md>) informa evidência, gravidade, localização e próxima ação. O [TESTES_VBA.md](<./TESTES_VBA.md>) contém o roteiro de compilação e regressão segura. Não promover esta revisão como aplicação validada até concluir essas verificações.

## 31/10/2026 — fase 2, execução 20261031_201621Z

Continuação da execução 20261031_174936Z, cujos 155 arquivos foram preservados. Leitura integral dos seis documentos obrigatórios e verificação por existência/hash precederam alterações.

1. Conferidos 39 originais/backs da fase 1, cinco revisões A01–A07 e replay exato dos 16 blocos anteriores; não houve modificação posterior detectada.
2. Criados [backup de 46 arquivos](./_auditoria/fase2_20261031_201621Z/backup/) e [manifesto SHA-256](./_auditoria/fase2_20261031_201621Z/manifesto-inicial-fase2.json); preparada nova [cópia de desenvolvimento](./_auditoria/fase2_20261031_201621Z/desenvolvimento/) e pasta de testes fora do OneDrive.
3. Tentativa de VBProject em pasta nova recusada pelo Office, HRESULT 0x800A03EC. Zero importações/compilações/macros. Não foram alteradas proteções. [Evidência](./_auditoria/fase2_20261031_201621Z/compilacao-probe.json).
4. Reproduzidos os sete Delete com sentinelas fictícias; protótipos mostraram troca de identidade por compactação e lacuna de filtro por limpeza. Solicitada decisão A08; comandos preservados.
5. Acrescentados clsOperacaoPlanilha/modIntegridadeOperacao e pré-validação/compensação limitada nos quatro gravadores. Limites Delete/Save/confirmação explícitos; edição anual e exclusões estruturais continuam pendentes.
6. Qualificadas referências conhecidas por ThisWorkbook; adicionados diagnósticos em navegação e conexões legadas; preservado contexto ativo de conexões/navegação intencional.
7. Ajustados estrutura/números XML, timeout de subprocesso próprio, espera Sleep, backup versionado e condição xlNotBetween/leitura de limites; nenhuma regra fiscal especializada alterada.
8. Preparador ampliado para classe, CodeNames/documentos e referências; criados esquema e suíte VBA de falhas. Não importados/compilados/executados. [Roteiro](./_auditoria/fase2_20261031_201621Z/COMPILACAO_MANUAL.md).
9. Testes finais: 12 verificações estáticas novas, 10 da suíte anterior com 35 cenários JS, 12 COM anteriores, 70 COM de recuperação, 20 COM de backup/contexto/condições/timeout, 18 de XML/condições e sete reproduções A08/dois protótipos. Modalidades e sobreposição constam de [TESTES_VBA](./TESTES_VBA.md); nenhum PASS funcional VBA.
10. Preservadas falhas intermediárias de marshaling, formato COM, hash bloqueado, homônimo de workbook e asserções/agrupador de testes; as evidências não foram reclassificadas retroativamente.
11. Atualizados relatório, mapa, testes, histórico e criado [plano por A01–A20](./PLANO_CORRECAO_FASE2.md). AGENTS e README mantidos; nenhuma nova regra permanente necessária.

Os 17 fontes existentes alterados e dois novos estão rastreados em [alteracoes-fase2.json](./_auditoria/fase2_20261031_201621Z/alteracoes-fase2.json), com blocos antes/depois. Para reversão, usar somente a cópia de desenvolvimento e os backups conferidos, preservando FRM/FRX/codificação. O XLSM da cópia segue intacto e não incorpora alterações.

**Entrega pendente de integração e homologação:** A08/A09 críticos, compilação/regressão VBA bloqueadas, Cancelado e integrações legadas exigindo decisão. Não houve commit, push, publicação, promoção, consulta externa, movimentação fiscal real ou alteração de segurança.

## 31/10/2026 — Fase 3, execução 20261031_213925Z

- Lidos integralmente os sete documentos requeridos; arquivos citados localizados por existência e hash. Confirmados 39 originais/backups da Fase 1, 155 arquivos anteriores, 146 da Fase 2 e 46 backups. Copiados 42 arquivos de desenvolvimento sem divergência; [novo backup](./_auditoria/fase3_20261031_213925Z/backup/) com 69 arquivos, conferido antes de alterações.
- Novo acesso VBProject recusado, 0x800A03EC. Registro consultado somente para diagnóstico de bibliotecas, sem modificações/instalações. [Sondagem](./_auditoria/fase3_20261031_213925Z/office-fase3.json). Esqueleto XLSM com 17 abas e pacote manual CP1252/FRX preparados em TEMP, fora do OneDrive; nenhuma integração/compilação/execução VBA declarada.
- C3-01, subitem A09: uso único do contexto e recusa de recuperação inativa. C3-02, subitem A09: preservação de Errors(xlNumberAsText).Ignore, após perda comprovada em COM. Somente uma classe de produto alterada; 665 assinaturas existentes preservadas. [Blocos](./_auditoria/fase3_20261031_213925Z/alteracoes-fase3.json).
- A suíte VBA existente foi ampliada para ciclo de vida/metadado, sem executar ou renomear eventos. Fases 1/2, FRM/FRX, fiscalidade e sete Delete preservados.
- A08: quatro experiências COM compararam três estratégias; exclusão física e limpeza/reutilização mantêm riscos. C é recomendação condicionada, sem identidade completa/persistência nem integração. Estatísticas OOXML agregadas demonstraram repetição na coluna A dos históricos e chave incompleta na filial; nenhum valor real de nota/placa/chave foi exportado.
- A09: 15 experiências de fronteira, com 200 células por snapshot; dez recompuseram campos, quatro exigiram conciliação e uma reproduziu recuperação incompleta. Não são eventos VBA. Corrigido metadado passou em quatro tipos de célula COM. XML: 12 cenários de parser/modelo; backup: cinco cenários COM com restauração/erros esperados. [Matriz](./MATRIZ_TESTES_FASE3.md).
- Criados [plano](./PLANO_CORRECAO_FASE3.md), [relatório](./RELATORIO_HOMOLOGACAO_FASE3.md), [matriz](./MATRIZ_TESTES_FASE3.md) e [pendências](./PENDENCIAS_DECISAO.md). Relatórios anteriores preservados no backup e em seus registros históricos. README, AGENTS e plano da Fase 2 mantidos.
- Nenhum crítico foi encerrado. Sistema NÃO HOMOLOGADO; fontes/pacote preparados para a etapa manual, com XLSM base intacto. Sem commit, push, publicação, consulta externa, segurança alterada, retenção automática ou execução de operações reais.


## 31/10/2026 — Fase 4, execução 20261031_230457Z

- Lidos integralmente os nove documentos requeridos e conferidos os artefatos da Fase 3. 211 comparações SHA-256 sem divergência: 39 originais, 42 arquivos de desenvolvimento, 69 backups, 42 itens de staging, 18 corpos de documento e fixture. [Manifesto](./_auditoria/fase4_20261031_230457Z/manifesto-inicial-fase4.json).
- Criados e verificados [114 backups](./_auditoria/fase4_20261031_230457Z/backup/) antes das atualizações; cópias de fonte/FRX idênticas à Fase 3 e nova fixture/pacote em TEMP fora do OneDrive. Nenhum código de produto/teste alterado; nenhum XLSM com dados reais aberto no Office.
- Não repetida automação VBProject nem testes equivalentes. Compilação e execução VBA continuam pendentes da intervenção solicitada; contagem real VBE não aferida e zero revisados comprovadamente integrados.
- Criados [relatório](./RELATORIO_INTEGRACAO_FASE4.md), [roteiro único](./COMPILACAO_VBA_FASE4.md), [matriz](./MATRIZ_REGRESSAO_VBA_FASE4.md) e [proposta de identidade](./PLANO_IDENTIDADE_REGISTROS.md). Inventário nominal preparado distingue 55 esperados de importação efetiva não comprovada.
- Preservadas as três suítes existentes. Roteiro instrui compilar, executar funções e devolver saídas/erros/contagens/exportados; não pressupõe que VBE, referências/ActiveX ou macros estejam permitidos.
- D01–D05 preservadas sem aprovação implícita. Especificados episódio de estoque, documento/item, veículo, movimentação, legado, retorno e leitores; nenhum campo/tabela/status/identificador ou nova política integrado.
- Atualizados os pontos de continuidade em relatório, mapa, testes, pendências e histórico; textos anteriores e backups preservados. AGENTS, README e planos anteriores mantidos.
- Sistema NÃO HOMOLOGADO, A08/A09 ainda críticos. Sem commit, push, publicação, instalações, segurança alterada, retenção automática, consultas externas ou operação sobre dados reais.


## 01/11/2026 — tentativa assistida pela interface (20261101_143216Z)

O usuário autorizou controle do computador e uso contínuo do Excel em fixtures. Foram preservados nove arquivos/backup da fixture, com hashes conferidos, antes de qualquer edição. [Manifesto](./_auditoria/fase4_20261031_230457Z/interface_20261101_143216Z/manifesto-inicial.json). A tentativa usou a interface Windows e não repetiu o acesso programático ao VBProject.

O controle localizou Excel e o diálogo Abrir. Houve expiração/reconexão, recusa do índice de edição em cache e recusa de clique por divergência entre a janela principal e modal. Após atualização/ativação e tentativa por teclado, o foco permaneceu na Caixa de Pesquisa; o campo Nome do arquivo não recebeu texto. Nenhum VBA foi importado, compilado ou executado. A fixture continua com SHA-256 inicial; isso não comprova compilação. [Erros exatos](./_auditoria/fase4_20261031_230457Z/interface_20261101_143216Z/resultado-interface.json).

**BLOQUEADO por controle da janela modal**, separado do bloqueio VBProject histórico. Foi solicitado somente abrir a fixture já preparada e deixá-la em primeiro plano, para retomar o VBE; não se solicitou nova aprovação nem mudança de segurança. A instância Excel não foi encerrada e nenhuma instância do usuário foi fechada. Continuam NÃO INTEGRADO/NÃO HOMOLOGADO e A08/A09 críticos.


## 01/11/2026 — fixture e VBE observados (20261101_150908Z)

Após a confirmação do usuário, a interface mostrou a fixture com as 17 abas esperadas e o VBE com **18 componentes reais, todos documentos**. Zero módulos padrão, classes ou UserForms revisados presentes. Os 17 CodeNames ainda são os nomes sequenciais da fixture, diferentes do mapeamento exigido; não foram ajustados. A contagem esperada de 55 não foi usada como resultado de importação. [Inventário observado](./_auditoria/fase4_20261031_230457Z/interface_20261101_150908Z/resultado-vbe.json).

Ctrl+M abriu Importar arquivo, e o campo Nome recebeu o caminho conferido de clsOperacaoPlanilha.cls. A primeira entrada tinha escape de barras incorreto; foi corrigida com separadores / antes de qualquer confirmação. O clique em Abrir foi recusado pela ferramenta: `point (1560, 1068) is over EXCEL.EXE "&Abrir", not target window EXCEL.EXE "Microsoft Visual Basic for Applications - Fixture_Auditoria_Fase4.xlsm"`. Enter/Alt+A mantiveram o diálogo aberto. Houve também `coordinate input geometry is unavailable` e timeout de SetValue; a atualização posterior confirmou o caminho correto, mas não a importação.

A consulta de caminho na Janela Imediata foi tentada, porém não houve saída verificável; não é PASS de VBA. Foram observados o mesmo bloqueio de geometria/ativação tanto em sky quanto no mecanismo cua_repl. Não se alterou segurança, não houve novo acesso VBProject e não se executaram consultas externas. **Importação revisada: 0; compilação: BLOQUEADO; suítes VBA: 0 execuções; sistema NÃO HOMOLOGADO.**

O arquivo da fixture em disco continua com o SHA-256 inicial `ff4f6b37d529e7404ae31bc65ceb1aeb3ba8f54aeb007b71c1e47c9bb9b9585d`. Seis arquivos foram respaldados e conferidos antes desta atualização, em [backup](./_auditoria/fase4_20261031_230457Z/interface_20261101_150908Z/backup/), com [manifesto](./_auditoria/fase4_20261031_230457Z/interface_20261101_150908Z/manifesto-backup.json). A interface ficou no diálogo de importação, sem exclusões, salvamento ou fechamento de instâncias do usuário. Para continuar, a confirmação/importação pelo VBE exige ação manual enquanto o controle de janelas não operar corretamente; seguir o [roteiro único](./COMPILACAO_VBA_FASE4.md).
