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


## 01/11/2026 — Fase 4: integração e execução reais (20261101_160244Z)

Após autorização para controle do Excel, importados pela interface VBE 37 componentes (34 do sistema + três de teste), configurados 17 CodeNames e 18 corpos de documento. Contagem real salva: 55 (30 BAS, duas classes, cinco UserForms, 18 documentos). Todos os corpos conferidos contra o pacote final, com normalização limitada descrita no [relatório](./RELATORIO_INTEGRACAO_FASE4.md). Depurar > Compilar foi executado e terminou sem erro após ajustes; estado compilado conferido após salvar.

Correções incrementais com backups/hashes: C4-01 sintaxe da comparação ExigirF2 na suíte; C4-02 normalização dos três limites TEMP curto/longo sem remover guardas; C4-03 reaplicação do corpo Option Explicit de Planilha5; C4-04 remoção do ClearContents redundante de clsOperacaoPlanilha.RestaurarCelula, que apagava dados antes de falhar em formatação protegida. Somente essa classe mudou entre os 34 fontes do sistema; demais 33 fontes/cinco FRX iguais à Fase 3. Nenhum Delete ou regra fiscal alterado.

Executados em VBA: suíte segura 9 PASS; suíte Fase 2, 377 PASS, repetida com outra pasta ativa e após C4-04; complemento A09, 11 PASS. Total 397 verificações únicas dessas suítes, não operações de estoque homologadas. Erro de ciclo de vida esperado foi interceptado pela preferência do depurador; sem PASS, opção não alterada. Teste de proteção ampla continua FAIL de recuperação integral, com quatro falhas e necessidade de conciliação; C4-04 evita apagar conteúdo adicional, sem prometer rollback. Foram carregados os cinco MSForms e testados pesquisa matriz/seleção/ListIndex=-1/sem resultado reais.

CriarBackupVersionado executado duas vezes com outra pasta ativa: duas versões distintas, dados em memória incluídos e origem sem salvamento inesperado. Ambas reabertas por COM somente leitura com fórmula/chave textual/Ignore preservados. Callback ChamarSalvar executado via Application.Run salvou proprietária e não salvou outra pasta ativa. XML real: quatro rejeições do validador e um parse malformado aprovados; demais casos fiscais e apresentação completa ainda não executados. Gravadores completos Venda/Transferência/Devolução/Cancelamento não executados nesta retomada.

Preservados por SHA-256 39 originais, 39 backups Fase 1, 46 Fase 2, 160 arquivos Fase 3 e 114 backups iniciais Fase 4. Novos manifestos/checkpoints antes das edições; relatórios anteriores respaldados. Sem alteração de segurança, instalação, consulta externa, commit/push ou promoção. A08/A09 críticos e D01–D04 abertos; NÃO HOMOLOGADO. [Evidências](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/), [matriz](./MATRIZ_REGRESSAO_VBA_FASE4.md), [entrega XLSM](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase4_20261031_230457Z/Entrega_Fase4_20261101/Fixture_Auditoria_Fase4_Integrada_Testada.xlsm>).


Complemento da mesma execução Fase 4: seis casos dos botões reais de entrada executados (cinco matriz: válida/duplicada/valor inválido/data inválida/número vazio; um filial válido), todos PASS nos critérios documentados. Matriz sem células alteradas nas recusas; filial preserva String 44/@, fórmulas sentinela M/N e referência anual. A:J corresponde aos dois históricos. A entrega fictícia final foi salva com SHA-256 dcbd72e46e8b182a9810254bcce460ea2cb17260b4bcb47b622ddc3d4f9879a0. Fontes dos gravadores não modificados; não equivale a testar Venda/Transferência/Devolução/Cancelamento nem falhas entre suas gravações. [Evidências](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/entradas-vba-real.json).


## Fase 5 — 20261101_202114Z — testes integrados reais

- Base Fase 4 localizada e preservada, SHA dcbd72e46e8b182a9810254bcce460ea2cb17260b4bcb47b622ddc3d4f9879a0. Checkpoint datado/manifesto conferido antes de alterações. Revalidados 39 originais, 329 itens da referência Fase 4 e 39 arquivos herdados de desenvolvimento antes de atualizar documentação.
- Criado somente componente de teste modTestesIntegradosFase5 (Windows-1252/CRLF), XML sintético e cópias por cenário em TEMP fora do OneDrive. Compilação real no VBE; removida duplicata de importação de teste; entrega com 56 componentes/17 planilhas, SHA f989289554a5c081cef8bd13de2e0b6f0e48ec4c1ba1dd5e75c349eff458f72c. Registro preliminar 57 preservado e explicitamente corrigido. Fontes de aplicação sem mudança semântica; serialização do Office/atributos documentados.
- Executados eventos reais: 26 cenários com snapshots; 20 PASS, cinco FAIL de integridade, um BLOQUEADO para injeção de Save. Asserções das suítes não somadas a cenários. Venda/Transferência/Devolução falharam após Delete; falha no Delete da devolução deixou quatro destinos e origem; transferência duplicada ambígua aceita. Quatro compensações limitadas e 15 rejeições comparadas em 7.560 células; segunda tentativa somente após recuperação verificada.
- Reexecutadas suítes VBA existentes: 377 + 11 asserções PASS; classe com outra pasta ativa PASS. Ciclo de vida BLOQUEADO por preferência de depuração preservada. FAIL de recuperação incompleta anterior mantido.
- Treze variantes XML/relatório reais retornaram o esperado. Cancelado preservou comportamento sem homologação fiscal. Sem dados reais, consultas externas, segurança alterada, instalações, commits ou publicações.
- Criados [relatório](./RELATORIO_TESTES_INTEGRADOS_FASE5.md), [matriz](./MATRIZ_REGRESSAO_FASE5.md), [mapa](./MAPA_INTEGRIDADE_ESTOQUE.md) e [plano de recuperação](./PLANO_RECUPERACAO_FALHAS.md). Atualizadas pendências e referências da auditoria. Backups/evidências preservados na [execução](./_auditoria/fase5_20261101_202114Z/). Fixtures XLSM excluídas do versionamento via regra local após backup de .gitignore.
- Reversão: checkpoint Fase 4 preservado; nenhum dos sete Delete ou regra fiscal foi modificado. A08/A09 críticos, D01/D02 sem decisão. **NÃO HOMOLOGADO; não substituir produção.**


## Fase 6 — 20261101_224949Z — arquitetura experimental

- Preservados 39 originais, referência F5 e 1.230 arquivos nas pastas das cinco fases; checkpoint/hash antes das alterações. 602 arquivos de execução sintética arquivados com cópias verificadas.
- Acrescentado somente módulo experimental em fixtures TEMP: IDs de veículo/episódio/movimentação/projeção/operação/tentativa, vínculos, versões, diário e bloqueio de conciliação. Fontes principais, FRM/FRX externos, sete Delete e regras fiscais preservados.
- Compilação real pelo VBE. Corrigidos erro de palavra reservada local, coleta Ignore em proteção e quatro chamadas Guardar interpretadas como rótulos, apenas no protótipo. R2 inclui selo de conteúdo e guarda de diário; R4 corrige três assertivas ausentes e acrescenta permutação preservando tipos.
- 43 cenários únicos consolidados: 28 PASS / 15 FAIL, mais dois checks de guarda PASS; 86 tentativas preservadas. Não somar repetições. Recuperação incompleta, leitores legados e referências físicas permanecem FAIL. Memória insuficiente/0x8000FFFF interrompeu uma reabertura R2; houve continuação em sessão Excel /x, sem kill global.
- Corrigidos metadados SHA intermediários que hashavam caminhos; cópias/manifesto inicial já usavam bytes. Versões dos metadados preservadas.
- Criados os seis documentos de arquitetura, impacto, protótipos, recuperação, regressão e decisões; pendências atualizadas. A aplicação continua NÃO HOMOLOGADA. [Resultados e evidências](./PROTOTIPOS_INTEGRIDADE_FASE6.md). Reversão documental pelo [checkpoint](./_auditoria/fase6_20261101_224949Z/backup/documentacao/); não houve migração no produto a desfazer.

A sessão R4 foi fechada normalmente. Duas janelas antigas de fixtures R2/C_COLISAO permaneceram após falha de memória, sem encerramento forçado; pendência operacional registrada nas evidências.


## Fase 7 — 20261102_025115Z — contrato e migração experimental

- Lidos os dez documentos obrigatórios e conferidos 1.986 hashes, sem divergência inicial; 39 originais e referência F6 preservados. Criado checkpoint de 28 arquivos, verificado antes de alterações. [Manifesto](./_auditoria/fase7_20261102_025115Z/manifesto-inicial-fase7.json).
- Contrato formal e mapa nominal de 57 procedimentos preparados antes de alterar fixtures. Nenhum esquema ou política novo aprovado para o principal.
- Acrescentado somente modPrototiposFase7 no mestre TEMP; compilação real pelo VBE, entrega com 58 componentes/17 abas. Os 57 componentes herdados preservam texto canônico, literais e sete Delete; FRM/FRX externos intactos.
- Adaptadores testados em VBA: 18 verificações PASS. Sessão fria em nova instância /x: 12 cenários/15 verificações PASS, sem nova escrita no escopo; 23 arquivos persistidos sem mudança de hash. Não houve crash/energia/concorrência ou recuperação integral comprovada.
- Preservadas tentativa interrompida por selo de data divergente e montagem inicial de fórmula textual. Preparação de formatos e reteste de fórmula somente nas cópias de teste. Quatro FAIL consolidados mantidos; primeiro PASS numérico do auxiliar não é prova de identidade. [Matriz e tentativas](./MATRIZ_REGRESSAO_FASE7.md).
- Criados os sete documentos da Fase 7, fontes/evidências e pacote de decisões. Atualizados pendências, auditoria, mapa e testes por acréscimo. .gitignore ampliado somente para excluir os XLSM sintéticos F7 do versionamento. Sem commit/push.
- As duas sessões exclusivas foram fechadas normalmente após conferir nome/caminho e um livro aberto; janela Excel preexistente mantida. Nenhuma segurança alterada, instalação, consulta externa ou operação real.
- Portão 2 aguarda aprovação; A08/A09 e cinco FAIL F5 permanecem abertos. NÃO HOMOLOGADO. [Relatório e restauração](./RELATORIO_TECNICO_FASE7.md).


## Fase 8 — 02/11/2026: preparação experimental do Portão 2

Execução [_auditoria/fase8_20261102_121422Z](./_auditoria/fase8_20261102_121422Z/). Base anterior:2.106 hashes conferidos sem divergência na abertura e35 backups novos verificados. Originais e sete Delete operacionais preservados. Mestre experimental final:60 componentes/17 abas, compilação real no Excel16.0/20527/64bits; fontes herdados sem alteração de literais/corpos. Portão2 não executado; sistema NÃO HOMOLOGADO.

Quatro defeitos F7 reproduzidos e corrigidos somente nos protótipos: data/selo persistente com política explícita, anual por episódio, consulta por ID após sort/lacuna/reuso e auxiliares com valores discriminantes. Guarda adicional comprovada: recusar selo ERRO coincidente; três FAIL preservados e três PASS após correção.70 asserções únicas aprovadas na seleção de resultados experimentais e4 volumes medidos (100/1000/5000/10000), com limites por revisão na matriz.26 integrados repetidos após guarda final;15 guardas F7 reexecutados em sessão fria. Não houve teste novo de recuperação integral ou aprovação de gravadores reais.

[Relatório técnico](./RELATORIO_TECNICO_FASE8.md), [matriz](./MATRIZ_REGRESSAO_FASE8.md), [correções](./CORRECOES_PROTOTIPOS_FASE8.md), [armazenamento](./PROPOSTA_ARMAZENAMENTO_FASE8.md), [persistência](./CONTRATO_PERSISTENCIA_FASE8.md), [M0–M9](./PLANO_IMPLANTACAO_FASE8.md), [decisões Portão2](./PACOTE_APROVACAO_PORTAO2.md). Cinco FAIL oficiais F5 e FAIL integrais F6 continuam pendentes. Nenhuma promoção, migração real, alteração fiscal, instalação ou mudança de segurança.


## Fase 9 — 20261102_142753Z — persistência experimental

Portão experimental somente; Portão 2 NÃO aprovado e sistema NÃO HOMOLOGADO. Entrega F8 R2 e 2.268 hashes conferidos; 43 backups novos verificados. ACE 16/ADODB-ADOX x64 do Office instalado executou transações reais em bancos sintéticos no TEMP. Mestre F9 compilado no VBE: 63 componentes/17 abas, com os 60 herdados preservados e os sete Delete intactos. Nenhuma integração no principal, migração real, regra fiscal alterada, instalação, segurança modificada, commit/push ou promoção.

65 asserções VBA: 64 PASS e 1 FAIL de Ignore após Sort, preservado com prova corretiva separada. Quatro benchmarks VBA PASS e três testes COM entre processos PASS, sem somar sondagens/hashes. Ciclo de sete movimentos/recibos, rollback pré-commit, bloqueio de pendência, escritor/fencing e reconstrução em sessão fria exercitados. Erro real de SaveAs e cancelamento BeforeSave preservaram a autoridade. Falha física de commit, energia/crash/disco cheio, rollback incompleto e snapshot sob escrita concorrente NÃO EXECUTADOS. Os sete FAIL oficiais F5/F6 permanecem nas classificações anteriores.

[Relatório](./RELATORIO_TECNICO_FASE9.md), [matriz](./MATRIZ_REGRESSAO_FASE9.md), [armazenamento](./COMPARATIVO_ARMAZENAMENTO_FASE9.md), [esquema](./ESQUEMA_TRANSACIONAL_FASE9.md), [durabilidade](./TESTES_DURABILIDADE_FASE9.md), [concorrência](./TESTES_CONCORRENCIA_FASE9.md), [reconstrução](./PROTOCOLO_RECONSTRUCAO_EXCEL_FASE9.md), [M0–M4](./PLANO_MIGRACAO_M0_M4_FASE9.md), [decisões](./DECISOES_PENDENTES_FASE9.md).


## 02/11/2026 — Fase10 (20261102_155920Z)

- Verificados2.417 arquivos anteriores, incluindo39 originais;50 backups novos conferidos por SHA-256. F9 final preservada com hash fbb127500d504e6ac9bf3f857dfc936f1f0e43449b3549a2473fff2a5b56386d.
- Somente laboratório: schema4 uniforme, FKs compostas, ponte CurrentEpisodeID, regras DAO, payload/recibo/tentativa SHA-256 e versão; WriterLease revalidado dentro das transações; auditoria de intenções e custódia.
- Corrigidos em protótipos: literal CryptoAPI SHA256, validação de data/dígitos, nulidade da ponte, Ignore por célula, fórmula Annual tratada como fórmula, baseline obrigatória e auditoria PREPARED incompleta. Reprodução de falhas preservada.
- Executadas121 asserções VBA únicas PASS,12 amostras benchmark PASS e10 experiências/verificações COM PASS. Regressão transacional19 casos repetida após a guarda final; não contada como novos IDs. Seis verificações frias VBA BLOQUEADAS por política do Office.
- Fixture final66 componentes/17 planilhas;63 fontes herdadas equivalentes com literais preservados. Designers internos reserializados pelo Office sem alegação de igualdade binária; FRM/FRX externos originais intactos.
- Entrega SHA256:7d1f119124e71a3314601e8c41475db22de9429010aa5424b1652d655b38b7e2. [Relatório](./RELATORIO_TECNICO_FASE10.md), [matriz](./MATRIZ_REGRESSAO_FASE10.md), [evidências/checkpoints](./_auditoria/fase10_20261102_155920Z/).
- Sem alteração dos sete Delete,57 leitores, gravadores reais, regras fiscais, políticas de segurança ou dados operacionais. Sem instalação, commit, push ou publicação. Sistema NÃO HOMOLOGADO e Portão2 não aprovado.


## 02/11/2026 — Entrega 1 integrada sintética (20261102T182823Z)

Autorização expressa P2-01 a P2-08 restrita à candidata fictícia. [Entrega](./ENTREGA_1_CANDIDATO_INTEGRADO.md). Base F10 preservada e manifesto anterior conferido (2.795 entradas, sem divergência). Os 39 originais permanecem com SHA-256 original.

Produzidos XLSM/ACCDB preferenciais; 69 componentes, 17 planilhas, cinco formulários; ACE esquema 5 como autoridade; identidade, versões, recibos, lease/fencing e oito projeções vinculadas. Sete Delete retirados somente da candidata; movimentações usam baixa lógica transacional, manutenção ainda incompatível bloqueada. Assinaturas dos eventos preservadas.

Entrada/Venda/Transferência/Devolução e nova transferência executadas por botões reais. 44 asserções registradas (38 PASS/6 FAIL históricos da própria execução, preservados); R3 possui 33 IDs únicos e uma repetição corretiva. Reentrada inicialmente teve confirmação pendente por chave opcional; corrigida verificação e reconstrução posterior. Falhas transacionais mantêm intenção pendente de conciliação, sem promessa de recuperação integral. Save cancelado por evento real e reconstrução posterior comprovados. [Matriz](./MATRIZ_REGRESSAO_ENTREGA.md).

Corrigidos conversão regional de moeda, formato literal de data, Ignore em projeção protegida, chamada de reconstrução interpretada como rótulo, ordem de preenchimento do formulário e validação da chave opcional da matriz. Backup pareado com lease/cópia fechada/hash e auditoria reaberto.

Compilação final real concluída. Reabertura interativa BLOQUEADA no aviso de macros, requer ação manual; script de preparação de cópias BLOQUEADO pela ExecutionPolicy, sem bypass. [Pendências](./REGISTRO_DEFEITOS_ABERTOS.md) e [manual](./MANUAL_INSTALACAO_RESTAURACAO.md). Sem commit/push/instalação/promoção. NÃO HOMOLOGADA.

## 02/11/2026 — continuação da Entrega 1: reabertura real

Checkpoint verificado em ./_desenvolvimento/entrega1_20261102T182823Z/continuacao_reabertura_20261102T205915Z/. O usuário habilitou macros manualmente. E1_SessaoReaberta revelou FAIL por diferença entre aliases MARCOS~1/MARCOSALEXANDREPERDO do mesmo arquivo. Corrigido somente modCandidatoIntegrado.ConnC1 mediante C1_MesmoArquivoAutoridade (existência + ShortPath); BaseID e vínculo persistido preservados. O Editor VBA recompilou sem erro. Três verificações imediatas retornaram True/False/False (mesmo arquivo/cópia distinta/ausente). Nova execução na sessão aberta PASS e nova instância Excel /x PID 50212 confirmou CACHE_AUDITADO automático e E1_SessaoReaberta PASS, sem repetir movimentos.

Log acumulado 47 execuções de asserção: 40 PASS/7 FAIL; falhas anteriores preservadas. As três consultas imediatas e a compilação são evidências separadas, não somadas como cenários de UserForm. Banco SHA-256 inalterado (3198e07699f644f64c325a9fb517470b64950aa5a262cc354629ed61f1503dc3), revisão 8, oito movimentos/recibos, três ativos, zero Pending. XLSM corrigido SHA-256 5b2a0d55fcea077dd034ace4c4892bbc9fdd46c69760d7d318d2162f54771749. Contagem binária 69 componentes; corpos dos cinco forms iguais ao checkpoint (VB_Base reserializado). 39 originais e backup F10 novamente conferidos, sem divergência. E1-04 encerrado no escopo de reabertura normal; E1-03, conciliação, fluxo D e homologação permanecem pendentes. Nenhuma mudança de segurança, produção, regra fiscal, commit ou publicação.


## 02/11/2026 — Entrega2, execução20261102T215146Z

- Conferida a revisão final E1 e continuação de reabertura:241 entradas do manifesto preservadas;39 originais e dois artefatos finais E1 sem divergência. Criados checkpoints E1/R0/R1/R3/R4 e backups de documentação/fontes antes de alterações.
- Criada candidata independente no TEMP, banco ACE sintético com BaseID próprio, esquema6/contrato2. Implementados ADJUST descritivo auditável, classificação de exclusão, reentrada via controles existentes e Cancelado matriz vinculado. Episódio/custódia permanecem hipóteses experimentais.
- Corrigidos ordem de Change/restauração da origem por token, contexto por instância/página, chave por contexto, verificação de data de HISTORY_OUT, chamada Call de reconstrução e precisão123,4567 nos leitores/envio.
- Executados eventos reais de Entrada, Venda, Transferência, Devolução, alteração, reentrada, Cancelado e backup. ReferênciaR0:10PASS. Continuação:61execuções/54PASS/7FAIL preservados,55IDs. Não somar repetições como cenários novos. Matriz discrimina chamadas auxiliares, erros da suíte e eventos.
- Falhas APPLY3/APPLY6: fatos/baixa/versão revertidos, intenção pendente e repetição bloqueada. Forenses preservados. Falha de projeção e erro após commit: recibos conservados e reconstrução verificada. Save cancelado por evento real e SaveAs inválido exercitados. Recibo ausente contido, sem recriação automática.
- Compilação real final concluída no VBE;69componentes/17planilhas/5forms/436assinaturas existentes preservadas. Designers reserializados pelo Office, sem alegação de igualdade binária. Fontes efetivamente embutidos e vbaProject.bin extraídos; análise estática final51fontes/1049procedimentos fora dos documentos, sem duplicidade pública detectada.
- Estado final ACE:rev16,6veículos,7episódios,16movimentos/recibos,40SheetProjections,0Pending,leaseFREE. AUDIT=OK/PROJECTION=VERIFIED no último VBA. Reabertura interativa E2 ainda BLOQUEADA no aviso de macros; não confundir com gerações reabertas pela reconstrução.
- Produzidos [candidata e banco](./_desenvolvimento/entrega2_20261102T215146Z/entrega/), [relatório](./ENTREGA_2_OPERACOES_CORRIGIDAS.md), [matriz](./MATRIZ_REGRESSAO_ENTREGA2.md), fontes finais, manifestos e instruções atualizadas. Entrega2 permanece incremental e NÃO HOMOLOGADA: conciliação operacional, portabilidade/restauração integral, casos de estorno e cobertura remanescente impedem conclusão integral.
- Nenhuma mudança de produção, regras fiscais, segurança, instalação, commit/push ou publicação. Os sete FAIL históricos não foram reclassificados.


## 02/11/2026,22:06 — continuação da Entrega2: reabertura

Macros habilitadas manualmente pelo usuário. Verificada a inicialização automática CACHE_AUDITADO no caminho E2 antes de carregar cache manualmente. E1_SessaoReaberta executado no projeto E2:PASS;AUDIT=OK;PROJECTION=VERIFIED;REV=16;MOV=16;RECEIPTS=16;ACTIVE=5;PENDING=0. Log acumulado62execuções/55PASS/7FAIL preservados,56IDs;R0separado. Estados Excel conferidos e17abas; XLSM/ACCDB com hashes inalterados. Sem alteração de fontes, Save ou repetição de movimentações. E2-09superado nesse escopo; demais pendências mantidas. [Checkpoint/evidências](./_desenvolvimento/entrega2_20261102T215146Z/continuacao_reabertura_20261103T010409Z/). Relatórios/manifestações anteriores copiados e verificados antes da atualização.
