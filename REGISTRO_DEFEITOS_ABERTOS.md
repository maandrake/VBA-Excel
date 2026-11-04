# Defeitos e pendências — candidata Entrega 1

Escopo: cópia sintética integrada. Nenhuma conclusão abaixo altera os sete FAIL históricos no produto original. A [matriz](./MATRIZ_REGRESSAO_ENTREGA.md) conserva execuções e limites.

| ID | Prioridade | Pendência / comportamento atual | Medida necessária |
|---|---|---|---|
| E1-01 | Alta | Recuperação operacional de PREPARED/Pending não está integrada. O ACE reverte efeitos, mas preserva intenção e rascunho; novas tentativas do episódio são recusadas. | Implementar conciliação auditável, decisão de abortar/retomar e recibo terminal. Não apagar Pending para liberar operação. Políticas ainda pendentes. |
| E1-02 | Alta | Fluxo D parcial: alteração, exclusão manual, retificação e filtros legados ainda incompatíveis são bloqueados. Nem todos os 57 leitores e relatórios foram migrados. | Migrar leitores/escritores por identidade, revisão e elegibilidade; habilitar somente após regressão. Não usar ActiveCell/primeiro Find para editar históricos. |
| E1-03 | Alta | Pacote preservado vincula AuthorityPath ao laboratório executado. Preparar-CopiaSintetica.ps1 foi criado, mas ExecutionPolicy bloqueou sua execução. Portabilidade/rebind e restauração integral em outra pasta NÃO testados. | Revisar/autorizar mecanismo de implantação sintética explícito e ensaiar cópia, rebind e reconstrução. Não usar bypass nem eleger uma cópia como autoridade automaticamente. |
| E1-05 | Alta | Concorrência dos novos gravadores C1 não foi exercitada em dois processos; fencing/transações são herdados do F10. | Reexecutar interleavings nos eventos integrados, fonte alterada após seleção, perda de lease e preparação órfã. Não extrapolar testes antigos. |
| E1-06 | Média | Contexto de formulário é indexado por TypeName. O fluxo testado usa as instâncias padrão; múltiplas instâncias do mesmo tipo não foram ensaiadas. | Identificar sessão/instância de formulário antes de permitir janelas concorrentes. |
| E1-07 | Média | Reentrada usa helper e VehicleID explícito fornecido pela suíte. Interface final para escolher episódio encerrado/proveniência ainda não está pronta. | Implementar escolha inequívoca e conciliação de dados ausentes; não deduzir VehicleID por placa. |
| E1-08 | Alta | Cancelado e outras operações além de Venda/Transferência matriz e Devolução filial não estão integradas. Cancelado é recusado na candidata; original permanece inalterado. | Decisão operacional/fiscal e integração preservando comportamento aprovado. Não declarar homologação fiscal por XML sintético. |
| E1-09 | Média | Reconstrução mantém 17 abas, mas somente oito são projeções dos fluxos integrados. Fórmulas anuais N/O de apresentação, relatórios gerais e demais abas não têm equivalência integral comprovada. | Inventariar saída esperada e migrar fórmulas/condicionais restantes; usar dados discriminantes por episódio. |
| E1-10 | Média | Backup C1 foi reaberto/auditado; restauração completa, perda de mídia e falha de disco ainda não ensaiadas. Botões de backup/manutenção legados fora dos caminhos C1 exigem adaptação. | Integrar todos os pontos de entrada ao backup pareado e verificar restauração sem descartar fatos posteriores. |
| E1-11 | Alta | Instalação do esquema 5 usa DDL em etapas; interrupção pode deixar instalação parcial. | Empacotar base pronta verificada ou migração retomável com checkpoints; nunca repetir DDL indiscriminadamente sobre fatos existentes. |
| E1-12 | Alta | Não há garantia contra energia, corrupção física, disco cheio, SQL externo ou adulteração intencional dos hashes. | Ensaios próprios e modelo de ameaça/permissões. Não prometer ACID entre ACE e XLSM. |
| E1-13 | Média | Precisão/conversão corrigida para texto canônico e entrada brasileira com vírgula; todos os extremos e entradas ambíguas como 1.000 não foram ensaiados. | Contrato explícito de entrada regional; testar limites monetários de quatro casas e arredondamento. Não aceitar inferências silenciosas. |
| E1-14 | Média | Designer embutido foi reserializado pelo Excel; comparação binária difere, embora assinaturas/eventos se mantenham e controles usados tenham funcionado. | Regressão visual/funcional dos cinco forms e inventário completo dos controles. Não declarar FRX embutido byte a byte igual. |
| E1-15 | Alta | Logs, retenção, limite de recuperação e semântica de episódio/custódia ainda são hipóteses/políticas não aprovadas para uso operacional. | Decisão expressa do responsável antes de qualquer dados reais ou promoção. |

## Defeitos encontrados e corrigidos nesta execução

1. Conversão regional CCur em payload canônico: primeira base incorreta preservada forense; nova base R3 confere 900 e 650,25.
2. Ignore em planilha protegida: acesso controlado com restauração da proteção; verificação persistida concluída.
3. Data normalizada regionalmente: formato literal e contrato F8; não se removeu NumberFormat do verificador.
4. Chamada de reconstrução interpretada como rótulo: Call explícito. R2 FAIL preservado; R3 venda PASS.
5. Campo apagado por evento Change ao reutilizar formulário: operação configurada antes dos campos; transferência sucessiva FAIL preservado e nova execução PASS.
6. VerifyBatch10 exigia chave em entrada matriz cujo contrato admite vazio: alinhado ao esquema 5; diagnóstico distingue PENDING de memória verificada; reconstrução posterior confirmada.
7. Novo backup liberava lease cedo: lease persistido mantido durante cópia fechada, verificação e manifesto; backup reaberto PASS.
8. Um dump de teste existente causou erro 58 antes de acionar o evento; prefixos por execução corrigidos, sem sobrescrever evidência.
9. Uma substituição de código do novo backup duplicou a assinatura; compilador real detectou, correção foi aplicada e compilação final passou. Não foi falha de estoque.

10. E1-04 — reabertura: GetAbsolutePathName comparava aliases curto/longo do mesmo banco como diferentes. A candidata usa ShortPath de arquivos existentes e mantém BaseID/AuthorityPath. Compilação real e três verificações de caminho passaram; novo processo inicializou CACHE_AUDITADO e E1_SessaoReaberta PASS na execução 47. O FAIL da execução 45 está preservado. Escopo encerrado: reabertura normal no mesmo laboratório; E1-03/portabilidade permanece aberto.

## Referências históricas preservadas

V_VALIDO, T_VALIDO, D_VALIDO, D_FALHA_BAIXA, T_DUP_DESTINO, C_FALHA_BAIXA e C_FALHA_COMPENSACAO continuam registrados como FAIL da implementação anterior. Há resultados próprios para a candidata, com escopo na matriz. A candidata não usa compensação de células depois de Delete; usa transação ACE e projeção reconstruível. A ausência desse caminho não converte testes legados em PASS.


# Atualização — Entrega 2 (02/11/2026)

Esta seção é o estado mais recente; o registro da Entrega 1 acima foi preservado. [Entrega executável](./ENTREGA_2_OPERACOES_CORRIGIDAS.md) e [matriz](./MATRIZ_REGRESSAO_ENTREGA2.md). Não há homologação nem conclusão integral da Entrega 2.

| ID | Gravidade | Estado atual e evidência | Ação necessária |
|---|---|---|---|
| E1-01 / E2-01 | Alta | Diagnóstico somente leitura implementado. APPLY3/APPLY6 reverteram fatos/baixa/versões, mas mantiveram Pending. Repetição recusada. | Definir responsável e protocolo auditável de retomada/aborto. Não apagar Pending ou gerar novo OperationID para liberar. |
| E1-02 / E2-02 | Alta | Alteração descritiva ativa ADJUST e exclusão classificada integradas. Histórico anterior imutável; botão real aprovado na matriz. Documento, dinheiro, datas, chave, placa e anulação confirmada permanecem bloqueados. | Aprovar contratos de estorno/retificação por caso; implementar fatos substitutivos e respectivos testes. Retificação filial ainda não executada. |
| E1-03 / E2-03 | Alta | Aliases do mesmo banco, BaseID divergente, caminho divergente e banco ausente ensaiados. Cópia em outro local não assume autoridade. | Instalador/rebind explícito e restauração integral ainda NÃO EXECUTADOS. Não executar pacote no OneDrive ou editar AuthorityPath para escolher automaticamente uma cópia. |
| E1-05 | Alta | Lease/fencing herdados; nenhum novo ensaio de dois processos nos eventos C2. | Concorrência, perda de lease e preparação órfã com dois gravadores reais. |
| E1-06 | Média | Contexto corrigido para TypeName+ObjPtr+página; chave deixou de ser global. | Teste de múltiplas instâncias simultâneas ainda não executado; correção de código não encerra esse cenário. |
| E1-07 / E2-04 | Média | Reentrada pelos controles reais passou: episódio encerrado escolhido na ListBox, mesmo VID/novo EP, sem VID injetado. Placa conhecida não vincula sozinha. | Testar múltiplos candidatos, evidência de Novo veículo com placa coincidente e identidade insuficiente nos dois locais. |
| E1-08 / E2-05 | Alta | Cancelado matriz vinculado integrado: dois históricos, zero e origem preservada, sem XML, botão real passou. | Fiscalidade permanece pendente; Cancelado sem origem/filial e demais operações não definidas ficam bloqueados. |
| E1-09 / E2-06 | Alta | Revisão estática dos57 pontos e migração dos afetados; anual e M/N/O por identidade conferidos nos fluxos ensaiados. Somente oito abas são projeções ACE. | Regressão dos demais relatórios, totais, condicionais, atalhos/OnAction e controles restantes. Não declarar migração completa das17abas. |
| E1-10 | Alta | Botão SalvarTudo produz backup pareado, banco reaberto/auditado e lease livre. | Recuperação completa em ambiente independente e cenários com fatos posteriores ainda pendentes. Não restaurar revisão14 sobre16 como se não houvesse movimentos novos. |
| E1-11 / E2-07 | Alta | Migração5→6 concluída com backup e auditoria. DDL continua em etapas. | Migração interrompida/retomada deve ser ensaiada; entregar base pronta não comprova atomicidade de DDL. |
| E1-12 | Alta | Energia, crash, disco cheio, SQL externo/adulteração intencional não demonstrados. | Plano específico de segurança e testes; não atribuir a hashes garantias de autenticação. |
| E1-13 / E2-08 | Média | Perda123,4567→123,46 reproduzida nos leitores/envio e corrigida para quatro casas, novas execuções PASS. | Extremos monetários e entrada ambígua1.000 continuam pendentes de contrato/regressão. |
| E1-14 | Média |436 assinaturas/CodeNames preservados. Designers reserializados: não são binariamente iguais. Controles usados nos eventos funcionaram. | Inventário/regressão visual de todas as propriedades e cinco formulários; reimportação do pacote de fontes ainda não ensaiada. |
| E1-15 | Alta | Episódio contínuo na transferência/devolução, logs/retenção e recuperação operacional seguem hipóteses/políticas de laboratório. | Decisão expressa antes de uso real. |
| E2-09 | Alta | Compilação final real PASS; reabertura interativa E2 BLOQUEADA no aviso de macros. | Clique manual do usuário e diagnóstico em novo contexto VBA; não substituir pelo PASS da Entrega1 ou de geração aberta programaticamente. |
| E2-10 | Média | Matriz dirigida executada, mas não todos os XMLs inválidos, saídas inválidas, seleção de arquivo normal e botões de filtro. | Completar cobertura pelos eventos reais e inspeção ACE em fixtures independentes. |
| E2-11 | Alta | Criação de veículo/episódio NEW antecede preparação em etapas. Diagnóstico detecta rascunho órfão, não o resolve. | Ensaiar interrupção nessa fronteira e especificar descarte/retomada auditável; nunca adivinhar vínculo. |

## Correções da Entrega 2 e resultados falhos preservados

- MONEY_FORMAT em HISTORY_OUT: ramo de data restaurado, erro de conexão fechada tratado; auditoria/reconstrução posteriores passaram.
- Retificação confirmada sem reconstrução: chamada rotulada involuntariamente corrigida com Call; FAIL revisão6 mantido, revisão7 PASS.
- Precisão dos leitores/envio: dois FAIL e duas novas execuções PASS com123,4567; fatos ACE não foram arredondados.
- Ordem de Change na saída: identidade/custo restaurados por token; transferência sucessiva/devolução confirmadas.
- Dois FAIL de guardas no primeiro ensaio decorreram de alterações não confirmadas/visibilidade de conexão. Nova metodologia confirma e fecha a conexão antes da leitura independente; duas recusas PASS. Nenhuma guarda foi removida para passar o teste.
- A suíte errou o nome de coluna ao conferir devolução já confirmada e serializou incorretamente snapshot antes de outro evento. Corrigida somente a instrumentação; consulta posterior e execução dirigida passaram. Não repetir a movimentação já confirmada.

Os sete casos históricos continuam FAIL no produto anterior. Na candidata, os novos testes distinguem evento aprovado, rollback de efeitos, reconstrução e contenção com conciliação pendente. A classe de compensação antiga não foi apresentada como corrigida.

E2-12 — Média: os sete casos históricos foram relacionados aos eventos/mecanismos novos, porém os fluxos normais compartilharam uma sequência de laboratório com snapshots. A repetição completa em sete fixtures físicas independentes permanece NÃO EXECUTADA. Não converter contenção em recuperação integral nem considerar encerrado o defeito da classe legada.


## Estado atualizado de E2-09 — reabertura

**Resolvido no escopo de reabertura interativa no laboratório existente.** Macros habilitadas manualmente; CACHE_AUDITADO automático e E1_SESSAO_REABERTA_E2_R3 PASS com revisão16/projeçãoVERIFIED. XLSM/ACCDB mantêm hashes. [Evidência](./_desenvolvimento/entrega2_20261102T215146Z/continuacao_reabertura_20261103T010409Z/evidencias_reabertura.json). A linha anterior de bloqueio é histórica; portabilidadeE2-03, conciliaçãoE2-01 e outras pendências não foram encerradas por esse PASS.


## Atualização do fechamento — 03/11/2026

| Identificação | Situação / risco | Responsável e próximo critério |
|---|---|---|
| E2-FECH-01 — TEMP por sessão | Corrigido na candidata; diagnóstico aberto passou. Reabertura fria final bloqueada por macros. | Engenharia: confirmar inicialização e F2_Restaurada em nova sessão. |
| E2-FECH-02 — restauração | Realocação ao vivo ensaiada e origem revogada. Recuperação quando origem/registro estão indisponíveis ainda não implementada/ensaiada. | Engenharia + responsável: protocolo sem duas autoridades e sem perda de fatos posteriores. |
| E2-FECH-03 — conciliação | Plano/recibo de abandono sem efeitos implementado; API passou. Interface humana, versão alterada e divergências ainda precisam de regressão. ENTRY/órfãos/incertos ficam bloqueados. | Engenharia: testar os limites; responsável: decidir somente as exceções operacionais necessárias. |
| E2-FECH-04 — regressão final | Novos eventos/filtros preparados mas não executados; Office recusou macros na sessão atual. | Proprietário: habilitação manual da cópia confiável; engenharia: executar, registrar e corrigir antes do aceite. |
| E2-FECH-05 — Saved=False | Causa exata ainda não isolada; nenhuma atribuição Saved=True usada para mascarar. | Engenharia: comparar cálculo, eventos e hashes antes/depois em sessão habilitada. |
| Exceções de manutenção | Anulação/estorno e retificação fiscal, monetária ou histórica NÃO DISPONÍVEIS sem contrato operacional. | Proprietário: regra específica; manter bloqueios. |

[Estado completo e artefatos](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_20261103T2147Z/FECHAMENTO_ENTREGA2_EM_ANDAMENTO.md). Não há autorização para produção, migração real ou homologação por protótipos.


## Estado corrente — continuação revisão 26

As linhas anteriores são histórico. [Evidências desta sessão](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_macros_20261103T2213Z/RELATORIO_CONTINUACAO_ENTREGA2.md).

| ID | Gravidade / estado | Responsável e próximo critério |
|---|---|---|
| E2-FECH-01 / 04 | Teste frio inicial, fluxos reais, filtros, XML e metadados executados. Reabertura final R26 BLOQUEADA no aviso de macros | Proprietário: habilitar conteúdo somente na candidata TEMP; engenharia: comprovar cache automático e fluxo após reabrir |
| E2-FECH-02 | Realocação ao vivo + sessão fria aprovadas; backup reaberto. Restauração quando origem/registro não existem continua NÃO EXECUTADA | Engenharia/responsável: protocolo explícito, sem duplicar autoridade ou perder fatos posteriores |
| E2-FECH-03 | Conciliação sem efeitos, interface, plano obsoleto, replay abortado e nova tentativa passaram | Engenharia: ampliar casos incertos; responsável: regras de ENTRY/órfãos/estorno. Não liberar automaticamente |
| E2-FECH-05 | Causa de Saved=False isolada: diagnóstico protege/desprotege para ler Ignore; hash em disco inalterado | Efeito conhecido em memória; orientar usuário. Não mascarar Saved nem prometer diagnóstico totalmente sem efeitos de interface |
| E2-FECH-06 | Média: erro de conflito ACE é apresentado como transação não iniciada. Teste recusou escrita e não criou fatos | Engenharia: preservar cadeia de erros e melhorar mensagem; reexecutar conflito após correção |
| E1-05 | Alta: evento real versus processo ADO independente aprovado; dois gravadores Excel não ensaiados nesta revisão | Engenharia: segundo cliente integrado, perda de lease, geração e retomada órfã |
| E2-11 | Alta: ENTRY/rascunho órfão continua sem resolução auditável | Engenharia + responsável: protocolo específico; não escolher VehicleID por placa |
| E1-07 / E2-04 | Reentrada por ListBox e placa conhecida sem vínculo aprovadas novamente | Engenharia: múltiplos candidatos, novo veículo com placa coincidente e identidade insuficiente nos dois locais |
| E1-09 / E2-06 | 66 projeções conferidas por identidade, fórmulas, formatos e Ignore; filtro Compra matriz/filial e Venda real aprovados | Engenharia: demais relatórios/atalhos e abrangência das 17 abas ainda não comprovados |
| E2-TESTE-IGNORE | FAIL de instrumentação preservado; CI corrigido e novo teste 66 projeções PASS | Não é defeito novo do gravador; manter evidência anterior |
| Exceções de manutenção | Anulação/estorno, retificação fiscal/monetária/histórica e exclusão de fato confirmado NÃO DISPONÍVEIS | Proprietário: decisão específica por operação. Nunca classificar bloqueio como implementação |

Energia/crash/corrupção física, escrita SQL externa, logs/retenção e validação operacional da custódia continuam pendentes. Entrega 2 não encerrada; não iniciar Entrega 3 como aprovada.


## Retomada de04/11/2026 — QueryClose, revisao26

[Relatorio da retomada](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_reabertura_final_20261103T233016Z/RELATORIO_RETOMADA_QUERYCLOSE.md). Os registros anteriores permanecem historicos. A referencia R26 reabriu com CACHE_AUDITADO e diagnostico VERIFIED; pesquisas reais matriz/filial confirmaram episodios e valores150,3333/456,7891. Foi reproduzido erro1004 no X de Cad_NFE: incremento arbitrario de U1 na planilha ativa protegida. Removida somente essa escrita em QueryClose de Cad_NFE e Cad_NFE_Saida na candidata. Compilacao real e regressao de X/botao Fechar dos dois formularios PASS.17tabelas/ACCDB inalterados;39originais com hashes iguais.

Novo XLSM preservado: SHA256 efd1916395159da3d30ecc99838ff6e0e44dd9e0ca2b13b29570a3b59fe247e0; ACCDB9fc0bfe8001120d2804b8f56836949249606e2030a6af431a71b1757c80267dc. Pacote privado Entrega2_R26_QueryClose_20261104T1221Z_PENDENTE_ACEITE.72componentes;436assinaturas dos5formularios preservadas; designers reserializados, sem alegar identidade binaria.

FORMULA_VALUE apos edicao/formularios preservado como FAIL de tentativa; recalc completo recuperou VERIFIED, sem alterar fatos. Causa fina ainda nao isolada. Arquivo salvo e instancia exclusiva encerrada normalmente. A nova abertura do XLSM corrigido concluiu, mas Office voltou a desabilitar macros: inicializacao/fluxo final BLOQUEADOS ate habilitacao manual. Entrega2 continua em fechamento, NAO HOMOLOGADA; nenhuma liberacao de Entrega3.

E2-FECH-07 (media): QueryClose1004 corrigido/testado no escopo dos dois formularios; reabertura funcional da correcao pendente por macros. E2-FECH-08 (media): formula em cache desatualizada apos edicao; recalc recuperou, causa fina por isolar. E2-FECH-06 (mensagem de conflito ACE), E1-05 (dois Excel), E2-11 (ENTRY orfao) e restauracao de desastre permanecem abertos, sob engenharia/responsavel.


## Retomada controlada — diagnóstico ACE em 04/11/2026, 18:21 UTC

[Evidências e resultado da etapa](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_diagnostico_ACE_20261104T182134Z/RESULTADO_ETAPA_CONTROLADA.md). Ponytail 5.3.3 Full; Git inicial main/b7b9d99 limpo. AGENTS vigente preservado. Apenas o Excel operacional estava aberto; nenhum comando foi enviado a ele. Candidata, ACCDB e registro ativo mantêm seus hashes (42b34e…/9fc0bf…/b9e0c8…). Revisão26, 26movimentos/26recibos e Pending0 confirmados por leitura ADO. Nenhum VBA ou arquivo operacional alterado.

**E2-FECH-06 permanece ABERTO (engenharia):** Seis verificações únicas COM/ADO em duas conexões de um processo, somente registro descartável TEMP: seis PASS. A primeira tentativa executou quatro; a segunda repetiu as quatro e acrescentou duas. Não são dez testes únicos nem execução funcional VBA. Conflito capturado com SQLState3218; rollback imediato após a recusa funcionou. SQLState3034 apareceu apenas no rollback após commit deliberado da sondagem. A causa do diagnóstico E2-FECH-06 no evento real continua não isolada, sem correção especulativa.

Próximo critério: capturar erro/etapa e cadeia ADO no evento real antes da limpeza, aplicar apenas correção comprovada, compilar e regredir. A ausência de janela da candidata impede essa etapa de interface; não foi observada uma nova recusa de macros nesta sessão. Não encerrar E2-FECH-07/reabertura por estes testes.


## Reabertura real da candidata corrigida — 04/11/2026, 22:04–22:15 UTC

[Resultado e evidências](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_vba_real_20261104T185156Z/RESULTADO_REABERTURA_VBA_REAL.md). Instância exclusiva PID7644; XLSM42b34e… reaberto com CACHE_AUDITADO automático antes de carga manual. C1_Diagnostico real antes/depois: AUDIT=OK;PROJECTION=VERIFIED;REV=26;MOV=26;RECEIPTS=26;ACTIVE=7;PENDING=0. X e CommandButtonFechar dos dois UserForms executados por interface, sem1004; ambos Visible=False. U1 vazia/protegida e estados Application preservados. Seis verificações dirigidas PASS, com modalidades separadas; não são seis fluxos de estoque.

**E2-FECH-07 resolvido no escopo dos dois formulários, incluindo reabertura funcional da revisão corrigida.** O bloqueio anterior de macros/reabertura permanece como tentativa histórica. E2-FECH-08 não foi reproduzido: diagnóstico VERIFIED sem CalculateFull; causa fina anterior continua não isolada. E2-FECH-06, E2-11/órfãos, dois Excel concorrentes e recuperação sem origem/registro permanecem abertos. Não encerrar a Entrega2 por este resultado. VBA, XLSM/ACCDB/registro e17tabelas persistidas permaneceram inalterados. Saved=False é alteração em memória; sem Save, hash em disco igual. Compilar VBAProject já desabilitado no projeto carregado; nova invocação explícita de compilação NÃO EXECUTADA. Nenhuma configuração de segurança alterada. Candidata permanece aberta no laboratório. E2-FECH-06 e demais pendências continuam abertos; Entrega2 ainda não encerrada/NÃO HOMOLOGADA. Sem commit/push/promoção.
