# Manual de teste, preservação e restauração — Entrega 1

**Somente laboratório sintético. NÃO utilizar com estoques, clientes ou XML reais. Não promover para produção.**

## Requisitos verificados

Windows e Excel 16.0.20527.20112 x64, ADO/ADOX/DAO e Microsoft.ACE.OLEDB.16.0 já instalados. ACE observado 16.0.20527.20016 no laboratório. Não instalar redistribuíveis nem alterar segurança. Compatibilidade de distribuição em outro computador continua pendente.

## Arquivos e integridade

[Pacote preservado](./_desenvolvimento/entrega1_20261102T182823Z/entrega/), [manifesto SHA-256](./_desenvolvimento/entrega1_20261102T182823Z/entrega/manifesto_artefatos.json), [checkpoint e cópias forenses](./_desenvolvimento/entrega1_20261102T182823Z/checkpoints/), [backup inicial](./_desenvolvimento/entrega1_20261102T182823Z/backup/).

XLSM: 5b2a0d55fcea077dd034ace4c4892bbc9fdd46c69760d7d318d2162f54771749

ACCDB: 3198e07699f644f64c325a9fb517470b64950aa5a262cc354629ed61f1503dc3

Os hashes referem-se à cópia entregue antes de novas execuções. Save, lease e recibos alteram hashes legitimamente; faça novo checkpoint ao continuar. A integridade não comprova autorização fiscal.

## Continuar no laboratório já testado

1. Abra somente [Estoque_Fera_Veiculos_Candidato_Integrado.xlsm](C:/Users/MARCOS~1/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/Entrega1_20261102T182823Z/Estoque_Fera_Veiculos_Candidato_Integrado.xlsm) no TEMP. O [banco sintético](C:/Users/MARCOS~1/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/Entrega1_20261102T182823Z/Banco_Candidato_Sintetico.accdb) deve permanecer ao lado, com esses nomes. Não abrir a cópia do OneDrive para executar testes.
2. Se aparecer Habilitar Conteúdo, a escolha precisa ser manual. Nesta continuação o usuário habilitou as macros e o teste em novo processo passou. Não alterar Trust Center, AccessVBOM, ExecutionPolicy ou locais confiáveis. A automação não efetuou esse clique.
3. Alt+F11 → Exibir → Verificação imediata. Execute uma linha por vez:

   `? C1_UltimoResultado`

   Deve indicar CACHE_AUDITADO, observado antes de qualquer carregador manual. Se indicar bloqueio ou vazio, preserve a mensagem e investigue; não carregue o cache apenas para mascarar uma falha de inicialização.

   `? C1_Diagnostico()`

   O esperado no checkpoint final é AUDIT=OK;PROJECTION=VERIFIED;REV=8;MOV=8;RECEIPTS=8;ACTIVE=3;PENDING=0. Qualquer divergência exige inspeção; não repetir movimentações.
4. Execute `Call E1_SessaoReaberta` somente depois dessa reabertura. O teste passou na execução 47, em nova instância Excel PID 50212, com cache automático e diagnóstico VERIFIED. O FAIL anterior por alias de caminho e o primeiro PASS da sessão aberta foram preservados. Não aceitar resultado se houver #N/D, identidade divergente ou confirmação pendente.
5. Para testar interface: Alt+F8 → CandidatoAbrirEntrada. A entrada matriz exige número, unidade, código, veículo, placa, ano/cor, valor e data; a filial exige chave textual de 44 dígitos. Use números fictícios novos e valores brasileiros explícitos, por exemplo 123,45. Não digitar chave em célula numérica.
6. Pesquise a placa e selecione a linha da ListBox. O token é resolvido pela autoridade e versão antes de enviar ao formulário de saída. Use Venda/Transferência na matriz ou Devolução na filial. As demais operações não estão integradas. Os XMLs do pacote são estritamente sintéticos e correspondem apenas às massas documentadas; não representam documentos autorizados.
7. Os cenários E1_Entrada/E1_Venda/E1_Transferencia/E1_Devolucao são dependentes de massa inicial específica. **Não executá-los indiscriminadamente na base já movimentada.** Consulte os pré-requisitos e snapshots na [matriz](./MATRIZ_REGRESSAO_ENTREGA.md). As falhas devem usar cópias descartáveis.

## Cópia em outra pasta TEMP

A cópia deslocada é recusada por AuthorityPath; isso é deliberado. Não editar caminho da autoridade no pacote mestre nem apontar duas candidatas para bases divergentes como se fossem produção.

[Preparar-CopiaSintetica.ps1](./_desenvolvimento/entrega1_20261102T182823Z/entrega/Preparar-CopiaSintetica.ps1) foi preparado para: conferir os dois hashes; criar diretório TEMP novo; copiar apenas a massa sintética; recusar lease ocupado/Pending/linhagem diferente; vincular explicitamente a nova cópia e exigir reconstrução. Sua execução foi **BLOQUEADA pela ExecutionPolicy atual**. Não está validado e não foi substituído por Invoke-Expression, Bypass ou alteração de Registro. Não executá-lo sem revisão do mecanismo permitido no ambiente. A portabilidade permanece E1-03.

## Backup reproduzido

A macro CandidatoBackup verifica a autoridade, adquire lease, fecha conexões, copia o ACCDB, compara hashes da fonte/cópia/fonte, reabre e audita a cópia, libera somente seu lease e salva uma cópia do XLSM. Produz manifesto. O teste reabriu o banco copiado e conferiu os recibos e lease FREE. Não é backup validado contra escritores SQL externos. Não exclui versões antigas e não aplica retenção automática.

Backup realmente produzido no TEMP: Backup_C1_DD7DBA137BAD453F95A7705CECA6B472 (revisão 5). O checkpoint final da entrega é revisão 8. Não restaurar revisão 5 sobre revisão 8 para esconder as movimentações posteriores.

## Restauração e conciliação

- **Sem fatos posteriores, laboratório descartável:** fechar somente a própria candidata e suas conexões; preservar primeiro a cópia atual com hash; restaurar um par conhecido em ambiente independente; conferir schema/BaseID/recibos/Pending; fazer rebind explícito autorizado; reconstruir; salvar/reabrir/verificar. A restauração completa em outro diretório não foi executada nesta entrega.
- **Com fatos posteriores:** não substituir o banco por versão antiga. Preservar ambos, identificar movimentos/recibos posteriores e conciliar. Excel é reconstruído da autoridade escolhida mediante decisão rastreada.
- **PREPARED/Pending:** não apagar intenção nem criar OperationID novo para contornar. Os testes demonstraram reversão dos fatos, mas o encerramento/retomada da intenção exige procedimento ainda pendente.
- **Commit confirmado, Excel/Save falhou:** não repetir venda/transferência/devolução. Conferir recibo e autoridade; executar CandidatoReconstruir. Ela revalida lease/auditoria, materializa, verifica, salva geração e a reabre antes do ACK. Não há atomicidade conjunta ACE/XLSM.
- **Estado incerto, lease órfão, hash divergente ou histórico alterado:** suspender escrita dependente, guardar arquivos forenses e conciliar. Não desbloquear por tempo.

A reconstrução integral após Save cancelado e após erro injetado depois de commit foi executada. Crash, energia, disco cheio e falha física de commit NÃO foram testados.

## Manutenção e reversão de código

O [backup F10](./_desenvolvimento/entrega1_20261102T182823Z/backup/Fixture_F10_Referencia.xlsm) preserva a base anterior. A reversão de código não é restauração cega de dados. Os 39 arquivos originais em Excel/Form/MDL ficaram intactos. Os [fontes](./_desenvolvimento/entrega1_20261102T182823Z/fontes/) incluem cinco pares FRM/FRX e módulos de documento separados; não importar documentos como BAS e não renomear CodeNames. A candidata executável já contém os componentes.

Nenhuma publicação, instalação, mudança fiscal ou promoção operacional foi realizada.

## Correção de caminho e checkpoint da continuação

O banco continuou com o mesmo hash e revisão. A candidata corrigida aceita os aliases 8.3 e longo do mesmo arquivo existente, sem alterar AuthorityPath ou BaseID. Uma cópia em outro diretório continua recusada; E1-03 não foi encerrado. [Backup anterior e evidências](./_desenvolvimento/entrega1_20261102T182823Z/continuacao_reabertura_20261102T205915Z/) preservam o XLSM de hash 9b24bec50e588551a744128912c062bdbb020905bccd41b40f62c0aee89dbad2 para reversão de código. Não substituir uma autoridade depois de novas operações sem conciliação.


# Continuação — Entrega 2

Esta seção prevalece para a candidata E2; as instruções E1 anteriores são histórico. **Somente dados fictícios, fora do OneDrive.**

## Par final e hashes

[Pacote E2](./_desenvolvimento/entrega2_20261102T215146Z/entrega/). XLSM:856c1c148451bae188d5539b3a344371579275268d7475dd535e1ade89ee39d0. ACCDB:1390a79c2b36fcea13843bb211dcb56ffbfe3fe8523e2200f70f32e1dbc35581. [Manifesto](./_desenvolvimento/entrega2_20261102T215146Z/entrega/manifesto_artefatos.json).

Ambiente executado: Excel16.0.20527.20112x64/ACE16.0.20527.20016 já instalados. Banco esquema6/contrato2, BaseID143C55EC-132B-4839-8413-268B265117AF, revisão16. Não instalar provedores ou mudar políticas.

## Retomar o laboratório existente

1. Abrir somente [XLSM E2 no TEMP](C:/Users/MARCOS~1/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/Entrega2_20261102T215146Z/Estoque_Fera_Veiculos_Candidato_Operacoes_Corrigidas.xlsm). O [ACCDB correspondente](C:/Users/MARCOS~1/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/Entrega2_20261102T215146Z/Banco_Candidato_Operacoes_Corrigidas.accdb) permanece ao lado. Não abrir a cópia de preservação do OneDrive.
2. O aviso Habilitar Macros deve ser tratado manualmente pelo usuário. A reabertura E2 aguarda esse clique. Nenhum bloqueio de segurança foi contornado.
3. Alt+F11, selecionar o projeto que contém exatamente o nome da candidata E2 e abrir um módulo dela. Ctrl+G exibe a janela Imediata. Se a Entrega1 estiver aberta, não executar nomes públicos sem conferir o projeto: existem rotinas de mesmo nome nas duas pastas.
4. Antes de carregar cache manualmente, consultar ? C1_UltimoResultado. Depois executar ? C1_Diagnostico(). O checkpoint16 espera AUDIT=OK;PROJECTION=VERIFIED;REV=16;MOV=16;RECEIPTS=16;ACTIVE=5;PENDING=0. Executar Call E2_Diagnostico gera relatório somente leitura de Pending/rascunhos, mas altera a mensagem C1_UltimoResultado; por isso consultar a inicialização antes dele. Registrar resultados reais, inclusive falhas.
5. Depurar > Compilar VBAProject foi executado e aprovado no arquivo entregue. Não importar todos os fontes sobre ele; já está integrado. Alterações futuras exigem novo checkpoint, compilação e testes afetados.
6. Alt+F8, selecionar a macro CandidatoAbrirEntrada da candidata. Compra cadastra veículo novo; placa já conhecida exige escolha explícita. Reentrada mostra episódios encerrados pela pesquisa e exige seleção da ListBox antes de Gravar. Não vincular por coincidência de placa/NF. Novo veículo com placa coincidente pede evidência humana e ainda possui cobertura incompleta.
7. Selecionar registro ativo na pesquisa e usar Copiar para saída. Venda/Transferência matriz e Devolução filial trabalham pelo token/versão; custos são reconstruídos da autoridade. Cancelado testado somente para origem matriz vinculada. Usar apenas XML sintético correspondente à operação; não usar documentos reais.
8. Alterar permite somente metadados descritivos de entrada ativa, com motivo e responsável. Cria ADJUST e preserva fato anterior. Campos congelados ou históricos encerrados devem gerar bloqueio. Excluir classifica o caso e não apaga fatos confirmados; anulação/estorno ainda exige decisão.

Não reexecutar E1_IniciarBaseLimpa ou a sequência de cadastros sobre a base final: ela contém os cenários já confirmados. E1_/E2_ são suítes dependentes de massa/revisão. Para regressão, usar checkpoint e fixture independente, conforme [matriz](./MATRIZ_REGRESSAO_ENTREGA2.md).

## Preservação e restauração

[Checkpoint da Entrega1](./_desenvolvimento/entrega2_20261102T215146Z/checkpoint_entrega1/), [referência R0](./_desenvolvimento/entrega2_20261102T215146Z/checkpoint_referencia_R0/), [antes das falhas revisão10](./_desenvolvimento/entrega2_20261102T215146Z/checkpoint_R3_rev10/), [revisão15](./_desenvolvimento/entrega2_20261102T215146Z/checkpoint_R4_rev15/) e [backup real revisão14](./_desenvolvimento/entrega2_20261102T215146Z/evidencias/backup_verificado/) estão preservados. As bases com Pending e recibo ausente estão em evidencias/D_FALHA_BAIXA, ANTES_COMMIT e RECIBO_AUSENTE. São forenses, não bases liberadas para novas operações.

Para backup, usar SalvarTudo no formulário de saída ou CandidatoBackup. Foi conferido: lease mantido, conexões fechadas durante cópia ACCDB, hashes fonte/cópia/fonte, abertura/auditoria da cópia, cópia XLSM e manifesto. Não é garantia contra SQL externo ou falha física. Não há retenção automática.

- Commit confirmado e projeção/Save falhou: não repetir movimento; conferir auditoria/recibo e executar CandidatoReconstruir. A rotina materializa, verifica, salva e reabre uma geração antes do ACK. Caso falhe, preservar a mensagem e suspender escrita dependente.
- Pending, rascunho órfão ou recibo ausente: CandidatoConsultarPendencias é diagnóstico, não resolução. Preservar banco/XLSM/logs/hashes e encaminhar ao responsável. Não apagar intenção, inventar recibo ou trocar OperationID.
- Sem fatos posteriores em ramificação descartável: preservar primeiro o estado atual; fechar somente a própria candidata; copiar par conhecido e verificar hashes/esquema/BaseID/recibos. O laboratório retomou checkpoints em seu caminho conhecido após preservar forenses. Isso não validou um restaurador operacional.
- Com fatos posteriores: nunca restaurar cegamente revisão14/15 sobre16. Comparar movimentos/recibos e conciliar; reversão de código não autoriza perda de fatos novos.
- Outra pasta/TEMP/computador: o vínculo de autoridade impede a cópia de se autoproclamar principal. Rebind/instalação explícita e restauração integral em outro local continuam NÃO EXECUTADOS. Se o TEMP deixar de existir, não editar caminho silenciosamente nem escolher o banco pela data/nome; preparar nova implantação sintética rastreada.

Save cancelado real, SaveAs inválido, reconstrução e banco copiado reaberto foram executados. Energia/crash, disco cheio, falha física de commit e recuperação operacional humana não foram demonstrados. Regras fiscais, logs/retenção e semântica de custódia continuam pendentes de aprovação operacional.


## Confirmação da reabertura E2

O usuário habilitou manualmente as macros e a conferência em novo contexto do projeto passou. C1_UltimoResultado indicou CACHE_AUDITADO antes de qualquer carregamento manual. E1_SessaoReaberta, executado no projeto da candidata E2, registrou PASS com revisão16,16movimentos/recibos,5ativos,0Pending ePROJECTION=VERIFIED. [Evidência](./_desenvolvimento/entrega2_20261102T215146Z/continuacao_reabertura_20261103T010409Z/evidencias_reabertura.json). O aviso já foi atendido nesta sessão; a referência anterior a aguardando clique é histórica.

Os hashes entregues não mudaram. A consulta encontrou Workbook.Saved=False em memória após inicialização/cálculo, sem novo Save. Não interpretar isso como persistência adicional nem como nova movimentação. A abertura em outra pasta/instalação e a conciliação operacional continuam pendentes.


## Continuação do fechamento: contrato3 / esquema7

**Pacote compilado em andamento, não liberado para homologação.** [Artefatos e hashes](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_20261103T2147Z/FECHAMENTO_ENTREGA2_EM_ANDAMENTO.md). Office somente no TEMP, sobre a candidata sintética. O pacote privado é preservação; abrir suas cópias não estabelece autoridade.

A autoridade atual está em C:/Users/MARCOS~1/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/Entrega2_Restauracao_Controlada_20261103T0146Z/. Requer o XLSM/ACCDB Final e o registro compartilhado Controle_E2_143C55EC132B48398413268B265117AF.accdb no caminho persistido. Não recriar ou copiar o registro para eleger uma autoridade paralela. Não executar C3_InstalarReferencia novamente: é bootstrap da referência congelada, não rotina de restauração.

Realocação sintética ensaiada: conferir Audit10=OK, READY, Pending0, WriterLeaseFREE, BaseID/linhagem/revisão e hash do banco fechado; preparar destino TEMP vazio; usar C3_Realocar com destino, responsável, justificativa e SHA esperado. O método preserva cópias e registra a nova geração; a origem é revogada. Reabrir o destino, habilitar macros manualmente se o Office exigir, diagnosticar e reconstruir/verificar a projeção antes de gravar. O ensaio preservou rev16, gerou geração2 e recusou a origem. Não repetir em caso de falha ou resultado incerto.

As etapas entre arquivos não são uma transação única. Interrupção exige diagnóstico das duas bases e do registro; não apagar evidências, zerar lease nem escolher por data/nome. Não usar esse procedimento como prova de restauração de backup antigo com fatos posteriores. Recuperação sem origem/registro disponível continua pendente.

Conciliação: CandidatoConciliarIntencao solicita OperationID, análise, responsável, motivo e confirmação do plano. Somente PREPARED não ENTRY, sem efeitos comprovados e com origem/versão intactas, pode receber CONCILIADO_SEM_EFEITOS. A API foi exercitada com confirmação sintética; o diálogo completo ainda não passou por teste. Fatos confirmados, recibos ausentes/divergentes, histórico alterado, ENTRY/rascunhos e lease órfão permanecem bloqueados. Não usar outro OperationID para contornar Pending. Preservar o recibo e revalidar auditoria/projeção depois.

Nova sessão atual: macros desativadas pelo Office. O proprietário deve habilitar conteúdo apenas nessa candidata confiável, sem alterar políticas globais. Compilar e executar não são equivalentes: a compilação passou; a regressão ampliada aguarda essa ação. [Sequência precisa de retomada](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_20261103T2147Z/FECHAMENTO_ENTREGA2_EM_ANDAMENTO.md#retomada-sem-reiniciar-auditoria).


## Procedimento vigente — candidata revisão 26 (03/11/2026)

Esta seção atualiza o estado corrente, preservando o histórico. Somente laboratório sintético; Entrega 2 em fechamento.

1. Usar o XLSM Final já aberto no TEMP Entrega2_Restauracao_Controlada_20261103T0146Z, com o ACCDB Final ao lado e registro compartilhado no caminho persistido. Não operar a cópia de preservação privada.
2. O proprietário deve clicar Habilitar Conteúdo manualmente. Nenhuma política global deve ser alterada.
3. Alt+F11, selecionar um módulo da candidata Final; Ctrl+G. Antes de carregar cache manualmente: `? C1_UltimoResultado`. Esperado CACHE_AUDITADO. Se vazio/bloqueado, registrar e investigar, sem mascarar a inicialização.
4. `? C1_Diagnostico()`. Checkpoint 26: AUDIT=OK;PROJECTION=VERIFIED;REV=26;MOV=26;RECEIPTS=26;ACTIVE=7;PENDING=0. O diagnóstico pode deixar Saved=False porque desprotege/reprotege para ler Ignore. O experimento confirmou cache sem esse efeito e hash do arquivo inalterado durante a consulta.
5. Alt+F8 → CandidatoAbrirEntrada; pesquisar F2A0001/F2F0002 e conferir o episódio/token e local. Essa consulta após a reabertura final ainda está bloqueada por macros. Não repetir as suítes de cadastro na massa final já movimentada.

### Realocação e restauração

A realocação ao vivo foi ensaiada: origem revogada, geração2 e teste frio inicial na nova pasta aprovados. Conferir auditoria, READY, FREE, Pending0, BaseID, linhagem, revisão e hash; preparar destino TEMP vazio. Na candidata de origem correta, a assinatura é `C3_Realocar(destino, responsavel, motivo, sha256EsperadoDoBanco)`. Todos os argumentos são obrigatórios e o hash deve ser obtido sem conexões de escrita abertas. Não executar C3_InstalarReferencia novamente.

A função deve devolver REALOCADO_REQUER_REABERTURA_E_VERIFICACAO. Fechar somente a origem, abrir destino, habilitar macros manualmente, conferir identidade/geração/revisão, reconstruir e verificar antes de gravar. Se devolver PENDENTE_OU_RECUSADO, conservar os dois conjuntos e o registro, interromper e diagnosticar. Não repetir automaticamente, apagar cópias ou editar caminhos/geração. Não há transação única entre os arquivos.

Esse procedimento exige origem e registro íntegros disponíveis. Recuperação de desastre sem eles e recuperação de backup antigo com fatos posteriores ainda não foram implementadas/ensaiadas. O backup real revisão24 foi reaberto/auditado; não substituir a autoridade26 por ele. Preservar ambos e conciliar fatos25/26 e recibos antes da decisão. Cópia de ACCDB com mesmo BaseID não assume autoridade; copiar o registro para criar autoridade paralela é proibido.

### Conciliação limitada pela interface

Alt+F8 → CandidatoConciliarIntencao. Informar OperationID exato, responsável experimental e motivo; ler o plano e confirmar somente se demonstrado PREPARED não ENTRY sem efeitos, origem/versão intactas e auditoria pronta. A decisão cria Attempt, recibo e ABORTED, preservando snapshot e justificativa. Plano obsoleto exige nova análise. Nunca trocar OperationID para contornar Pending.

O diálogo completo passou para AA80FD0A-5A9A-4113-921C-12C9046D1A26, já conciliado; não repeti-lo como intenção pendente. Somente depois do recibo houve nova tentativa34021 confirmada uma vez. O procedimento não desfaz commit, não inventa recibo e não resolve ENTRY/rascunho, lease órfão, histórico alterado ou estado incerto. Esses casos permanecem PENDENTE_CONCILIACAO, sob decisão específica.

### Manutenção e limites

ADJUST descritivo foi testado em matriz e filial, com motivo/responsável e fato substitutivo, preservando o anterior. Retificação fiscal, dinheiro, chave, placa/data, histórico encerrado, anulação/estorno e exclusão confirmada são operações diferentes e permanecem NÃO DISPONÍVEIS sem regra aprovada. Não corrigir diretamente por SQL.

Após commit confirmado e falha no Excel, não repetir o movimento: auditar recibo e executar CandidatoReconstruir. Falha pré-commit pode reverter efeitos e manter intenção; isso não é recuperação integral. Não liberar lease por idade. Backup produz conjunto verificado sem retenção automática. Energia, corrupção física, dois escritores Excel e restauração de desastre continuam sem aprovação.


## Retomada de04/11/2026 — QueryClose, revisao26

[Relatorio da retomada](./_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_reabertura_final_20261103T233016Z/RELATORIO_RETOMADA_QUERYCLOSE.md). Os registros anteriores permanecem historicos. A referencia R26 reabriu com CACHE_AUDITADO e diagnostico VERIFIED; pesquisas reais matriz/filial confirmaram episodios e valores150,3333/456,7891. Foi reproduzido erro1004 no X de Cad_NFE: incremento arbitrario de U1 na planilha ativa protegida. Removida somente essa escrita em QueryClose de Cad_NFE e Cad_NFE_Saida na candidata. Compilacao real e regressao de X/botao Fechar dos dois formularios PASS.17tabelas/ACCDB inalterados;39originais com hashes iguais.

Novo XLSM preservado: SHA256 efd1916395159da3d30ecc99838ff6e0e44dd9e0ca2b13b29570a3b59fe247e0; ACCDB9fc0bfe8001120d2804b8f56836949249606e2030a6af431a71b1757c80267dc. Pacote privado Entrega2_R26_QueryClose_20261104T1221Z_PENDENTE_ACEITE.72componentes;436assinaturas dos5formularios preservadas; designers reserializados, sem alegar identidade binaria.

FORMULA_VALUE apos edicao/formularios preservado como FAIL de tentativa; recalc completo recuperou VERIFIED, sem alterar fatos. Causa fina ainda nao isolada. Arquivo salvo e instancia exclusiva encerrada normalmente. A nova abertura do XLSM corrigido concluiu, mas Office voltou a desabilitar macros: inicializacao/fluxo final BLOQUEADOS ate habilitacao manual. Entrega2 continua em fechamento, NAO HOMOLOGADA; nenhuma liberacao de Entrega3.

Na candidata Final corrigida ja aberta no TEMP, proprietario: clicar Habilitar Conteudo somente para esse arquivo sintetico reconhecido, sem alterar politicas. Depois informar a habilitacao. Retomada: ler C1_UltimoResultado antes de carregar cache; executar C1_Diagnostico; conferir REV26/MOV26/RECEIPTS26/ACTIVE7/PENDING0 e pesquisa. Nao repetir cadastramentos na massa final. O X dos formularios continua orientando o uso do botao Fechar, por comportamento preservado. Copias no pacote privado sao preservacao, nao autoridades.


## Continuação controlada — 05/11/2026: preparação do teste ACE

Laboratório exclusivo criado em `C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/Entrega2_Restauracao_Controlada_ACE_20261105T101001Z_A2`, com BaseID e linhagem novos, sem fatos operacionais, esquema 7 preservado e registro de autoridade próprio. Não representa realocação da autoridade R26 nem autorização de instalação operacional. A projeção inicial está PENDING e requer reconstrução explícita antes da escrita. Para retomar o teste real, o usuário deve habilitar o conteúdo somente nesse XLSM sintético já aberto. Não alterar segurança do Office nem abrir a tentativa A1.

Evidências, hashes, checkpoint, primeira tentativa revertida e roteiro de captura: [resultado da preparação ACE](_desenvolvimento/entrega2_20261102T215146Z/fechamento_20261103T011847Z/continuacao_ACE_20261105T101001Z/RESULTADO_PREPARACAO_ACE.md). Registros anteriores preservados integralmente.
