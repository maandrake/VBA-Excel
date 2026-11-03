# Manual de teste, preservação e restauração — Entrega 1

**Somente laboratório sintético. NÃO utilizar com estoques, clientes ou XML reais. Não promover para produção.**

## Requisitos verificados

Windows e Excel 16.0.20527.20112 x64, ADO/ADOX/DAO e Microsoft.ACE.OLEDB.16.0 já instalados. ACE observado 16.0.20527.20016 no laboratório. Não instalar redistribuíveis nem alterar segurança. Compatibilidade de distribuição em outro computador continua pendente.

## Arquivos e integridade

[Pacote preservado](../entrega/), [manifesto SHA-256](../entrega/manifesto_artefatos.json), [checkpoint e cópias forenses](../checkpoints/), [backup inicial](../backup/).

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

[Preparar-CopiaSintetica.ps1](../entrega/Preparar-CopiaSintetica.ps1) foi preparado para: conferir os dois hashes; criar diretório TEMP novo; copiar apenas a massa sintética; recusar lease ocupado/Pending/linhagem diferente; vincular explicitamente a nova cópia e exigir reconstrução. Sua execução foi **BLOQUEADA pela ExecutionPolicy atual**. Não está validado e não foi substituído por Invoke-Expression, Bypass ou alteração de Registro. Não executá-lo sem revisão do mecanismo permitido no ambiente. A portabilidade permanece E1-03.

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

O [backup F10](../backup/Fixture_F10_Referencia.xlsm) preserva a base anterior. A reversão de código não é restauração cega de dados. Os 39 arquivos originais em Excel/Form/MDL ficaram intactos. Os [fontes](../fontes/) incluem cinco pares FRM/FRX e módulos de documento separados; não importar documentos como BAS e não renomear CodeNames. A candidata executável já contém os componentes.

Nenhuma publicação, instalação, mudança fiscal ou promoção operacional foi realizada.

## Correção de caminho e checkpoint da continuação

O banco continuou com o mesmo hash e revisão. A candidata corrigida aceita os aliases 8.3 e longo do mesmo arquivo existente, sem alterar AuthorityPath ou BaseID. Uma cópia em outro diretório continua recusada; E1-03 não foi encerrado. [Backup anterior e evidências](../continuacao_reabertura_20261102T205915Z/) preservam o XLSM de hash 9b24bec50e588551a744128912c062bdbb020905bccd41b40f62c0aee89dbad2 para reversão de código. Não substituir uma autoridade depois de novas operações sem conciliação.
