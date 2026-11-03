# Protótipos de integridade — Fase 6

**Entrega experimental, NÃO HOMOLOGADA e não integrada à versão principal.** A recomendação é condicional: identidade estável e baixa lógica só podem ser adotadas junto com leitores, fórmulas, históricos e recuperação aprovados.

## Ambiente, preservação e compilação

Os 39 originais e a referência compilada Fase 5 foram verificados por SHA-256 antes da execução; os 1.230 arquivos das cinco pastas anteriores receberam inventário de preservação. Novo checkpoint verificado em [backup](./_auditoria/fase6_20261101_224949Z/backup/). A referência F5 mantém 56 componentes e 17 abas; o módulo experimental adicional resulta em 57 componentes no mestre e 62 nas fixtures com cinco abas experimentais, contados nos arquivos efetivamente salvos.

A importação, edição incremental, compilação e execução ocorreram no Editor VBA nativo do Excel em TEMP fora do OneDrive. Não se repetiu a sondagem bloqueada VBProject/COM. Compilação real comprovada nos registros [R2](./_auditoria/fase6_20261101_224949Z/evidencias/compilacao-real-r2.json), [R3](./_auditoria/fase6_20261101_224949Z/evidencias/compilacao-real-r3.json) e [R4](./_auditoria/fase6_20261101_224949Z/evidencias/compilacao-real-r4.json). O arquivo foi salvo e seu código embutido inspecionado; a revisão final corresponde ao fonte esperado, considerando somente recasing e espaçamento do VBE. Os 56 componentes herdados preservam o texto funcional com literais intactos, descontado VB_Base regenerado pelo editor.

Os designers foram reserializados pelo Office: na comparação R2, 75 dos 127 streams examinados mudaram bytes. Não se declara igualdade binária dos designers ou nova regressão funcional dos UserForms nesta fase. Os FRM/FRX externos e a fixture F5 de referência permanecem intactos. Nenhum controle foi editado deliberadamente. Código de produto, regras fiscais e os sete Delete não receberam alterações.

Ocorrências técnicas preservadas: parâmetro experimental local era palavra reservada e provocou erro real de compilação, corrigido para localEstoque; coletor Ignore falhou em célula protegida, corrigido para registrar NAO_LIDO:1004; quatro Guardar: eram rótulos, corrigidos para Call Guardar: na R3. A guarda também existia em linha isolada em Preparar e foi executada nos cenários anteriores. R3 acrescentou teste do marcador inválido/restaurado, ambos executados.

Um registro preliminar de compilação e a cópia chamada Fixture_Mestre_Compilada_57.xlsm foram gerados antes do salvamento efetivo e ainda possuíam 56 componentes. Estão preservados como preparatórios/obsoletos; **não são evidência da compilação integrada**. Utilizar os registros persistidos e R2/R3 citados.

Ao trocar o módulo da R2, uma importação entrou temporariamente no projeto R1 aberto. Essa janela foi fechada sem salvar; o disco R1 foi inspecionado com 57 componentes, sem o módulo adicional automático. A importação correta foi repetida com o projeto R2 selecionado. O incidente não envolveu fontes principais.

Os resultados consolidados são **43 cenários únicos: 28 PASS e 15 FAIL**, além de dois checks da guarda PASS. As 86 tentativas das quatro revisões estão preservadas. Três nomes de cenários R3 careciam de assertivas específicas; sua cobertura foi registrada como NÃO EXECUTADO e a instrumentação corrigida/reexecutada na R4. Não se contabilizou o PASS genérico como prova da guarda ausente.

Metadados SHA intermediários que usavam texto do caminho foram identificados na conferência final e corrigidos para hash dos bytes; versões anteriores foram preservadas. Isso não afetou a verificação inicial/cópias, que já usava bytes. Veja [correção dos metadados](./_auditoria/fase6_20261101_224949Z/evidencias/correcao-metadados-sha256.json).

Mestre experimental arquivado: [Fixture_Arquitetura_Fase6_R4_EXPERIMENTAL_NAO_HOMOLOGADA.xlsm](./_auditoria/fase6_20261101_224949Z/entrega/Fixture_Arquitetura_Fase6_R4_EXPERIMENTAL_NAO_HOMOLOGADA.xlsm). SHA-256 dos bytes: 7f0f41f7336ebae8f7cfea2b698335be4ef346e78743cde34d899299dd848de1. Arquivado no workspace somente para preservação; executar exclusivamente nova cópia em TEMP.

## Dados e coleta

Cada cenário usa cópia independente do mestre. Entradas sintéticas matriz A/B/C nas linhas 4–6 e filial D/E/F nas linhas 4–6; NF-e 11001–11003 e 21001–21003, saída 22001, valor de entrada 900, placas FIC6A01 etc., data sintética 01/11/2026 e chave String de 44 dígitos com zero inicial. IDs são GUIDs locais; conteúdo não representa documento fiscal autorizado.

M/N/O são fórmulas sentinela de valor, NF-e e veículo. Demonstração Q/R referencia diretamente o registro; S usa INDIRECT por posição; T resolve por ProjectionID. Esses sentinelas simplificam o diagnóstico e não equivalem a todas as fórmulas/UDF operacionais. O registro filial D é diferente do A matriz, evitando afirmar que dois veículos sem vínculo comprovado sejam a mesma identidade.

Cada antes/depois completo contém 28.320 células: A1:X40 das 17 abas e A1:X100 das cinco abas experimentais. Registra tipo, valor, fórmula, NumberFormat, Ignore e proteção. Não captura fonte/borda/objetos nem todas as células da pasta. Ignore não legível em proteção é NAO_LIDO:1004, nunca False presumido. Assinaturas de compensação excluem o diário, que deve mudar para registrar a tentativa.

O arquivo principal do cenário representa a persistência feita pelo motor. estado_memoria_nao_promover.xlsm é cópia forense de memória, potencialmente inconsistente. O verificador fecha e reabre o principal sem salvar alterações posteriores e testa recibos/pendências. Isso é reabertura controlada, não crash ou queda de energia. A reentrada isolada após a operação é verificada em memória/imagem forense; a sequência de transferências subsequentes provoca persistência das novas etapas. Não se infere validação fria de todo registro adicional a partir de uma única assertiva de conclusão.

## Comparação das alternativas

| Critério | A — Delete físico | B — Limpeza sem deslocamento | C — Baixa lógica por identidade |
|---|---|---|---|
| Fórmulas diretas/indiretas | Quebra #REF! e troca silenciosa de veículo | Endereço permanece, mas conteúdo some; reutilização mostra novo veículo | Preserva a projeção; consultas por ID resistem à troca de posição, mas fórmulas físicas continuam vulneráveis |
| Auxiliares filial | Delete A:M não acompanha N/O | Preserva fórmula, porém referência a campo limpo deixa de representar o registro | M/N/O sentinelas preservadas nos fluxos normais do modelo |
| Histórico anual/reentrada | Busca por número/placa ambígua | Histórico perde fonte; linha reutilizada não prova nova passagem | Novo episódio após venda; exige retirar COUNTIF por placa como vínculo de encerramento |
| Pesquisa/filtros/ListBox | Linhas e seleções mudam | Lacunas interrompem filtros legados | Pesquisa rápida atual continua mostrando baixado; elegibilidade e IDs precisam entrar em todos os leitores |
| Duplicidade | NF-e/placa não distinguem episódio | Mesmo risco, acrescido de reutilização | IDs declarados distinguem contextos; projeção/histórico ambíguo bloqueados; legado ainda exige decisão humana |
| Recuperação | Alteração estrutural impede compensação segura por célula | Campos podem ser compensados antes do Save, mas vínculo precisa ser conferido | Snapshot limitado e estado persistente permitem contenção; proteção/Save/crash continuam limites |
| Compatibilidade/complexidade | Compatível com gravadores legados, insegura nos casos provados | Exige adaptar filtros e reserva de linha; não resolve o histórico | Maior esforço; migração conjunta de 17 abas/leitores; não pode ser integrada parcialmente |

Não foi feita troca indiscriminada por EntireRow.Delete. A reprodução usa somente intervalos legados identificados nos fluxos de Venda/Transferência/Devolução; os demais pontos estão no mapa estático, não tiveram política reescrita.

## Resultados decisivos

A e B falham na integridade dos vínculos. Em B_REUSO, a referência antiga passa a FIC6G01. Em B_LACUNAS, o laço legado para na primeira posição limpa e deixa de percorrer os registros seguintes.

C preserva a associação no modelo de Venda, Transferência e Devolução, mantém a chave textual e fórmulas sentinela, separa reentrada do mesmo veículo em novo EpisodeID e permite transferências sucessivas com versão. Same OperationID não grava novamente; payload diferente e versão obsoleta são recusados. Zero/múltiplas projeções ou histórico ambíguo bloqueiam a operação correspondente.

Isso não torna C pronta: C_REORDENACAO R3 falha também no resolvedor depois de permutar via Formula, alterando a interpretação de textos. O caso adicional C_REORDENACAO_TIPOS preserva tipos/formatos/Ignore e passa no resolvedor por ID, mas a referência física continua apontando a outro veículo — **FAIL**. C_LEITOR_LEGADO executa a pesquisa rápida existente com AuditoriaListaFake, conta registros físicos e reproduz a fórmula anual por placa; esses subtestes expõem leitores incompatíveis com baixa lógica/reentrada. Não se deve apresentar toda a estratégia C como PASS.

Os ensaios de falha no primeiro ao quarto destino, antes da baixa e antes do Save comprovam compensação limitada nas condições sintéticas testadas. C_FALHA_BAIXA e C_FALHA_COMPENSACAO conservam FAIL de recuperação integral; a contenção exige pendência e bloqueio de repetição. C_FALHA_SAVE provoca erro real 1004 em SaveCopyAs para pai inexistente; não equivale ao Save legado. C_APOS_SAVE demonstra a janela entre XLSM salvo e recibo externo, com conciliação após reabrir.

Os cinco FAIL F5 permanecem FAIL da aplicação legada. Os testes F6 são execuções funcionais de um **motor experimental VBA compilado**, usando a classe de compensação existente; não são nova execução dos eventos privados dos gravadores do produto. A pesquisa rápida existente é chamada diretamente somente no cenário específico de leitor. Não houve teste COM equivalente nem simulação JS usado como prova de macro nesta fase; JavaScript apenas organiza/hash/compara evidências locais.

## Rastreabilidade e repetição

Fonte final: [modPrototiposFase6_R4.bas](./_auditoria/fase6_20261101_224949Z/prototipos/modPrototiposFase6_R4.bas); cópia do [texto embutido final](./_auditoria/fase6_20261101_224949Z/prototipos/modPrototiposFase6_R4_embutido.bas). Revisões anteriores e checkpoints não foram apagados. A [matriz](./MATRIZ_REGRESSAO_FASE6.md) aponta cada tentativa e seu estado, incluindo evidência parcial.

Copiar o mestre final para nova raiz TEMP marcada, fora do OneDrive. O arquivo precisa manter prefixo Fixture_Arquitetura_Fase6_, marcador de pasta e Menu Principal!Z1=AUDITORIA_VBA_FICTICIO. Executar F6_TestarGuardas e F6_RodarLote com índices existentes no F6_Casos; a suíte recusa sobrescrita de pasta de cenário. Usar lotes pequenos e sessão exclusiva; a R2 apresentou 0x8000FFFF, erro Run 1004 e Memória insuficiente depois de vários ciclos de abertura. Não repetir sobre uma fixture incerta; criar cópia nova. Não habilitar políticas globais de macro nem usar VBProject bloqueado.

## Recomendação e limitações

Recomenda-se C condicionada a D01/D02 e à escolha de autoridade/diário D04/D06, com migração reversível dos leitores e fórmulas. A identidade estável deve entrar antes de qualquer nova política de baixa e cada efeito precisa ter vínculo explícito e recibo conferido. Operação incerta fica bloqueada para conciliação, inclusive diante de novo OperationID.

Não houve migração legada real, teste multiwriter, crash/energia, corrupção, edição versionada de históricos, validação de todos os formulários ou implantação de armazenamento durável. O modelo de diário não é write-ahead de produção, não persiste todas as pré-imagens e não promete ACID/rollback. Nenhuma mudança fiscal, XML, Selenium, consulta externa ou publicação foi realizada.


## Estado das sessões ao finalizar

A sessão exclusiva R4 foi fechada normalmente depois de confirmar que continha somente a fixture mestre. A sessão anterior com falha de memória conservou duas janelas de teste (mestre R2 e C_COLISAO); as tentativas de fechamento normal não foram concluídas. O Editor VBA antigo fechou. Não houve encerramento forçado/global. As cópias já estão preservadas; eventual fechamento manual deve atingir somente essas duas fixtures e descartar alterações posteriores. [Registro das sessões](./_auditoria/fase6_20261101_224949Z/evidencias/encerramento-sessoes-teste.json).
