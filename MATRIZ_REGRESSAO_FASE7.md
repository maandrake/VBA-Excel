# Matriz de regressão — Fase 7

**Execução VBA experimental real no Excel, sem integração ao produto.** Consolidação por ID da assertiva: 38 verificações únicas, 34 PASS e quatro FAIL; 40 execuções de assertivas nos três arquivos finais (repetições R06 e BASE_FORMATO preservadas). Uma tentativa anterior interrompida conserva ERRO_EXECUCAO FAIL, separada dessa contagem. Não são 38 operações de estoque.

Os adaptadores possuem 18 verificações R01–R18 PASS. Sessão fria: 12 cenários, 15 verificações PASS (três recusas adicionais da mesma operação). Quatro FAIL são referências incompatíveis/portabilidade; um PASS numérico BASE_AUXILIAR_FISICO usou valores iguais e **não discrimina identidade**: o reteste com valores diferentes falhou. Não usar esse PASS como aprovação do auxiliar.

[Consolidação estruturada](./_auditoria/fase7_20261102_025115Z/evidencias/resultados-consolidados-fase7.json), [fontes/compilação](./_auditoria/fase7_20261102_025115Z/evidencias/compilacao-real-final.json), [manifesto das cópias](./_auditoria/fase7_20261102_025115Z/evidencias/manifesto-copias-execucoes.json).

| ID | Resultado | Esperado | Observado | Evidência VBA direta |
|---|---|---|---|---|
| BASE_FORMATO_REABERTURA | FAIL | selo estrito preservado ao reabrir | ANTES=Falso;formatos restaurados na fixture=14 | [formulas_reteste.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/formulas_reteste.tsv) |
| R01_PESQUISA_ATIVA_REENTRADA | PASS | somente novo episodio | 1;{69CE760D-0001-4C3F-A262-B213036CBC67} | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R02_CONTAGEM_ATIVA | PASS | 6 ativos, excluindo antigo A | 6 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R03_ANUAL_POR_EPISODIO | PASS | novo NAO; antigo SIM | NAO;SIM | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| BASE_ANUAL_PLACA | FAIL | NAO para nova entrada | SIM | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R04_LISTA_APOS_REORDENAR | PASS | token resolve linha10 A | 10 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R05_TIPOS_FORMATOS_IGNORE | PASS | selo A:L exato e String44/@/Ignore | 32:Double:11001:11001:General:Falso38:String:FICTICIO:FICTICIO:General:Falso40:String:SINTETICO:SINTETICO:General:Falso52:String:VEICULO FIC6A01:VEICULO FIC6A01:General:Falso42:String:F6_FIC6A01:F6_FIC6A01:General:Falso36:String:FIC6A01:FIC6A01:General:Falso30:Double:2026:2026:General:Falso34:String:BRANCO:BRANCO:General:Falso28:Double:900:900:General:Falso35:Double:46327:46327:dd/mm/yyyy:Falso28:Empty:<EMPTY>::General:Falso109:String:01234567890123456789012345678901234567890123:01234567890123456789012345678901234567890123:@:Verdadeiro | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R06_FORMULAS_ID_REORDENACAO | PASS | A/chave/1800/11001 | FIC6A01;01234567890123456789012345678901234567890123;1800 | [formulas_reteste.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/formulas_reteste.tsv) |
| BASE_REFERENCIA_FISICA | FAIL | A | FIC6C01 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| BASE_AUXILIAR_FISICO | PASS | 1800 | 1800 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R07_LACUNAS_CONTAGEM | PASS | 6 atraves das lacunas | 6 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R08_LISTINDEX_NEGATIVO | PASS | sem selecao | 0 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R09_SELECAO_OBSOLETA | PASS | versao divergente recusa | 0 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R10_PROJECAO_DUPLICADA | PASS | recusa, sem primeiro candidato | 0 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R11_ID_AUSENTE | PASS | nao usa placa/NFe/linha | 0 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R12_COERCAO_DETECTADA | PASS | conteudo tipado divergente bloqueia | 0 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R13_HISTORICO_AMBIGUO | PASS | nenhum primeiro historico | 0 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R14_CUSTODIA_SUCESSIVA | PASS | um atual filial, versao4 apos T/D/T | 1;Entradas Filial;4 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R15_OUTRA_PASTA_ATIVA | PASS | leitor usa proprietario explicito | 6 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R16_ESTADOS_APPLICATION | PASS | estados preservados | Verdadeiro;-4105 | [leitores_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/leitores_resultados.tsv) |
| R17_CHAVE_FORMULA_STRING | PASS | String44/@/Ignore | String;44 | [formulas_reteste.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/formulas_reteste.tsv) |
| BASE_AUXILIAR_VALOR_DISTINTO | FAIL | 1800 do episodio A | 2400 | [formulas_reteste.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/formulas_reteste.tsv) |
| R18_FORMULA_CONTAGEM_ATIVA | PASS | 3 matriz e 6 globais validados | 3;6 | [formulas_reteste.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/formulas_reteste.tsv) |
| FRIO_OK | PASS | JA_CONCLUIDA; sem escrita de negocio | JA_CONCLUIDA;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_ORFAO | PASS | CONCILIAR_ORFAO; sem escrita de negocio | CONCILIAR_ORFAO;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_TRUNCADO | PASS | CONCILIAR_TRUNCADO; sem escrita de negocio | CONCILIAR_TRUNCADO;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_PENDENTE | PASS | CONCILIAR_PENDENTE; sem escrita de negocio | CONCILIAR_PENDENTE;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_PENDENTE_MESMA_OP | PASS | CONCILIAR_PENDENTE | CONCILIAR_PENDENTE | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_DIVERGENTE | PASS | CONCILIAR_BASE; sem escrita de negocio | CONCILIAR_BASE;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_POS_SAVE | PASS | CONCILIAR_ESTADO; sem escrita de negocio | CONCILIAR_ESTADO;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_HISTORICO | PASS | CONCILIAR_EFEITOS; sem escrita de negocio | CONCILIAR_EFEITOS;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_AUSENTE | PASS | CONCILIAR_AUSENTE; sem escrita de negocio | CONCILIAR_AUSENTE;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_PAYLOAD | PASS | RECUSADA_PAYLOAD; sem escrita de negocio | RECUSADA_PAYLOAD;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_VERSAO | PASS | RECUSADA_VERSAO; sem escrita de negocio | RECUSADA_VERSAO;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_BAIXA | PASS | CONCILIAR_PENDENTE; sem escrita de negocio | CONCILIAR_PENDENTE;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_BAIXA_MESMA_OP | PASS | CONCILIAR_PENDENTE | CONCILIAR_PENDENTE | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_COMPENSACAO | PASS | CONCILIAR_PENDENTE; sem escrita de negocio | CONCILIAR_PENDENTE;IGUAL=Verdadeiro | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |
| FRIO_COMPENSACAO_MESMA_OP | PASS | CONCILIAR_PENDENTE | CONCILIAR_PENDENTE | [frio_resultados.tsv](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/frio_resultados.tsv) |

## Reprodução e estado anterior/posterior

F7_Leitores usa sete insumos forenses sintéticos da F6, copiados com hash, sem chamar os gravadores do produto. Cria cópias exec independentes, compara dados e resolve vínculos em memória. No ensaio de reordenação, A sai da linha4 para10, C ocupa4 e6 fica vazia. Token da ListBox experimental encontra A na10; a referência física S4 passa a C. A chave preserva String44/@/Ignore em A:L e na consulta corrigida.

F7_RetestarFormulas corrige a montagem da fórmula de chave no teste: inserir fórmula com formato General, depois aplicar @/Ignore, conservando retorno String. A atribuição inicial em @ deixou fórmula como texto e R06 FAIL na tentativa anterior. Esse FAIL permanece no TSV original; R06 PASS é somente o reteste específico. O procedimento F7_Leitores arquivado mantém a montagem inicial para reprodução; o roteiro exige o reteste. Não anunciar que esse primeiro procedimento isolado passa integralmente.

No reteste, o veículo C recebe valor1200 apenas na cópia, enquanto A mantém900. A consulta por ID resulta1800; M10 físico calcula2400 usando C. A alteração deliberada de C invalida seu selo e não representa edição operacional validada. As imagens de memória preservam esse estado somente para evidência.

Pesquisa/contagem: A antigo ENCERRADO é omitido e reentrada nova exibida; seis ativos globais e três matriz. A fórmula COUNTIFS do protótipo confere a quantidade junto ao adaptador validado; COUNTIFS sozinha não valida todos os vínculos. A fórmula anual por episódio resulta NAO para novo e SIM para antigo; P8 legado por placa permanece SIM/FAIL.

Snapshots de leitor_antes/depois, semente_antes/depois e formula_reteste_antes/depois possuem A1:X40 das22 abas: 21.120 células por snapshot com tipo/valor/fórmula/NumberFormat/Ignore. Não capturam toda a pasta, fonte/bordas/objetos nem estado visual completo. [Execuções](./_auditoria/fase7_20261102_025115Z/evidencias/execucoes/).

## Portabilidade do selo

BASE_FORMATO_REABERTURA FAIL: Excel normalizou dd/mm/yyyy para m/d/yyyy ao reabrir. A guarda estrita recusou a primeira leitura. Preparação experimental restaura somente NumberFormat J explicitamente conhecido no selo, sem atualizar selo nem preencher identidade. Outros erros, inclusive coerção de texto/chave em C_REORDENACAO, continuam recusados. Não há normalização silenciosa no produto; estabilidade persistente do formato/envelope segue requisito. [Diagnóstico](./_auditoria/fase7_20261102_025115Z/evidencias/falha-selo-formato-data.json).

## Sessão fria e limites

F7_PrepararFrio salva 12 livros, cada um com23 abas/63 componentes herdados+controle sintético. Fecha a sessão1 normalmente após confirmar um livro aberto e caminho do mestre. Nova instância excel /x, janela Excel36637574/VBE2497612, executa F7_VerificarFrio lendo somente arquivos persistidos. Nenhum estado global de preparação é reutilizado. [Antes](./_auditoria/fase7_20261102_025115Z/evidencias/sessao-fria-antes-fechamento.json), [depois](./_auditoria/fase7_20261102_025115Z/evidencias/sessao-fria-depois.json).

As15 guardas comparam assinatura de5.280 células/11 abas A1:X20 sem novas escritas. Os23 arquivos de fixture/diário permanecem byte a byte iguais por SHA-256. O contrato frio sintético tem nove campos TSV e um selo mínimo de Entradas!A4:L4 + Todas Saidas!A2:L2; não cobre todos os efeitos de produção. O modo novaOperacao é uma sonda da guarda, não executor completo de outro OperationID. Casos BAIXA/COMPENSACAO usam imagens dos FAIL F6 e estado pendente declarado; não tentam reparar esses dados.

## Referências obrigatórias preservadas

| Referência | Situação do produto / evidência F7 |
|---|---|
| V_VALIDO, T_VALIDO, D_VALIDO (F5) | FAIL anteriores mantidos; eventos não reexecutados. Guardas/leitores/fórmulas experimentais não corrigem gravadores. |
| D_FALHA_BAIXA (F5) | FAIL mantido; FRIO_BAIXA comprova recusa do contrato experimental, não recuperação integral. |
| T_DUP_DESTINO (F5) | FAIL mantido; R10/R13 recusam duplicidade por ID declarada, não conciliam legado. |
| C_LEITOR_LEGADO | R01/R02/R03/R18 demonstram adaptação parcial; fórmula anual legada continua FAIL. |
| C_REORDENACAO_TIPOS | R04/R05/R06 PASS nos acessos migrados; BASE_REFERENCIA_FISICA e auxiliar com valor distinto FAIL. |
| C_REORDENACAO | R12 recusa conteúdo tipado divergente; não sela dados corrompidos novamente para obter PASS. |
| C_FALHA_BAIXA / C_FALHA_COMPENSACAO | FAIL de recuperação integral F6 mantidos. Guardas frias PASS de contenção, não reparo. |

## Cobertura por modalidade

- PASS — análise estática/hash:1.986 conferências iniciais,57 procedimentos dirigidos e57 componentes herdados sem mudança canônica. Não equivalem a teste funcional.
- PASS — compilação real VBA: módulo F7 no mestre de58 componentes/17 abas, VBE sem erro e menu Compilar desabilitado depois de reabrir.
- PASS/FAIL — execução funcional VBA experimental: resultados acima, ListBox em objeto AuditoriaListaFake. Não é interação com todos os MSForms reais.
- NÃO EXECUTADO — eventos reais migrados, UI completa/UserForms, todas as43.232 fórmulas do produto, migração legada, armazenamento autoritativo de produção, escritor concorrente/lock/fencing, crash, energia, disco cheio e corrupção.
- NÃO EXECUTADO — novos testes COM equivalentes ou simulação JS/.NET; JavaScript somente analisa/hash/copia evidências.
- BLOQUEADO por aprovação do Portão2 — aplicação de IDs/baixa/diário e novos gravadores ao principal. Não foi solicitado desvio de segurança.

XML/fiscal, backup/Ribbon e fontes estabilizados não foram alterados nem reexecutados sem hipótese nova. Resultados anteriores são referência histórica. Sistema NÃO HOMOLOGADO.
