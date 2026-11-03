# Matriz de regressão — Fase 8

Sistema principal NÃO HOMOLOGADO; Portão 2 NÃO autorizado. PASS abaixo significa execução funcional de VBA experimental no Excel real, nunca aprovação de gravador operacional. Não foram usados COM equivalente ou simulação como evidência funcional nova.

## Contabilidade e revisões

70 asserções únicas com PASS na seleção de resultados mais recente de cada cenário; 4 volumes de benchmark com PASS separados. As 26 asserções integradas foram repetidas após a guarda final; a repetição não aumenta a quantidade de cenários. Cold/warm de datas são hipóteses distintas. Os 15 guardas F7 são regressão dirigida, identificada como repetida. Resultados de datas, Save e benchmark/cold anteriores à guarda final são evidências daquela revisão; a guarda final foi verificada por 3 negativos e 26 integrados, sem mudar a serialização de datas.

Quatro FAIL de reprodução F7, três FAIL da guarda de selo antes da correção, ERRO_INTEGRADO/1004 e duas tentativas de benchmark100/1004 permanecem arquivados. A instrumentação do teste de selo também parou em erro 58 após as asserções por conflito de nome log/snapshot; corrigido e registrado. Compilação duplicada e bloqueio inicial de macros da sessão fria são ocorrências técnicas, não testes aprovados por inferência.

## Fixtures e estados de referência

| Grupo | Dados sintéticos / estado anterior | Ação e comparação posterior | Evidência de estado |
|---|---|---|---|
| Integrado | VID_A com EP_A1 encerrado/venda e EP_A2 ativo; VID_B/VID_C ativos; mesma NF-e88001 em contextos diferentes; chave44 com zero inicial; entradas1000/2200/1500/3100 e saídas1370/2790/1940/4110; fatores2/3/4/5 | Pesquisa por elegibilidade, token resolvido, controles reais, anual por EP; Sort nativo, lacuna/linha12 e novo ocupante9900; versões/duplicidades/conteúdo/pendências; T/D/T sintéticas | [Antes](./_auditoria/fase8_20261102_121422Z/evidencias/revisao_final/integrado_antes_r2.tsv), [depois](./_auditoria/fase8_20261102_121422Z/evidencias/revisao_final/integrado_depois_r2.tsv):7 abas A1:R40 com tipo/valor/fórmula/formato/Ignore |
| Datas | 12 casos: duas políticas locais equivalentes, ISO literal/escape, data2030, formato 0.00, texto inválido, texto ISO, hora, overflow, BR literal e BR recusado | Salvar/fechar; novo processo Excel; comparar canônico sem restaurar formato | [Fixture](./_auditoria/fase8_20261102_121422Z/evidencias/execucoes/Fixture_F8_Datas.xlsm), D/E de _F8_Datas e TSVs. E7 converteu rótulo0.00 em0; D7 conserva o canônico correto |
| Save | Base salva antes da intenção; nova intenção PREPARADA em memória | BeforeSave exclusivo cancela uma chamada real de Save; sem recibo, pendência impede outro OperationID; memória forense preservada | [Memória](./_auditoria/fase8_20261102_121422Z/evidencias/execucoes/save_recusado_memoria_nao_promover.xlsm). Não prova disco cheio ou durabilidade de intenção no disco |
| Benchmark | n episódios ativos, IDs explícitos/únicos, chave44, data ISO e valores distintos | 200 consultas lineares/índice/verificadas, lista validada, fake e ListBox real; projeção em array; recálculo nativo min(n,200), Save; novo processo reabre e valida último ID/quantidade/tipo/chave/Ignore | [Métricas estruturadas](./_auditoria/fase8_20261102_121422Z/benchmark-consolidado.json); amostra de recálculo não representa todas as fórmulas do produto |
| F7 frio | 12 pares de cenários/diários copiados e hash conferido | 15 guardas incluindo mesmo/novo ID para pendência; conteúdo anterior=posterior; sem escrita de negócio | [Manifesto de insumos](./_auditoria/fase8_20261102_121422Z/insumos-regressao-f7.json). Guarda não é recuperação |
| Selo de erro | PID_B1 válido é alterado para data textual/formato material/chave curta; catálogo recebe deliberadamente ERRO coincidente | Antes ENCONTRADO/FAIL; depois CONTEUDO_DIVERGENTE/PASS, sem autorizar mutação | [Antes](./_auditoria/fase8_20261102_121422Z/evidencias/revisao_final/selo_erro_antes.tsv), [depois](./_auditoria/fase8_20261102_121422Z/evidencias/revisao_final/selo_erro_corrigido.tsv), snapshot e fixtures no mesmo diretório |

## Resultados executados por cenário

### Leitores/rep/UI e formulas integrados

Procedimento: F8_Integrado. Modalidade: execução funcional VBA experimental. Fixture: Fixture_F8_Integrada_R2.xlsm. Evidência: [TSV](./_auditoria/fase8_20261102_121422Z/evidencias/revisao_final/integrado_f8_r2.tsv).

| Cenário | Resultado | Esperado | Observado |
|---|---|---|---|
| F8_02_REENTRADA_PESQUISA | PASS | somente EP_A2 ativo | ENCONTRADO;EP_A2 |
| F8_FORMULARIO_REAL | PASS | controles reais com EP_A2 | N_Nfe;PLACA;txt_Valor;1500 |
| F8_02_ANUAL_EPISODIOS | PASS | EP_A1 SIM; EP_A2 NAO | SIM;NAO |
| F8_02_ANUAL_AUSENTE | PASS | erro explicito | Erro 2042 |
| F8_03_TOKEN_APOS_SORT_REUSO | PASS | PID_A2 linha12, indice invalidado | ENCONTRADO;12 |
| F8_03_FORMULA_ID | PASS | mesmo veiculo/episodio apos sort e reuso | FIC8A01;1500 |
| F8_04_AUX_DISCRIMINANTE | PASS | A2 margem440/aux6000; B6600; C15500 | 440;6000;6600;15500 |
| F8_CHAVE_TIPOS_FORMATOS | PASS | String44/@/Ignore | String;44;@ |
| F8_LISTINDEX_NEGATIVO | PASS | nenhuma selecao autoriza escrita | -1 |
| F8_LACUNAS_CONTAGEM | PASS | 4 ativos incluindo novo ocupante e lacuna | ENCONTRADO;4 |
| F8_RETORNO_BAIXADO | PASS | BAIXADO | BAIXADO |
| F8_HISTORICO_ENTRADA | PASS | historico preservado | ENCONTRADO |
| F8_TOKEN_OBSOLETO_SEM_MUTACAO | PASS | VERSAO_OBSOLETA | VERSAO_OBSOLETA |
| F8_EDICAO_FORA_CONTRATO | PASS | CONTEUDO_DIVERGENTE | CONTEUDO_DIVERGENTE |
| F8_02_EPISODIO_DUPLICADO | PASS | erro explicito sem primeiro resultado | IDENTIDADE_AMBIGUA |
| F8_DUPLICADO_SEM_EVENTO | PASS | indice nao oculta duplicado | DUPLICADO |
| F8_RETORNO_INEXISTENTE | PASS | INEXISTENTE | INEXISTENTE |
| F8_RETORNO_AMBIGUO | PASS | IDENTIDADE_AMBIGUA | IDENTIDADE_AMBIGUA |
| F8_TRANSFERENCIAS_DEVOLUCAO | PASS | T/D/T mesmo EP, versao5 | CONCLUIDA_FIXTURE;CONCLUIDA_FIXTURE;CONCLUIDA_FIXTURE;ENCONTRADO |
| F8_02_ANUAL_TRANSFERIDO | PASS | transferencia/devolucao nao herdam Venda A1 | NAO |
| F8_OPERATIONID_REPETIDO | PASS | mesmo recibo sem duplicar intencao | PREPARADA;JA_CONCLUIDA |
| F8_PAYLOAD_DIVERGENTE | PASS | RECUSADA_PAYLOAD | RECUSADA_PAYLOAD |
| F8_RETORNO_PENDENTE | PASS | OPERACAO_PENDENTE | OPERACAO_PENDENTE |
| F8_NOVO_ID_NAO_CONTORNA | PASS | OPERACAO_PENDENTE | OPERACAO_PENDENTE |
| F8_RETORNO_CONCILIACAO | PASS | AGUARDANDO_CONCILIACAO | AGUARDANDO_CONCILIACAO |
| F8_ESTADOS_APPLICATION | PASS | estado de eventos anterior | Verdadeiro |

### Contrato data antes Save

Procedimento: F8_DatasPreparar. Modalidade: execução funcional VBA experimental. Fixture: Fixture_F8_Datas.xlsm. Evidência: [TSV](./_auditoria/fase8_20261102_121422Z/evidencias/execucoes/datas_sessao1.tsv).

| Cenário | Resultado | Esperado | Observado |
|---|---|---|---|
| F8_01_FORMATOS_EQUIVALENTES | PASS | equivalencia somente pela politica explicita | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=LEGACY_D14_LOCAL\|FMT=D14_LOCAL;DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=ISO_LITERAL\|FMT=ISO_YMD |
| F8_01_DATAS_DISTINTAS | PASS | datas distintas possuem canonico distinto | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E7392000000000\|POLICY=ISO_LITERAL\|FMT=ISO_YMD |
| F8_01_FORMATO_MATERIAL_RECUSADO | PASS | erro de formato | ERRO:FORMATO_0.00 |
| F8_01_INVALIDAS_RECUSADAS | PASS | texto, hora e overflow recusados | ERRO:DATA_TIPO_String;ERRO:DATA_INVALIDA |
| F8_01_BR_ESTRITO | PASS | BR literal aceito; local recusado sob BR estrito | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=BR_LITERAL\|FMT=BR_DMY;ERRO:FORMATO_m/d/yyyy |

### Nova instancia sem restaurar formato

Procedimento: F8_DatasFrio. Modalidade: execução funcional VBA experimental. Fixture: Fixture_F8_Datas.xlsm. Evidência: [TSV](./_auditoria/fase8_20261102_121422Z/evidencias/execucoes/datas_sessao_fria.tsv).

| Cenário | Resultado | Esperado | Observado |
|---|---|---|---|
| F8_01_FRIO_LEGADO_BR | PASS | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=LEGACY_D14_LOCAL\|FMT=D14_LOCAL | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=LEGACY_D14_LOCAL\|FMT=D14_LOCAL;FMT_ANTES=dd/mm/yyyy;FMT_ATUAL=m/d/yyyy |
| F8_01_FRIO_LEGADO_LOCAL | PASS | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=LEGACY_D14_LOCAL\|FMT=D14_LOCAL | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=LEGACY_D14_LOCAL\|FMT=D14_LOCAL;FMT_ANTES=m/d/yyyy;FMT_ATUAL=m/d/yyyy |
| F8_01_FRIO_ISO_LITERAL | PASS | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=ISO_LITERAL\|FMT=ISO_YMD | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=ISO_LITERAL\|FMT=ISO_YMD;FMT_ANTES=yyyy"-"mm"-"dd;FMT_ATUAL=yyyy"-"mm"-"dd |
| F8_01_FRIO_ISO_ESCAPE | PASS | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=ISO_LITERAL\|FMT=ISO_YMD | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=ISO_LITERAL\|FMT=ISO_YMD;FMT_ANTES=yyyy-mm-dd;FMT_ATUAL=yyyy-mm-dd |
| F8_01_FRIO_OUTRA_DATA | PASS | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E7392000000000\|POLICY=ISO_LITERAL\|FMT=ISO_YMD | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E7392000000000\|POLICY=ISO_LITERAL\|FMT=ISO_YMD;FMT_ANTES=yyyy"-"mm"-"dd;FMT_ATUAL=yyyy"-"mm"-"dd |
| F8_01_FRIO_FORMATO_MATERIAL | PASS | ERRO:FORMATO_0.00 | ERRO:FORMATO_0.00;FMT_ANTES=0;FMT_ATUAL=0.00 |
| F8_01_FRIO_DATA_INVALIDA_TEXTO | PASS | ERRO:DATA_TIPO_String | ERRO:DATA_TIPO_String;FMT_ANTES=@;FMT_ATUAL=@ |
| F8_01_FRIO_TEXTO_PARECE_DATA | PASS | ERRO:DATA_TIPO_String | ERRO:DATA_TIPO_String;FMT_ANTES=@;FMT_ATUAL=@ |
| F8_01_FRIO_DATA_COM_HORA | PASS | ERRO:DATA_INVALIDA | ERRO:DATA_INVALIDA;FMT_ANTES=yyyy"-"mm"-"dd;FMT_ATUAL=yyyy"-"mm"-"dd |
| F8_01_FRIO_DATA_OVERFLOW | PASS | ERRO:DATA_INVALIDA | ERRO:DATA_INVALIDA;FMT_ANTES=yyyy"-"mm"-"dd;FMT_ATUAL=yyyy"-"mm"-"dd |
| F8_01_FRIO_BR_LITERAL | PASS | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=BR_LITERAL\|FMT=BR_DMY | DATE_ONLY\|EXCEL1900\|DOUBLE_HEX=40E69F0000000000\|POLICY=BR_LITERAL\|FMT=BR_DMY;FMT_ANTES=dd"/"mm"/"yyyy;FMT_ATUAL=dd"/"mm"/"yyyy |
| F8_01_FRIO_BR_ESTRITO_NAO_LOCAL | PASS | ERRO:FORMATO_m/d/yyyy | ERRO:FORMATO_m/d/yyyy;FMT_ANTES=m/d/yyyy;FMT_ATUAL=m/d/yyyy |

### Cancelamento real de Save por evento

Procedimento: F8_SaveRecusado. Modalidade: execução funcional VBA experimental. Fixture: Fixture_F8_Save_Recusado.xlsm. Evidência: [TSV](./_auditoria/fase8_20261102_121422Z/evidencias/execucoes/save_recusado_f8.tsv).

| Cenário | Resultado | Esperado | Observado |
|---|---|---|---|
| F8_SAVE_REAL_CANCELADO | PASS | Save chamado, cancelado por evento e sem recibo | PREPARADA;EVENTO=Verdadeiro;SAVED=Falso |
| F8_SAVE_PENDENCIA_NOVA_OP | PASS | conciliare antes de repetir | OPERACAO_PENDENTE |

### Reabertura de vinculos e pendencia

Procedimento: F8_IntegradoFrio. Modalidade: execução funcional VBA experimental. Fixture: Fixture_F8_Integrada_R2.xlsm. Evidência: [TSV](./_auditoria/fase8_20261102_121422Z/evidencias/execucoes/integrado_frio.tsv).

| Cenário | Resultado | Esperado | Observado |
|---|---|---|---|
| F8_FRIO_PENDENCIA | PASS | AGUARDANDO_CONCILIACAO | AGUARDANDO_CONCILIACAO |
| F8_FRIO_CONTEUDO | PASS | ENCONTRADO sem restaurar formatos | ENCONTRADO |
| F8_FRIO_ANUAL | PASS | SIM/NAO por episodio | anual por identidade |

### Reabertura dos quatro volumes

Procedimento: F8_BenchFrio. Modalidade: execução funcional VBA experimental. Fixture: Fixture_F8_Bench_{100,1000,5000,10000}.xlsm. Evidência: [TSV](./_auditoria/fase8_20261102_121422Z/evidencias/execucoes/benchmark_frio.tsv).

| Cenário | Resultado | Esperado | Observado |
|---|---|---|---|
| F8_BENCH_FRIO_100 | PASS | identidade/quantidade/String44/formato/Ignore persistentes | ENCONTRADO;100;FMT=yyyy"-"mm"-"dd |
| F8_BENCH_FRIO_1000 | PASS | identidade/quantidade/String44/formato/Ignore persistentes | ENCONTRADO;1000;FMT=yyyy"-"mm"-"dd |
| F8_BENCH_FRIO_5000 | PASS | identidade/quantidade/String44/formato/Ignore persistentes | ENCONTRADO;5000;FMT=yyyy"-"mm"-"dd |
| F8_BENCH_FRIO_10000 | PASS | identidade/quantidade/String44/formato/Ignore persistentes | ENCONTRADO;10000;FMT=yyyy"-"mm"-"dd |

### Reexecucao dirigida guardas F7

Procedimento: F7_VerificarFrio. Modalidade: execução funcional VBA experimental. Fixture: regressao/VBA_Fase7_20261102_025115Z/frio_*/Fixture_Fria.xlsm. Evidência: [TSV](./_auditoria/fase8_20261102_121422Z/evidencias/execucoes/regressao/VBA_Fase7_20261102_025115Z/frio_resultados.tsv).

| Cenário | Resultado | Esperado | Observado |
|---|---|---|---|
| FRIO_OK | PASS | JA_CONCLUIDA; sem escrita de negocio | JA_CONCLUIDA;IGUAL=Verdadeiro |
| FRIO_ORFAO | PASS | CONCILIAR_ORFAO; sem escrita de negocio | CONCILIAR_ORFAO;IGUAL=Verdadeiro |
| FRIO_TRUNCADO | PASS | CONCILIAR_TRUNCADO; sem escrita de negocio | CONCILIAR_TRUNCADO;IGUAL=Verdadeiro |
| FRIO_PENDENTE | PASS | CONCILIAR_PENDENTE; sem escrita de negocio | CONCILIAR_PENDENTE;IGUAL=Verdadeiro |
| FRIO_PENDENTE_MESMA_OP | PASS | CONCILIAR_PENDENTE | CONCILIAR_PENDENTE |
| FRIO_DIVERGENTE | PASS | CONCILIAR_BASE; sem escrita de negocio | CONCILIAR_BASE;IGUAL=Verdadeiro |
| FRIO_POS_SAVE | PASS | CONCILIAR_ESTADO; sem escrita de negocio | CONCILIAR_ESTADO;IGUAL=Verdadeiro |
| FRIO_HISTORICO | PASS | CONCILIAR_EFEITOS; sem escrita de negocio | CONCILIAR_EFEITOS;IGUAL=Verdadeiro |
| FRIO_AUSENTE | PASS | CONCILIAR_AUSENTE; sem escrita de negocio | CONCILIAR_AUSENTE;IGUAL=Verdadeiro |
| FRIO_PAYLOAD | PASS | RECUSADA_PAYLOAD; sem escrita de negocio | RECUSADA_PAYLOAD;IGUAL=Verdadeiro |
| FRIO_VERSAO | PASS | RECUSADA_VERSAO; sem escrita de negocio | RECUSADA_VERSAO;IGUAL=Verdadeiro |
| FRIO_BAIXA | PASS | CONCILIAR_PENDENTE; sem escrita de negocio | CONCILIAR_PENDENTE;IGUAL=Verdadeiro |
| FRIO_BAIXA_MESMA_OP | PASS | CONCILIAR_PENDENTE | CONCILIAR_PENDENTE |
| FRIO_COMPENSACAO | PASS | CONCILIAR_PENDENTE; sem escrita de negocio | CONCILIAR_PENDENTE;IGUAL=Verdadeiro |
| FRIO_COMPENSACAO_MESMA_OP | PASS | CONCILIAR_PENDENTE | CONCILIAR_PENDENTE |

### Guarda final do selo

Procedimento: F8_SeloErro("corrigido"). Modalidade: execução funcional VBA experimental. Fixture: Fixture_F8_Selo_corrigido.xlsm. Evidência: [TSV](./_auditoria/fase8_20261102_121422Z/evidencias/revisao_final/selo_erro_corrigido.tsv).

| Cenário | Resultado | Esperado | Observado |
|---|---|---|---|
| F8_SELO_ERRO_1 | PASS | CONTEUDO_DIVERGENTE mesmo com selo ERRO coincidente | CONTEUDO_DIVERGENTE;ERRO:DATA_TIPO_String |
| F8_SELO_ERRO_2 | PASS | CONTEUDO_DIVERGENTE mesmo com selo ERRO coincidente | CONTEUDO_DIVERGENTE;ERRO:FORMATO_0.00 |
| F8_SELO_ERRO_3 | PASS | CONTEUDO_DIVERGENTE mesmo com selo ERRO coincidente | CONTEUDO_DIVERGENTE;ERRO:CHAVE |

## Benchmark medido no Excel16.0/20527/64bits

Uma execução por volume após corrigir a preparação. Tempos em segundos;200 consultas por método; índices por revisão não eliminam as verificações de unicidade da consulta autorizada. Memória é working set do processo antes/depois, não pico nem vazamento medido. Aplicação permaneceu estável nessas execuções e foi encerrada normalmente.

| n | Seed | Construção índice | Linear200 | Índice200 (candidato) | Verificada200 | Lista validada | ListBox real | Projeção | Recálculo amostra | Save | MiB antes/depois | Total |
|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| 100 | 0.4168 | 0.0089 | 0.0018 | 0.0151 | 0.4302 | 0.0975 | 0.0035 | 0.0217 | 0.0013 | 0.8075 | 599.7/707.6 | 3.8159 |
| 1000 | 1.2650 | 0.0262 | 0.0099 | 0.0124 | 0.4665 | 0.9699 | 0.0054 | 0.0330 | 0.0179 | 0.7952 | 564.6/746.2 | 5.6648 |
| 5000 | 4.7821 | 0.1222 | 0.0496 | 0.0128 | 0.5962 | 4.8682 | 0.0126 | 0.0998 | 0.0239 | 0.8783 | 567.4/721.6 | 13.6668 |
| 10000 | 9.3825 | 0.2432 | 0.0978 | 0.0131 | 0.8083 | 9.7135 | 0.0283 | 0.2715 | 0.0122 | 1.0629 | 597.3/778.7 | 24.1331 |

Fonte: benchmark_100/1000/5000/10000.tsv em evidencias/execucoes. Carga MSForms é atribuição da matriz ao ListBoxLista existente sem selecionar linha; não testa todos os eventos legados de clique. Recálculo é100 fórmulas para n100 e200 nos demais, com COUNTIF/INDEX/MATCH delimitados e unicidade/EP/VID; não extrapolar para dezenas de milhares de fórmulas. A listagem valida todos os n selos por célula; ~9,71s em10mil sugere custo de materialização/validação. Índice candidato não é autorização. Benchmark precede a guarda final de erro de selo e não foi repetido por mudança constante dessa guarda.

## Referências que continuam FAIL / fora do escopo

| Referência | Estado preservado | Execução nesta fase |
|---|---|---|
| V_VALIDO — Venda | FAIL de integridade F5 | NÃO EXECUTADO nos eventos reais; protótipo não corrige Delete |
| T_VALIDO — Transferência | FAIL de integridade F5 | NÃO EXECUTADO no gravador real |
| D_VALIDO — Devolução | FAIL de integridade F5 | NÃO EXECUTADO no gravador real |
| D_FALHA_BAIXA | FAIL de integridade F5 | NÃO EXECUTADO no gravador real |
| T_DUP_DESTINO | FAIL de identidade F5 | NÃO EXECUTADO no gravador real |
| C_FALHA_BAIXA / C_FALHA_COMPENSACAO | FAIL integral F6 | Guardas F7 reexecutadas e PASS de contenção; nenhuma recuperação integral nova |
| Portão 2 / migração de IDs, leitores, gravadores, sete Delete e históricos | BLOQUEADO por aprovação específica | Nenhuma integração realizada |
| Crash/energia/disco cheio/dois escritores/mudança regional/benchmark de banco | NÃO EXECUTADO | Sem substituição por modelo equivalente |
| XML/fiscal/Selenium/DETRAN/dados reais | NÃO EXECUTADO nesta fase | Sem alteração de regras nem consultas reais |

A compilação real foi executada pelo VBE; o menu Compilar ficou desabilitado após compilação e as macros rodaram. [Sessões/evidência de interface](./_auditoria/fase8_20261102_121422Z/sessao-excel-fase8.json). A inspeção de ZIP/CFB e hashes é análise estática, registrada separadamente no manifesto, não prova de execução.

## Verificação auxiliar de entrega

O script PowerShell de integridade foi tentado: **BLOQUEADO**, PSSecurityException/UnauthorizedAccess, pois scripts estão desabilitados pela política vigente. Não houve alteração da política nem bypass. Modalidade: verificador de arquivos; fora das 70 asserções VBA. A alternativa independente de SHA-256 em Node foi executada e aprovada. [Mensagem registrada](./_auditoria/fase8_20261102_121422Z/evidencias/verificador-powershell-bloqueado.json).
