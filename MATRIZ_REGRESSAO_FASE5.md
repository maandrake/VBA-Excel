# Matriz de regressão — Fase 5

Execução local iniciada em 01/11/2026, identificador 20261101_202114Z. **Sistema NÃO HOMOLOGADO.** [Relatório](./RELATORIO_TESTES_INTEGRADOS_FASE5.md).

PASS significa que o critério específico da linha foi observado; rejeitar uma entrada inválida pode ser PASS. FAIL indica divergência de integridade, mesmo que o evento tenha terminado. BLOQUEADO indica impedimento técnico; NÃO EXECUTADO indica ausência de execução. Não somar asserções das suítes a cenários de negócio.

## Método e dados

Os eventos privados reais de Cad_NFE_Saida foram acionados pelo botão do UserForm ou pela propriedade Value do controle no Editor VBA. Não foram renomeados nem substituídos. Venda/Transferência/Cancelado usam CommandButton_Gravar_Click; Devolução usa CommandButton_GravarFilial_Click. Os quatro casos válidos tiveram interação com os botões na interface. As seleções XML e o fechamento dos relatórios modais foram controlados pela interface. As falhas foram injetadas exclusivamente por auxiliares de teste e pontos de interrupção.

Cada cenário recebeu uma cópia independente da base compilada em TEMP, fora do OneDrive. [Fonte da instrumentação](./_auditoria/fase5_20261101_202114Z/scripts/modTestesIntegradosFase5.bas). Dados comuns: veículo FICTICIO A, código F5A, placa FIC5A01, entrada matriz 11001/filial 21001, saída 22001, entrada 900, saída 1000, data 01/11/2026. Foram semeados também veículos B/C para detectar deslocamento. Chave sintética: texto de 44 dígitos iniciado por zero.

Variantes: DUP insere a NF de saída no destino; T_DUP_DESTINO semeia 22001 em Entradas Filial A4, com a mesma placa A; VAZIO remove o número; VALOR usa ABC na entrada; DATA usa 31/02/2026; SEMORIGEM usa entrada 99999; XML_DIVERGE usa valor 999 no XML contra 1000 no formulário. V_VAZIO_XML usa ide/nNF vazio também no XML. FALHA_INTER troca o valor de entrada por ABC depois de destinos já gravados; FALHA_BAIXA protege a origem somente após a pré-validação; FALHA_SAVE tentou indisponibilizar a pasta descartável, mas o Windows recusou.

Cada par antes/depois contém **7.560 células: nove abas, A1:U40**. Compara tipo, valor, fórmula, NumberFormat, Ignore, Locked e proteção da planilha. Não é comparação de todo o XLSM nem de fontes, bordas e cores das células. [Resultados estruturados com diferenças por célula](./_auditoria/fase5_20261101_202114Z/evidencias/matriz-cenarios-executados.json). Os TXT registram o status e os estados de Application produzidos pelo VBA. Mensagens e pausas de depuração são observações transcritas da interface, identificadas como tal.

## Eventos integrados reais

São 26 cenários com estado anterior/posterior: **20 PASS, 5 FAIL e 1 BLOQUEADO para injeção no Save**. Dentro dos 20 PASS: 15 rejeições sem escrita, quatro compensações limitadas e um Cancelado válido. A segunda tentativa de D_FALHA_INTER é repetição controlada do mesmo cenário, não um teste adicional.

| Cenário / fixture e evidência | Resultado | Observado / limite |
|---|---|---|
| [C_DATA](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/C_DATA/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [C_DUP](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/C_DUP/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [C_FALHA_INTER](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/C_FALHA_INTER/) | PASS | Erro 13 injetado durante historico; compensacao com igualdade exata das 7560 celulas monitoradas. |
| [C_FALHA_SAVE](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/C_FALHA_SAVE/) | BLOQUEADO | Injecao de indisponibilidade do diretorio recusada pelo Windows EPERM. Save real concluiu normalmente; falha no Save NAO foi executada. |
| [C_VALIDO](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/C_VALIDO/) | PASS | Dois historicos de saida gravados; entrada preservada; sem XML. Verificacao do comportamento legado, pendencia fiscal D03. |
| [C_VALOR](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/C_VALOR/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [C_VAZIO](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/C_VAZIO/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [D_DATA](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/D_DATA/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [D_DUP](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/D_DUP/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [D_FALHA_BAIXA](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/D_FALHA_BAIXA/) | FAIL | Protecao aplicada somente na fixture depois da prevalidacao; Delete falha 1004. Quatro destinos persistem e origem permanece; conciliacao obrigatoria, sem nova tentativa. |
| [D_FALHA_INTER](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/D_FALHA_INTER/) | PASS | Erro 13 injetado durante historico; compensacao com igualdade exata das 7560 celulas monitoradas. |
| [D_SEMORIGEM](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/D_SEMORIGEM/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [D_VALIDO](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/D_VALIDO/) | FAIL | Quatro destinos gravados. Delete A:M deixa N4=#REF!, O4 do veiculo A junto do B; Demonstracao Q5/R5 quebradas e S5 passa a B. |
| [D_VALOR](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/D_VALOR/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [D_VAZIO](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/D_VAZIO/) | PASS | Saída antecipada sem mensagem; status inicial permanece. Número vazio não gerou escrita; 7.560 células iguais. |
| [T_DUP_DESTINO](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/T_DUP_DESTINO/) | FAIL | Filial aceita NF 22001 em A4 e A7, mesma placa; passagem/identidade ambigua. Baixa matriz tambem quebra formulas. Nao define regra fiscal de unicidade. |
| [T_FALHA_INTER](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/T_FALHA_INTER/) | PASS | Erro 13 injetado durante historico; compensacao com igualdade exata das 7560 celulas monitoradas. |
| [T_VALIDO](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/T_VALIDO/) | FAIL | Quatro destinos gravados e chave L7 String/44/@/Ignore preservada; formulas M7/N7 preservadas. Baixa da matriz quebra Q4/R4 e reatribui S4 ao veiculo B. |
| [V_DATA](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/V_DATA/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [V_DUP](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/V_DUP/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [V_FALHA_INTER](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/V_FALHA_INTER/) | PASS | Erro 13 injetado durante historico; compensacao com igualdade exata das 7560 celulas monitoradas. |
| [V_SEMORIGEM](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/V_SEMORIGEM/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [V_VALIDO](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/V_VALIDO/) | FAIL | Destinos gravados; Delete A:N gera #REF! em Demonstracao Q4/R4 e referencia indireta S4 passa ao veiculo B; coluna O permanece do veiculo A. |
| [V_VALOR](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/V_VALOR/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |
| [V_VAZIO_XML](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/V_VAZIO_XML/) | PASS | CarregarXMLNFe rejeitou ide/nNF ausente ou inválido; XML não importado e nenhuma escrita. Não prova uma validação independente de número vazio após carregar XML válido. |
| [V_XML_DIVERGE](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/V_XML_DIVERGE/) | PASS | Rejeicao confirmada; igualdade exata das 7560 celulas monitoradas. |

Em todos os 19 casos de rejeição/compensação acima, as 7.560 células monitoradas coincidiram exatamente com a referência anterior. Em D_FALHA_INTER, isso foi verificado antes de autorizar a única repetição: a falha foi reinjetada e a compensação tornou a coincidir, sem duplicidade. Nenhuma repetição foi feita depois de baixa com estado inconsistente/incerto. O número 891 de diferenças em D_FALHA_BAIXA inclui 840 indicadores de proteção; não representa 891 células de dados gravadas.

## XML e relatório reais

Fixture [XML_SUITE](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/XML_SUITE/), procedimento F5_XML → modValidacaoXML.ValidarXMLNFe; modalidade **execução funcional VBA**, com MSXML6 e frmRelatorioXML reais. O esperado é True para válido/múltiplos itens/tributos e False nos demais. Todos os 13 retornos corresponderam. [Retornos emitidos pelo VBA](./_auditoria/fase5_20261101_202114Z/evidencias/xml-resultados-vba.json).

| XML sintético | Esperado / observado | Resultado | Interface observada |
|---|---|---|---|
| valido | True / True | PASS | XML VALIDADO |
| malformado | False / False | PASS | MSXML ParseError -1072896685; DOM incompleto rejeitado por ausência de um infNFe |
| sem_inf | False / False | PASS | Exatamente uma NF-e exigida |
| dois_inf | False / False | PASS | Exatamente uma NF-e exigida |
| chave_invalida | False / False | PASS | Id exige NFe + 44 dígitos |
| numero_invalido | False / False | PASS | ide/nNF ausente ou inválido |
| vazio | False / False | PASS | Relatório: três erros |
| numerico_invalido | False / False | PASS | Relatório: três erros; exportação real preservada |
| overflow | False / False | PASS | Relatório: três erros |
| multiplos_itens | True / True | PASS | XML VALIDADO |
| soma_divergente | False / False | PASS | Relatório: um erro |
| tributos | True / True | PASS | Soma de dois itens; base 100, ICMS 12, PIS 1,65, COFINS 7,60 |
| tributos_divergentes | False / False | PASS | ICMS total 11 em vez de 12; um erro |

O wrapper carrega o DOM e chama o validador diretamente. O caso malformado não comprova a mensagem do carregador de arquivos da aplicação para XML malformado. As mensagens da interface foram observadas e transcritas; os retornos/ParseError vieram dos arquivos da suíte. O [relatório exportado](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/XML_SUITE/Relatorio_XML_numerico_invalido.xlsx.txt) é texto Windows-1252, apesar do nome .xlsx.txt. Não é uma pasta Excel. O seletor oferece extensão de planilha e o gravador acrescenta .txt; melhoria de apresentação pendente. Nenhum XML tem autorização ou assinatura fiscal comprovada; cStat/tpAmb sintéticos exercitam apenas as regras existentes.

## A09: regressão VBA reutilizada

| Procedimento / evidência | Resultado | Escopo efetivamente observado |
|---|---|---|
| [AuditoriaFase2_Executar](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/A09_SUITE/A09_regressao.txt) | PASS — 377 asserções | Auxiliares, snapshots, valores, fórmulas, formatos numéricos, Ignore, vazios, estados Application, fronteiras Delete/Save e formatação condicional; não equivale aos eventos de negócio |
| [AuditoriaFase4_IntegridadeComplementar](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/A09_SUITE/A09_complementar.txt) | PASS — 11 asserções | Sobreposição/ordem inversa, falha individual reconhecida, continuidade dos outros snapshots, chave String 44/@/Ignore e segunda recuperação após desproteger |
| [Classe com duas pastas](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/A09_2WB/A09_duas_pastas.txt) | PASS | Pasta1 ativa permanece intocada; alvo da fixture recupera ANTES e =7+8; Application True/True/True/-4105 |
| AuditoriaFase3_CicloVida | BLOQUEADO | Depurador interrompe na exceção esperada “Operacao ja iniciada”, -2147219404; continuação repetiu a parada. Teste encerrado, configuração preservada, sem PASS |
| Recuperação incompleta da Fase 4 | FAIL conhecido, mantido | O PASS complementar reconhece falha individual; não transforma recuperação incompleta em completa |

A primeira tentativa de preparar o teste com duas pastas encontrou erro 424 por contexto da janela Imediata. A preparação foi corrigida para referências explícitas por nome, sem alteração de código de aplicação. A tentativa concluída gerou o TXT acima. O TXT do ciclo de vida contém somente BOM, sem resultado; não é evidência de aprovação.

## Cobertura ainda pendente

| Escopo | Situação nesta Fase 5 | Motivo / próximo critério |
|---|---|---|
| Falha antes da primeira escrita dos eventos de cada fluxo | NÃO EXECUTADO como falha injetada | Houve rejeições antes de escrita; elas não substituem falha artificial em cada etapa |
| Falha em cada destino/coluna, antes da baixa em Venda/Transferência e após cada estágio | NÃO EXECUTADO em todas as posições | Uma falha intermediária por fluxo e tentativa de Delete na Devolução executadas; ampliar após decisão de integridade |
| Falha real de Save | BLOQUEADO | EPERM na preparação; Save terminou normalmente. Não houve alteração de permissões/segurança |
| Falha após Save, encerramento inesperado | NÃO EXECUTADO | Sem mecanismo comprovado de reconciliação/idempotência; nenhum encerramento forçado |
| Alterar, ExcluirFilial, edição anual e exclusões duplas ambíguas | BLOQUEADO para correção por D01/D02; NÃO EXECUTADO funcionalmente | Não há vínculo aprovado para selecionar registro histórico |
| Operações de negócio com outra pasta ativa | NÃO EXECUTADO | Somente a classe foi exercitada com duas pastas |
| Entrada matriz/filial, pesquisa/ListBox, backup versionado, Ribbon | NÃO EXECUTADO novamente | Referência da Fase 4 preservada; fontes sem alteração semântica. Não renovar PASS antigo como teste desta fase |
| Devolução com XML divergente / Cancelado com comprovação fiscal | NÃO EXECUTADO como validação fiscal | Ramo atual da Devolução/Cancelado não ganhou regra nova; D03 pendente |
| Selenium, DETRAN, consultas externas e dados reais | NÃO EXECUTADO | Fora do escopo autorizado destes testes |

Os testes falhos de integridade permanecem falhos mesmo que representem fielmente a lógica legada. Não promover qualquer uma das fixtures de cenário.
