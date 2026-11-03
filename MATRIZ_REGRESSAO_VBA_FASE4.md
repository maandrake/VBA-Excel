# Matriz de regressão VBA — Fase 4

Atualização 01/11/2026. Fixture fictícia integrada e compilada no Excel 16.0 x64. **Regressão parcial, sistema NÃO HOMOLOGADO.** [Relatório](./RELATORIO_INTEGRACAO_FASE4.md), [inventário](./_auditoria/fase4_20261031_230457Z/inventario-integracao.json).

PASS = execução aprovada; FAIL = execução reprovada; BLOQUEADO = impedimento técnico/permissão/decisão; NÃO EXECUTADO = sem execução. Um teste do auxiliar real não prova os eventos gravadores. São 397 verificações das três suítes concluídas, não 397 cenários de estoque. Repetições não aumentam a contagem. Nenhuma simulação JS/.NET ou teste COM equivalente novo substituiu VBA.

## Integração, compilação e suítes

| ID | Objeto/cenário | Modalidade | Estado | Observação/evidência |
|---|---|---|---|---|
| I01 | 37 importados + 18 documentos, 17 CodeNames | Inspeção do VBE e extração local do VBA salvo | PASS | 55 reais; 37 corpos e 18 documentos correspondem ao pacote após normalização documentada; [fontes](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/fontes-integradas-entrega.json) |
| I02 | Depurar > Compilar VBAProject após correções | Compilação real VBA | PASS | Comando executado e depois desabilitado; erro C4-01 corrigido |
| I03 | VBProject programático | Sondagem COM exclusiva | BLOQUEADO | 0x800A03EC; nenhuma importação COM; sem mudar segurança |
| S01 | Auditoria_ExecutarTestesSeguros | Execução funcional VBA com fake de lista | PASS | Nove verificações; pesquisa não equivale sozinha ao MSForms |
| S02 | AuditoriaFase2_Executar | Execução funcional VBA; classe/helper e planilhas reais fictícias | PASS | 377; repetida com outra pasta ativa e após C4-04; não chama eventos gravadores |
| S03 | AuditoriaFase3_CicloVida | Execução funcional VBA iniciada | BLOQUEADO | Depurador em Interromper em módulo de classe para no Err.Raise esperado +2100; sem retorno PASS; preferência não alterada |
| S04 | AuditoriaFase4_IntegridadeComplementar | Execução funcional VBA | PASS | 11; [saída e escopo](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/testes-vba-complementares.json) |

## A09 — compensação anterior a Delete/Save

| ID | Cenário | Modalidade | Estado | Evidência/limite |
|---|---|---|---|---|
| T01 | Captura/restauração de valores, fórmula, texto literal e chave 44 | Execução funcional VBA S02/S04 | PASS | String, NumberFormat @, zeros e Ignore conferidos |
| T02 | Snapshots sobrepostos restaurados em ordem inversa | Execução funcional VBA S04 | PASS | ANTES → INTERMEDIARIO → DEPOIS; recuperação final ANTES |
| T03 | Falhas artificiais após escritas em cada etapa dos seis fluxos do auxiliar | Execução funcional VBA S02 | PASS | Snapshots das células sentinela recompostos; não é chamada dos gravadores privados |
| T04 | Um destino protegido em outra aba | Execução funcional VBA S04 | PASS | Recuperar retorna False, diagnóstico uma falha; outros destinos restaurados |
| T05 | Nova tentativa do mesmo snapshot após liberar proteção | Execução funcional VBA S04 | PASS | Recuperação antes de Delete/Save, número/formato/fórmula/chave/Ignore restaurados; não prova idempotência de uma operação comercial |
| T06 | Proteção ampla da aba durante recuperação | Execução funcional VBA, antes/depois C4-04 | **FAIL** | Quatro falhas; antes três células apagadas, depois conteúdo alterado preservado; recuperação integral continua incompleta |
| T07 | Application e outra pasta ativa | Execução funcional VBA S02 + inspeção final | PASS | Propriedades True/True/True/-4105 restauradas; auxiliares operam na proprietária |
| T08 | Recusa de compensação após marcar baixa/Save | Execução funcional VBA S02 | PASS | Mecanismo retorna recusa; não executa nem recupera Delete/Save real |
| T09 | Reinicialização e ciclo de vida completo | Execução funcional VBA S03 | BLOQUEADO | Erro esperado interceptado pelo depurador; não ampliar garantia |
| T10 | Valor/data inválidos antes da primeira escrita matriz | Execução funcional VBA pelo botão real | PASS | Erros 13 tratados; quatro contagens zero e células das duas abas sem diferenças; outras falhas anteriores à escrita não testadas |
| T11 | Falha real em Delete, Save, encerramento ou retorno após crash | Execução funcional VBA | NÃO EXECUTADO | Não há mecanismo comprovado de recuperação estrutural/persistente |

A primeira versão do teste complementar supunha falha em só uma célula ao proteger toda a aba; NumberFormat/Ignore também falharam em outros destinos. Esse FAIL motivou C4-04 e foi preservado. O teste final separa o destino protegido em outra aba para medir falha individual sem ocultar o FAIL amplo. [Preparação](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/preparacao-falha-individual.json), [antes/depois](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/C4-04.json), [regressão após correção](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/regressao-pos-C4-04.json).

## Interface e pesquisa

| ID | Cenário | Modalidade | Estado | Resultado |
|---|---|---|---|---|
| U01 | Carregar os cinco UserForms | Execução funcional VBA/MSForms | PASS | Controles top-level: Cad_NFE 108, Cad_NFE_Saida 130, Cad_NCM 14, frmRelatorioXML 5, UserFormFiltro 21; não certifica layout completo |
| U02 | Pesquisa matriz com dois resultados e seleção | Execução funcional VBA/MSForms | PASS | PesquisarEntradasRapido encontra dois; ListIndex 0 aciona evento e preenche NF TESTE_1 |
| U03 | ListIndex = -1 | Execução funcional VBA/MSForms | PASS | Sem erro; NF TESTE_1 preservada |
| U04 | Pesquisa sem resultado | Execução funcional VBA/MSForms | PASS | Zero resultados |
| U05 | Pesquisa filial real, navegação completa, todos controles e filtros legados | Execução funcional VBA | NÃO EXECUTADO | S01/fake não substitui esses casos |

[Saídas observadas](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/userforms-listbox-vba.json).

## XML e apresentação

| ID | Cenário | Modalidade | Estado | Resultado |
|---|---|---|---|---|
| X01 | Sem infNFe | Execução funcional VBA / ValidarXMLNFe | PASS | False; mensagem exige exatamente uma NF-e |
| X02 | Duas infNFe | Execução funcional VBA / ValidarXMLNFe | PASS | False; mesma mensagem de unicidade |
| X03 | Id com chave curta | Execução funcional VBA / ValidarXMLNFe | PASS | False; mensagem NFe seguido de 44 dígitos |
| X04 | nNF ABC, Id estrutural com 44 dígitos | Execução funcional VBA / ValidarXMLNFe | PASS | False; mensagem número inválido |
| X05 | Texto malformado <raiz> | MSXML6 executado pelo VBA | PASS | LoadXML=False; parseError=-1072896685; sem fluxo completo do importador |
| X06 | XML completo válido, vários itens, somas fiscais, inválidos numéricos e overflow | Execução funcional VBA | NÃO EXECUTADO | Resultados COM anteriores continuam equivalentes; não foram promovidos a PASS VBA |
| X07 | Exibição/população real de frmRelatorioXML | Execução funcional VBA/UI | NÃO EXECUTADO | U01 comprova carregamento, não apresentação do relatório |

[Entradas, mensagens e saídas](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/xml-vba.json). Nenhuma validação de autorização fiscal/assinatura/autenticidade foi feita; regras tributárias preservadas.

## Backup e Ribbon

| ID | Cenário | Modalidade | Estado | Resultado |
|---|---|---|---|---|
| B01 | CriarBackupVersionado duas vezes com outra pasta ativa | Execução funcional VBA por Application.Run | PASS | Dois diretórios/arquivos distintos; sentinela de memória da proprietária incluída |
| B02 | Ausência de salvamento inesperado da origem | Macro VBA + comparação SHA-256 em disco | PASS | Hash da origem igual antes/depois das cópias |
| B03 | Abrir os dois backups e verificar conteúdo | Operação COM somente leitura | PASS | 17 abas, sentinela, fórmula =1+2, String 44, formato @ e Ignore; não é nova execução VBA |
| B04 | Restaurar versão anterior em fluxo operacional | Execução funcional VBA | NÃO EXECUTADO | Reabertura B03 prova legibilidade/conteúdo, não conciliação de movimentações |
| B05 | SaveWorkbookBackup / SaveWorkbookBackupToFloppy | Execução funcional VBA | NÃO EXECUTADO | Wrappers/destinos legados não chamados |
| B06 | Destino inexistente, falha de gravação e colisão simultânea | Execução funcional VBA | NÃO EXECUTADO | Sem prova desses ramos; nenhuma retenção automática implementada |
| R01 | ChamarSalvar com outra pasta ativa alterada | Execução funcional VBA por Application.Run | PASS | Hash da fixture mudou ao salvar sentinela; hash da outra ficou igual, alteração dela descartada só ao fechar a pasta fictícia |
| R02 | Clique no Ribbon e binding customUI | Interface VBA/Ribbon | NÃO EXECUTADO | Teste R01 chama o callback diretamente com Nothing |

[Backups](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/backups-vba.json), [reabertura](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/reabertura-backups-com.json), [Ribbon antes](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/ribbon-antes.json), [depois](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/ribbon-depois.json).

## Regressão dos fluxos de negócio e decisões

| Fluxo | Casos requeridos | Estado atual da macro completa | Impedimento/continuação |
|---|---|---|---|
| Entrada matriz/filial | Válida, duplicada, ausentes, datas/valores inválidos, chave, fórmulas, anual | PASS nos seis casos abaixo; demais NÃO EXECUTADOS | Eventos reais acionados na interface; não comprova injeção entre escritas ou todos os campos obrigatórios |
| Venda | Válida, duplicada, origem, histórico e falhas entre escritas | NÃO EXECUTADO | Teste integrado do gravador ainda necessário; baixa depende D01/D02 |
| Transferência | Saidas, Entradas Filial, anual filial, históricos, fórmula M e vínculo | NÃO EXECUTADO | Sem aprovação da identidade/baixa; nenhum Delete alterado |
| Devolução | Origem existente/ausente, quatro abas, fórmula e baixa | NÃO EXECUTADO | D01/D02 e configuração fictícia específica |
| Cancelamento | Dois históricos, entrada preservada e evidência fiscal | NÃO EXECUTADO | Regra atual preservada; D03 continua aberta |
| Alterar/ExcluirFilial/históricos | Duplicidades e vínculo inequívoco | BLOQUEADO | D01/D02; não usar primeiro Find/registro ambíguo |
| A08 / sete Delete | Política de baixa aprovada e migração | BLOQUEADO | Contrato no plano de identidade é proposta, não implementado |
| Logs A20 | Destino, privacidade e retenção | BLOQUEADO | D04; sem consultas externas ou logs reais |

Para executar os eventos privados futuramente, use a interface dos UserForms em fixture nova com dados fictícios e identidade aprovada, ou instrumentação restrita à cópia de teste, sem renomear eventos de produção. Registre antes/depois de todas as abas e pontos de falha. Os PASS das suítes acima não autorizam promover o sistema.


### Verificação dirigida adicional de C4-04

**PASS, execução funcional VBA:** célula Saidas!A30 inicialmente Empty, snapshot, escrita TEMP_FICTICIO e recuperação retornando True com IsEmpty=True. A tentativa inicial pela Janela Imediata teve erro 13 por argumento Variant implícito para ByRef String; repetida com expressão String explícita, concluída e Finalizar executado. Estado final Application True/True/True/-4105. Não altera a contagem 397 das suítes. [Evidência](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/celula-vazia-vba.json).


## Eventos reais dos gravadores de entrada — complemento final

Modalidade de todos os seis casos: **execução funcional VBA pela interface**, com controles reais de Cad_NFE e clique no botão Salvar. Nenhum evento foi renomeado e nenhum fonte de negócio foi alterado para execução.

| Caso | Entrada fictícia | Estado | Resultado observado |
|---|---|---|---|
| ENT-M-VALIDA | 900001, Compra, dados de veículo fictício, 1234,56 e 01/11/2026 | PASS | Um registro por aba; valor/data corretos e A:J igual ao anual |
| ENT-M-DUPLICADA | Repetir 900001 | PASS | Mensagem já cadastrado; continua um registro por aba |
| ENT-M-VALOR | 900002, valor ABC | PASS | Erro 13 tratado; zero registros nos dois destinos |
| ENT-M-DATA | 900003, DATA_INVALIDA e valor 100,00 | PASS | Erro 13 tratado; zero registros nos dois destinos |
| ENT-M-NUMERO-VAZIO | Número vazio | PASS | Retorno sem gravação; A5 vazia em ambos os destinos |
| ENT-F-VALIDA | 900101, Compra, veículo filial fictício e chave sintética 44 | PASS | Um registro por aba, A:J igual ao anual, L4 String exata 44/@, M4/N4 e referência anual preservadas |

[Dados/saídas](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/entradas-vba-real.json), [comparação de todas as células das duas abas matriz](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/comparacao-invalidos-matriz.json). A comparação resolve sharedStrings e confere valor/tipo/fórmula/índice de estilo; sem diferenças entre checkpoint após sucesso e estado final após recusas. Não compara todos os objetos/estilos globais. A chave não é persistida no anual pelo código atual; não se adicionou campo para mudar isso. Os seis casos não entram na contagem 397 das suítes.

Continuam NÃO EXECUTADOS: filial duplicada/inválida, todos os outros campos obrigatórios, falhas entre gravações dos botões e Save falhando. A baixa e as ligações ambíguas continuam dependentes de D01/D02.
