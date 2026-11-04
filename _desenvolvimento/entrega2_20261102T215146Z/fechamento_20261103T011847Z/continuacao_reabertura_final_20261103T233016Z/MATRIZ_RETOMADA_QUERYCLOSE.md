# Regressao dirigida — reabertura e QueryClose

Execucao iniciada em03/11, retomada em04/11/2026. Preserva a matriz anterior; nao soma assercoes a fluxos completos.

## R26-COLD-BASE — PASS

- Modalidade: Execucao VBA real apos habilitacao manual.
- Procedimento: C1_UltimoResultado e C1_Diagnostico.
- Esperado: Preservar identidade, fatos e metadados no escopo descrito.
- Observado: CACHE_AUDITADO automatico; AUDIT=OK;PROJECTION=VERIFIED;REV=26;MOV=26;RECEIPTS=26;ACTIVE=7;PENDING=0.
- Origem: Transcricao manual da interface nativa observada durante esta execucao; nao e log automatico da suite.
- Evidencia: [RELATORIO_RETOMADA_QUERYCLOSE.md](./RELATORIO_RETOMADA_QUERYCLOSE.md).

## R26-PESQ-MATRIZ — PASS

- Modalidade: Eventos reais de TextBoxPesquisa/ListBoxLista.
- Procedimento: Cad_NFE pesquisa F2A0001 e selecao real.
- Esperado: Preservar identidade, fatos e metadados no escopo descrito.
- Observado: NF12003;150,3333;03/11/2026;EpisodeID68414E81-9543-414B-810A-F16CAFA4F450; nao selecionou episodio encerrado NF12001.
- Origem: Transcricao manual da interface nativa observada durante esta execucao; nao e log automatico da suite.
- Evidencia: [vinculos_pesquisa_interface.json](./vinculos_pesquisa_interface.json).

## R26-PESQ-FILIAL — PASS

- Modalidade: Eventos reais de pesquisa e ListBox filial.
- Procedimento: Cad_NFE pesquisa F2F0002 e selecao real.
- Esperado: Preservar identidade, fatos e metadados no escopo descrito.
- Observado: NF12002;456,7891;AMARELO QA;EpisodeIDF42BDB7A-F143-4BBB-9148-AE47568DC96D; chave44/String conferida no banco.
- Origem: Transcricao manual da interface nativa observada durante esta execucao; nao e log automatico da suite.
- Evidencia: [vinculos_pesquisa_interface.json](./vinculos_pesquisa_interface.json).

## R26-QC-ORIGINAL — FAIL

- Modalidade: Evento VBA real por X.
- Procedimento: Cad_NFE.UserForm_QueryClose.
- Esperado: Fechar sem erro ou escrita arbitraria.
- Observado: Erro1004 ao incrementar U1 em Entradas Filial protegida. FAIL anterior preservado.
- Origem: Transcricao manual da interface nativa observada durante esta execucao; nao e log automatico da suite.
- Evidencia: [RELATORIO_RETOMADA_QUERYCLOSE.md](./RELATORIO_RETOMADA_QUERYCLOSE.md).

## R26-QC-COMPILE — PASS

- Modalidade: Compilacao VBA real.
- Procedimento: Depurar > Compilar VBAProject.
- Esperado: Preservar identidade, fatos e metadados no escopo descrito.
- Observado: Compilacao sem mensagem de erro; item Compilar desabilitado apos execucao.
- Origem: Transcricao manual da interface nativa observada durante esta execucao; nao e log automatico da suite.
- Evidencia: [RELATORIO_RETOMADA_QUERYCLOSE.md](./RELATORIO_RETOMADA_QUERYCLOSE.md).

## R26-QC-ENTRADA — PASS

- Modalidade: Eventos VBA reais por X e botao Fechar.
- Procedimento: Cad_NFE.QueryClose / CommandButtonFechar_Click.
- Esperado: Preservar identidade, fatos e metadados no escopo descrito.
- Observado: X mostra orientacao informativa; botao oculta formulario; Visible=False; U1 vazia e protecao preservadas.
- Origem: Transcricao manual da interface nativa observada durante esta execucao; nao e log automatico da suite.
- Evidencia: [RELATORIO_RETOMADA_QUERYCLOSE.md](./RELATORIO_RETOMADA_QUERYCLOSE.md).

## R26-QC-SAIDA — PASS

- Modalidade: Eventos VBA reais por X e botao Fechar.
- Procedimento: Cad_NFE_Saida.QueryClose / CommandButtonFechar_Click.
- Esperado: Preservar identidade, fatos e metadados no escopo descrito.
- Observado: X mostra orientacao; botao oculta; Visible=False; U1 vazia; EnableEvents/ScreenUpdating/DisplayAlerts=True; Calculation=-4105.
- Origem: Transcricao manual da interface nativa observada durante esta execucao; nao e log automatico da suite.
- Evidencia: [RELATORIO_RETOMADA_QUERYCLOSE.md](./RELATORIO_RETOMADA_QUERYCLOSE.md).

## R26-FORMULA-CACHE — FAIL

- Modalidade: Diagnostico VBA real.
- Procedimento: C1_Diagnostico apos edicao/fechamento.
- Esperado: Projecao verificada.
- Observado: AUDIT=OK mas PROJECTION=FORMULA_VALUE. Nao aprovado como projecao integra nessa tentativa; causa fina ainda nao isolada.
- Origem: Transcricao manual da interface nativa observada durante esta execucao; nao e log automatico da suite.
- Evidencia: [RELATORIO_RETOMADA_QUERYCLOSE.md](./RELATORIO_RETOMADA_QUERYCLOSE.md).

## R26-FORMULA-RECALC — PASS

- Modalidade: VBA real, recalc e diagnostico.
- Procedimento: Application.CalculateFull; C1_Diagnostico.
- Esperado: Preservar identidade, fatos e metadados no escopo descrito.
- Observado: PROJECTION=VERIFIED; REV26/MOV26/RECEIPTS26/ACTIVE7/PENDING0. Recalculo recuperou resultados; nao alterou formulas nem fatos ACE.
- Origem: Transcricao manual da interface nativa observada durante esta execucao; nao e log automatico da suite.
- Evidencia: [RELATORIO_RETOMADA_QUERYCLOSE.md](./RELATORIO_RETOMADA_QUERYCLOSE.md).

## R26-QC-SAVE — PASS

- Modalidade: Salvamento real no Excel.
- Procedimento: Salvar candidata; ?ThisWorkbook.Saved,Workbooks.Count.
- Esperado: Preservar identidade, fatos e metadados no escopo descrito.
- Observado: Saved=True; Workbooks.Count=1; hash XLSM efd1916395159da3d30ecc99838ff6e0e44dd9e0ca2b13b29570a3b59fe247e0.
- Origem: Transcricao manual da interface nativa observada durante esta execucao; nao e log automatico da suite.
- Evidencia: [RELATORIO_RETOMADA_QUERYCLOSE.md](./RELATORIO_RETOMADA_QUERYCLOSE.md).

## R26-QC-BANCO — PASS

- Modalidade: Inspecao ADO ACE somente leitura e SHA256.
- Procedimento: Capturar-Estado-Sintetico.ps1 e comparacao de17tabelas.
- Esperado: Preservar identidade, fatos e metadados no escopo descrito.
- Observado: Zero tabelas divergentes; hash ACCDB9fc0bfe8...; sem nova operacao, recibo, pending ou lease.
- Origem: Resultado estruturado local.
- Evidencia: [comparacao_banco_queryclose.json](./comparacao_banco_queryclose.json).

## R26-QC-COLD-FINAL — BLOQUEADO

- Modalidade: Tentativa real em nova instancia Excel.
- Procedimento: Encerrar normalmente instancia exclusiva; abrir caminho exato.
- Esperado: Cache automatico e diagnostico em sessao fria.
- Observado: Arquivo abriu; barra As macros foram desabilitadas. Inicializacao VBA final nao exercitada; sem alterar seguranca.
- Origem: Transcricao manual da interface nativa observada durante esta execucao; nao e log automatico da suite.
- Evidencia: [RELATORIO_RETOMADA_QUERYCLOSE.md](./RELATORIO_RETOMADA_QUERYCLOSE.md).

## R26-QC-FONTES — PASS

- Modalidade: Analise estatica do binario salvo.
- Procedimento: inspectEmbedded; comparacao de72componentes e assinaturas.
- Esperado: Preservar identidade, fatos e metadados no escopo descrito.
- Observado: Somente Cad_NFE/Cad_NFE_Saida com mudanca de corpo; patch esperado confere;436assinaturas dos5formularios iguais.127streams de designer presentes;43bytes-streams diferem, sem alegar igualdade binaria.
- Origem: Resultado estruturado local.
- Evidencia: [comparacao_fontes_salvos.json](./comparacao_fontes_salvos.json).

## R26-ORIGINAIS — PASS

- Modalidade: SHA256 dos arquivos originais.
- Procedimento: Conferencia39arquivos contra manifesto existente.
- Esperado: Preservar identidade, fatos e metadados no escopo descrito.
- Observado: 39 correspondencias, zero divergencias.
- Origem: Resultado estruturado local.
- Evidencia: [integridade_originais_reconferida.json](./integridade_originais_reconferida.json).
