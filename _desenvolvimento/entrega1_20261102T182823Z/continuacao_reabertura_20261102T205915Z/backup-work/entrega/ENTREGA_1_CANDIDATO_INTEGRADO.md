# Entrega 1 — candidata integrada sintética

Data: 02/11/2026. Execução: 20261102T182823Z. **Uso exclusivo em laboratório. NÃO HOMOLOGADA.**

## Software produzido

- [Estoque_Fera_Veiculos_Candidato_Integrado.xlsm](./Estoque_Fera_Veiculos_Candidato_Integrado.xlsm): 69 componentes VBA, 17 planilhas e cinco UserForms herdados. Compilação real pelo Editor VBA concluída; [evidência do menu após compilar](../evidencias/compilacao-final-menu.txt).
- [Banco_Candidato_Sintetico.accdb](./Banco_Candidato_Sintetico.accdb): ACE 16, esquema 5, fonte autoritativa experimental. BaseID 69C72E8F-65EF-4904-8433-A1B2006CC968; revisão 8; 3 veículos, 4 episódios, 8 movimentos, 8 recibos, 22 registros em SheetProjections e zero Pending. WriterLease FREE e ProjectionState VERIFIED, observados por [consulta ADO direta](../evidencias/ace-final-revisao8.json).
- [Fontes extraídos da candidata salva](../fontes/) e [manifesto dos artefatos](./manifesto_artefatos.json).

A cópia executada está no TEMP: C:\Users\MARCOS~1\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\Entrega1_20261102T182823Z. Os arquivos do pacote no workspace são preservação; não executar Office no OneDrive. O banco verifica o caminho da autoridade e uma cópia deslocada não se elege autoridade automaticamente. Consulte o [manual](./MANUAL_INSTALACAO_RESTAURACAO.md).

## Resultado funcional efetivamente observado

O primeiro fluxo vertical foi executado por controles reais: entrada 11001/FIC5A01/900 → pesquisa → seleção por token → botão de copiar → venda 22001/1000 com XML sintético → baixa lógica no ACE → históricos → reconstrução e verificação de arquivo de projeção reaberto. E1_ENTRADA_EVENTO_R3 e E1_V_VALIDO_R3 passaram. A origem permanece preservada e inelegível; nenhum deslocamento físico foi usado.

Também passaram pelos eventos reais: entrada B por 700; transferência por 800 para filial; devolução de 800 à matriz; segunda transferência do mesmo episódio após correção da ordem de preenchimento do formulário; recusa de destino duplicado. A hipótese experimental mantém o episódio na transferência/devolução e muda a custódia. Ela ainda exige validação operacional.

Reentrada: o teste forneceu explicitamente o VehicleID do episódio encerrado e acionou o botão de cadastro. Foi criado outro EpisodeID e preservado o anterior CLOSED; 650,25 foi conservado. A seleção desse VehicleID foi instrumentação da suíte, e não uma interface definitiva de conciliação. A primeira confirmação da projeção falhou por chave opcional; o defeito foi corrigido e a projeção foi reconstruída posteriormente. O PASS original dessa asserção comprova identidade, não sucesso integral daquela primeira execução.

## Integração realizada

1. ACE é a única autoridade dos fluxos integrados. Excel guarda projeções verificáveis; não há transação conjunta entre commit ACE e Save do XLSM.
2. VehicleID/EpisodeID/MovementID/ProjectionID/OperationID/AttemptID, versões, lease/fencing, intenção, recibos e auditoria do laboratório F10 foram incorporados. SheetProjections liga os oito destinos pelos IDs explícitos.
3. Pesquisa ativa, seleção, preenchimento, envio para saída e navegação dos fluxos integrados resolvem identidade novamente. Versão obsoleta e identidade ambígua são recusadas.
4. Fórmulas auxiliares M/N/O da filial, meses/prazos e indicação anual de saída usam ViewID/EpisodeID; não buscam a primeira placa coincidente.
5. Os sete Delete dos dois formulários foram retirados exclusivamente da candidata. As movimentações usam mudança transacional de elegibilidade/custódia. Exclusão manual e alteração de históricos ainda incompatíveis foram bloqueadas, sem troca por ClearContents. O Clear da reconstrução limpa apenas a projeção descartável para rematerializá-la da autoridade.
6. Os eventos e assinaturas dos cinco formulários foram preservados por comparação estática. Os designers não foram editados deliberadamente; o Excel reserializou streams/GUIDs internos, portanto não se declara igualdade binária dos designers embutidos. Os cinco pares FRM/FRX de referência permanecem preservados no pacote de fontes; a equivalência completa de todos os controles exige regressão adicional.
7. O único procedimento alterado no validador XML foi SelecionarArquivoXML para permitir o provedor local de teste. As demais 46 rotinas mantêm seus corpos, desconsiderando a normalização de capitalização do VBE. Nenhuma regra fiscal foi alterada. Os XMLs são sintéticos sem validade fiscal.

## Correções técnicas dirigidas

- Conversão de moeda: CCur aplicado a texto canônico com ponto era interpretado pela regionalização brasileira. AmountC1 passou a validar e converter explicitamente; 900 e 650,25 foram conferidos no ACE.
- Data: formato literal brasileiro e F8_CanonData evitam confundir normalização regional com alteração material.
- Ignore: leitura em projeção protegida faz liberação/restauração controlada da proteção da própria fixture.
- Reconstrução da saída: chamada Call explícita elimina o rótulo VBA involuntário causado por dois-pontos.
- Reutilização do formulário: a operação é definida antes dos campos, pois o evento Change de Venda limpa valores. O FAIL da segunda transferência foi preservado e sua repetição passou.
- Chave opcional da matriz: VerifyBatch10 alinhado ao esquema 5 (vazia ou 44 dígitos textuais). A filial e saídas continuam exigindo chave. O diagnóstico agora distingue verificação em memória de confirmação persistente.
- Backup: lease mantido ocupado durante cópia com conexões fechadas; hashes antes/depois; auditoria do banco copiado; liberação somente do lease da cópia; manifesto. A coordenação protege somente escritores conformes, não SQL externo.

## Evidências e limites

O [log bruto](../evidencias/execucao/resultados_entrega1.tsv) preserva 44 asserções executadas: 38 PASS e 6 FAIL, incluindo revisões descartadas. R3 tem 34 execuções de asserção (33 PASS/1 FAIL), 33 IDs únicos; o FAIL da transferência sucessiva possui nova execução PASS. Não somar versões repetidas como novos cenários. A [matriz](./MATRIZ_REGRESSAO_ENTREGA.md) delimita cada modalidade e a ressalva da reentrada.

Falhas após inserção de movimento, após projeção e antes do commit reverteram fatos/recibos e mantiveram intenção e episódio NEW pendentes. Falha durante baixa da devolução reverteu os quatro destinos e a alteração da origem; a repetição foi bloqueada. Isso não encerra a conciliação. Erro injetado após commit foi tratado por reconstrução. Save cancelado por evento real preservou a autoridade e exigiu reconstrução.

A compilação está comprovada. A reabertura independente do XLSM chegou ao aviso Habilitar Macros e depende de intervenção manual; não foi marcada PASS. As gerações de projeção foram abertas e verificadas pelo VBA com eventos desabilitados, o que é uma evidência diferente de uma sessão interativa fria.

O script de preparação de cópia foi criado, mas sua execução foi BLOQUEADA pela ExecutionPolicy do ambiente. Não houve mudança nem bypass. Políticas operacionais, conciliação durável, todos os caminhos de manutenção, retorno de energia/crash e regressão integral permanecem em [defeitos abertos](./REGISTRO_DEFEITOS_ABERTOS.md).

Os 39 originais mantêm SHA-256; [inventário](../evidencias/inventario_integracao.json). Nenhuma promoção, commit, push ou operação real foi feita. O critério mínimo de Entrada/Venda foi demonstrado em execução; a entrega é incremental e não equivale à conclusão da homologação ou do fluxo D.
