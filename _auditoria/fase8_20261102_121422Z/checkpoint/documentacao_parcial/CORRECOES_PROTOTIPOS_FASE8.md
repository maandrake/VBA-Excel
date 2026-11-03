# Correções dos protótipos — Fase 8

Portão 1 exclusivamente. Os quatro FAIL F7 foram novamente executados e preservados em [reprodução](./_auditoria/fase8_20261102_121422Z/evidencias/execucoes/reproducao_f7.tsv). Os resultados seguintes referem-se ao novo modelo sintético, sem corrigir gravadores ou fórmulas do produto.

| Item | Causa raiz e reprodução | Alteração experimental mínima | Evidência nova / resultado | Limite e reversão |
|---|---|---|---|---|
| F8-01 | Selo dependia da string de NumberFormat embutido; após reabrir dd/mm/yyyy vira m/d/yyyy | Serial Value2 Double/Excel1900 em IEEE754 hexadecimal; política de apresentação explícita no selo; ISO/BR literais e legado local separados | 5 verificações na preparação + 12 na nova sessão; PASS. Mudança de formato efetivo continua registrada; formato material/texto/hora/overflow recusados | Somente esta região/Office testados. Política local não garante DMY estrito. Não substituir selos do produto. Descartar candidato/restaurar checkpoint |
| F8-02 | Anual por placa herdava venda do episódio anterior: SIM na reentrada | F8_Anual exige EpisodeID único e movimentos concluídos/IDs únicos; 0/múltiplos retorna erro | EP_A1 SIM, EP_A2 NAO; ausente/duplicado explícitos; T/D/T e reabertura preservam NAO. PASS | Não migrou históricos antigos; vínculos sintéticos declarados. Requer decisão de custódia e correspondências do legado |
| F8-03 | INDIRECT por endereço continuava válido apontando para outro veículo | Repositório por ProjectionID + Episode/Vehicle/Version; UDF por ID com faixas delimitadas, token revalidado; índice só localiza | Sort nativo, lacuna e reuso com novo PID; PID_A2 encontrado na linha12 e veículo correto; duplicado sem evento bloqueado. PASS | Endereço é pista, não autoridade. Índice não torna qualquer leitura O(1). Fórmulas físicas do principal continuam defeituosas |
| F8-04 | Auxiliar passou a usar entrada de outro veículo; igualdade numérica inicial ocultava associação | Fórmulas de margem/auxiliar resolvem os campos pelo mesmo PID/EP/VID | Valores distintos: A2 entrada1500/saída1940/margem440/aux6000; B aux6600; C aux15500 após reordenação. PASS | Não converter PASS experimental em aprovação dos auxiliares M/N/O do produto |

## Repositório experimental

[clsRepositorioFase8](./_auditoria/fase8_20261102_121422Z/prototipos/embutidos/clsRepositorioFase8.cls) mantém índice por revisão e invalidação por SheetChange; reconsulta a unicidade corrente, identidade, versão, episódio/custódia, pendências e selo antes da ação. Retorna ENCONTRADO, INEXISTENTE, DUPLICADO, BAIXADO, IDENTIDADE_AMBIGUA, VERSAO_OBSOLETA, OPERACAO_PENDENTE, AGUARDANDO_CONCILIACAO e CONTEUDO_DIVERGENTE. Não seleciona a primeira duplicidade. Se eventos forem suprimidos, a consulta revalida contagens/linha atual; isso não constitui proteção contra edição maliciosa de todos os dados e selos.

Os sete IDs/campos de controle provisórios aparecem somente nas novas abas _F8_*. IDs são literais fictícios; não há migração nem inferência de veículos reais. O registro é selado com tipos, valores, NumberFormat e Ignore; datas têm política específica. Não há selo criptográfico assinado nem autenticação do autor. Relacionamentos Movement/Operation completos e restrições de banco continuam especificação de M2/M7.

[modExperimentosFase8](./_auditoria/fase8_20261102_121422Z/prototipos/embutidos/modExperimentosFase8.bas) integra pesquisa ativa/histórica, lista/token, resolução, preenchimento dos controles reais N_Nfe/PLACA/txt_Valor de Cad_NFE, navegação por linha novamente resolvida, contagem e fórmulas. Eventos privados e leitores principais não foram alterados. Carregamento em massa do ListBox real foi medido; não representa clique/navegação de todos os eventos legados.

## Falhas encontradas na própria infraestrutura F8

1. Uma importação repetida criou clsRepositorioFase81/modExperimentosFase81 e compilação acusou nome F8_CanonRegistro repetido. Checkpoint preservado; removidas exclusivamente as duas duplicatas experimentais pelo VBE. Compilação seguinte aprovou. Não houve renomeação de procedimentos do produto.
2. Primeiro integrado: 4 PASS e ERRO_INTEGRADO FAIL/1004 após Sort. A classificação nativa removeu Ignore; o selo recusou o conteúdo e o teste tentou usar Row sem ENCONTRADO. O adaptador OrdenarComMetadados8 captura Ignore por ID único, classifica e restaura o metadado por ID, sem regenerar o selo. A revisão integrada R2 passou 26 verificações. A execução falha permanece arquivada, não foi reclassificada. Classificação manual fora do adaptador pode bloquear a operação; não deve autorizar troca silenciosa.
3. Benchmark de100: primeira execução FAIL/1004; repetição instrumentada localizou SEED. Range.Errors(...).Ignore em múltiplas células não era aceito. Correção pontual: aplicar Ignore por célula, mantendo escrita dos valores em array. Os dois FAIL anteriores foram preservados; 100/1000/5000/10000 passaram depois.
4. Metadado auxiliar de formato na célula E7 de datas: a string 0.00 foi convertida pelo Excel em número0 porque a coluna estava Geral. O canônico persistido em D7 conserva ERRO:FORMATO_0.00 e a sessão fria observa 0.00; a comparação de recusa permanece válida. O rótulo FMT_ANTES=0 do TSV não prova formato0; usar o canônico D7 e a preparação VBA como evidência. Limitação de registro declarada, sem alterar a evidência histórica.

## Preservação e riscos

58 componentes herdados comparados com F7: nenhum corpo semanticamente alterado (comparação ignora VB_Base gerado pelo Office, caixa e espaços finais). Mestre F8 tem60 componentes/17 abas; fixtures com sete abas de dados adicionais têm67/24. Preservar fonte externo não prova binário FRX interno idêntico após serialização Office; pares externos não foram editados. Todos os sete Delete herdados permanecem.

A admissão/repetição usa recibo sintético marcado pelo próprio teste; não executa um motor completo de gravação de estoque. MudarCustodia8 é transformação de fixture, não evento de transferência/devolução do produto. Save recusado por evento é execução real de recusa, sem simular falha física do disco. Reverter experimentalmente exige checkpoint compatível; não usar estas fixtures como operação real.

Ver [matriz](./MATRIZ_REGRESSAO_FASE8.md), [persistência](./CONTRATO_PERSISTENCIA_FASE8.md) e [plano](./PLANO_IMPLANTACAO_FASE8.md).
