# Contrato de persistência e conciliação — Fase 8

Especificação proposta, não protocolo implantado no sistema principal. A fonte autoritativa depende de D06; a operação experimental admite intenção/recibo de teste e bloqueios, sem motor completo de movimentação. Não existe alegação de ACID entre Excel e armazenamento externo.

## Envelope canônico proposto

BaseID, SchemaVersion, OperationID imutável, AttemptID novo por tentativa, tipo da operação, VehicleID, EpisodeID, ProjectionIDs de origem/destino, MovementIDs, ExpectedVersion de cada agregado, payload canônico, hash de payload, revisão anterior, fases realizadas, evidências de efeitos, revisão/recibo final e motivo de pendência. Quantias com escala explícita; datas somente data civil com calendário declarado; chaves NF-e String44. Hash detecta divergência, não autentica usuário nem assinatura fiscal. Não inserir XML integral ou dados pessoais desnecessários no diário.

OperationID não é número de NF-e. Mesmo OperationID + mesmo envelope + recibo e efeitos íntegros retorna o recibo anterior; não executa novamente. Mesmo ID com payload/identidade/versão diferentes é conflito. Novo OperationID não contorna episódio pendente. AttemptID não redefine a identidade da operação. Todos os campos do envelope deverão ser cobertos pela serialização; o protótipo F8 usa payload sintético explícito, e não prova a canonicalização de todos os campos dos gravadores reais.

## Fluxo de 12 passos

1. Identificar BaseID, operação, tentativa, sessão e responsável.
2. Resolver IDs únicos e conferir ExpectedVersion na autoridade.
3. Validar as regras operacionais/fiscais já aprovadas, sem inventar novas.
4. Verificar pendências do episódio, da operação e da base; estado incerto bloqueia escrita.
5. Conferir OperationID, envelope, payload e recibos: repetição íntegra não gera efeitos.
6. Preparar plano explícito dos destinos, versões e pré-imagens, sem mutação operacional.
7. Persistir intenção e verificar sua leitura segundo o armazenamento aprovado, antes dos efeitos. Se a durabilidade não foi confirmada, não avançar.
8. Executar efeitos por ID, com registro de etapas e sem depender da posição da linha.
9. Reconsultar e comparar todos os destinos, vínculos, tipos, fórmulas, formatos e contagens.
10. Aplicar baixa lógica somente conforme política aprovada, persistir efeitos e obter recibo da autoridade. Save da projeção não substitui commit da autoridade.
11. Confirmar operação e revisão; registrar recibo íntegro e tornar a projeção coerente com a autoridade.
12. Qualquer divergência ou resultado de persistência incerto entra em AGUARDANDO_CONCILIACAO, preserva evidências e impede nova tentativa.

Estados propostos: PREPARADA, EM_EXECUCAO, EFEITOS_VERIFICADOS, PERSISTENCIA_PENDENTE, CONCLUIDA, ABORTADA_COMPROVADA, AGUARDANDO_CONCILIACAO. Sem recibo não presumir CONCLUIDA; sem efeitos verificados não presumir ABORTADA. O protótipo não implementa toda essa máquina.

## Matriz de falhas e tratamento

| Falha | Tratamento seguro proposto | Limite/evidência desta fase |
|---|---|---|
| Antes da primeira escrita | Abortar sem efeitos; liberar apenas após prova de nenhuma alteração | Pré-validação experimental testada; protocolo produtivo NÃO EXECUTADO |
| Durante cada destino | Compensar apenas as células/pre-imagens comprovadas antes da fronteira; verificar recuperação; senão conciliar | Recuperação integral F6 continua FAIL; não houve novo gravador completo F8 |
| Antes da baixa | Verificar destinos e pré-imagens; compensação limitada se comprovadamente íntegra | Guardas de pendência testadas, não equivalem à recuperação |
| Durante a baixa | Estado incerto: preservar origem/destinos e conciliar, sem repetir | C_FALHA_BAIXA continua FAIL de recuperação integral |
| Durante Save | Não emitir recibo de sucesso; comparar disco e memória em cópias; bloquear | F8 executou Save real cancelado por evento; não simulou disco cheio/permissão de arquivo |
| Após Save, antes do recibo | Verificar autoridade e efeitos por OperationID; emitir recibo somente se inequívocos; senão conciliar | Guarda POS_SAVE reaproveitada; não prova durabilidade de produção |
| Após recibo | Revalidar recibo/efeitos/revisão; devolver resultado, sem novas gravações | Repetição sintética em F8; recibo marcado pelo teste, não por motor de negócio |
| Diário truncado | Recusar parser/etapa incompleta; nenhuma aplicação automática | Sessão fria da suíte F7; guarda, não reconstrução |
| Diário órfão | Identificar BaseID/geração ausente; quarentena e conciliação | Sessão fria da suíte F7 |
| Histórico alterado | Divergência de selo/vínculo bloqueia; preservar ambas versões | Guarda F7 e edição sem selo F8 |
| Dados/diário divergentes | Nenhum arquivo eleito pela data; verificar linhagem, hashes e recibos | Contenção fria F7; autoridade ainda pendente |
| Compensação incompleta | Registrar cada falha, interromper autorização de nova tentativa | C_FALHA_COMPENSACAO continua FAIL integral |
| Nova tentativa em estado incerto | Bloquear tanto mesmo ID quanto ID novo no mesmo episódio | Guardas F8 testadas; bloqueio não é recuperação |
| Crash/energia/concorrência | Exigir protocolo/motor específico, recuperação na abertura e conciliação | NÃO EXECUTADO, sem encerramento forçado nem dois escritores |

## Roteiro de conciliação

Congelar comandos de escrita do escopo afetado; copiar disco, memória quando disponível, diário e recibos para checkpoint datado; conferir hashes. Identificar BaseID, operação, episódio e versões. Comparar origem, cada destino, histórico e baixa por ID e conteúdo — sem escolher a primeira placa/NF-e. Classificar efeitos ausentes, completos, parciais ou contraditórios. Responsável aprova conclusão/compensação em cópia de teste, com relatório antes/depois e nova tentativa segura comprovada. Só então aplicar a decisão no escopo explicitamente aprovado e registrar recibo de conciliação separado. Não reescrever o passado para parecer sucesso.

Se houver movimentos posteriores, não restaurar XLSM antigo cegamente. Preservar a autoridade atual, reconciliar os movimentos posteriores e reconstruir projeções. Evidência insuficiente mantém operação bloqueada.

## Data, selo e apresentação

F8 usa Value2 Double, serial Excel1900, somente dias inteiros entre 61 e 2958465; o bit pattern IEEE754 é serializado em hexadecimal, sem separador regional. Date1904, texto que parece data, fração/hora e overflow são recusados por este contrato experimental. Isso não é regra de negócio retroativamente aprovada para todos os documentos.

ISO_LITERAL aceita apenas formatos ISO explícitos equivalentes; BR_LITERAL exige dia/mês/ano literal e recusa m/d/yyyy. LEGACY_D14_LOCAL aceita conscientemente a apresentação local do formato embutido 14 (dd/mm/yyyy ou m/d/yyyy); não garante exibição brasileira. A política faz parte do selo. Formatos materiais diferentes são recusados. Nos demais campos, NumberFormat observado e Ignore continuam no selo, além de tipo e valor. Selos não são regenerados durante leitura. Formato efetivo antes/depois é registrado mesmo quando canonicamente equivalente.

A equivalência foi testada no Office/região desta máquina; mudança real de idioma/região permanece NÃO EXECUTADA. Datas diferentes não podem ter o mesmo canônico; células com fórmula na fonte experimental são recusadas. Fórmulas de projeção são verificadas separadamente.

Ver [correções](./CORRECOES_PROTOTIPOS_FASE8.md), [matriz](./MATRIZ_REGRESSAO_FASE8.md) e [armazenamento](./PROPOSTA_ARMAZENAMENTO_FASE8.md).
