# Autoridade e reconstrução do Excel — Fase 9

Hipótese demonstrada no laboratório: **ACCDB experimental autoritativo → projeções Excel descartáveis/reconstruíveis**. Nenhum banco foi ligado ao estoque operacional. Commit do banco e Save do XLSM são fronteiras separadas.

## Fluxo executado

1. Resolver episódio/versão e conferir sessão/Fence, identidade, estado e pendências.
2. Canonicalizar comando com BaseID, operação, episódio, ExpectedVersion, tipo, quantias, data/documento/chave sintéticos e payload. Preparar ProjectionID/MovementID uma vez.
3. Confirmar intenção, tentativa e Pending numa transação própria. Mesmo OperationID/envelope retorna recibo íntegro; outro payload recusa; episódio pendente bloqueia ID novo.
4. Iniciar transação de efeitos; validar propriedade/fonte; inserir movimento/projeção, atualizar episódio por CAS, gravar ChangeLog, revisão e recibo, remover Pending e confirmar. Interrupção anterior reverte efeitos; intenção continua rastreável.
5. Consultar relações em lote, verificar PK/unicidade/identidade, canônico tipado, Currency com quatro casas e chave String de 44 caracteres. Nenhuma linha do Excel determina qual registro é autoritativo.
6. Materializar array no XLSM novo: IDs/nota/chave/data como texto; chave @ antes da atribuição e Ignore; quantias com quatro casas; elegibilidade por estado e CurrentPID; fórmulas derivadas no conjunto de linhas do mesmo PID/EP/VID.
7. Footer BaseID/RevisionNo identifica revisão projetada; conferir conteúdo e metadados antes de utilizar um comando. Token antigo resolve de novo na autoridade e recusa versão obsoleta. ListBox real recebeu arrays de14 colunas no benchmark, sem acionar eventos privados operacionais.
8. Se Excel interromper ou Save falhar, preservar banco/recibo; projeção permanece pendente. Reconstruir a partir da autoridade e conferir novamente. Não repetir movimento para “refazer a planilha”.

Recibo de negócio é persistido antes da projeção, no commit do banco. Ack persistente por consumidor/projeção, seu protocolo de divergência e falha após Save continuam requisitos não implementados. Uma falha depois de commit não autoriza apagar fatos confirmados para imitar rollback do Excel.

## Provas e resultados

F9_Core:7 movimentos/recibos e2 episódios ativos; A1 encerrado e A2 ativo do mesmo VID. F9_APOS_COMMIT:8 projeções/recibos, Excel desatualizado. F9_ProjectionFaults interrompe após cabeçalho, verifica banco igual e reconstrói8/2. Valores de entrada1000,1234/1500,1111/2200,4321 e saídas1370,9876/1940,2222/2100,3333/2500,4444 distinguem os vínculos. Sort comparou PID/EP/VID/valores/margem/anual antes/depois, não apenas ausência de#REF!. [Resultados](./MATRIZ_REGRESSAO_FASE9.md).

F9_Complemento recusou footer de outra revisão e edição local de valor. A verificação explícita detectou **Ignore perdido pelo Sort nativo**: F9_METADATA_APOS_SORT_ANTES é FAIL preservado. Restaurar o metadado derivado e revalidar conteúdo/identidade passou. Isso não aprova classificar manualmente qualquer projeção nem os leitores antigos. A cópia Projecao_Inspecao_Metadata foi depois deliberadamente adulterada e salva para prova negativa; não usá-la como versão de referência.

## Contrato de verificação ainda a completar

VerifyProjection9X cobre IDs/versão/estado/CurrentPID, valores, chave, formato @, Ignore, presença das fórmulas e margem. Não confere todas as expressões de fórmulas, estilos, formatos monetários/data e todas as relações da aplicação. A anual do modelo usa o estado do episódio derivado da venda sintética; não implementa toda regra anual/fiscal existente. Não chamar essa conferência parcial de comparação integral das17 abas.

Índice em memória localiza candidatos; consulta verificada ainda acessa a autoridade e valida identidade, versão, canônico e pendência. Uma projeção pode estar visualmente ordenada e ainda ter metadados divergentes. Para o produto, exigir validação completa do contrato de apresentação e um snapshot/revisão estável sob escritor exclusivo ou transação de leitura; leitura concorrente durante reconstrução não foi testada.

Restauração: copiar banco fechado e mestre final para TEMP novo, preservar evidências, conferir hashes e nome exigido pela guarda. Nunca abrir o XLSM arquivado no OneDrive para executar macros. Depois de movimentos novos, não restaurar banco/XLSM antigos cegamente: congelar, preservar ambas gerações, conciliar fatos/recibos por IDs, reconstruir projeções e aprovar a revisão de retorno.
