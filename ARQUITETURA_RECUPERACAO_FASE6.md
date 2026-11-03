# Arquitetura de recuperação operacional — Fase 6

**Proposta para aprovação; não integrada no produto.** A09 permanece crítico. A classe atual oferece compensação limitada em memória, anterior à fronteira Delete/Save. Os testes reais do protótipo não criam transação ACID entre planilhas e arquivos.

## Operação proposta

1. Identificar base/revisão, intenção persistente, episódio, veículo, projeção de origem e versão esperada.
2. Pré-validar campos, datas, valores, documentos conforme regras existentes, origem inequívoca, proteção e destinos. Capturar dados da interface antes de converter/gravar.
3. Consultar OperationID e conteúdo canônico. Rejeitar conteúdo divergente, versão obsoleta e qualquer pendência daquele episódio, mesmo com novo OperationID.
4. Calcular todos os destinos e suas identidades. Não executar nova busca por placa/NF-e entre gravações.
5. Persistir intenção e plano de efeitos, com pré-imagens suficientes e versão do esquema. Se a preparação durável não for comprovada, não escrever dados de negócio.
6. Gravar de modo controlado com checkpoints de etapa; capturar/restaurar valores, fórmulas e formatos suportados. Registrar falha individual e continuar a compensação dos demais itens quando seguro.
7. Verificar identidades, conteúdo, quantidade, formato textual das chaves, fórmulas e versões em todos os destinos.
8. Aplicar a baixa aprovada. No modelo C: transição do episódio/custódia, preservando projeções. Não reutilizar posição histórica nem apagar vínculos.
9. Persistir e reabrir/verificar o recibo e os efeitos antes de considerar concluído. Reconhecer a janela entre persistência do XLSM e registro externo da conclusão.
10. Registrar CONCLUIDA, COMPENSADA_VERIFICADA ou PENDENTE_CONCILIACAO. Estado desconhecido bloqueia repetição; a interface precisa apresentar essa restrição.

PREPARADA, GRAVANDO, VERIFICADA, BAIXA_APLICADA e SAVE_INICIADO são etapas de diagnóstico, não prova isolada de conclusão. Uma operação nunca pode ser considerada concluída apenas porque o VBA terminou sem erro.

## Diário mínimo proposto

Identificação da base e revisão; versão do esquema; OperationID/IntentID, AttemptID, EpisodeID, MovementID e ProjectionIDs; tipo de operação; versão esperada; hash canônico do envelope e referência ao conteúdo autorizado; sequência e etapa; resultado/erro; instante e sua origem; lista de efeitos esperados; pré-imagens/referências de recuperação estritamente necessárias; recibo de persistência e resultado da conciliação. Registrar responsável/motivo em correções aprovadas. Não copiar XML completo ou dados pessoais indiscriminadamente.

O diário experimental TSV registra sequência, operação, tentativa, etapa, resultado, erro e, na R2, episódio. A aba experimental de operações contém conteúdo canônico e versão esperada; snapshots de células permanecem em memória e os TSV antes/depois são evidência de teste. **O protótipo não implementa um diário de pré-imagens suficiente para recuperação autônoma após crash.** Fechar o arquivo após append não demonstra flush físico, durabilidade em queda de energia, exclusão mútua ou resistência à adulteração.

R2 verifica linhas truncadas, sequência crescente, operação órfã/desconhecida e unicidade do OperationID. Esse parser conservador não é um protocolo de armazenamento de produção. O selo textual dos vínculos também não é assinatura/autenticidade.

## Opções de armazenamento

| Opção | Vantagens | Limitações / condição |
|---|---|---|
| Aba auxiliar no XLSM | Distribuição simples; vínculos e dados no mesmo pacote | Depende do mesmo Save; pode desaparecer com corrupção/versão antiga; edição manual; não é write-ahead externo. Exige esquema/proteção aprovados. |
| Arquivo local separado, fora do OneDrive | Permite intenção anterior ao Save e diagnóstico independente | Dois arquivos sem commit conjunto; lacuna Save/recibo; acesso/privacidade, disco cheio, locks, caminho perdido, backups coordenados. Exige escritor único e protocolo durável testado. |
| Armazenamento relacional local ou serviço aprovado | Restrições únicas, transações internas e consultas de conciliação | Exige infraestrutura, permissões e segurança. Uma transação no banco não abrange automaticamente Excel. Access eventualmente disponível é alternativa a avaliar, não dependência testada/instalada nesta fase. |

Recomendação condicionada: selecionar uma fonte autoritativa única para fatos/diário e tratar planilhas como projeções verificáveis. Enquanto a escolha não ocorrer, limitar o modelo a escritor único, bloquear divergências e exigir conciliação. Não distribuir um TSV experimental como solução transacional pronta. Destino, dados, acesso, retenção e recuperação dependem de D04/D06. Não foi criada retenção nem exclusão automática de logs/backups.

## Idempotência e concorrência

A interface precisa guardar a intenção antes de qualquer efeito e reutilizá-la após interrupção. Gerar novo GUID em cada clique destrói a idempotência. Mesmo ID e mesmo envelope/efeitos íntegros retorna recibo anterior. Mesmo ID com dados diferentes é recusado. Pendência do episódio bloqueia também uma intenção nova. COMPENSADA só libera nova tentativa depois de conferir pré-imagens e persistência; mantém AttemptID anterior no diário.

A versão otimista do protótipo recusa comandos obsoletos em uma sessão. Não protege duas instâncias simultâneas, cópias concorrentes, sincronização OneDrive ou edição manual entre validação e Save. Futuro contrato deve definir exclusão de escritor, autoridade da base, fencing/revisão e tratamento de conflitos; testes multiwriter estão NÃO EXECUTADOS.

## Matriz de falhas e recuperação

| Ponto | Ação segura proposta | Evidência e limite |
|---|---|---|
| Antes da primeira escrita | Recusar; não criar efeito de negócio | Protótipo injeta falha e compara assinatura de dados. |
| Durante cada destino / entre destinos | Compensar em ordem inversa, conferir pré-imagem; persistir resultado | Testes de quatro gravações do protótipo; não equivalem aos eventos privados do produto. |
| Antes da baixa | Compensar quando todas as células forem recuperáveis | Teste sintético antes da fronteira. |
| Durante Delete estrutural | Não reconstruir automaticamente linhas/referências; isolar e conciliar | A_D_FALHA_BAIXA reproduz destinos com origem; A08/A09 continuam FAIL. |
| Durante baixa lógica | Tentar compensação limitada; se houver qualquer falha, marcar pendência | Proteção da aba de episódio impede restauração completa no ensaio; FAIL de integridade preservado. |
| Antes do Save | Recuperar e verificar enquanto ainda não iniciou persistência | C_ANTES_SAVE; sem garantia após a fronteira. |
| Save retorna erro | Tratar estado persistido como incerto; comparar disco, diário e memória | C_FALHA_SAVE usa SaveCopyAs para caminho pai inexistente, erro real 1004. Não é execução do Save legado nem simulação de disco cheio. |
| Após Save / antes do recibo externo | Bloquear repetição se workbook e diário divergem | C_APOS_SAVE injeta falha nessa janela e reabre a pasta salva. |
| Fechar/reabrir controladamente | Revalidar diário, recibo e efeitos; bloquear pendências persistidas | Verificador VBA real sobre o arquivo principal de cada cenário. Não equivale a crash. |
| Encerramento inesperado / energia / corrupção | Abrir em quarentena, identificar última revisão confiável e conciliar | NÃO EXECUTADO; não houve encerramento forçado nem garantia de recuperação. |
| Erro na compensação individual | Continuar itens recuperáveis; registrar endereço/erro; manter bloqueio | C_FALHA_COMPENSACAO protege um destino; recuperação integral FAIL, bloqueio subsequente conferido. |

## Conciliação proposta

Congelar o episódio e todos os seus destinos; preservar XLSM principal, imagem de memória, diário e versões com hash. Identificar a intenção e listar efeitos por ProjectionID, não por primeira coincidência de NF-e. Comparar origem, destinos, histórico, versões e fórmulas contra checkpoint/pré-imagens. Classificar cada efeito como ausente, íntegro, duplicado, divergente ou desconhecido. Quando a identidade for ambígua, decisão humana obrigatória.

Preparar um plano de reparo explícito e reversível, aprovado pelo responsável, com identificação exata de cada efeito e nova revisão. Só liberar repetição depois de confirmar todos os efeitos/compensações e persistir recibo de conciliação. Nunca repetir automaticamente sobre estado incerto nem interpretar Cancelado como desfazimento fiscal autorizado.

## Restauração das evidências desta execução

A referência Fase 5 e os checkpoints estão preservados no [manifesto inicial](./_auditoria/fase6_20261101_224949Z/manifesto-inicial-fase6.json). A pasta de [backup da Fase 6](./_auditoria/fase6_20261101_224949Z/backup/) contém documentos anteriores, fontes de desenvolvimento e fixture de referência.

Para repetir ensaios: copiar o mestre experimental final para nova pasta TEMP fora do OneDrive que preserve o marcador VBA_Fase6_20261101_224949Z e AUDITORIA_VBA_FICTICIO; conferir hashes; executar F6_RodarLote somente em raiz nova. A suíte recusa sobrescrever cenário existente. Não executar cópias arquivadas no workspace/OneDrive. estado_memoria_nao_promover.xlsm é imagem forense de estado possivelmente inconsistente, não backup operacional aprovado.

Para desfazer somente a documentação: conferir hashes e restaurar os arquivos correspondentes de backup/documentacao; preservar os relatórios e evidências desta fase antes. Nenhum fonte funcional principal foi alterado e não há política nova de baixa a reverter no produto.
