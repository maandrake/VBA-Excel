# Defeitos e pendências — candidata Entrega 1

Escopo: cópia sintética integrada. Nenhuma conclusão abaixo altera os sete FAIL históricos no produto original. A [matriz](./MATRIZ_REGRESSAO_ENTREGA.md) conserva execuções e limites.

| ID | Prioridade | Pendência / comportamento atual | Medida necessária |
|---|---|---|---|
| E1-01 | Alta | Recuperação operacional de PREPARED/Pending não está integrada. O ACE reverte efeitos, mas preserva intenção e rascunho; novas tentativas do episódio são recusadas. | Implementar conciliação auditável, decisão de abortar/retomar e recibo terminal. Não apagar Pending para liberar operação. Políticas ainda pendentes. |
| E1-02 | Alta | Fluxo D parcial: alteração, exclusão manual, retificação e filtros legados ainda incompatíveis são bloqueados. Nem todos os 57 leitores e relatórios foram migrados. | Migrar leitores/escritores por identidade, revisão e elegibilidade; habilitar somente após regressão. Não usar ActiveCell/primeiro Find para editar históricos. |
| E1-03 | Alta | Pacote preservado vincula AuthorityPath ao laboratório executado. Preparar-CopiaSintetica.ps1 foi criado, mas ExecutionPolicy bloqueou sua execução. Portabilidade/rebind e restauração integral em outra pasta NÃO testados. | Revisar/autorizar mecanismo de implantação sintética explícito e ensaiar cópia, rebind e reconstrução. Não usar bypass nem eleger uma cópia como autoridade automaticamente. |
| E1-05 | Alta | Concorrência dos novos gravadores C1 não foi exercitada em dois processos; fencing/transações são herdados do F10. | Reexecutar interleavings nos eventos integrados, fonte alterada após seleção, perda de lease e preparação órfã. Não extrapolar testes antigos. |
| E1-06 | Média | Contexto de formulário é indexado por TypeName. O fluxo testado usa as instâncias padrão; múltiplas instâncias do mesmo tipo não foram ensaiadas. | Identificar sessão/instância de formulário antes de permitir janelas concorrentes. |
| E1-07 | Média | Reentrada usa helper e VehicleID explícito fornecido pela suíte. Interface final para escolher episódio encerrado/proveniência ainda não está pronta. | Implementar escolha inequívoca e conciliação de dados ausentes; não deduzir VehicleID por placa. |
| E1-08 | Alta | Cancelado e outras operações além de Venda/Transferência matriz e Devolução filial não estão integradas. Cancelado é recusado na candidata; original permanece inalterado. | Decisão operacional/fiscal e integração preservando comportamento aprovado. Não declarar homologação fiscal por XML sintético. |
| E1-09 | Média | Reconstrução mantém 17 abas, mas somente oito são projeções dos fluxos integrados. Fórmulas anuais N/O de apresentação, relatórios gerais e demais abas não têm equivalência integral comprovada. | Inventariar saída esperada e migrar fórmulas/condicionais restantes; usar dados discriminantes por episódio. |
| E1-10 | Média | Backup C1 foi reaberto/auditado; restauração completa, perda de mídia e falha de disco ainda não ensaiadas. Botões de backup/manutenção legados fora dos caminhos C1 exigem adaptação. | Integrar todos os pontos de entrada ao backup pareado e verificar restauração sem descartar fatos posteriores. |
| E1-11 | Alta | Instalação do esquema 5 usa DDL em etapas; interrupção pode deixar instalação parcial. | Empacotar base pronta verificada ou migração retomável com checkpoints; nunca repetir DDL indiscriminadamente sobre fatos existentes. |
| E1-12 | Alta | Não há garantia contra energia, corrupção física, disco cheio, SQL externo ou adulteração intencional dos hashes. | Ensaios próprios e modelo de ameaça/permissões. Não prometer ACID entre ACE e XLSM. |
| E1-13 | Média | Precisão/conversão corrigida para texto canônico e entrada brasileira com vírgula; todos os extremos e entradas ambíguas como 1.000 não foram ensaiados. | Contrato explícito de entrada regional; testar limites monetários de quatro casas e arredondamento. Não aceitar inferências silenciosas. |
| E1-14 | Média | Designer embutido foi reserializado pelo Excel; comparação binária difere, embora assinaturas/eventos se mantenham e controles usados tenham funcionado. | Regressão visual/funcional dos cinco forms e inventário completo dos controles. Não declarar FRX embutido byte a byte igual. |
| E1-15 | Alta | Logs, retenção, limite de recuperação e semântica de episódio/custódia ainda são hipóteses/políticas não aprovadas para uso operacional. | Decisão expressa do responsável antes de qualquer dados reais ou promoção. |

## Defeitos encontrados e corrigidos nesta execução

1. Conversão regional CCur em payload canônico: primeira base incorreta preservada forense; nova base R3 confere 900 e 650,25.
2. Ignore em planilha protegida: acesso controlado com restauração da proteção; verificação persistida concluída.
3. Data normalizada regionalmente: formato literal e contrato F8; não se removeu NumberFormat do verificador.
4. Chamada de reconstrução interpretada como rótulo: Call explícito. R2 FAIL preservado; R3 venda PASS.
5. Campo apagado por evento Change ao reutilizar formulário: operação configurada antes dos campos; transferência sucessiva FAIL preservado e nova execução PASS.
6. VerifyBatch10 exigia chave em entrada matriz cujo contrato admite vazio: alinhado ao esquema 5; diagnóstico distingue PENDING de memória verificada; reconstrução posterior confirmada.
7. Novo backup liberava lease cedo: lease persistido mantido durante cópia fechada, verificação e manifesto; backup reaberto PASS.
8. Um dump de teste existente causou erro 58 antes de acionar o evento; prefixos por execução corrigidos, sem sobrescrever evidência.
9. Uma substituição de código do novo backup duplicou a assinatura; compilador real detectou, correção foi aplicada e compilação final passou. Não foi falha de estoque.

10. E1-04 — reabertura: GetAbsolutePathName comparava aliases curto/longo do mesmo banco como diferentes. A candidata usa ShortPath de arquivos existentes e mantém BaseID/AuthorityPath. Compilação real e três verificações de caminho passaram; novo processo inicializou CACHE_AUDITADO e E1_SessaoReaberta PASS na execução 47. O FAIL da execução 45 está preservado. Escopo encerrado: reabertura normal no mesmo laboratório; E1-03/portabilidade permanece aberto.

## Referências históricas preservadas

V_VALIDO, T_VALIDO, D_VALIDO, D_FALHA_BAIXA, T_DUP_DESTINO, C_FALHA_BAIXA e C_FALHA_COMPENSACAO continuam registrados como FAIL da implementação anterior. Há resultados próprios para a candidata, com escopo na matriz. A candidata não usa compensação de células depois de Delete; usa transação ACE e projeção reconstruível. A ausência desse caminho não converte testes legados em PASS.
