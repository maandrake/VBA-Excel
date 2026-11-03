# Plano de recuperação de falhas — Fase 5

**Proposta técnica; não autoriza alterações em produção.** [Resultados reais](./MATRIZ_REGRESSAO_FASE5.md), [decisões D01/D02](./PENDENCIAS_DECISAO.md). Excel/VBA não oferece transação ACID entre essas planilhas. Não há rollback completo depois de Delete, Save ou encerramento inesperado.

## Fronteiras comprovadas

| Estado | Evidência | Conduta |
|---|---|---|
| Pré-validação rejeita dados | 15 casos sem diferenças nas 7.560 células monitoradas | Corrigir a entrada; isso não cobre toda falha artificial antes da primeira escrita |
| Escritas anteriores a MarcarBaixaIniciada/MarcarSalvamentoIniciado | Quatro erros 13 intermediários, um por fluxo; igualdade exata e estados Application restaurados | Conferir recuperação e todas as abas; repetir somente depois de estado anterior comprovado |
| Recuperação individual falha | Suíte complementar reconhece falha e continua nos outros snapshots; célula bloqueada fica alterada | Recuperação incompleta; preservar diagnóstico/checkpoint e conciliar. Não repetir o negócio automaticamente |
| Baixa marcada/tentativa de Delete | Devolução erro 1004: quatro destinos gravados e origem existente | Conciliação obrigatória, mesmo sem exclusão efetiva; o marcador limita a garantia |
| Delete concluiu | Venda/Transferência/Devolução quebraram fórmulas ou associações | Preservar a tentativa; não compensar por endereços antigos nem executar novamente |
| Save iniciado/concluído ou Excel encerrado | Falha real de Save não provocada; encerramento não testado | Não presumir estado em disco/memória; comparar cópias e checkpoint sob decisão responsável |

Snapshots são locais à sessão e aos intervalos registrados, sem diário persistente. Não cobrem alterações estruturais nem todas as propriedades de formatação. NumberFormat/Ignore e fórmulas foram examinados; bordas/fontes/cores não integram o comparador. EnableEvents, ScreenUpdating, DisplayAlerts e Calculation devem retornar aos valores anteriores; os TXT observados mostram True/True/True/-4105. O teste com outra pasta ativa recuperou a fixture e preservou Pasta1.

## Conciliação proposta para estado incerto

1. Suspender novas tentativas e guardar operação, etapa, erro, horário, parâmetros mínimos autorizados e retorno da recuperação. Não encerrar globalmente Excel nem forçar salvamento de uma pasta real.
2. Preservar cópias distintas do estado disponível e do checkpoint anterior com SHA-256. A captura diagnóstica não deve sobrescrever a pasta principal ou o backup válido. Destino/privacidade de logs reais dependem de D04.
3. Em cópias, localizar **todos os candidatos** de origem e destinos: Saidas/Todas Saidas; também Entradas Filial/Entradas Anual Filial na transferência; Saidas Filial/Todas Saidas Filial/Entradas/Entradas Anual e origem filial na devolução.
4. Comparar NF, chave como texto, item/veículo, valor/data, operação e históricos. Não presumir placa, nNF ou chave isolada como identidade da passagem. Guardar ambiguidades sem escolher o primeiro Find.
5. Conferir fórmulas e associação ao veículo: INDIRECT pode mostrar o veículo seguinte sem erro. Examinar M/N/O e vínculos externos à linha excluída.
6. O responsável aprova a passagem correta e quais escritas serão mantidas, corrigidas ou restauradas. Conciliar em cópia antes de qualquer ação operacional; não excluir históricos automaticamente.
7. Somente depois de consistência, persistência verificada e decisão sobre identidade/idempotência autorizar nova tentativa. Nenhuma repetição automática após Delete/Save/recuperação parcial foi integrada.

Diário persistente por operação e protocolo de reconciliação são propostas futuras, dependentes de identidade, armazenamento, privacidade e retenção aprovados. Não foram implementados.

## Restauração das fixtures

A [base limpa compilada](./_auditoria/fase5_20261101_202114Z/entrega/Fixture_Auditoria_Fase5_Base_Compilada_NAO_HOMOLOGADA.xlsm) tem **56 componentes**, SHA-256 **f989289554a5c081cef8bd13de2e0b6f0e48ec4c1ba1dd5e75c349eff458f72c**. A [base Fase 4 preservada](./_auditoria/fase5_20261101_202114Z/backup/Fixture_Auditoria_Fase4_Integrada_Testada.xlsm) tem SHA **dcbd72e46e8b182a9810254bcce460ea2cb17260b4bcb47b622ddc3d4f9879a0**. Ambas são sintéticas. Nenhuma substitui o XLSM original.

Os [cenários preservados](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/) contêm estados posteriores, inclusive inconsistentes. Não reutilizá-los como base. Para repetir, conferir o hash da base limpa com Get-FileHash -LiteralPath no PowerShell; copiar para um diretório **novo** em TEMP, fora do OneDrive, cujo caminho contenha VBA_Fase5_20261101_202114Z. Manter o marcador fictício Menu Principal!Z1. Nunca sobrescrever uma evidência existente.

Exemplo de Venda: nomear a nova cópia Fixture_Auditoria_Fase5_V_VALIDO.xlsm, copiar a pasta xml da [evidência V_VALIDO](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/V_VALIDO/) para o mesmo diretório novo. Abrir somente essa cópia no Excel, Alt+F11, selecionar seu projeto e conferir Depurar > Compilar. O comando já foi executado nesta fase; desabilitado na conferência final do projeto selecionado indicou ausência de compilação pendente. Registrar qualquer erro/módulo/linha ou bloqueio sem alterar segurança.

Na janela Imediata do projeto correto, executar F5_Preparar "V_VALIDO": F5_Mostrar. Conferir os dados fictícios; clicar Gravar, selecionar xml/valido.xml da nova pasta, observar e fechar o relatório XML. Ao terminar, fechar somente os formulários de teste e executar F5_Concluir no mesmo projeto. Esse helper salva a fixture para evidência; não comprova sucesso do Save da aplicação em cenários de erro. O resultado atual esperado é **FAIL de integridade** conforme a matriz, sem nova tentativa sobre o estado resultante.

Para outros cenários usar nova cópia com nome Fixture_Auditoria_Fase5_<CENARIO>.xlsm e F5_Preparar com o mesmo identificador. O helper rejeita caminho OneDrive, ausência de marcador e cenário com antes.tsv existente. D_FALHA_INTER foi repetido uma vez nesta execução somente após igualdade verificada; arquivos da primeira tentativa foram preservados.

Para reproduzir a falha intermediária, colocar breakpoint na instrução ThisWorkbook.Worksheets("Todas Saidas").Activate, dentro do gravador matriz, ou "Todas Saidas Filial" no gravador filial. Acionar o botão real. Quando pausado, verificar os destinos com 22001, executar F5_FalhaEntreGravacoes pela janela Imediata e continuar F5. Para falha de Delete da Devolução, parar antes de opContexto.MarcarBaixaIniciada, executar F5_ProtegerOrigem somente nessa fixture e continuar. A senha dessa proteção sintética é F5QA; não alterar proteção de dados reais. Remover os breakpoints de teste ao terminar.

As suítes AuditoriaFase2_Executar e AuditoriaFase4_IntegridadeComplementar são reutilizadas em nova cópia sintética, com caminho TEMP de saída conforme assinatura no [fonte real](./_auditoria/fase5_20261101_202114Z/scripts/fase4/modTestesFase2.bas). AuditoriaFase3_CicloVida permanece BLOQUEADA pela preferência do depurador; nenhuma configuração foi alterada e nenhum PASS atribuído.

Para reverter fontes, recuperar o checkpoint Fase 4, conferir os hashes e importar somente em nova fixture pelo VBE. A Fase 5 adicionou um módulo de teste e não alterou a lógica da aplicação. Nenhuma restauração de produção está autorizada por este roteiro.
