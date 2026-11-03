# Relatório técnico — Fase 7

**Portão1 executado como especificação e protótipo; Portão2 não aprovado. Sistema NÃO HOMOLOGADO.** A08/A09 continuam críticos na aplicação principal. Não houve novo esquema, política de baixa, diário ou regra fiscal integrado ao sistema principal.

Execução 20261102_025115Z, iniciada em 01/11/2026 no horário local e continuada em 02/11. [Checkpoint/manifesto inicial](./_auditoria/fase7_20261102_025115Z/manifesto-inicial-fase7.json). Os dez documentos obrigatórios foram lidos; arquivos e evidências F6 foram conferidos por existência/hash, não apenas por caminhos documentados.

## Resultado técnico

O [contrato de identidade](./CONTRATO_IDENTIDADE_FASE7.md) define IDs imutáveis de veículo, episódio, movimento, projeção e intenção, separando AttemptID de OperationID e documento fiscal de passagem pelo estoque. A continuidade do episódio em transferência/devolução interna é recomendação específica D01-C, ainda não aprovada. Reentrada após encerramento cria episódio novo somente com identidade comprovada.

O [mapa de leitores](./MAPA_MIGRACAO_LEITORES_FASE7.md) reconfirma57 procedimentos dos componentes salvos e as dependências das17 abas já inventariadas. Cada item indica origem, identificação atual, dependência física, risco, adaptação, legado e regressão. Não é alegação de inexistência de chamadores dinâmicos/externos não disponíveis.

Os adaptadores VBA experimentais passaram em18 verificações dirigidas, incluindo elegibilidade após baixa/reentrada, lacunas, seleção por token após reordenação, versão obsoleta, ausência/duplicidade, histórico ambíguo, custódia T/D/T, chave e fórmulas por IDs. ListBox é o objeto de teste AuditoriaListaFake; os eventos reais dos formulários não foram migrados nem reexecutados nesta fase.

Sessão fria real:12 cenários/15 verificações PASS em uma nova instância Excel /x depois de fechar normalmente a preparação. Guardas recusam contratos órfãos/truncados/ausentes/divergentes, estados pendentes, histórico alterado, payload divergente e versão obsoleta, sem escrita no conjunto comparado. Os23 arquivos permaneceram idênticos por SHA-256. Isso é contenção experimental, não recuperação durável ou transação ACID.

Contagem consolidada:38 verificações únicas,34 PASS/quatro FAIL;40 execuções de assertivas preservadas, sem somar repetição. Uma das34 é comparação numérica não discriminante com valores iguais; não vale aprovação de identidade. A [matriz](./MATRIZ_REGRESSAO_FASE7.md) separa essa observação, os18 checks de adaptadores, os15 de guardas e as quatro falhas. Tentativa inicial interrompida e FAIL anterior da montagem da fórmula permanecem arquivados.

## Falhas e correções somente da suíte

| Arquivo/procedimento | Causa / alteração / dependência | Evidência e resultado | Risco e reversão |
|---|---|---|---|
| modPrototiposFase7 / Item7 e F7_Leitores | AuditoriaListaFake.List retorna array sem argumentos; instrumentação ajustada antes de importar para ler token sem presumir API de MSForms | R01/R04/R09 PASS; assinatura da classe herdada mantida | Não valida controle visual real. R1/R2 preservados no checkpoint |
| OpenInput7 / RestaurarFormatoSintetico7 | Reabertura normaliza formato de data embutido; selo estrito F6 passa a divergir. Preparação das cópias restaura somente formato J já explicitamente selado, sem recriar selo/identidade | BASE_FORMATO_REABERTURA continua FAIL; primeira tentativa ERRO5 preservada. Demais leitores puderam ser testados sob formato restaurado | Contrato persistente de formatos precisa solução futura; não aplicado ao produto. Backup mestre R2 e R3 |
| F7_RetestarFormulas | Fórmula de chave inserida em célula @ permaneceu texto. Reteste insere fórmula em General e aplica @/Ignore ao resultado String | R06 inicialmente FAIL, reteste PASS; R17 String44/@/Ignore PASS | A rotina F7_Leitores original conserva montagem inicial e requer reteste; não declarar sua execução isolada integralmente aprovada. Não houve alteração do gravador |
| F7_Campo / navegação | Consulta por ProjectionID+EpisodeID+VehicleID, dependências de dados/IDs explícitas; token é resolvido novamente, sem linha como autoridade | A movida4→10, C na4, lacunas e duplicação: acessos por ID PASS | UDF experimental depende do mestre aberto; não migra fórmulas do produto nem valida sozinha todo o envelope. Reversão descartando candidato |
| GuardCold7 / F7_VerificarFrio | Contrato mínimo em aba sintética e TSV de nove campos; recusa conservadora de incoerência em sessão nova |15 checks PASS, assinatura5.280 células iguais e23 hashes iguais | Selo mínimo de duas linhas, sem lock, pré-imagens duráveis ou flush; não executor de movimentações. Imagens F6 de falha continuam inconsistentes |

Quatro FAIL consolidados: (1) selo/formato após reabrir; (2) fórmula anual por placa para reentrada; (3) INDIRECT físico passa a C; (4) auxiliar físico de A calcula2400 de C no reteste com valores distintos. O primeiro PASS do auxiliar,1800=1800 com valores iguais, era insuficiente para revelar a associação; não foi usado para ocultar o FAIL posterior.

Os cinco FAIL oficiais F5 V_VALIDO/T_VALIDO/D_VALIDO/D_FALHA_BAIXA/T_DUP_DESTINO continuam como referência obrigatória e não foram reclassificados. C_FALHA_BAIXA e C_FALHA_COMPENSACAO F6 continuam FAIL de recuperação integral; a nova guarda PASS não conserta seus efeitos.

## Abordagens comparadas para eliminar posição como autoridade

| Abordagem | Avaliação / recomendação |
|---|---|
| Resolver campos por IDs em VBA | Adaptação testada; exige unicidade, versão, envelope e elegibilidade. Construir índice por revisão em volume, pois protótipo usa laços/células e pode ser quadrático |
| Tabela de projeções/mapeamento | IDs/FKs persistentes e índice descartável de localização; LinhaDica não autoriza fallback. Necessária no esquema aprovado, sem ocupar U:X automaticamente |
| UDF por ID com ranges explícitos | Prova real resiste ao deslocamento e conserva identidade/chave. Consulta linear por fórmula; evitar colunas inteiras. Depende de distribuição/compilação local, cálculo e cache versionado. Não benchmarkada em volume |
| INDEX/MATCH ou funções nativas por ID | Compatibilidade e legibilidade potencialmente melhores, mas primeira ocorrência pode ocultar duplicidade. Exigir teste explícito de unicidade/erro antes de retornar e adaptar todas as dependências |
| Estoque ativo como projeção separada do histórico | Recomendada com fonte autoritativa; leitura operacional usa mesma elegibilidade em pesquisa/contagem/totais. Reconstrução precisa preservar formatos/fórmulas/vínculos e nunca reusar identidade |
| Referência direta/INDIRECT / COUNTIF placa | Permanecem FAIL nos casos demonstrados. Nenhum resolvedor aprovado as corrige automaticamente |

Recomendação: C com identidade, projeções e leitores migrados conjuntamente; diário/fonte autoritativa aprovados antes de persistência operacional. O [plano reversível](./PLANO_MIGRACAO_REVERSIVEL_FASE7.md) exige leitores/fórmulas/históricos antes da baixa. [Protocolo](./PROTOCOLO_PERSISTENCIA_CONCILIACAO_FASE7.md) compara XLSM, arquivo local e banco, sem assumir commit conjunto com Excel.

## Preservação, ambiente e entrega efetiva

1.986 hashes iniciais coincidiram: 39 originais,1.230 arquivos das cinco fases anteriores e717 itens da entrega/fechamento F6, incluindo documentos.28 backups novos verificados antes de alterar arquivos. O manifesto F6 mantém SHA627b46c182776f340fc850641fabca0b924aba6f79ecb29ce10c934ba92e5de4.

Referência F6 preservada:7f0f41f7336ebae8f7cfea2b698335be4ef346e78743cde34d899299dd848de1, 57 componentes/17 abas. Original XLSM:2ad1c635da68b51a68d5d2c0d3cd638f1ff8064942b299699f2594d7a9ca8850. Nenhum desses arquivos foi aberto para movimentar dados reais.

[Entrega experimental F7](./_auditoria/fase7_20261102_025115Z/entrega/Fixture_Arquitetura_Fase7_EXPERIMENTAL_NAO_HOMOLOGADA.xlsm): SHA-256 **c7858e3682403f719784db4f817dcd7810ee1ff617fb6c3a7d647d8050d16b1a**,58 componentes reais (33 BAS,2 classes,5 UserForms,18 documentos),17 abas. Um módulo experimental adicional com27 procedimentos;57 herdados canonicamente iguais com literais preservados, descontando atributos VB_Base e recasing/espaçamento do VBE. Não afirmar igualdade binária dos designers reserializados pelo Office. FRM/FRX externos, CodeNames, nomes/assinaturas e sete Delete herdados preservados.

Compilação e execução pelo VBE nativo do Excel16.0 build20527, Windows64bits. Menu Depurar > Compilar desabilitado após conferência na sessão fria, sem erro. [Registro](./_auditoria/fase7_20261102_025115Z/evidencias/compilacao-real-final.json) é transcrição da interface observada, combinada com inspeção do arquivo salvo; não screenshot persistido ou COM VBProject.

Testes somente em TEMP fora do OneDrive.51 arquivos sintéticos de execução foram arquivados com hash (71.305.646 bytes); existem backups adicionais em separado. A cópia no workspace é preservação de evidência, não local para executar Office. Nenhuma instalação, consulta externa, alteração de segurança, commit/push, migração real ou promoção. As duas sessões novas foram encerradas normalmente após confirmar um livro e caminho do mestre; a janela Excel preexistente8589906 não foi encerrada.

## Repetição, restauração e limites

Copiar mestre e sete insumos do arquivo de evidências para diretório TEMP novo contendo o marcador VBA_Fase7_20261102_025115Z; manter o nome Fixture_Arquitetura_Fase7_Mestre.xlsm e Menu Principal!Z1=AUDITORIA_VBA_FICTICIO. Não copiar resultados/exec/frio já existentes: as rotinas recusam sobrescrita. O mestre já contém o módulo F7; não importar novamente. Se necessário reconstruir fonte, usar [R4 embutido](./_auditoria/fase7_20261102_025115Z/prototipos/modPrototiposFase7_R4_embutido.bas), não os fontes R1–R3 preparatórios.

Executar F7_Leitores e registrar seus FAIL esperados; F7_RetestarFormulas é o reteste dirigido da montagem inicial. Executar F7_PrepararFrio, fechar normalmente a sessão exclusiva e abrir novo Excel /x para F7_VerificarFrio. Habilitar apenas o arquivo sintético revisado conforme aviso do Office, sem mudar políticas globais. Guardas usam caminho/marcador fixo; qualquer adaptação experimental requer novo checkpoint/compilação.

Para restauração documental, conferir manifesto inicial e copiar os correspondentes de backup/documentacao somente após preservar esta fase. Para experimentos, clonar novamente referência/insumos, nunca usar uma imagem forense inconsistente como estoque atual. Não houve mudança principal a reverter.

NÃO EXECUTADOS: migração real, todos os57 leitores adaptados em produto, eventos migrados, todos os formulários/43.232 fórmulas, escritor concorrente/lock/fencing, crash/energia/corrupção, Save recusado no protocolo definitivo e restauração autônoma por diário completo. Não houve medição de desempenho. O contrato e os ensaios não satisfazem aceitação funcional de produção.

## Decisão necessária

O [pacote de decisões](./PACOTE_DECISOES_FASE7.md) separa identidade/esquema, episódio/custódia, baixa conjunta, histórico/conciliação e autoridade/privacidade/retenção. O responsável deve aprovar escolhas concretas antes de qualquer Portão2. A recomendação técnica não é autorização implícita. Homologação continua impedida por A08/A09, pelos leitores/fórmulas principais não migrados e pela regressão operacional essencial pendente.
