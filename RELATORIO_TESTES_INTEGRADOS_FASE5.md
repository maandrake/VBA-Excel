# Relatório de testes integrados — Fase 5

**Resultado: aplicação NÃO HOMOLOGADA.** A compilação e a execução VBA reais foram mantidas, mas Venda, Transferência e Devolução falharam na preservação de fórmulas/associação após a baixa legada. A08 e A09 continuam críticos. Nenhum arquivo de produção foi alterado e nenhuma nova política de estoque/fiscal foi integrada.

Execução identificada por 20261101_202114Z, iniciada em 01/11/2026, continuação da Fase 4. [Matriz de regressão](./MATRIZ_REGRESSAO_FASE5.md), [mapa de integridade](./MAPA_INTEGRIDADE_ESTOQUE.md), [plano de recuperação](./PLANO_RECUPERACAO_FALHAS.md), [decisões](./PENDENCIAS_DECISAO.md).

## Base, preservação e integração

A fixture Fase 4 efetivamente existente foi localizada na área TEMP da execução anterior, com SHA-256 **dcbd72e46e8b182a9810254bcce460ea2cb17260b4bcb47b622ddc3d4f9879a0**, 55 componentes e 17 planilhas. Antes de editar houve checkpoint/manifesto; os 39 arquivos originais, 329 itens da referência Fase 4 e 11 backups iniciais foram conferidos. Nova verificação antes de atualizar esta documentação confirmou 39/39 originais, 329/329 itens da Fase 4 e 39/39 arquivos de desenvolvimento herdados intactos. [Evidência](./_auditoria/fase5_20261101_202114Z/evidencias/integridade-antes-documentacao-final.json).

Os testes ocorreram exclusivamente em cópias sintéticas descartáveis, independentes por cenário, no diretório TEMP VBA_Fase5_20261101_202114Z fora do OneDrive. As cópias no workspace são armazenamento de evidência, não local de execução do Office. Não houve instalação, mudança de Trust Center/AccessVBOM/ExecutionPolicy/Registro, consulta externa, Selenium, operação fiscal real, commit ou push. Apenas pastas de teste identificadas foram fechadas; nenhuma instância do usuário foi encerrada globalmente.

A [entrega de desenvolvimento sintética](./_auditoria/fase5_20261101_202114Z/entrega/Fixture_Auditoria_Fase5_Base_Compilada_NAO_HOMOLOGADA.xlsm) tem SHA-256 **f989289554a5c081cef8bd13de2e0b6f0e48ec4c1ba1dd5e75c349eff458f72c**, **56 componentes: 31 módulos padrão, duas classes, cinco formulários e 18 módulos de documento**, com 17 planilhas e CodeNames herdados. Acrescentou-se somente o [módulo de testes integrados](./_auditoria/fase5_20261101_202114Z/scripts/modTestesIntegradosFase5.bas), SHA **a8facf72c2c30f6bc9af45e39636160a1e39b66cfa74c5f9ab656f3fb8dbcbc4**, Windows-1252/CRLF. clsOperacaoPlanilha e modIntegridadeOperacao foram preservados. As 29 fixtures com execução/evidência preservadas foram inspecionadas em disco: todas com 56 componentes. Isso não significa 29 cenários de negócio aprovados.

O módulo foi importado pelo Editor VBA, a compilação real foi executada sem erro, e a base reaberta apresentou Depurar > Compilar VBAProject desabilitado na conferência final. A evidência de interface é transcrição da observação direta; a contagem veio da inspeção do arquivo salvo, não de acesso COM a VBProject. [Registro final](./_auditoria/fase5_20261101_202114Z/evidencias/compilacao-final-56.json), [inventário integrado](./_auditoria/fase5_20261101_202114Z/evidencias/inventario-integrado-final.json). Nenhuma sondagem COM bloqueada foi repetida.

Na preparação, a importação duplicada gerou modTestesIntegradosFase51. Foi removida somente essa duplicata de teste, e a base foi salva/compilada com 56 componentes. O registro preliminar compilacao-inicial.json com 57 foi preservado como **obsoleto**, acompanhado da correção acima; não representa a entrega. O checkpoint Fixture_Preparacao_Duplicada_NAO_USAR.xlsm é exclusivamente rastreabilidade, não base de teste.

A comparação exata dos fontes embutidos Fase 4/Fase 5 mostrou somente capitalização de identificadores MSXML pelo VBE e regeneração dos atributos VB_Base dos cinco formulários; os corpos de aplicação não receberam alteração semântica. [Diferenças](./_auditoria/fase5_20261101_202114Z/evidencias/comparacao-fontes-embutidos.json). Os cinco FRX exportados e todos os 39 arquivos de desenvolvimento continuam byte a byte iguais à Fase 4. Os streams de designer embutidos foram novamente serializados pelo Office e **não são byte a byte iguais**; não se afirma equivalência binária de todos os designers. Os controles exercitados e os relatórios reais abriram/resolveram suas dependências; uma regressão visual completa de todos os UserForms continua pendente. Referências/CodeNames foram herdados, não inventados ou renomeados.

## Resultados executados

| Grupo / modalidade | Resultado observado | Limite |
|---|---|---|
| Eventos reais de negócio, 26 cenários com antes/depois | 20 PASS, cinco FAIL, um BLOQUEADO para injeção de Save | 20 PASS = 15 rejeições, quatro compensações antes da baixa e Cancelado válido |
| XML/frmRelatorioXML reais, 13 variantes | 13 PASS de retorno esperado | Três aceitos, dez rejeitados; sem autenticidade/assinatura/autorização fiscal |
| AuditoriaFase2_Executar, VBA real reutilizado | 377 asserções PASS | Auxiliares, não eventos dos formulários |
| AuditoriaFase4_IntegridadeComplementar, VBA real reutilizado | 11 asserções PASS | Reconhece falha individual, sem prometer recuperação completa |
| Classe com outra pasta ativa, VBA real | PASS | Recupera alvo da fixture e preserva Pasta1; não é teste de todos os fluxos com duas pastas |
| AuditoriaFase3_CicloVida | BLOQUEADO | Depurador para no erro esperado de contexto já iniciado; preferência não modificada |
| Recuperação incompleta conhecida da Fase 4 | FAIL mantido | Não foi apagado/reclassificado pelo PASS complementar |

Asserções e cenários não foram somados para inflar cobertura. Não foram criados novos testes COM equivalentes ou simulações para substituir eventos reais. A matriz identifica lacunas e os componentes da Fase 4 não repetidos nesta fase.

## Falhas críticas demonstradas

**A08 — Venda e Transferência:** os destinos receberam o veículo esperado, mas Delete A:N em Entradas deslocou o veículo B para a linha 4. As referências diretas sintéticas Demonstração Q4/R4 viraram #REF!, a indireta S4 passou silenciosamente a exibir B e O4 permaneceu MATRIZ_A. Transferência conservou a chave no novo destino como String de 44 caracteres, formato @ e Ignore=True, além de M7/N7 e O7 sentinela; isso não compensa a falha de origem.

**A08 — Devolução:** as quatro gravações ocorreram, mas Delete A:M na filial deixou N4=#REF!, O4 do veículo A associado à linha de B e as referências Q5/R5/S5 inconsistentes. Mesmo comportamento legado concluído não atende integridade. As sentinelas são sintéticas e não se presume que Q:S do original tenham essas mesmas fórmulas.

**A09 — falha no Delete da Devolução:** proteção aplicada apenas na fixture, depois da pré-validação, provocou erro 1004. Os quatro destinos ficaram gravados e a origem permaneceu. O mecanismo informou conciliação manual e restaurou os estados de Application. Não houve nova tentativa. Não há compensação garantida após o marcador da baixa.

**Identidade/duplicidade:** Transferência aceitou 22001 em A4 e A7 de Entradas Filial, mesma placa, sem resolver a passagem. O teste reprova a prevenção de ambiguidade necessária à integridade; não estabelece que nNF isolado deva ser fiscalmente único. A escolha de uma nova regra depende de D01/D02.

São cinco cenários FAIL concentrados nesses riscos de A08/A09, e não cinco novos grupos independentes de auditoria. Nenhum dos sete Delete foi alterado.

## Recuperação e validação

Cada comparação monitora 7.560 células em nove abas, A1:U40, incluindo tipo/valor/fórmula/NumberFormat/Ignore/Locked/proteção. Quinze rejeições e quatro falhas intermediárias recuperadas produziram igualdade exata nesse conjunto. Em cada fluxo, a falha intermediária aconteceu depois de destinos escritos e durante gravação de histórico, por conversão ABC → número. A mensagem indicou restauração antes de Delete. D_FALHA_INTER teve uma segunda tentativa somente depois de igualdade verificada, novamente sem diferenças/duplicidades; os arquivos da primeira execução foram preservados.

O teste complementar de recuperação individual manteve o valor bloqueado enquanto outros snapshots foram recuperados; reconheceu o erro em vez de anunciar recuperação total. Depois de remover a proteção **de teste**, outra chamada recuperou a célula/formato. O FAIL amplo da Fase 4 continua documentado. Não há atomicidade, compensação garantida após Save ou recuperação depois de encerramento inesperado.

A tentativa de provocar Save com pasta indisponível foi bloqueada por EPERM ao renomear o diretório descartável. Nada foi movido; não houve alteração de ACL/segurança. O evento prosseguiu com Save normal. Portanto **falha real de Save não foi executada**, e o caso está BLOQUEADO, não PASS.

Os XML sintéticos exercitaram estrutura, múltiplos infNFe, chave, número, vazios, número inválido, overflow, vários itens, somatórios e tributos sem mudar regras fiscais ou tolerâncias. O relatório real exibiu três erros nos casos de valor inválido e um nas divergências de soma; uma exportação real foi preservada como texto Windows-1252. O nome .xlsx.txt evidencia uma apresentação de extensão a melhorar, sem alteração nesta fase. [Evidências XML](./_auditoria/fase5_20261101_202114Z/evidencias/cenarios/XML_SUITE/).

Cancelado gravou as duas saídas e manteve a entrada sem XML, conforme código atual; rejeitou número vazio/duplicado, valor e data inválidos. Esse PASS não comprova evento fiscal de cancelamento. D03 permanece aberta.

## Alterações e reversão

Não houve refatoração ampla ou correção de regra de negócio. Foram criados: módulo de teste com guardas de caminho/marcador, XML integralmente sintético, snapshots/TXT, comparação estruturada, inventário, quatro relatórios e atualização rastreável de documentos anteriores. A preparação duplicada foi corrigida no componente de teste. A regra local de ignorar fixtures XLSM da Fase 5 no Git foi acrescentada após backup; nenhum arquivo foi adicionado ao índice ou publicado.

Os [backups](./_auditoria/fase5_20261101_202114Z/backup/), [manifesto inicial](./_auditoria/fase5_20261101_202114Z/manifesto-inicial-fase5.json), [evidências preservadas](./_auditoria/fase5_20261101_202114Z/evidencias/preservacao-cenarios.json) e [manifesto final](./_auditoria/fase5_20261101_202114Z/manifesto-final-fase5.json) permitem conferir/restaurar a base sintética. A base limpa da entrega não contém os estados posteriores de testes destrutivos. As fixtures de cenário são forenses, não versão candidata à produção. Instruções no [plano de recuperação](./PLANO_RECUPERACAO_FALHAS.md).

## Próxima etapa necessária

D01/D02 precisam definir identidade de passagem, política de baixa e vínculo com históricos antes de integrar nova estratégia. A recomendação arquitetural de identidade estável continua condicional. Após decisão, implementar em incrementos pequenos e repetir os casos FAIL, todos os leitores de lacunas/linhas, falhas em cada fronteira e idempotência sob conciliação. D03 deve definir a comprovação fiscal de cancelamento; D04, destino/privacidade/retenção dos diagnósticos; D05, a execução do ciclo de vida e cobertura de dependências tardias/UI.

Permanecem não executados testes de falha real no Save, encerramento inesperado, todas as posições de escrita, alterações anuais ambíguas e nova regressão completa de entrada/pesquisa/backup/Ribbon. Resultados Fase 4 desses componentes são referência histórica, não PASS renovado. Nenhuma promoção, migração de dados reais ou homologação foi realizada.
