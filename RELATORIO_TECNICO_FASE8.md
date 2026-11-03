# Relatório técnico — Fase 8

**Portão 1 concluído como preparação experimental. Portão 2 NÃO executado. Sistema principal NÃO HOMOLOGADO.** Os quatro defeitos de protótipos F7 foram reproduzidos e possuem novas evidências de correção no modelo sintético. Isso não resolve os cinco FAIL dos gravadores principais F5, nem demonstra recuperação integral após baixa/Save/crash.

Execução: 02/11/2026, [_auditoria/fase8_20261102_121422Z](./_auditoria/fase8_20261102_121422Z/). Office executado exclusivamente no TEMP, fora do OneDrive. Leitura obrigatória integral concluída; nenhum dos 12 documentos distintos exigidos estava ausente. Relatórios e evidências F7 existem e foram conferidos, sem reiniciar a auditoria geral.

## Preservação e integração experimental

2.106 hashes da base anterior conferidos na abertura, sem divergências. Manifesto F7 b812ac512988b2956fcf873de0e8783569896b7a3b55ee9e74492efd9b185bb4 validado contra seu digest. Manifesto F6 627b46c182776f340fc850641fabca0b924aba6f79ecb29ce10c934ba92e5de4 também conferido. 35 backups novos de documentos/base F7/fontes foram copiados e todos os hashes rechecados. [Manifesto inicial](./_auditoria/fase8_20261102_121422Z/manifesto-inicial-fase8.json), [backup](./_auditoria/fase8_20261102_121422Z/backup/) e [checkpoints](./_auditoria/fase8_20261102_121422Z/checkpoint/).

Original XLSM preservado: 2ad1c635da68b51a68d5d2c0d3cd638f1ff8064942b299699f2594d7a9ca8850. 39 arquivos originais e checkpoints/evidências anteriores mantidos. Os arquivos externos de formulários/FRX não foram editados. Os 58 componentes herdados da F7 foram comparados com o mestre final preservando literais; nenhuma diferença além de atributos gerados/capitalização externa a strings/espaços finais. Os sete Delete operacionais permanecem: 3 em Cad_NFE e 4 em Cad_NFE_Saida. Há mais 2 Delete em componentes de teste herdados; não confundir contagem textual total 9 com os sete operacionais.

Dois componentes novos somente no laboratório: clsRepositorioFase8 e modExperimentosFase8. Mestre: 60 componentes/17 planilhas; fixture integrada:67/24, devido às sete abas experimentais _F8_*. A identidade e estrutura operacional principal não foram alteradas. CodeNames e assinaturas herdados preservados. A serialização interna de designers pelo Office não é anunciada como igualdade binária FRX.

Entrega final compilada: [Fixture_Integridade_Fase8_R2_EXPERIMENTAL_NAO_HOMOLOGADA.xlsm](./_auditoria/fase8_20261102_121422Z/entrega/Fixture_Integridade_Fase8_R2_EXPERIMENTAL_NAO_HOMOLOGADA.xlsm). SHA-256 **a64b1a81bfba47000d7f69f353ef7fa697746f78bded65f53685c5dee4a0fa3d**. [Fontes finais importáveis e textos efetivamente embutidos](./_auditoria/fase8_20261102_121422Z/prototipos/revisao_final/). A entrega intermediária anterior continua arquivada e não deve ser confundida com R2.

## Execução comprovada

Excel 16.0, build 20527, Windows/Office 64 bits. Importação pelo Editor VBA nativo; Depurar > Compilar executado sem erro após retirar duas importações experimentais duplicadas. Menu Compilar desabilitado ao final e execução efetiva das macros. Não houve sondagem VBProject COM nem modificação de AccessVBOM, Trust Center, ExecutionPolicy ou Registro. A segunda sessão inicialmente recusou macros; o mesmo arquivo sintético foi reaberto e habilitado no aviso específico, sem configuração global.

Três processos exclusivos de laboratório identificados: 10488, 51860 e 49544. Foram encerrados normalmente, após confirmar apenas 1 workbook de teste aberto. Processo preexistente 42180/janela 8589906 preservado. Estados finais observados em cada sessão: EnableEvents=True, ScreenUpdating=True, Calculation=-4105, DisplayAlerts=True. [Registro de interface e processos](./_auditoria/fase8_20261102_121422Z/sessao-excel-fase8.json); contém transcrição identificada da janela Imediata, não extração COM do projeto.

70 asserções únicas com PASS na seleção de resultados mais recente, mais 4 volumes de benchmark PASS separados. Os 26 integrados foram repetidos depois da última correção compartilhada; não contam como cenários novos. Os 15 guardas F7 foram reexecutados por hipótese de regressão fria. As contagens não incluem preservação de hashes/compilação como testes funcionais adicionais. Todos os FAIL históricos e falhas intermediárias permanecem nas evidências. [Matriz completa](./MATRIZ_REGRESSAO_FASE8.md) e [resultados estruturados](./_auditoria/fase8_20261102_121422Z/resultados-consolidados-fase8.json).

## Resultado dos quatro defeitos e guarda adicional

| Problema | Nova prova no laboratório | Situação do produto |
|---|---|---|
| F8-01 formato de data | Canônico de Double/serial1900 em hexadecimal + política de apresentação explícita; 12 comparações após nova instância PASS sem restaurar formatos; formatos materiais e datas textuais recusados | Selo antigo não foi migrado |
| F8-02 anual por placa | EP_A1 vendido=SIM e reentrada EP_A2=NAO; ausente/duplicado erro; T/D/T e histórico por EP preservados | Fórmulas anuais legadas não foram substituídas |
| F8-03 INDIRECT físico | Resolver por PID/EP/VID/versão, índice invalidado, token revalidado; Sort, lacuna, linha12 e reuso com novo ocupante mantêm associação | Referências físicas principais continuam pendentes |
| F8-04 auxiliar incorreto | A2 margem 440/aux 6000, B 6600, C 15500 após sort; identidade e valores discriminantes aprovados | Auxiliares reais M/N/O ainda exigem M6/regressão |
| F8-05 selo contendo erro | 3 FAIL reais antes: data textual, formato 0.00 e chave curta aceitos por erro coincidente; guarda explícita recusa e 3 PASS depois; 26 integrados novamente PASS | Componente experimental somente; sem alteração fiscal |

Ver [causas, correções e reversão](./CORRECOES_PROTOTIPOS_FASE8.md). O Sort nativo removeu Ignore e foi envolvido por preservação explícita do metadado por ID, sem refazer selos; primeira execução falha preservada. Benchmark usava Ignore multicélula e falhou; corrigido por célula. O diagnóstico de selo teve conflito de nome entre log e snapshot (erro 58) após asserções; nomes separados e evidência conservada. Um rótulo de formato 0.00 foi convertido em0 na coluna auxiliar de datas; o canônico D7 e o formato atual guardam a prova correta. Limitação registrada, sem editar resultados para aparentar sucesso.

## Desempenho medido e decisão de leitura

100/1000/5000/10000 registros concluíram e foram reabertos em novo processo; quantidade, ID final, String44 e Ignore preservados. Em 10 mil: 200 consultas verificadas em 0,8083 s; listagem validada em 9,7135 s; atribuição ao ListBox real em 0,0283 s; escrita da projeção em 0,2715 s; Save em 1,0629 s; total de 24,1331 s. Working set antes/depois de 597,3/778,7 MiB, sem alegação de pico ou teste prolongado.

200 consultas lineares em array: 0,0978 s; 200 consultas ao índice candidato: 0,0131 s, em 10 mil registros; em volumes menores o overhead da consulta de revisão pode superar a busca linear. O índice apenas localiza; consulta autorizada ainda verifica unicidade/identidade/conteúdo. Não prometer O(1) integral nem usar o tempo do candidato como tempo de autorização. Listagem varre todos os selos; recomenda-se futura validação por lote/revisão com invalidação confiável e projeções delimitadas, sem relaxar identidade.

Recálculo medido em amostra máxima de 200 fórmulas nativas, não no catálogo operacional inteiro. UDFs por ID com faixa delimitada provam associação no caso pequeno; para massa, preferir projeção em array e fórmulas nativas com unicidade comprovada. A orientação oficial recomenda reduzir transferências célula/VBA e evitar obstruções de recálculo: [Microsoft — desempenho Excel](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions). As medições desta máquina são a evidência quantitativa; a recomendação de migração não extrapola esses números para produção. Benchmarks antecedem a guarda final de selo; não medem exatamente seu custo constante adicional.

## Persistência, recuperação e decisões

A tentativa de Save foi realmente cancelada pelo evento exclusivo da fixture: Saved=False, sem recibo, nova operação bloqueada. Trata-se de contenção em memória; não prova intenção durável em disco. Guardas frias de diário órfão/truncado, base/histórico divergente, pós-Save, payload/versão e pendência passaram sem escrita de negócio. Não recuperaram os dados nem corrigiram C_FALHA_BAIXA/C_FALHA_COMPENSACAO da F6.

Recomendação condicional: identidade estável + baixa lógica somente após migração completa dos leitores/fórmulas e aprovação; como alvo de armazenamento, banco local transacional autoritativo e Excel como projeção reconstruível. Motor, provedor, privacidade, concorrência e durabilidade precisam aprovação e provas próprias. Nenhuma transação conjunta banco/XLSM foi demonstrada. [Alternativas](./PROPOSTA_ARMAZENAMENTO_FASE8.md), [contrato de persistência e conciliação](./CONTRATO_PERSISTENCIA_FASE8.md), [M0–M9 e reversão](./PLANO_IMPLANTACAO_FASE8.md) e [pacote formal do Portão 2](./PACOTE_APROVACAO_PORTAO2.md).

Pendências essenciais: cinco FAIL F5 (V_VALIDO/T_VALIDO/D_VALIDO/D_FALHA_BAIXA/T_DUP_DESTINO), recuperação integral F6, semântica de custódia, vínculos ambíguos do legado, decisão de armazenamento/autoridade, protocolo durável completo, concorrência, adaptação dos 57 procedimentos mapeados e catálogo de fórmulas, eventos reais e homologação. Crash, energia, disco cheio, dois escritores, outro idioma/região e benchmark de banco NÃO EXECUTADOS. Nenhuma regra fiscal foi alterada.

## Restauração e próxima etapa

Originais não precisam ser restaurados porque não foram modificados. Para retomar o laboratório, copiar a revisão final e insumos para TEMP, conferir o manifesto e seguir as guardas de nome/caminho/marcador antes de compilar e executar em novas cópias. Não abrir arquivos de evidência do OneDrive para executar Office. Documentos anteriores podem ser restaurados dos backups após SHA-256; conservar o histórico e evidências. Não restaurar cegamente uma base antiga após fatos novos.

Próxima etapa recomendada: decisões nominais do pacote e prova do armazenamento/esquema em candidato sintético, com M0–M4 condicionados à autorização específica. M7/M8/promoção não estão liberados. A entrega é preparação arquitetural testada, sem distribuição ou substituição operacional.

## Conferência final do pacote

Os 39 originais e os 35 backups passaram na rechecagem SHA-256. Dos 2.106 caminhos da linha de base, 2.100 permanecem iguais; somente cinco documentos de acompanhamento e o .gitignore receberam acréscimos autorizados, com prefixos originais integralmente preservados. Nenhum fonte operacional foi modificado. O [manifesto final](./_auditoria/fase8_20261102_121422Z/manifesto-final-fase8.json), seu [digest SHA-256](./_auditoria/fase8_20261102_121422Z/manifesto-final-fase8.sha256) e o [verificador somente leitura](./_auditoria/fase8_20261102_121422Z/scripts/Verificar-Fase8.ps1) permitem conferir novamente o conjunto.

O verificador PowerShell auxiliar foi tentado e ficou **BLOQUEADO** pela política existente: PSSecurityException/UnauthorizedAccess, “a execução de scripts foi desabilitada neste sistema”. Nenhuma política foi alterada e não houve bypass. A conferência independente de SHA-256 em Node permanece a evidência de integridade; não se trata de teste VBA. [Registro do bloqueio](./_auditoria/fase8_20261102_121422Z/evidencias/verificador-powershell-bloqueado.json).
