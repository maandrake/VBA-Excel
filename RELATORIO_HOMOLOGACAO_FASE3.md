# Relatório de homologação técnica — Fase 3

**Resultado: NÃO HOMOLOGADO.** Execução 20261031_213925Z, em 31/10/2026. Entrega de fontes e infraestrutura rastreável, com bloqueio de importação/compilação e riscos críticos A08/A09 ainda abertos. Nenhum XLSM original foi substituído ou promovido.

## O que está comprovado

- 39 originais/backups da Fase 1, 155 arquivos da execução anterior, 146 da Fase 2 e seus 46 backups íntegros. Novo backup de **69 arquivos** conferido antes das alterações. [Manifesto](./_auditoria/fase3_20261031_213925Z/manifesto-inicial-fase3.json).
- A cópia começou com os 42 arquivos da entrega da Fase 2, sem divergência. Somente uma classe do sistema recebeu dois ajustes técnicos. Os outros 33 fontes e os cinco FRX permanecem idênticos; assinaturas e 665 procedimentos exportados preservados. [Contratos estáticos: 12 PASS](./_auditoria/fase3_20261031_213925Z/verificacao-estatica-fase3.json).
- O XLSM base continua com 17 abas/50 componentes/648 procedimentos e sem incorporar os fontes revisados. A projeção de 52 componentes/666 procedimentos de desenvolvimento, ou 55 com três componentes de teste, não é uma contagem de importação realizada.
- Excel 16.0 build 20527, x64: acesso VBProject novamente recusado, **0x800A03EC**. Zero componentes importados, zero compilações, zero execuções VBA. [Erro original](./_auditoria/fase3_20261031_213925Z/office-fase3.json).
- Um esqueleto XLSM de 17 abas e um pacote de importação CP1252/CRLF com pares FRX foram realmente preparados fora do OneDrive. Não contêm dados reais; não são uma aplicação integrada. [Pacote e hashes](./_auditoria/fase3_20261031_213925Z/pacote-importacao.json), [ações manuais](./_auditoria/fase3_20261031_213925Z/COMPILACAO_MANUAL.md).

## Dependências

A inspeção do Registro foi somente leitura. Os caminhos declarados antigos de Office e Forms não existem naquele local, mas foram encontrados arquivos registrados de Office 2.8 em VFS e de Forms 2.0 em VFS/System. Portanto, ausência do caminho antigo não comprova biblioteca ausente. MSXML 6 e RegExp funcionaram por COM. Selenium possui registro e arquivos TLB, mas não foi instanciado e não houve navegador, login ou consulta externa.

O recorte de Common Controls não resolveu caminho de OCX; confirmar no VBE/host antes de concluir incompatibilidade. Nenhuma dessas leituras substitui References/IsBroken do projeto, abertura de designer ou teste de ativação ActiveX. [Registros reais](./_auditoria/fase3_20261031_213925Z/bibliotecas-registradas.json). Não instalar, registrar ou remover bibliotecas sem a decisão necessária.

A sondagem devolveu CodeNames vazios antes da inicialização do projeto no VBE. Os nomes desejados estão no esquema; eles não foram certificados/aplicados automaticamente. RowSources, nomes definidos, fórmulas operacionais, objetos e Ribbon continuam pendentes na fixture.

## A09 — revisão da infraestrutura e integração

Foram revistos integralmente classe, helper, quatro gravadores, Alterar e ExcluirFilial. As pré-conversões usam os mesmos CDbl/CDate do código original, antes das escritas; não foram criadas novas exigências fiscais. A chave preenchida permanece texto de 44 dígitos, com formato antes da atribuição. Origem é localizada antes da devolução filial e destinos são capturados antes das escritas.

Os destinos planejados correspondem à regra dos escritores: primeira vaga realmente Empty em cadastros/saídas, última linha da coluna A + 1 nos acréscimos de transferência/devolução. Essa correspondência foi examinada em fonte; não foi executada no evento real. O código ainda usa Activate/ActiveCell depois de ativar a aba proprietária; isso exige regressão de interface, especialmente porque eventos MSForms não são desativados por Application.EnableEvents.

A classe captura referência da célula, HasFormula, Formula, Value2 e NumberFormat. Nesta fase também captura Ignore do erro de número como texto. Compensa por célula, em ordem inversa; erro de uma célula é contado e não autoriza anunciar recuperação completa. Restaura os valores anteriores de Calculation/DisplayAlerts/ScreenUpdating/EnableEvents por tentativas independentes; erros desses setters são diagnósticos na Janela Imediata, sem garantia de restauração se o próprio Office recusar.

**C3-01:** Iniciar passou a recusar reutilização inclusive de contexto vazio confirmado/finalizado; Recuperar não informa sucesso para contexto inativo. **C3-02:** recuperar conteúdo e NumberFormat reseta Ignore, observado True → False em COM. RegistrarDestino/RestaurarCelula agora preservam esse metadado. Quatro casos COM corrigidos (vazio, fórmula, texto e chave) passaram. A suíte VBA existente foi ampliada, sem renomear eventos. [Blocos](./_auditoria/fase3_20261031_213925Z/alteracoes-fase3.json), [perda observada](./_auditoria/fase3_20261031_213925Z/a09-metadado-ignore-com.json), [regressão equivalente](./_auditoria/fase3_20261031_213925Z/a09-metadado-corrigido-com.json).

**Limites:** snapshot não é imagem integral da célula ou do workbook. Não captura fontes, bordas, larguras, validações, objetos, todos os metadados ou referências estruturais. Não restaura interface nem arquivo em disco. Recusa recuperação depois de iniciar Delete, Save ou confirmar; não sobrevive a encerramento inesperado. Células matriciais/derramadas/mescladas, proteção e pasta somente leitura possuem guardas em fonte, mas esses caminhos da classe não foram validados no runtime VBA. Não há ACID, diário persistente ou idempotência garantida.

### Falhas nas fronteiras

A extensão do teste COM reaproveita Snapshot/Restore e compara **200 células A4:T5 em cinco abas**, com valores, fórmulas, HasFormula, NumberFormat e os quatro estados de Application. É um recorte com sentinelas, não a totalidade dos 51/62 campos de formulário e fórmulas operacionais reais. As 200 células não são contadas como 200 testes.

Das 15 experiências, dez restauraram os campos comparados antes de Delete/Save. Quatro ficaram dependentes de conciliação: Delete recusado após a pré-validação, erro após Delete antes de Save, falha de SaveAs sem Delete e erro depois de Save. Na falha da própria compensação, 40 células da aba artificialmente protegida não puderam ser restauradas; outras tentativas continuaram. **A recuperação incompleta foi reproduzida e não classificada como integridade aprovada.**

Os estados de Application foram restaurados nos 15 ensaios; a outra pasta fictícia ficou intocada. Uma nova tentativa após compensação anterior à fronteira manteve um registro por destino no modelo. Não houve nova tentativa automática depois de estado parcial. [Evidência completa antes/depois](./_auditoria/fase3_20261031_213925Z/a09-fronteiras-com.json).

A falha de salvamento usa SaveAs para destino inexistente: demonstra erro de escrita e limite do modelo, mas não executa ThisWorkbook.Save da macro. Houve também Save COM bem-sucedido seguido de erro controlado; o hash do disco mudou e a compensação não foi tentada. O caso de falha do Save da rotina e a recusa pela classe VBA continuam BLOQUEADOS.

A listagem resumida no console PowerShell exibiu propriedades nulas por Select-Object sobre hashtable; o JSON completo possui os 15 registros e foi conferido diretamente. Não houve perda/reclassificação da evidência.

### Edição anual e exclusões relacionadas

CommandButtonAlterar_Click localiza Entradas, limpa o registro e usa a seleção anterior de Entradas Anual para limpar campos, sem localizar um vínculo histórico. CommandButton_ExcluirFilial_Click faz duas exclusões e procura o anual somente por número da nota. Ambas permanecem sem reescrita, pois não há identidade inequívoca aprovada.

A leitura local OOXML encontrou 18 grupos de valores repetidos na coluna A do histórico matriz e um no filial. A ausência de duplicidade nas entradas atuais não estabelece unicidade histórica/fiscal. Não presumir primeiro Find, placa ou número isolado como identidade. Estatísticas foram registradas sem exportar número, placa ou chave reais. [Evidência](./_auditoria/fase3_20261031_213925Z/identidade-existente.json). **A09 continua crítico.**

## A08 — comparação das estratégias

Quatro experiências cobriram três estratégias, com a exclusão física dividida em parcial/linha inteira. Todas alcançaram a observação esperada do ensaio; isso não representa aprovação de integridade da aplicação. [Resultados, fórmulas e hashes](./_auditoria/fase3_20261031_213925Z/a08-estrategias-com.json).

| Critério | A — exclusão física | B — limpar campos | C — identidade estável |
|---|---|---|---|
| Fórmulas | Parcial A:M quebrou N/O e dependência externa; linha inteira removeu sentinela lateral | Fórmulas M:O preservadas, mas dependência pode apontar para campo limpo | Fórmulas e identidade mantidas no protótipo sem mover linhas |
| Históricos | Valores do histórico fictício permaneceram; referência direta quebrou | Valor histórico ficou, mas referência passou a outro veículo após reutilizar | Histórico imutável ficou associado ao registro no modelo |
| NF-e/veículo | Deslocamento exige reconstrução correta das dependências | Reutilização trocou silenciosamente o veículo da referência | Chave fictícia não depende da linha; contrato real ainda indefinido |
| Pesquisas | Contiguidade pode ser mantida, com risco de referências | Pesquisa rápida atravessa lacunas; filtros legados que usam While exigem tratamento separado | Protótipo retorna dois ativos; pesquisa atual não adaptada retornaria três |
| Registros invisíveis | Risco de referência inválida/registro removido | Modelo While parou na lacuna; varredura completa encontrou dois | Evita ocultação somente com todos os leitores adaptados |
| Reutilização | Troca de posição exige corrigir vínculos | Não segura com referências físicas antigas | Precisa de novo episódio/versão e regra aprovada; não implementados |
| Compatibilidade | Próxima do fluxo atual, com defeitos demonstrados | Exige revisar edição, filtros e referências | Exige migração completa de leitores, identidade e persistência |
| Complexidade | Aparentemente baixa; reparo de dependências pode ser alto | Média/alta pelos vínculos e lacunas | Alta, com contrato e migração explícitos |
| Regressão | Alta, risco estrutural | Alta, risco silencioso de associação | Alta na migração; potencial de reduzir risco depois de validada |
| Recuperação | Não coberta pelo snapshot atual após Delete | Conteúdo compensável antes de salvar, mas vínculos precisam de política | Pode permitir conciliação/idempotência com registro persistente; protótipo é só memória |

A pesquisa rápida **já percorre lacunas** em MóduloPesquisa; o risco de parada na primeira célula vazia refere-se aos filtros legados de saídas/NCM. Esta precisão substitui a generalização dos relatórios anteriores.

C é a recomendação **condicional de arquitetura**, sem integração. O protótipo usa chaves fictícias existentes na sua coluna L e mapa em memória, rejeitando identidade ausente/duplicada. Não criou status/coluna de produto e não possui armazenamento durável. Não é possível aplicar esse identificador automaticamente ao sistema: o cadastro matriz/anual não grava a chave; somente 17 dos 42 registros lidos da filial têm chave com 44 dígitos armazenada como texto. O vínculo da saída matriz com a entrada filial também exige semântica confirmada. **A08 permanece crítica, com os sete Delete intactos.**

## XML, backup e decisões

**A14/A15:** 12 cenários MSXML/RegExp e modelo .NET passaram: estrutura com dois itens, malformação, nenhum/dois infNFe, chave/nNF inválidos/zero/longo, número inválido/vazio/overflow em segundo item e DTD. As somas fictícias de produtos/base/ICMS/PIS/COFINS corresponderam aos totais do caso válido no modelo, usando tolerância existente 0,01. O relatório INCORRETO foi modelado, não executado. XML fornecido é estrutural para os critérios testados, não uma NF-e completa fiscalmente autorizada. Nenhuma regra, tolerância, assinatura ou CFOP/CST/cBenef foi alterado. [Resultados](./_auditoria/fase3_20261031_213925Z/xml-fase3-com.json).

**Backup:** cinco cenários COM cobriram proprietário com outra pasta ativa, raiz inexistente criada, duas versões com colisão, restauração da versão anterior e atual, hashes, fórmula/formato/chave, pai inexistente, raiz como arquivo, falha de SaveCopyAs com arquivo travado e ausência de Save inesperado. Erros nativos foram preservados; arquivos de falha não apagados. Todas as condições esperadas passaram; a rotina VBA continua não executada. [Resultados e hashes](./_auditoria/fase3_20261031_213925Z/backup-fase3-com.json).

**A13:** comportamento Cancelado preservado; exige decisão sobre comprovação de evento e efeitos de estoque. **A20:** caminhos/logs legados mantidos; proposta de destino autorizado, minimização/mascaramento e retenção aguarda responsável. Nenhuma consulta, autenticação ou leitura de log real foi feita. [Decisões pendentes](./PENDENCIAS_DECISAO.md).

## Plano de conciliação e restauração

1. Após falha estrutural, Save, recuperação incompleta ou encerramento inesperado, bloquear a repetição operacional até conferir o estado. Não tentar um segundo Gravar como se a falha não tivesse produzido efeitos.
2. O responsável deve preservar separadamente o último arquivo em disco e a imagem de falha em memória, quando disponível, com hashes e contexto de etapa. Nesta execução isso só foi feito em fixtures; nenhuma ação sobre dados reais está autorizada por este relatório.
3. Relacionar origem, saídas, filial e históricos pelos vínculos aprovados; confrontar campos, fórmulas, chave como texto e quantidade de registros. Havendo mais de um candidato, não escolher o primeiro nem compensar automaticamente.
4. Depois de Delete, comparar também fórmulas/dependências fora dos destinos capturados. O snapshot por célula não desfaz referências quebradas ou trocadas. Depois de Save, conferir disco e memória; depois de crash não há snapshot persistente para usar.
5. Restaurar uma versão conhecida em nova cópia, sem sobrescrever a evidência de falha, e reconciliar operações legítimas posteriores com o responsável. A restauração COM das versões fictícias foi demonstrada; recuperação completa de uma operação fiscal não foi.
6. Executar regressão e obter liberação expressa antes de uso/promover. Não encerrar instâncias Excel do usuário nem presumir configurações Application anteriores depois de um crash.
7. Como evolução condicionada, propor registro persistente de etapa/identidade e tentativas idempotentes. Local, conteúdo, retenção e modelo de identidade devem ser aprovados e testados antes da implementação.

## Entrega e riscos

A [matriz](./MATRIZ_TESTES_FASE3.md) distingue PASS, FAIL de integridade, BLOQUEADO e NÃO EXECUTADO. Parser PowerShell e Problems são verificações auxiliares, nunca compilação. Os ensaios antigos não foram repetidos só para ampliar a contagem; suas evidências permanecem preservadas.

Riscos remanescentes: A08/A09, ambiguidade anual, falhas depois de fronteiras, compensação incompleta, contexto de interface, tipos de célula/arquiteturas não exercitados, ActiveX e referências não resolvidas, fixture integrada incompleta, execução dos eventos/validador/backup/Ribbon não comprovada e decisões A13/A20.

Apenas os fontes e o pacote estão preparados para a próxima etapa manual. **Nenhuma versão está aprovada para homologação ou produção.** Sem commit, push, publicação, instalação, mudança de segurança ou movimentação sobre dados reais.
