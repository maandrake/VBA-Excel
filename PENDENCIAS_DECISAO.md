# Atualização — Fase 5 (20261101_202114Z)

**Sistema NÃO HOMOLOGADO. Nenhuma decisão de negócio/fiscal foi presumida.** [Relatório integrado](./RELATORIO_TESTES_INTEGRADOS_FASE5.md), [matriz real](./MATRIZ_REGRESSAO_FASE5.md), [mapa de impacto](./MAPA_INTEGRIDADE_ESTOQUE.md), [plano de recuperação](./PLANO_RECUPERACAO_FALHAS.md).

| Decisão | Situação atual | Evidência / definição necessária |
|---|---|---|
| D01 — identidade e baixa / A08 | ABERTA, crítica | Eventos reais de Venda/Transferência/Devolução quebraram fórmulas e/ou associaram posição ao veículo seguinte. Transferência aceitou NF 22001 em duas posições da filial. Aprovar contrato da passagem, domínio de identificação, duplicidades, retorno, legado sem chave e política de baixa. Estratégia C continua proposta; sete Delete intactos |
| D02 — históricos e conciliação / A09 | ABERTA, crítica | Falha 1004 no Delete deixou quatro destinos gravados e origem existente; recuperação individual pode ser incompleta. Aprovar vínculo anual e responsável/protocolo de conciliação. Não repetir após Delete/Save ou recuperação incompleta sem decisão |
| D03 — Cancelado / A13 | ABERTA | Comportamento atual testado: duas saídas gravadas, entrada mantida e XML dispensado. PASS restrito de comportamento, sem comprovação de evento fiscal. Responsável fiscal precisa definir evidência e efeito operacional |
| D04 — logs / A20 | ABERTA | Nenhuma mudança nos caminhos operacionais/logs ou retenção. Evidências atuais são sintéticas e locais. Aprovar destino, minimização, mascaramento, acesso e retenção antes de diagnósticos reais; sem limpeza automática |
| D05 — compilação e ambiente | Compilação real PASS na base de 56 componentes; cobertura ainda incompleta | Sem alterações de segurança e sem repetição de COM VBProject. Ciclo de vida segue BLOQUEADO pela interrupção do depurador no erro esperado. A autorização genérica para usar Excel não foi tratada como aprovação de configuração. Falha real de Save não foi provocada por EPERM na preparação |

A correção dos erros de associação/baixa não foi implementada sem definição de regra. O FAIL de recuperação incompleta da Fase 4 foi mantido. Nenhuma nova coluna, status, identificador, migração real ou promoção foi autorizada. Os registros históricos abaixo estão preservados; estados antigos de compilação bloqueada não representam a base compilada atual.

---

# Pendências atuais — atualização da Fase 4 em 01/11/2026

**Sistema NÃO HOMOLOGADO. Nenhuma decisão D01–D04 foi presumida aprovada.** A importação e compilação foram concluídas por interface legítima, sem mudar segurança. [Relatório](./RELATORIO_INTEGRACAO_FASE4.md), [matriz](./MATRIZ_REGRESSAO_VBA_FASE4.md).

| Decisão | Situação atual | Ação específica ainda necessária |
|---|---|---|
| D01 — identidade e baixa / A08 | ABERTA; sete Delete preservados | Aprovar contrato de passagem/veículo/documento-item, legado, migração e leitores no [plano de identidade](./PLANO_IDENTIDADE_REGISTROS.md); não criar campos/colunas automaticamente |
| D02 — históricos, integridade e conciliação / A09 | ABERTA; risco crítico confirmado em VBA real | Definir vínculo inequívoco anual, responsabilidade e procedimento de conciliação. C4-04 evita apagar conteúdo adicional; recuperação com proteção ampla continua FAIL. PASS do auxiliar não valida os gravadores completos |
| D03 — cancelamento / A13 | ABERTA; comportamento preservado | Responsável fiscal define evidência de evento, vínculo e efeito operacional; sem modificação fiscal automática |
| D04 — logs / A20 | ABERTA | Aprovar destino, dados mínimos, mascaramento e retenção; não migrar/excluir logs nem consultar serviços externos |
| D05 — compilação e dependências | Importação/compilação real PASS na fixture; COM VBProject continua BLOQUEADO | Nenhuma mudança de segurança necessária para a via utilizada. Completar regressão funcional, dependências tardias/ActiveX e preferência do depurador para ciclo de vida |

**D05, preferência de depuração:** foi solicitada autorização para trocar temporariamente Interromper em módulo de classe por Interromper em erros não tratados, executar AuditoriaFase3_CicloVida e restaurar. Sem resposta registrada; configuração mantida. Não é alteração de Trust Center/AccessVBOM. A suíte não retornou PASS e permanece BLOQUEADA. Os outros testes independentes foram continuados.

Depois de Delete/Save, falha de recuperação ou encerramento inesperado, não repetir automaticamente a operação: preservar checkpoint/diagnóstico, localizar todos os registros candidatos sem escolher o primeiro, e conciliar sob responsabilidade aprovada. Nenhum diário persistente, idempotência comercial ou rollback garantido foi implantado. [FAIL protegido e correção](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/C4-04.json), [teste complementar](./_auditoria/fase4_20261031_230457Z/integracao_20261101_160244Z/testes-vba-complementares.json).

Os registros históricos abaixo permanecem para rastreabilidade; seus estados de importação bloqueada foram superados apenas na fixture desta Fase 4.

---

# Atualização — Fase 4 (20261031_230457Z)

As decisões D01–D05 abaixo continuam **sem aprovação**. [Relatório atual](./RELATORIO_INTEGRACAO_FASE4.md). D01/D02 receberam uma [especificação de identidade e migração](./PLANO_IDENTIDADE_REGISTROS.md), ainda proposta: não há esquema criado nem alteração dos sete Delete. D03/D04 permanecem sem mudança fiscal/logs.

D05 agora tem um [roteiro assistido único](./COMPILACAO_VBA_FASE4.md), com cópia TEMP exclusiva, componentes nominais, CodeNames, referências e três funções preparadas. A ação pendente é o operador importar e compilar pelo VBE e devolver os resultados; não foi solicitada mudança de segurança nem repetida sondagem VBProject. Nenhum resultado manual foi recebido nesta preparação. A ausência de resposta não é aprovação.

Registro anterior preservado integralmente abaixo.

---

# Pendências de decisão — Fase 3

Nenhuma decisão abaixo foi presumida. As permissões concedidas permitem análise, backup, testes fictícios e correções técnicas; não equivalem a aprovar uma nova política de baixa ou regra fiscal.

## D01 — identidade e política de baixa A08

**Evidência:** [experiências A/B/C](./_auditoria/fase3_20261031_213925Z/a08-estrategias-com.json), [identidade existente, somente estatísticas](./_auditoria/fase3_20261031_213925Z/identidade-existente.json). A08 continua crítica.

**Alternativas:** A mantém exclusão física e seu risco de #REF!/deslocamento; B limpa campos, preserva fórmulas, mas permite referência antiga apontar para outro veículo ao reutilizar a posição; C separa identidade de posição, exigindo identidade aprovada, leitores adaptados e registro persistente. Recomendação condicional: C.

**Definição solicitada ao responsável:** qual combinação identifica unicamente uma entrada e uma nova passagem do mesmo veículo pelo estoque? Número da nota isolado não é inequívoco no histórico. A chave de 44 dígitos não está disponível de forma completa nos cadastros/históricos atuais. Uma chave da saída matriz recebida pela filial também não deve ser confundida automaticamente com a identidade da entrada original da matriz.

Antes de integrar, aprovar expressamente: identidade/versão, ligação matriz–filial–histórico, tratamento de registros ambíguos/legados, manutenção ou exclusão do anual em cada operação, reutilização de posições e migração de leitores/fórmulas. Nenhuma coluna, status, tabela ou identificador novo foi criado no sistema. A pergunta sobre identidade foi enviada durante a execução e permanece sem resposta registrada.

## D02 — vínculo anual e conciliação A09

**Evidência:** CommandButtonAlterar_Click limpa a seleção anterior de Entradas Anual sem localizar o registro histórico; ExcluirFilial faz Find por número e executa duas exclusões em sequência. Existem valores repetidos na coluna A dos históricos. Não escolher automaticamente o primeiro resultado nem presumir placa como identidade da passagem pelo estoque.

**Decisão necessária:** aprovar o vínculo inequívoco e o destino histórico de cada edição/baixa. Depois de Delete/Save ou falha de compensação, aprovar procedimento de conciliação e responsável antes de repetir a operação. [Fronteiras observadas](./_auditoria/fase3_20261031_213925Z/a09-fronteiras-com.json): quatro cenários com conciliação necessária e um de recuperação incompleta. O plano de recuperação está no [relatório](./RELATORIO_HOMOLOGACAO_FASE3.md).

Um diário persistente/registro de operação e mecanismo de idempotência são propostas futuras, dependentes de identidade e armazenamento aprovados. Não foram implantados e não há transação ACID.

## D03 — Cancelado, A13

O ramo atual evita carregamento/validação XML, grava os dois históricos de saída e não baixa a entrada nesse ramo. Foi preservado. O responsável fiscal deve definir qual evidência de evento é exigida, sua ligação à NF-e/veículo e quando/como afeta o estoque. XML estruturalmente válido ou XML normal da NF-e não comprova evento de cancelamento. Nenhuma regra fiscal foi alterada nem operação fiscal real executada.

## D04 — logs e caminhos, A20

WriteLog/LogNFE mantêm caminhos operacionais legados, inclusive C:\Temp, e podem receber identificadores. Nenhum log real, autenticação ou consulta externa foi executado.

Aprovar destino autorizado, quais dados podem ser registrados, mascaramento e retenção. Proposta: códigos de operação/etapa/erro, sem XML completo, credenciais ou identificadores pessoais integrais; caminho configurável aprovado. Não migrar logs existentes nem excluir arquivos automaticamente. O destino/retencão do backup também depende de política; nenhuma limpeza foi aplicada.

## D05 — compilação e dependências

Acesso automatizado VBProject continua bloqueado, 0x800A03EC. A ação pendente é a importação/compilação **manual com as proteções atuais**, conforme [roteiro](./_auditoria/fase3_20261031_213925Z/COMPILACAO_MANUAL.md), e conferência real de referências/UserForms/ActiveX.

O registro parcial de Common Controls e a ausência de um caminho declarado não bastam para concluir biblioteca ausente. Office/MSForms têm caminhos registrados diferentes dos antigos. Não remover referências, instalar/registrar OCX ou mudar políticas sem autorização específica. Se a via manual também for impedida, registrar a mensagem e submeter a decisão ao responsável pelo ambiente.

**Situação:** sistema NÃO HOMOLOGADO; não promover ou substituir o XLSM original.


## Atualização — Fase 6 (execução 20261101_224949Z)

A08/A09 continuam críticos. Proposta C validada parcialmente em protótipos VBA, sem integração: 43 cenários únicos, 28 PASS e 15 FAIL; dois checks de guarda PASS separados. A/B quebram vínculos; C exige adaptação conjunta de leitores/fórmulas e tratamento de recuperação incompleta. Pesquisa legada retorna baixado, contagem física diverge e fórmula anual por placa confunde reentrada.

D01, D02, D03, D04 e D05 permanecem sem aprovação implícita. D05 teve compilação experimental real confirmada pelo VBE; isso não aprova migração/dependências futuras. Acrescentada D06: fonte autoritativa, diário durável, escritor único e cópias concorrentes. Alternativas, critérios e pacote concreto para aprovação estão em [DECISOES_ARQUITETURA_FASE6.md](./DECISOES_ARQUITETURA_FASE6.md). Nenhuma nova coluna, política de baixa, regra fiscal ou estrutura persistente entrou no produto.


## Atualização — Fase 7 (20261102_025115Z)

Portão 1: contrato, mapa dirigido de 57 procedimentos, plano reversível e protótipos VBA compilados. Portão 2 NÃO APROVADO: nenhum ID, leitor, gravador, diário ou baixa novo integrado ao principal. Sete Delete e regras fiscais preservados. Sistema NÃO HOMOLOGADO.

Resultados: 18 verificações de adaptadores PASS; 12 cenários frios/15 verificações PASS de contenção; quatro FAIL consolidados em formato reaberto, anual por placa, referência física e auxiliar com valor distinto. Recuperação integral F6 e cinco FAIL F5 continuam abertos. Não transformar recusa de replay em reparo dos dados. [Relatório](./RELATORIO_TECNICO_FASE7.md), [matriz](./MATRIZ_REGRESSAO_FASE7.md).

D01-I/D01-C/D01-B: aprovar separadamente IDs/esquema, semântica global ou local de episódio/custódia e corte conjunto de baixa/leitores/fórmulas. D02: vínculo histórico, correções versionadas e responsável/plano de conciliação. D04/D06: escolher autoridade/armazenamento, escritor/recuperação, dados mínimos, ACL, backup e retenção. D03 sem mudança fiscal. D05 compilação experimental real confirmada, sem alteração de segurança; cobertura de integração futura ainda pendente. Alternativas, risco, critérios e reversibilidade em [PACOTE_DECISOES_FASE7.md](./PACOTE_DECISOES_FASE7.md).


## Atualização Fase 8 — Portão 2 permanece NÃO autorizado

Os resultados experimentais e suas limitações estão no [relatório F8](./RELATORIO_TECNICO_FASE8.md). Nenhuma decisão D01–D06 foi implicitamente aprovada. O [pacote P2-01 a P2-11](./PACOTE_APROVACAO_PORTAO2.md) separa identidade, custódia, baixa, históricos, armazenamento, autoridade, concorrência, conciliação, logs/retenção, recuperação e homologação. Escolha de apresentação de datas também requer política explícita.

Quatro defeitos F7 têm provas novas de correção em laboratório, mas A08/A09 continuam críticos no produto. V_VALIDO, T_VALIDO, D_VALIDO, D_FALHA_BAIXA e T_DUP_DESTINO seguem FAIL F5; C_FALHA_BAIXA/C_FALHA_COMPENSACAO seguem FAIL integral F6. Guardas de repetição não restauram dados. D03 fiscal/Cancelado permanece preservada e pendente.

Recomendação técnica condicional: IDs estáveis/baixa lógica depois de leitores e fórmulas migrados; banco local autoritativo como alvo, motor/provedor e protocolo ainda a provar/aprovar; Excel reconstruível por revisão. Próximo escopo proposto M0–M4 em candidato sintético, somente depois de aprovação nominal. Não foi autorizado alterar os sete Delete, gravadores, estrutura operacional ou dados reais.


## Fase 9 — armazenamento e migração continuam experimentais

ACE 16/ADODB-ADOX x64 foi executado com transações, concorrência entre processos e sessão fria. Motor/provedor são candidatos de laboratório; P2-01 a P2-11 não aprovados. [Decisões F9](./DECISOES_PENDENTES_FASE9.md) preservam identidade/custódia, baixa, históricos, autoridade, armazenamento/licença, escritor, conciliação, logs/retenção, recuperação e homologação como decisões distintas. D03 fiscal/Cancelado permanece intacta. F5/F6 seguem FAIL; guardas e rollback ACE não validam os eventos antigos.
