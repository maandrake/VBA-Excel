# Pacote de decisões para o Portão 2 — Fase 8

**NÃO APROVADO.** Este documento prepara uma decisão concreta; nenhuma recomendação abaixo constitui autorização. O laboratório corrige quatro defeitos de protótipos, mas os cinco FAIL dos eventos principais F5 e a recuperação integral F6 permanecem pendentes. Sistema NÃO HOMOLOGADO.

## Escopo a decidir separadamente

| ID / problema | Alternativas e recomendação condicional | Riscos, consequências e dados existentes | Requisitos / critério de aceite | Reversibilidade |
|---|---|---|---|---|
| P2-01 Identidade (D01) | Manter posição/placa/NF-e ou adotar Vehicle/Episode/Movement/Projection/Operation imutáveis. Recomendar IDs internos + atributos documentais separados | ID novo não resolve identidade legada; união incorreta perpetua erro | Aprovar contrato, gerador/constraints/versões, proveniência; zero vínculo automático ambíguo; chave44 texto | Staging e mapa antes/depois; sem fatos novos restaurar conjunto; depois conciliar |
| P2-02 Episódio/custódia (D01-C) | Episódio global com mudanças de custódia versus novo episódio por local. Recomendar global na transferência/devolução interna, nova projeção e versão; reentrada após encerramento cria episódio novo | Semântica operacional ainda não aprovada; não confundir devolução interna com nova compra/retorno pós-venda | Responsável define eventos de abertura/encerramento/custódia e exceções; testar T/D/T e retorno do mesmo veículo | Mapa explícito de relações, sem fundir episódios antigos por placa |
| P2-03 Baixa (A08/D01) | A física, B limpeza, C lógica por ID. Recomendar C somente após M5–M7 | Leitor antigo pode exibir baixado ou atribuir outra linha; sete Delete atuais continuam | Aprovar regra de elegibilidade e bloqueio; adaptar leitores/fórmulas antes do corte; cinco FAIL F5 resolvidos nos eventos reais | Corte por pacote compatível; suspender escrita e conciliar após fatos novos |
| P2-04 Históricos (D02) | Edição por posição/placa versus vínculo explícito Movement/Episode/Projection e revisões. Recomendar vínculos + fatos substitutivos auditáveis | Números repetidos e chaves ausentes impedem vínculo automático; histórico não pode ser deduplicado silenciosamente | Aprovar correspondências comprovadas e tratamento de ambíguos; fórmulas anuais por episódio, erro em 0/múltiplos | Preservar origem e decisão de vínculo; retorno coordenado com movimentos posteriores |
| P2-05 Armazenamento (D04/D06) | A XLSM, B arquivos/diário separados, C banco local transacional. Recomendar C como alvo; motor/provedor ainda a selecionar e provar | Dependências, operação de backup, permissões e migração; nenhum benchmark de banco nesta fase | Aprovar motor/versão/interface/licença local, esquema e testes de commit/falha/restauração; instalação somente com autorização | Exportação consistente + backup do motor + schema; projeção reconstruível |
| P2-06 Autoridade (D06) | XLSM autoritativo ou armazenamento externo autoritativo. Recomendar autoridade única do banco se C aprovado e Excel como projeção | Commit externo e Save XLSM são fronteiras distintas; cópia divergente não pode comandar escrita | BaseID/revisão/linhagem/recibo; reconstrução e comparação completas; sem eleição pelo nome/data do arquivo | Congelar comando, preservar ambas fontes e conciliar antes de reconstruir |
| P2-07 Concorrência | Escritor único versus múltiplos escritores com protocolo próprio. Recomendar escritor único inicialmente + CAS de versão | OneDrive/cópias não são controle concorrente; dois escritores NÃO EXECUTADOS no laboratório | Aprovar dono/sessão/lock; provar segunda sessão recusada e lock órfão conciliado; não apagar lock por timeout sem prova | Desabilitar escrita concorrente; manter recibos e conciliar pendências |
| P2-08 Conciliação (D02/D06) | Repetição automática versus estado pendente e decisão rastreável. Recomendar bloqueio por episódio/operação e conciliação aprovada | Bloquear não restaura dados; pendência pode exigir intervenção e indisponibilidade | Aprovar responsável, escopo e roteiro; provar efeitos por ID, restauração/conclusão e replay seguro em cópia | Plano de conciliação versionado e recibo próprio, sem apagar evidências |
| P2-09 Logs/retenção (D04) | Diário mínimo local com ACL versus conteúdo fiscal integral. Recomendar mínimo técnico, sem XML integral/credenciais | Logs podem conter identificadores e valores; prazo e acesso ainda não definidos | Aprovar caminho, responsáveis, ACL, backups, retenção e descarte; nenhuma exclusão automática agora | Exportar/verificar trilha antes de mudança; descarte só sob política específica |
| P2-10 Limites de recuperação (A09) | Prometer rollback amplo versus fronteiras e garantias comprovadas. Recomendar compensação limitada verificada + pendência após fronteira incerta | Crash/energia/disco cheio/concorrência e recuperação integral ainda não comprovados | Aprovar RPO/RTO, matriz de falhas, teste em cada destino/baixa/Save/recibo; F6 FAIL não pode desaparecer por guarda PASS | Checkpoints coerentes; com fatos novos conciliar, nunca restaurar XLSM antigo cegamente |
| P2-11 Homologação/promoção | Integrar tudo de uma vez versus M0–M9 com portões. Recomendar candidato isolado por etapa e promoção separada | PASS experimental não valida gravadores fiscais/operacionais | Compilação real, eventos reais, fórmulas/identidade/backup/segurança; nenhum crítico sem decisão; responsável aprova evidências | Pacotes compatíveis, backup restaurado em ensaio; promoção revogável somente com conciliação dos fatos novos |

D03/cancelamento fiscal permanece independente: preservar comportamento atual, sem exigir novo evento fiscal ou alterar tributos. Este pacote não pede nem concede mudança de CFOP/CST/cBenef/ICMS/PIS/COFINS/tolerâncias/XML.

## Decisão de apresentação de datas

Escolher explicitamente ISO_LITERAL, BR_LITERAL estrita ou LEGACY_D14_LOCAL. Recomendar formato literal para campos cuja apresentação é exigida; aceitar local apenas onde a apresentação local foi aprovada. Tipo/serial/calendário e política são persistidos; formatos materiais não são ignorados. Não substituir o selo antigo do produto sem migração de contrato e regressão.

## Menor autorização futura proposta

Após revisar os [resultados](./RELATORIO_TECNICO_FASE8.md), aprovar nominalmente P2-01/P2-02/P2-04/P2-05/P2-06/P2-07/P2-08/P2-09/P2-10 e o esquema de M2, e autorizar somente M0–M4 em candidato sintético com staging e validação, ainda sem dados reais, sem mudança dos sete Delete e sem promoção. Se qualquer escolha permanecer pendente, registrar a pendência e limitar o candidato ao escopo independente. A adaptação dos gravadores e o corte M8 exigem escopo posterior explícito e evidências próprias.

Uma aprovação deve citar IDs escolhidos, alternativa, autoridade, responsável, ambiente e etapas autorizadas; “aprovar a fase” não deve ser interpretado como liberação de migração real ou produção. Este documento não solicita nova aprovação para concluir os trabalhos experimentais já autorizados.


## Evidências adicionais — Fase 9 (sem aprovação do Portão 2)

O [laboratório F9](./RELATORIO_TECNICO_FASE9.md) demonstrou ACE 16 instalado com ADO x64, transações, escritor controlado, recibos e reconstrução Excel. A recomendação de motor/provedor agora possui prova de ambiente, mas escolha/implantação/licença/autoridade no produto continuam pendentes P2-05/P2-06. Todos os itens P2-01 a P2-11 anteriores permanecem sem aprovação. [Decisões concretas F9](./DECISOES_PENDENTES_FASE9.md) e [plano M0–M4](./PLANO_MIGRACAO_M0_M4_FASE9.md).

Antes de integrar: consolidar schema/constraints e imutabilidade, tornar preparação/revogação concorrente segura, provar snapshot estável, verificador completo e ack persistente de projeção, falhas reais de commit/recuperação e restauração. O FAIL de Ignore do Sort e os FAIL oficiais anteriores não foram apagados. Não autoriza modificar gravadores, fórmulas, sete Delete, históricos ou dados reais.


## Adendo Fase10 — candidato experimental, sem aprovação do Portão2

O [pacote F10](./PACOTE_DECISAO_INTEGRACAO_FASE10.md) atualiza as evidências dos itens P2-01 a P2-11, sem alterar seu estado NÃO APROVADO. ACE16 já instalado foi exercitado com schema4 uniforme em base nova/migrada, transações, fencing e reconstrução de projeção. Isso não autoriza distribuição ou banco operacional.

A entrega contém121 asserções VBA únicas PASS,12 amostras benchmark PASS e10 evidências COM PASS; seis verificações frias VBA BLOQUEADAS. Permanecem limites de autoridade, ACKs, acesso SQL externo, conciliação e falhas físicas. Aprovação futura deve indicar itens P2, alternativa, versão, ambiente, responsável e etapas M0–M9. Nenhuma migração real, alteração de leitores/gravadores, Delete, fiscal ou promoção foi feita.
