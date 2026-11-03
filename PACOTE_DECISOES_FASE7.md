# Pacote de decisões — Fase 7

**Nenhuma decisão do Portão 2 foi aprovada.** Este pacote define o escopo revisável para uma autorização futura; aprovação de arquitetura não promove o XLSM e não autoriza mudança fiscal. A aplicação permanece NÃO HOMOLOGADA.

## D01-I — identidade e armazenamento do esquema

Problema: placa/NF-e/linha não distinguem passagens e históricos. Alternativas: continuar dependência física; chave composta documental sujeita a ausências/repetição; IDs internos imutáveis com FKs e evidência de identidade física. Recomendação: o [contrato](./CONTRATO_IDENTIDADE_FASE7.md) separa VehicleID, EpisodeID, MovementID, ProjectionID, OperationID e AttemptID/Revision, sem inferir identidade fiscal. Justificativa: reentrada e duplicidades exigem identidade de fato separada do veículo/documento.

Impacto: registros novos recebem IDs, legados exigem mapa de correspondências e podem ficar pendentes. Risco: GUIDs podem formalizar vínculos errados se a migração presumir coincidência de placa/nota. Testes: ausência/duplicidade/colisão, reentrada, histórico e seleção obsoleta. Critério de aprovação: aprovar semântica, esquema versionado e localização junto com D06, tratamento humano de ambiguidades e invariantes. Reversibilidade: mapa/checkpoint antes de novos fatos; depois, conciliação dos fatos novos.

## D01-C — episódio global ou por estabelecimento

Problema: transferência/devolução podem significar continuidade de custódia ou novo episódio local. Alternativas: um EpisodeID global e projeções de custódia sucessivas; episódios por estabelecimento ligados por relação explícita. Recomendação: episódio global para a passagem, preservado em transferência/devolução interna; nova entrada após venda cria episódio novo. É decisão operacional específica, não resultado automático do protótipo.

Impacto: regras de origem atual, contagem matriz/filial e vínculo com histórico. Risco: confundir devolução filial com retorno de cliente ou duplicar estoque global. Testes: T/D/T, retorno após venda e uma única custódia atual. Aprovação: confirmar expressamente o significado de episódio/custódia e quais operações encerram a passagem. Reversão: ensaio em cópia e mapa de relações; depois de novos fatos, exige remapeamento aprovado.

## D01-B — baixa lógica e migração conjunta dos leitores

Problema: Delete quebra/desloca e limpeza/reuso reatribui identidade; leitores antigos continuam exibindo baixados e fórmulas físicas trocam de veículo. Alternativas: A exclusão, B limpeza, C baixa lógica com identidade e leitores/fórmulas migrados. Recomendação condicional: C com estoque elegível separado da consulta histórica, sem reutilizar projeção.

Impacto: 57 procedimentos inventariados e dependências das 17 abas, mais revisão final de fórmulas/UDF/nomes/listas/OnAction. Risco: implantação parcial mantém corrupção silenciosa. Testes: cinco FAIL F5 em eventos reais, todos os leitores, reentrada, reordenação, auxiliares, contagens e histórico. Aprovação: autorizar candidato integrado somente após M0–M7 do [plano](./PLANO_MIGRACAO_REVERSIVEL_FASE7.md), com sete pontos de baixa tratados no corte conjunto. Reversibilidade: nenhum principal alterado nesta fase; futuro corte segue checkpoints completos, sem retorno cego após novas movimentações.

## D02 — histórico, correções e conciliação

Problema: Find por nota/ActiveCell anual e falha parcial podem atingir histórico errado. Alternativas: editar primeiro candidato; exigir vínculo único e correção versionada; bloqueio/manual para legado sem evidência. Recomendação: vínculo por MovementID/EpisodeID/ProjectionID, revisão preservada e conciliação por responsável designado.

Impacto: edição/exclusão passam a exigir identidade e versão; registros ambíguos não são apagados nem escondidos. Risco: indisponibilidade operacional até conciliar. Testes: históricos duplicados, pós-Save alterado, falha na baixa/compensação e nova tentativa. Aprovação: definir responsável, evidências aceitas, plano de reparo, autoridade de liberação e trilha mínima. Reversibilidade: cada reparo com pré-imagem/checkpoint e efeitos exatos; não prometer reconstrução automática de Delete.

## D04/D06 — autoridade, diário, privacidade e retenção

Problema: Excel e diário separado podem divergir; snapshots em memória não sobrevivem à interrupção. Alternativas A/B/C comparadas no [protocolo](./PROTOCOLO_PERSISTENCIA_CONCILIACAO_FASE7.md). Recomendação condicional: banco local aprovado como autoridade de fatos/IDs e Excel como projeção, escritor único inicialmente. Se escolher arquivo separado, aprovar explicitamente limites e protocolo antes de uso. Não existe transação conjunta automática banco/XLSM.

Impacto: dependência de armazenamento, backup coordenado, exclusividade/fencing, recibos de revisão e reconstrução. Risco: fonte dupla, cópia antiga, perda/truncamento de diário e dados expostos. Testes: sessão fria, lock/concorrência real, Save recusado, corrupção/truncamento, perda de processo/energia sob plano seguro, restauração e reconstrução. Parte desses testes permanece NÃO EXECUTADA.

Aprovação: escolher A/B/C; definir caminho/tecnologia sem instalação implícita, BaseID/autoridade, escritor, protocolo de recuperação, dados mínimos, ACL, responsáveis, backup, retenção e privacidade. Nenhuma limpeza automática de logs/backups é proposta como já autorizada. Reversibilidade: conjunto consistente autoridade/diário/projeções/mapa; logs sintéticos atuais não migram dados operacionais.

## D03 e D05 — fiscalidade e ambiente

D03: manter Cancelado como hoje. Alternativas fiscais só pelo responsável fiscal, com evidência/documentos e testes específicos. Impacto/risco: mudar exigência documental ou estoque sem decisão altera negócio; reversão exige análise fiscal. Nenhuma exigência nova foi criada.

D05: VBE nativo permite compilar os protótipos sem AccessVBOM. Manter via legítima e dependências atuais; nenhuma alteração de Trust Center/ExecutionPolicy/Registro. Novo candidato integrado exige compilação real e regressão dos eventos/UserForms. Formatos de datas/selo precisam de contrato persistente estável; o FAIL novo de reabertura não foi encerrado no produto. Preferência de depuração histórica e cobertura de outras arquiteturas continuam pendentes; não mudar configuração para obter PASS.

## Autorizações concretas para uma próxima etapa

A decisão do responsável deve indicar separadamente: (1) identidade/esquema; (2) semântica de episódio/custódia; (3) armazenamento autoritativo e protocolo; (4) migração conjunta dos leitores/fórmulas/gravadores e baixa em candidato isolado; (5) critérios e responsável pela conciliação do legado. A promoção para produção exige outra aprovação após regressão. Não interpretar silêncio ou aprovação genérica de Excel como aceite desses itens.
