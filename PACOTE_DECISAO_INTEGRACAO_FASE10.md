# Pacote de decisão para integração — Fase 10

**Portão 2 não autorizado. Sistema NÃO HOMOLOGADO.** Esta entrega é laboratório sintético. Os resultados experimentais não alteram V_VALIDO, T_VALIDO, D_VALIDO, D_FALHA_BAIXA, T_DUP_DESTINO, C_FALHA_BAIXA e C_FALHA_COMPENSACAO no produto.

## Recomendação técnica

Prosseguir com autoridade transacional externa e Excel como projeção reconstruível. ACE16 já instalado é candidato comprovado para este laboratório local de escritor único. A decisão de distribuição permanece pendente: verificar arquitetura, versão, suporte, instalação existente, direitos de uso e ambiente de destino. Não implantar automaticamente o esquema4 no sistema principal.

O Office local é x64, Excel16.0.20527.20112; ACEOLEDB.DLL16.0.20527.20016, ADO10.0, ADOX/DAO disponíveis. Nenhum provedor foi instalado. A designação ACE16 não significa que se recomenda o Redistributable2016: seu suporte terminou em14/10/2025 segundo a [Microsoft](https://www.microsoft.com/en-us/download/details.aspx?id=54920). O [Microsoft365 Access Runtime](https://support.microsoft.com/en-us/access/download-and-install-microsoft-365-access-runtime) inclui ADE; eventual implantação exige autorização e avaliação dos termos, arquitetura e coexistência no destino. Esta observação não certifica licenciamento de outro computador.

## Decisões nominais preservadas

|Item|Alternativa recomendada, sujeita a aprovação|Risco/requisito e critério de aceite|Reversão|
|---|---|---|---|
|P2-01 Identidade|GUIDs internos e documentos separados|Aprovar esquema/constraints, ID imutável no protocolo, proveniência e rejeição de ambiguidades; nenhuma união automática por placa/NF-e|Staging e mapa de correspondências; preservar origem|
|P2-02 Episódio/custódia|Episódio contínuo em transferência/devolução interna; reentrada pós-encerramento abre novo episódio|Definir casos operacionais e exceções; ensaio T/D/T não aprova a regra|Relações explícitas antes/depois, sem fundir legado|
|P2-03 Baixa|C lógica por identidade, somente após leitores e fórmulas|M5–M7 validados antes de M8; resolver cinco FAIL F5 nos eventos reais|Corte coordenado; com fatos novos exigir conciliação|
|P2-04 Históricos|Movimentos/projeções imutáveis por protocolo e retificações auditáveis|Não escolher primeiro duplicado; aprovar mapeamento e tratamento humano de ambíguos|Conservar fatos e decisão de vínculo|
|P2-05 Armazenamento|Avaliar ACE16 local como candidato, sem instalação nesta fase|Aprovar versão suportada no destino, contratos monetários, esquema4, migração, backup, ACL e limites SQL externos|Backup fechado verificado, exportação e schema versionado|
|P2-06 Autoridade|ACCDB único controlado; Excel reconstruível|Definir localização e identidade autoritativas; cópia do banco com mesmo BaseID não pode eleger-se autoridade; completar protocolo de ACK e exclusão de corrida de arquivo|Congelar escrita, preservar fontes divergentes e conciliar|
|P2-07 Concorrência|Escritor único, fencing e CAS dentro de cada transação|Provar todos os caminhos de comando; definir takeover humano e token antigo recusado; teste local não comprova isolamento serializável nem SQL hostil|Suspender sessões, preservar lease/intenção e recibos|
|P2-08 Conciliação|Pendência durável e decisão rastreável|Definir responsável, evidências e recibo de resolução; bloqueio não equivale a recuperar|Nunca apagar intenção/lease só por idade|
|P2-09 Logs/retenção|Mínimo técnico local, acesso restrito|Aprovar destino, ACL, responsáveis, conteúdo, retenção e backup; hashes sem segredo não autenticam administrador|Exportação verificada; nenhum descarte automático|
|P2-10 Recuperação|Rollback interno verificado; reconstrução da projeção após commit|Aprovar RPO/RTO, backup consistente e fronteiras ainda não ensaiadas; não prometer energia/crash/disco cheio/commit físico|Com fatos posteriores não restaurar cópia antiga cegamente|
|P2-11 Homologação|M0–M9 com aprovação separada por etapa|Integrar e testar gravadores/leitores reais, 17 abas, controles e fórmulas; resolver críticos e obter aprovação expressa|Pacote compatível e conciliação de fatos novos|

D03 fiscal/Cancelado continua independente, sem mudança de CFOP, CST, cBenef, ICMS, PIS, COFINS ou validação fiscal. Não é solicitada decisão fiscal implícita.

## Ensaio M0–M4 e próximos portões

|Etapa|Insumo e execução de laboratório|Saída exigida para eventual autorização|Estado|
|---|---|---|---|
|M0|Hash da entrega F9, 2.417 arquivos anteriores e checkpoint50 arquivos|Responsável, congelamento e backup restaurável|Ensaiado em sintéticos; congelamento operacional não autorizado|
|M1|Catálogo13 tabelas, 17 CodeNames herdados e fontes de laboratório; mapa57 leitores anterior preservado|Inventário completo de leitores/escritores, fórmulas e dependências do pacote operacional|Produto não migrado|
|M2|Schema4 em base nova e migrada, FKs compostas e contratos|Aprovar esquema e semântica de custódia; resolver limites de imutabilidade/auditoria/ACK|Candidato experimental|
|M3|Provenance separa origem, hash e decisão; dados fictícios e casos conhecidos|Correspondências comprovadas; ambíguos ficam pendentes humanos|Nenhum vínculo real atribuído|
|M4|Ciclo com GUIDs, dois episódios do mesmo veículo e vínculos explícitos|Mapa de IDs/históricos validado e reversível; conferir antes/depois|Somente sintético|

M5 leitores em sombra → M6 fórmulas/relatórios → M7 gravadores/recuperação → M8 política de baixa → M9 homologação/promoção continuam condicionados a aprovação nominal. Não adaptar apenas a baixa deixando leitores posicionais ativos.

## Limites a tratar antes de autorizar integração

Acesso SQL externo e alterações coordenadas de hashes; auditoria de toda recusa sem operação válida; identidade autoritativa de cópias do banco; imutabilidade e integridade completa dos ACKs; corrida de arquivo entre verificação e confirmação; migração interrompida; takeover de lease/intenção órfã; compatibilidade do ambiente de destino; backup quente; falha física de commit, crash e energia. O candidato reduz riscos no escopo ensaiado, sem encerrar essas pendências por mera contenção.

Uma autorização futura deve citar P2 aplicáveis, alternativa, versão do esquema/protocolo, ambiente, responsável e etapas. Esta entrega não executou nenhuma decisão do Portão2.
