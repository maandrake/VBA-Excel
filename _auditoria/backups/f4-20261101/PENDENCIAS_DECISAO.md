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
