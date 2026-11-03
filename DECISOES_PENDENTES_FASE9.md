# Decisões pendentes — Fase 9

**Portão 2 NÃO aprovado. Sistema NÃO HOMOLOGADO.** A recomendação ACE 16 é do laboratório; nenhuma instalação/implantação ou migração real foi autorizada.

| Itens P2 | Decisão concreta | Recomendação condicional e risco | Aceite e reversão |
|---|---|---|---|
|01 identidade | Gerador/escopo/imutabilidade/versões e legado sem evidência | IDs internos opacos com constraints e proveniência; GUID novo não resolve identidade histórica | Completar formato/colisões/relações; staging e mapa reversível |
|02 custódia | Episódio global ou por estabelecimento; fechamento/reentrada | Modelo ensaiado global mantém EP em T/D interna, nova projeção; retorno após venda cria EP | Responsável define semântica/exceções antes de migrar; não fundir legado por placa |
|03 baixa | Critério de ativo e substituição futura dos7 Delete | C lógica apenas depois de leitores/fórmulas/gravadores adaptados | Cinco FAIL F5 precisam eventos reais corrigidos; suspender corte diante de divergência |
|04 históricos | Vinculação/correção e imutabilidade | Movement/Episode/Projection explícitos, fatos substitutivos | Constraints completas e revisão; ambiguidades requerem decisão, sem primeiro Find |
|05 armazenamento | ACE 16/ADO candidato versus outro motor; licença/distribuição/suporte | ACE instalado demonstrou transações x64 sem instalar; não escolher definitivamente sem ambiente destino/EULA | Aprovar motor/versão/provider e testes de recuperação; backup consistente e schema versionado |
|06 autoridade | Banco único ou XLSM; recibos e projeções | SeC aprovado, banco autoritativo; Excel reconstruível. Nenhuma transação conjunta | BaseID/linhagem/revisão e ack de projeção; reconstrução validada; não eleger arquivo por data/nome |
|07 concorrência | Escritor único, revogação e órfãos | Owner+fence+CAS; sem TTL automático | Provar preparação/revogação concorrente, snapshot e segundo escritor em ambiente destino; ACL contra acesso direto |
|08 conciliação | Responsável e decisão de conclusão/aborto de intenção | Bloquear Pending mesmo com novo ID; verificar efeitos/recibos e aprovar plano | Não liberar órfão por relógio; recibo de conciliação separado e revisão reversível |
|09 logs/retenção | Caminho, dados, ACL, acesso, prazo e descarte | Mínimo técnico sem XML/credenciais; nenhum prazo inventado | Aprovar política; backups e descarte testados; sem exclusão automática nesta fase |
|10 recuperação | RPO/RTO e limites aceitos | Rollback demonstrado antes de commit; depois reconstruir projeção; resultado incerto exige conciliação | Falha física de commit/energia/crash/rollback incompleto ainda NÃO EXECUTADA; não prometer ACID Excel+banco |
|11 homologação | Escopo e etapas autorizadas | Candidato por etapas e promoção separada | Eventos reais e críticos resolvidos; arquivo operacional nunca substituído por aprovação genérica |

D03/Cancelado e regras fiscais continuam preservados; a semântica interna não cria nova exigência fiscal. O laboratório não resolveu V_VALIDO/T_VALIDO/D_VALIDO/D_FALHA_BAIXA/T_DUP_DESTINO/C_FALHA_BAIXA/C_FALHA_COMPENSACAO.

Antes de pedir integração de gravadores, completar as lacunas do esquema e da concorrência descritas no [relatório](./RELATORIO_TECNICO_FASE9.md). O próximo escopo tecnicamente recomendável é consolidar o candidato sintético e a versão do esquema, com testes de fronteira específicos; M7/M8 e promoção permanecem vedados. Aprovação deve citar P2, alternativa, ambiente, responsável e etapas; não foi solicitada aprovação para continuar experimentos já autorizados.


## Adendo Fase10 — 02/11/2026

Candidato experimental schema4,121 asserções VBA únicas PASS,12 amostras benchmark PASS e10 experiências/verificações COM PASS, com modalidades separadas. A sessão fria VBA foi BLOQUEADA pelo Office; não houve alteração de segurança. Novas guardas corrigiram a aceitação de intenções incompletas e divergência de custódia, mantendo as reproduções FAIL.

Permanecem pendentes: fonte autoritativa de cópias divergentes; ACK íntegro e corrida do arquivo; imutabilidade contra SQL externo; auditoria durável de toda recusa e retificações cadastrais; lease órfão/recibo de conciliação; backup quente; falhas físicas; migração interrompida; ambiente de destino. Nenhuma decisão P2 foi aprovada. Ver [relatório F10](./RELATORIO_TECNICO_FASE10.md) e [pacote nominal](./PACOTE_DECISAO_INTEGRACAO_FASE10.md). Os sete FAIL históricos do produto e sete Delete continuam preservados.
