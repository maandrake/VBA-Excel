# Protocolo de persistência e conciliação — Fase 7

Proposta para D04/D06; não implantada. O motor F6 e os eventos principais continuam com compensação limitada. Guardas experimentais F7 recusam estados incoerentes; recusar não recupera os dados.

## Autoridade e alternativas

| Requisito | A — diário dentro do XLSM | B — arquivo local separado | C — banco local aprovado |
|---|---|---|---|
| Autoridade / IDs | XLSM único identificado por BaseID e revisão; cópias antigas não podem operar | Definir explicitamente se XLSM ou repositório de fatos é autoridade. TSV de log não vira autoridade por existir | Recomenda-se banco autoritativo de fatos/IDs e Excel como projeção regenerável, se infraestrutura for aprovada |
| Escritor / concorrência | Um escritor; arquivo aberto/read-only não garante trava de negócio entre cópias | Lock exclusivo de base + token de fencing, proprietário e revisão; não inferir autoridade só pela existência de lockfile | Transação e restrição única no banco; escritor controlado para materialização. Exige driver/arquitetura/permissões compatíveis |
| Etapas / atomicidade | Save inclui abas/diário no mesmo pacote; não há write-ahead externo nem garantia de resistência a crash | Registro preparado antes do Save e recibo após conferência; dois arquivos sem commit conjunto | Atomicidade somente dos fatos/diário no banco; Save do XLSM permanece separado. Recibo de projeção/revisão permite detectar atraso |
| Save / falha | Erro torna persistência incerta; reabrir cópia isolada e comparar revisão/efeitos | Janela entre Save e recibo é obrigatoriamente PENDENTE até conciliar; não escrever CONCLUIDA só porque Save retornou | Commit dos fatos é ponto de autoridade; projeção Excel falha pode ser reconstruída conforme contrato aprovado; não executar novamente fato para consertar projeção |
| Divergência / órfão / truncado | Cópia antiga/zip corrompido exige checkpoint e conciliação | BaseID/revisão/schema/seq/tamanho/checksum/recibo e plano de efeitos; parser rejeita incompleto. Arquivo órfão não é descartado ou aplicado automaticamente | FK/unique e integridade do banco; arquivo Excel de outra revisão fica em quarentena. Backups divergentes também precisam de base/revisão |
| Hash / privacidade | Hash de bytes em manifesto externo, conteúdo mínimo e acesso restrito | Envelopes canônicos versionados, hash de conteúdo e cadeia verificada. Hash detecta alteração, não autentica autor; não protege de atacante que reescreve ambos | Hash/checksum e trilha com autorização; minimização e ACL aprovadas, sem presumir criptografia/segurança só por usar banco |
| Backup / restauração | Checkpoint XLSM fechado e manifesto; não sobrescrever base ao abrir versão antiga | Snapshot coordenado XLSM + diário + recibos + mapa de revisão. Restaurar conjunto, nunca componente isolado como atual | Backup consistente do banco + versão/protocolo e XLSM de projeção compatível; testar reconstrução/recibos |
| Retenção / compatibilidade | Simples distribuição, mas crescimento, edição manual e corrupção conjunta | Caminho local fora de sincronização, permissões e espaço; recuperação e truncamento exigem protocolo real | Mais esforço de instalação/administração/migração. Nenhuma ferramenta/banco foi instalado/testado neste portão |

Recomendação arquitetural condicional: C com autoridade única dos fatos, escritor único inicialmente, projeções verificáveis. Se B for escolhido como etapa intermediária, manter autoridade única declarada, exclusão mútua testada, envelopes/pré-imagens duráveis e janela Save/recibo em conciliação. A é aceitável somente com limites explicitados e checkpoints externos; não resolve sozinha durabilidade. A decisão de armazenamento não está aprovada.

## Protocolo implementável proposto

1. Abrir base em modo de verificação; conferir BaseID, SchemaVersion, Revision, identidade da cópia autoritativa e recibos. Divergência impede movimentação.
2. Adquirir exclusividade no armazenamento escolhido; validar fencing/revisão. Trava vencida não é removida automaticamente sem verificar autoridade e conciliação.
3. Persistir OperationID e envelope íntegro, ExpectedVersion e plano completo dos ProjectionIDs, pré-imagens mínimas e efeitos esperados antes de escrever. Falha de preparação aborta sem efeitos.
4. Comparar OperationID/envelope: igual concluído e íntegro retorna recibo, divergente recusa; qualquer pendência no episódio bloqueia também novo OperationID. A tentativa usa novo AttemptID.
5. Executar gravações em sequência, registrar etapa, verificar conteúdo/identidade, formatos e fórmulas. Compensar antes da fronteira somente quando pré-imagens e identidades forem verificadas.
6. Aplicar baixa aprovada e conferir invariantes. Marcar SAVE_INICIADO antes da persistência; salvar e verificar o arquivo efetivamente persistido em leitura isolada. Erro ou lacuna mantém pendência.
7. Persistir recibo da revisão/efeitos e conclusão no armazenamento autoritativo. Duas persistências não são uma transação conjunta. Após conclusão, validar qualquer replay contra efeitos/versões originais.
8. Liberar exclusividade somente após estado conhecido. A interface deve exibir OPERAÇÃO PENDENTE e impedir repetição, sem esconder o registro.

Envelope de produção ainda exige definição de esquema completo, serialização, hashes, pré-imagens, escritor/fencing e teste de durabilidade. O protótipo F7 tem contrato mínimo sintético, nove campos TSV e selo textual de duas linhas, sem pré-imagens duráveis, lock, fsync, autenticação ou transação. Não é implementação do protocolo completo.

## Tratamento de falhas

| Situação | Classificação / procedimento |
|---|---|
| Erro antes de efeito | RECUSADA; conservar intenção/tentativa e permitir correção do comando conforme contrato |
| Durante gravações, antes da baixa | COMPENSAÇÃO LIMITADA; restaurar em ordem inversa, conferir todos os efeitos e persistir recibo. Se uma restauração falha, PENDENTE |
| Durante baixa / compensação incompleta | PENDENTE_CONCILIAÇÃO; bloquear episódio/destinos e também novo OperationID. FAIL F6 de recuperação integral permanece |
| Save recusado/incompleto | Estado persistido incerto; não sobrescrever nem repetir. Preservar disco, memória e diário separadamente e comparar revisão/efeitos |
| Erro depois de Save | Conferir arquivo salvo vs diário. Ausência de recibo íntegro exige conciliação, mesmo que arquivo diga CONCLUIDA |
| Diário ausente/órfão/truncado/base divergente | Quarentena do conjunto, sem recriar diário automaticamente ou escolher arquivo mais recente pelo relógio |
| Histórico alterado ou duplicidade | Comparar IDs/FKs e envelope esperado. Bloquear correção/replay, manter todas as versões; nunca primeira NF-e/placa |
| Nova tentativa | Mesmo OperationID somente após conclusão íntegra ou compensação verificada persistida; pendência impede nova intenção para contornar bloqueio |
| Crash/energia/concorrência | NÃO EXECUTADOS. Sessão fria controlada não comprova resistência a esses eventos |

## Conciliação e liberação

Preservar evidências com hashes; apontar autoridade/revisão confiável; inventariar por IDs cada efeito ausente, íntegro, divergente, duplicado ou desconhecido. Conferir origem, projeções atuais, históricos, versões, fórmulas e documentos disponíveis sem inferir regra fiscal. Responsável aprova plano exato de reparo com pré/pós-condições, reversão e registro de decisão. Executar em cópia, testar e reabrir; liberar episódio somente com recibo de conciliação e novo checkpoint. Não repetir automaticamente para descobrir se já gravou.

Os testes frios F7 comparam 5.280 células em 11 abas (A1:X20), inclusive tipos/valores/fórmulas/NumberFormat/Ignore, antes/depois da guarda. Esse escopo não é estado completo do livro. Arquivos são preparados, fechados normalmente e lidos em nova instância /x; sem encerramento forçado. Evidências e limites finais constarão da [matriz](./MATRIZ_REGRESSAO_FASE7.md).
