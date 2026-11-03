# Verificação das projeções — Fase 10

## Resultado final do ensaio dirigido

F10_Projecoes executou 23 asserções PASS no VBA real. A verificação inicial exigiu sete projeções históricas, dois registros ativos, três episódios e dois veículos. O ciclo contém reentrada, transferências sucessivas, devolução, NF-e repetida e valores discriminantes. Nenhum resultado numérico igual foi usado sozinho como prova de vínculo.

A autoridade é o ACCDB sintético. O XLSM produzido é uma projeção reconstruível; commit no ACE não confirma Save do Excel.

## Contrato integral da projeção sintética

A planilha Projecao_F10 contém 19 colunas: ProjectionID, EpisodeID, VehicleID, MovementID, Version, EpisodeState, CurrentPID, Place, Plate, NFe, Key, AmountIn, AmountOut, CivilDate, Canon, Active, Margin, Annual e ReceiptHash.

VerifyFull10 obtém um snapshot do banco sob a fronteira de escritor, verifica a auditoria e compara por ProjectionID, sem tomar a posição física como identidade. Confere BaseID, revisão, linhagem, esquema, selo da autoridade, BuildID, estado BUILT, cabeçalhos, quantidade, tipos, valores, chave, fórmulas exatas e resultados. Verifica formatos efetivos por intervalo e Ignore de xlNumberAsText por célula da chave; os totais U1:U4 exigem a fórmula contratada e o valor esperado. Duplicidades, ausências e mudanças de episódio são recusadas.

A fórmula Annual deriva do estado do episódio recuperado por identidade; não é migração das fórmulas anuais operacionais. As fórmulas Active/Margin/Annual pertencem à projeção sintética e são reconstruídas e verificadas junto aos IDs. O contrato não cobre todos os estilos, erros ignorados, controles ou fórmulas das 17 abas do produto.

## Adulterações e reordenação

Foram executadas 13 alterações independentes, sempre reconstruindo a referência antes do cenário seguinte: valor, fórmula de margem, formato, chave, identidade duplicada, registro sem ID, revisão, linhagem, hash de recibo, Ignore, EpisodeID, estado de atualização e fórmula de contagem. Todas foram recusadas, com o estado factual do banco preservado.

O Sort nativo removeu Ignore da chave, reproduzindo o defeito da Fase 9. O verificador recusou a planilha ordenada. A reconstrução refez vínculos, fórmulas e metadados e passou na verificação integral. O FAIL histórico permanece registrado; a correção experimental não muda o produto.

## Correções durante esta execução

1. A atribuição de Ignore num Range falhou com erro 1004 neste Excel. A rotina passou a usar a propriedade por célula da chave. O registro R1 foi preservado.
2. NumberFormat=@ em Annual fez o Excel persistir a expressão como texto, embora FormulaR1C1 devolvesse a expressão. O verificador de valores derivados recusou o resultado. A correção usa General nessa coluna de fórmula, conserva @ nas chaves e nos campos textuais e exige cálculo e resultado reais. Logs R2 preservados.
3. A suíte agora interrompe as adulterações se a referência inicial não estiver VERIFIED. Assim uma planilha já inválida não produz aprovação enganosa dos testes negativos.

## Confirmação persistente

Sequência: fato confirmado no banco → INTENT de reconstrução → BUILT em memória → Save → fechamento → reabertura somente leitura → verificação integral → hash do arquivo estável → confirmação no banco por BaseID/linhagem/revisão/BuildID. O arquivo não é regravado depois do hash para dizer CONFIRMED; essa confirmação pertence à autoridade.

Sem arquivo, retorna PENDING_FILE. Save isolado mantém INTENT. O arquivo normal foi reaberto, validado e recebeu CONFIRMED com SHA-256 f98e65232ad0527fccaadee4f76ddaa9951b430acc863243ef97295c6f9efe71. A sessão fria é registrada separadamente na matriz.

Um BeforeSave real cancelou Save por clsSaveFase9: Observado=True e Saved=False. O BuildID em disco permaneceu antigo e a confirmação retornou PENDING_VERIFY:BUILD_ID. SaveAs para pai inexistente retornou erro real1004. Uma interrupção artificial7100 após cabeçalhos não foi aceita como reconstrução completa. Fatos e recibos permaneceram na autoridade.

## Limitações de confirmação

Não existe transação conjunta ACE/XLSM nem bloqueio exclusivo do arquivo pelo sistema operacional durante toda a confirmação. O hash antes/depois reduz divergências observáveis, mas não prova ausência de corrida com um agente externo. ProjectionAcks não é uma âncora autenticada externa; adulteração coordenada do ACK e de seu arquivo não foi protegida. A CAS do ACK confere identidade/revisão, mas não constitui imutabilidade de todo o histórico de confirmações. Antes do Portão2, exigir protocolo de recibo de projeção, acesso controlado, caminho autoritativo e testes dessas fronteiras adicionais.

[Evidências e fontes](./_auditoria/fase10_20261102_155920Z/), [matriz](./MATRIZ_REGRESSAO_FASE10.md), [recuperação](./TESTES_RECUPERACAO_FASE10.md).
