# Importação e compilação manual — Fase 3

**Estado observado:** importação automatizada BLOQUEADA por 0x800A03EC. Nenhum componente importado, nenhuma compilação ou macro executada. [Mensagem original](./office-fase3.json). Excel 16.0 build 20527, x64. Não alterar AccessVBOM, Trust Center, ExecutionPolicy, Registro, locais confiáveis ou política de macros.

## Arquivos já preparados e conferidos

- [Esqueleto XLSM novo](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase3_20261031_213925Z/Fixture_Auditoria_Fase3.xlsm>): 17 abas, marcador fictício, sem fontes revisados incorporados. Não é versão compilada nem fixture funcional completa.
- [Pacote para importar no VBE](<C:/Users/MARCOSALEXANDREPERDO/AppData/Local/Temp/vscode-agent-codex-sandbox-JTJOQt/VBA_Fase3_20261031_213925Z/Importar_No_VBE>), fora do OneDrive: 34 componentes do sistema, cinco FRX e três componentes de teste. Os fontes de importação estão em Windows-1252/CRLF; os arquivos de referência não foram recodificados.
- [Manifesto do pacote](./pacote-importacao.json) e [esquema dos documentos](./scripts/fixture-esquema.json). A cópia anterior de dois arquivos do pacote foi preservada antes da correção C3-02.
- [Fontes revisados](./desenvolvimento/) permanecem separados da base. O XLSM dessa pasta contém dados da base, está intacto e **não deve ser usado para executar movimentações**.

## Ações que dependem do operador no Excel

1. Abrir somente o esqueleto indicado acima. Conferir caminho TEMP, prefixo Fixture_Auditoria_ e Menu Principal!Z1 = AUDITORIA_VBA_FICTICIO. Se a política impedir a abertura/execução, registrar o bloqueio e parar essa etapa, sem contornar a proteção.
2. Alt+F11: selecionar o projeto do esqueleto pelo nome completo. Não selecionar o projeto da aplicação original. Ao inicializar o VBE, conferir os 18 módulos de documento gerados; a sondagem anterior à inicialização devolveu CodeNames vazios, portanto nenhum CodeName foi certificado automaticamente.
3. Em Ferramentas > Referências, registrar a lista real e AUSENTE/MISSING. Office possui versões registradas 2.6/2.7/2.8; MSForms 2.0, MSXML 6 e Selenium 2 possuem caminhos registrados existentes, mas isso não comprova referência resolvida ou ativação. O recorte consultado de Common Controls não resolveu um arquivo OCX. [Diagnóstico somente leitura](./bibliotecas-registradas.json). Não remover uma referência nem registrar/instalar OCX para forçar compilação.
4. No painel Propriedades (F4), atribuir os CodeNames exatamente conforme o esquema. Renomear primeiro para nomes temporários apenas os componentes novos da fixture, se necessário para evitar colisão, e depois aplicar os nomes reais. Não criar abas ou controles para substituir referências ausentes.
5. No módulo EstaPastaDeTrabalho e em cada planilha, substituir somente o corpo de código pelo TXT correspondente de Documentos no pacote. Não importar esses TXT como classes. Preservar o evento existente em Planilha15; não criar CommandButton1.
6. Arquivo > Importar arquivo: importar os 29 componentes de MDL (28 BAS e uma CLS), os cinco FRM de Form com seus FRX ao lado, e os três de Testes. Não importar duas versões do mesmo formulário. O resultado esperado é **55 componentes**: 18 documentos + 34 do sistema + três de teste. Essa é expectativa, não contagem observada. Investigar qualquer diferença antes de executar.
7. Nos designers, conferir nomes, tipos, eventos e RowSource/ControlSource. Preencher somente os intervalos realmente referenciados com dados fictícios. A fixture ainda precisa de nomes definidos, fórmulas, listas, vínculos de botões/Ribbon e eventuais objetos de planilha compatíveis. Não copiar registros reais do XLSM para resolver dependências.
8. Depurar > Compilar VBAProject. Registrar cada mensagem com componente, procedimento, linha, data e referências resolvidas. Corrigir apenas na nova cópia respaldada e repetir a compilação. Se houver bloqueio ou incompatibilidade, registrar BLOQUEADO.
9. Salvar a fixture, fechar/reabrir e registrar SHA-256, contagem real de componentes e evidência da compilação. Não confundir importação, salvar como XLSM ou ausência de Problems com compilar.

## Execução VBA preparada, ainda não realizada

Depois de compilar e conferir a fixture, executar na Janela Imediata:

    ? Auditoria_ExecutarTestesSeguros()
    ? AuditoriaFase2_Executar()
    ? AuditoriaFase3_CicloVida()

A suíte existente foi reaproveitada. A extensão testa uso único do contexto e recuperação recusada quando inativo. O teste da chave agora exige restaurar também Errors(xlNumberAsText).Ignore. Esses testes usam pequenos intervalos fictícios e não executam os eventos fiscais dos formulários.

Para eventos privados, usar a interface: abrir Cad_NFE.Show ou Cad_NFE_Saida.Show na fixture, preencher os controles reais e acionar o botão correspondente conforme [matriz](../../MATRIZ_TESTES_FASE3.md). Não renomear ou tornar público um evento de produção. Diálogos de arquivo/relatório devem ser atendidos pelo operador com XML fictício. Os XML já fornecidos testam estrutura e somas; não constituem uma NF-e fiscalmente completa para aprovar Venda. Preparar os demais campos exigidos pelas regras atuais antes do caso integrado válido.

A planilha isolada não reconstrói automaticamente vínculos Ribbon. Para salvamento, duas pastas fictícias devem estar abertas; deixar a outra ativa e chamar o callback ChamarSalvar com controle Nothing, se a assinatura permitir, verificando a pasta proprietária e os arquivos em disco. O callback real continua NÃO EXECUTADO.

Falha de SaveAs por caminho inexistente, usada no ensaio COM, **não substitui falha de ThisWorkbook.Save na macro**. Este caso e as falhas estruturais precisam de instrumentação controlada na fixture depois de compilar; não provocar indisponibilidade em arquivo real.

Não promover: A08/A09 continuam críticos. O roteiro não autoriza mudança de baixa, identificadores, campos, fiscalidade ou segurança.
