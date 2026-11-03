# Cópia de desenvolvimento — validação pendente

Esta pasta contém os fontes com as correções da auditoria 20261031_174936Z.

**O arquivo em [Excel/](./Excel/) é uma cópia intacta da base. Ele NÃO recebeu os fontes revisados**, pois o Office bloqueou o acesso programático a VBProject. Não distribua esse XLSM como versão corrigida.

Os ajustes estão em [Form/](./Form/) e [MDL/](./MDL/). Preserve os pares FRM/FRX e a codificação dos fontes. Não importe módulos de documento como classes comuns.

Consulte [relatório](../../../RELATORIO_AUDITORIA.md), [testes](../../../TESTES_VBA.md), [mapa](../../../MAPA_PROJETO.md) e [histórico](../../../HISTORICO_ALTERACOES.md). A preparação de fixture está em [Preparar-TestesVBA.ps1](../scripts/Preparar-TestesVBA.ps1); ela não altera segurança e deve parar se VBProject for recusado.

Não execute operações fiscais ou de estoque em dados reais. O próximo passo é importar/compilar em arquivo novo de teste, executar a regressão com dados fictícios e resolver os riscos críticos de exclusão e gravação parcial.
