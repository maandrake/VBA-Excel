# Regras permanentes de manutenção VBA

## Escopo e preservação

- Trate [Excel/](./Excel/), [MDL/](./MDL/) e [Form/](./Form/) como a base original até haver promoção explícita de uma revisão validada. A auditoria de 31/10/2026 manteve esses arquivos intactos.
- Trabalhe em cópia de desenvolvimento identificada. Antes de editar, gere backup datado e manifesto SHA-256; confira todos os hashes. Nunca sobrescreva originais de produção.
- Execute Office somente com arquivos de teste em pasta temporária fora do OneDrive. Não execute movimentações fiscais ou de estoque sobre dados reais, mesmo em uma cópia contendo esses dados.
- Não instale programas nem altere ExecutionPolicy, Trust Center, AccessVBOM, locais confiáveis ou segurança de macros sem autorização específica. Não use opções de bypass ou edições de registro para contornar bloqueios.
- Não envie fontes, XML, dados pessoais ou credenciais a serviços externos. Não publique nem envie ao Git remoto automaticamente.
- Feche apenas instâncias do Office criadas e identificadas pelo teste; nunca encerre globalmente Excel/Word.

## Contratos do projeto

- Preserve a lógica de negócio. Mudanças de CFOP, CST, cBenef, ICMS, PIS ou COFINS exigem evidência e aprovação específica.
- Chaves NF-e são texto de 44 dígitos. Defina NumberFormat = "@" antes da atribuição como String; não converta chaves em números.
- Confira nomes de controles no FRX/designer e planilhas no Excel antes de alterar referências. Não invente componentes ou substituições para nomes ausentes.
- Preserve pares FRM/FRX, atributos VB_Name, assinaturas e nomes de eventos. Mantenha a codificação detectada e as quebras de linha; muitos fontes usam Windows-1252/CRLF.
- Antes de remover código aparentemente inativo, confira eventos, Ribbon, botões, OnAction, chamadas dinâmicas, fórmulas e módulos de documento embutidos. Corpos duplicados de eventos não provam inutilidade.
- Não altere intervalos Delete/Insert sem validar fórmulas e colunas auxiliares de todas as abas envolvidas. Há riscos críticos de exclusão parcial registrados nesta auditoria.
- Pré-valide entradas e destinos e desenhe recuperação para gravações entre várias abas. Não prometa atomicidade inexistente.
- Restaure os valores anteriores de EnableEvents, ScreenUpdating, Calculation e DisplayAlerts inclusive em falhas.

## Verificação e entrega

- Análise textual e painel Problems do VS Code não equivalem à compilação VBA.
- VBProject pode retornar null silenciosamente no PowerShell. Verifique o objeto e conserve a exceção original do Office; não registre importação sem componentes reais.
- Importe somente em fixture nova, com dados fictícios. Pare quando o Office bloquear o acesso ao projeto e registre a pendência.
- Separe PASS, FAIL, BLOQUEADO e NÃO EXECUTADO. Distinga teste COM equivalente, simulação JS, compilação real e execução da rotina VBA.
- Compare assinaturas, regras fiscais, FRX e resultados antes/depois. Nunca promova um XLSM que ainda não incorporou os fontes revisados.
- Atualize [RELATORIO_AUDITORIA.md](./RELATORIO_AUDITORIA.md), [MAPA_PROJETO.md](./MAPA_PROJETO.md), [TESTES_VBA.md](./TESTES_VBA.md) e [HISTORICO_ALTERACOES.md](./HISTORICO_ALTERACOES.md).
- Use Git para fontes, scripts e documentação; mantenha backups e novas cópias de dados privados fora do versionamento.
