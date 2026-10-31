# VBA-Excel — Estoque Fera Veículos

Projeto em VBA para automatizar rotinas administrativas e fiscais da operação de veículos no Microsoft Excel. A solução utiliza UserForms, planilhas estruturadas e módulos VBA para cadastrar, pesquisar, consultar e validar informações de entradas e saídas de veículos e documentos fiscais.

## Funcionalidades

- Cadastro de NF-e de entrada e de saída.
- Cadastro e pesquisa de veículos.
- Pesquisa de NCM e CEST na base da pasta de trabalho.
- Pesquisa rápida de registros nas planilhas `Entradas` e `Entradas Filial`.
- Navegação do registro encontrado diretamente para a célula de origem.
- Consulta de informações de veículos e NF-e em serviços do Detran/Renavam usando Selenium/ChromeDriver.
- Preenchimento automático de formulários a partir dos dados consultados.
- Importação, comparação e validação de XML de NF-e.
- Relatório visual das validações realizadas, com indicação de erros e avisos.
- Rotinas de backup da pasta de trabalho.
- Controles de interface: tela cheia, ocultação da barra de título, cursor personalizado e navegação entre células.
- Integração com o Office Ribbon para abrir formulários e salvar a pasta de trabalho.

## Estrutura do projeto

```text
VBA-Excel/
├── Excel/
│   └── Estoque Fera Veículos 2026.xlsm   # Pasta de trabalho principal
├── Form/                                 # UserForms exportados do VBA
│   ├── Cad_NCM.frm                       # Pesquisa NCM/CEST
│   ├── Cad_NFE.frm                       # Cadastro de NF-e de entrada
│   ├── Cad_NFE_Saida.frm                 # Cadastro de NF-e de saída
│   ├── frmRelatorioXML.frm                # Relatório de validação XML
│   └── UserFormFiltro.frm                 # Filtros e pesquisa
├── MDL/                                  # Módulos padrão exportados do VBA
└── README.md
```

Os arquivos `.frx` acompanham os UserForms e armazenam seus recursos visuais. Eles devem permanecer junto dos respectivos arquivos `.frm` quando os formulários forem importados em outro projeto.

## Principais módulos

| Módulo | Responsabilidade |
| --- | --- |
| `MóduloPesquisa.bas` | Pesquisa em lote nas abas de entradas e navegação para o registro localizado. |
| `modPreencherFormulario.bas` | Transfere dados consultados para os formulários. |
| `modSeleniumCommon.bas` | Configuração do ChromeDriver, espera de carregamento, pausas e logs. |
| `modSeleniumConsultaNFE.bas` | Consulta de dados de NF-e e integração com o Detran. |
| `modSeleniumRenavamConsulta.bas` | Consulta de dados por placa/Renavam. |
| `modValidacaoXML.bas` | Leitura e validação de campos do XML da NF-e. |
| `mdlCriaBackup.bas` | Criação de cópias de segurança da pasta de trabalho. |
| `MóduloOfficeRibbonEditor.bas` | Comandos chamados pelo Ribbon do Office. |
| `mdlCopiar.bas` | Cópia de valores entre formulário, planilha e painel de navegação. |
| `mdlTelaCheia.bas` / `mdlOcultarTítulo.bas` | Ajustes de apresentação dos formulários. |
| `mdlCondCel.bas` | Leitura de condições e cores de formatação condicional. |

## Pré-requisitos

- Microsoft Excel para Windows com suporte a macros VBA.
- Permissão para executar conteúdo habilitado para macro e, se aplicável, acesso às referências usadas pelo projeto.
- Google Chrome instalado para as rotinas de consulta automatizada.
- SeleniumBasic/VBA Selenium e ChromeDriver compatíveis com a versão instalada do Chrome, quando as consultas Detran/Renavam forem utilizadas.
- Acesso à internet e disponibilidade dos sites consultados.
- Permissão para chamadas da API do Windows usadas por alguns módulos de interface.

> As consultas externas dependem do site de destino, de sua disponibilidade e de alterações no HTML. Se o site mudar, os seletores e os fluxos Selenium poderão precisar de atualização.

## Instalação e abertura

1. Clone ou copie o projeto mantendo a estrutura de pastas.
2. Abra `Excel/Estoque Fera Veículos 2026.xlsm` no Excel para Windows.
3. Habilite o conteúdo e as macros somente se a origem do arquivo for confiável.
4. Confira em **VBA > Ferramentas > Referências** se as bibliotecas exigidas pelo projeto estão disponíveis.
5. Configure o ChromeDriver conforme a versão do Chrome e os caminhos definidos em `modSeleniumCommon.bas`.
6. Teste primeiro as funções locais — cadastro, pesquisa e validação de XML — antes de executar as consultas externas.

## Uso básico

### Entradas e saídas

Use os comandos disponibilizados no Ribbon ou as macros de abertura dos formulários:

- `ChamarFrmEntrada`: abre o cadastro de NF-e de entrada.
- `ChamarFrmSaida`: abre o cadastro de NF-e de saída.
- `ChamarSalvar`: salva a pasta de trabalho.

Os dados são mantidos nas abas de entradas e saídas da pasta de trabalho. O formulário de entrada também possui fluxo específico para registros da filial e permite copiar dados entre os contextos principal e filial.

### Pesquisa

O módulo `MóduloPesquisa` percorre os registros das abas `Entradas` e `Entradas Filial`, apresenta os resultados em uma ListBox e mantém a planilha e a linha de origem para permitir a navegação até o registro selecionado.

### NCM/CEST

O formulário **Pesquisa NCM** usa a aba `NCM e CEST`. Os cabeçalhos devem estar na primeira linha; a pesquisa filtra os registros a partir do campo selecionado.

### Validação de XML

A validação compara informações do formulário com tags do XML, incluindo número da NF-e, data de emissão, tipo de operação, ambiente e campos tributários. O resultado é exibido no formulário de relatório com os estados `CORRETO`, `INCORRETO`, `AVISO` ou informação complementar.

### Consultas Detran/Renavam

As consultas são acionadas pelos botões dos formulários e usam Selenium para controlar o Chrome. Antes de executar:

- confirme que o ChromeDriver está acessível;
- verifique se o site permite a consulta e se os dados informados estão completos;
- encerre a sessão do navegador pelas rotinas do projeto quando necessário;
- não armazene credenciais ou dados sensíveis no código-fonte.

## Desenvolvimento

Os fontes `.bas` e `.frm` são exportações do editor do VBA e devem ser editados preferencialmente no próprio projeto VBA. Ao alterar UserForms, mantenha o par `.frm`/`.frx` e valide a abertura do formulário no Excel.

Antes de distribuir uma alteração:

1. Faça uma cópia de backup da pasta de trabalho.
2. Compile o projeto em **VBA > Depurar > Compilar Projeto VBA**.
3. Teste cadastro, edição, exclusão e pesquisa nas abas principal e filial.
4. Teste a validação de XML com arquivos válidos e inválidos.
5. Teste as consultas Selenium em ambiente controlado.
6. Salve e abra novamente o `.xlsm` para confirmar que o Ribbon e os formulários continuam funcionando.

## Segurança e limitações

- Arquivos `.xlsm` podem executar código com acesso aos dados locais e à internet. Abra-os somente de fontes confiáveis.
- O projeto altera planilhas e pode excluir registros; faça backup antes de operações de manutenção.
- As consultas automatizadas são sensíveis a mudanças de navegador, ChromeDriver, rede, autenticação e layout dos sites externos.
- A compatibilidade foi pensada para Excel no Windows; APIs do Windows e componentes ActiveX/Selenium podem não funcionar no macOS ou em ambientes web.
- Dados fiscais e de clientes devem ser tratados conforme as políticas internas e a legislação aplicável.

## Licença

Este repositório não contém uma licença de uso explícita. Consulte o responsável pelo projeto antes de redistribuir, modificar ou utilizar os componentes fora do ambiente autorizado.
