Attribute VB_Name = "modValidacaoXML"
Option Explicit

Private Const TOLERANCIA_MOEDA As Double = 0.01
Private Const TOLERANCIA_ALIQUOTA As Double = 0.0001

Private mResultados As Collection
Private mQuantidadeErros As Long
Private mQuantidadeAvisos As Long

'===============================================================================
'CARREGAR UMA ÚNICA VEZ E PREENCHER OS CONTROLES EXISTENTES
'===============================================================================
Public Function CarregarXMLNFe( _
    ByVal formulario As Object, _
    ByRef documentoXML As Object, _
    ByRef arquivoXML As String) As Boolean

    Dim xmlDoc As Object
    Dim grupos As Object
    Dim infNFe As Object
    Dim idInfNFe As String
    Dim chave As String
    Dim numero As String

    On Error GoTo TratarErro

    CarregarXMLNFe = False
    Set documentoXML = Nothing

    arquivoXML = SelecionarArquivoXML()

    If arquivoXML = vbNullString Then
        MsgBox "A seleção do arquivo XML foi cancelada." & vbCrLf & _
               "A NF-e não será gravada.", _
               vbExclamation, "Validação do XML"
        Exit Function
    End If

    Set xmlDoc = CriarDocumentoXML()

    If xmlDoc Is Nothing Then
        MsgBox "Não foi possível iniciar o componente de leitura do XML.", _
               vbCritical, "Validação do XML"
        Exit Function
    End If

    With xmlDoc

        .async = False
        .validateOnParse = False
        .resolveExternals = False
        .SetProperty "ProhibitDTD", True
        .SetProperty "SelectionLanguage", "XPath"

        'Única leitura do arquivo em disco.
        If Not .Load(arquivoXML) Then
            MsgBox "Não foi possível carregar o arquivo XML." & vbCrLf & _
                   "Motivo: " & .parseError.reason, _
                   vbCritical, "Validação do XML"
            Exit Function
        End If

    End With

    Set grupos = xmlDoc.SelectNodes("//*[local-name()='infNFe']")

    If grupos.length <> 1 Then
        MsgBox "Selecione um XML contendo exatamente uma NF-e (grupo infNFe).", _
               vbCritical, "Validação do XML"
        Exit Function
    End If

    Set infNFe = grupos.item(0)

    idInfNFe = TextoAtributoXML(infNFe, "Id")
    chave = RemoverPrefixoNFe(idInfNFe)

    numero = TextoNoXML( _
        infNFe, "./*[local-name()='ide']/*[local-name()='nNF']")

    If UCase$(Left$(idInfNFe, 3)) <> "NFE" Or _
       Len(chave) <> 44 Or SomenteDigitos(chave) <> chave Then

        MsgBox "O atributo infNFe/@Id deve conter NFe seguido de 44 dígitos.", _
               vbCritical, "Validação do XML"
        Exit Function

    End If

    If Len(numero) = 0 Or Len(numero) > 9 Or _
       SomenteDigitos(numero) <> numero Then

        MsgBox "O grupo ide/nNF está ausente ou contém um número inválido.", _
               vbCritical, "Validação do XML"
        Exit Function

    End If

    If Val(numero) <= 0 Then
        MsgBox "O número da NF-e deve ser maior que zero.", _
               vbCritical, "Validação do XML"
        Exit Function
    End If

    'A chave permanece como texto, sem conversão numérica.
    formulario.TextBoxChaveDeAcesso.value = chave
    formulario.TextBox_N_Saida.value = numero

    Set documentoXML = xmlDoc

    CarregarXMLNFe = True
    Exit Function

TratarErro:

    Set documentoXML = Nothing

    MsgBox "Ocorreu um erro durante a leitura do XML." & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, _
           vbCritical, "Validação do XML"

End Function

'===============================================================================
'VALIDAR O DOCUMENTO JÁ CARREGADO
'
'A chamada antiga ValidarXMLNFe(Me) também permanece disponível.
'Quando o documento é informado, não ocorre nova leitura do arquivo.
'===============================================================================
Public Function ValidarXMLNFe( _
    ByVal formulario As Object, _
    Optional ByVal documentoXML As Object = Nothing, _
    Optional ByVal caminhoXML As String = vbNullString) As Boolean

    Dim arquivoXML As String
    Dim xmlDoc As Object
    Dim infNFe As Object
    Dim itemNFe As Object
    Dim numeroNFeXML As String
    Dim codigoProdutoXML As String
    Dim cfopXML As String
    Dim dataXML As String
    Dim tpNFXML As String
    Dim naturezaOperacaoXML As String

    On Error GoTo TratarErro

    ValidarXMLNFe = False

    Set mResultados = New Collection
    mQuantidadeErros = 0
    mQuantidadeAvisos = 0

    If documentoXML Is Nothing Then

        If Not CarregarXMLNFe(formulario, xmlDoc, arquivoXML) Then
            Exit Function
        End If

    Else

        Set xmlDoc = documentoXML
        arquivoXML = caminhoXML

    End If

    Set infNFe = xmlDoc.SelectSingleNode("//*[local-name()='infNFe']")

    If infNFe Is Nothing Then
        MsgBox "O arquivo selecionado não contém o grupo infNFe." & vbCrLf & _
               "O arquivo não foi reconhecido como uma NF-e válida.", _
               vbCritical, "Validação do XML"
        Exit Function
    End If

    ValidarEstruturaEAutorizacao xmlDoc, infNFe

    tpNFXML = TextoNoXML( _
        xmlDoc, "//*[local-name()='ide']/*[local-name()='tpNF']")

    AdicionarResultado _
        "Tipo da NF-e - tpNF", _
        "1 - Saída", _
        DescricaoTipoNF(tpNFXML), _
        Trim$(tpNFXML) = "1", _
        "O formulário de saídas somente aceita NF-e com tpNF igual a 1."

    AdicionarResultado _
        "Chave de acesso no formulário", _
        ObterValorControle(formulario, "TextBoxChaveDeAcesso"), _
        RemoverPrefixoNFe(TextoAtributoXML(infNFe, "Id")), _
        StrComp(ObterValorControle(formulario, "TextBoxChaveDeAcesso"), _
                RemoverPrefixoNFe(TextoAtributoXML(infNFe, "Id")), _
                vbBinaryCompare) = 0, _
        "A chave preenchida deve corresponder ao atributo Id da infNFe."

    numeroNFeXML = TextoNoXML( _
        infNFe, "./*[local-name()='ide']/*[local-name()='nNF']")

    AdicionarResultado _
        "Número da NF-e - nNF", _
        ObterValorControle(formulario, "TextBox_N_Saida"), _
        numeroNFeXML, _
        NumerosNFeIguais( _
            ObterValorControle(formulario, "TextBox_N_Saida"), numeroNFeXML), _
        vbNullString

    dataXML = TextoNoXML( _
        xmlDoc, "//*[local-name()='ide']/*[local-name()='dhEmi']")

    ValidarDataEmissao _
        ObterValorControle(formulario, "TextBox_Data"), dataXML

    naturezaOperacaoXML = TextoNoXML( _
        xmlDoc, "//*[local-name()='ide']/*[local-name()='natOp']")

    Set itemNFe = LocalizarItemPorCodigo( _
        xmlDoc, ObterValorControle(formulario, "COD"))

    If itemNFe Is Nothing Then

        AdicionarResultado _
            "Código do produto - cProd", _
            ObterValorControle(formulario, "COD"), _
            ListaCodigosProdutos(xmlDoc), _
            False, _
            "Não foi localizado no XML um item com o código informado."

        AdicionarDependenciasNaoValidadas

    Else

        codigoProdutoXML = TextoNoXML( _
            itemNFe, "./*[local-name()='prod']/*[local-name()='cProd']")

        AdicionarResultado _
            "Código do produto - cProd", _
            ObterValorControle(formulario, "COD"), _
            codigoProdutoXML, _
            StrComp(Trim$(ObterValorControle(formulario, "COD")), _
                    Trim$(codigoProdutoXML), vbTextCompare) = 0, _
            vbNullString

        ValidarPlacaNoProduto formulario, itemNFe

        cfopXML = TextoNoXML( _
            itemNFe, "./*[local-name()='prod']/*[local-name()='CFOP']")

        ValidarOperacaoECFOP _
            ObterValorControle(formulario, "ComboBox_Operação"), _
            cfopXML, naturezaOperacaoXML

        ValidarTransferenciaEntreEstabelecimentos _
            xmlDoc, ObterValorControle(formulario, "ComboBox_Operação"), cfopXML

        ValidarBeneficioFiscal itemNFe, cfopXML

        ValidarNumeroCampo _
            "Valor do produto - vProd", _
            ObterValorControle(formulario, "TextBox_Valor_Saida"), _
            TextoNoXML(itemNFe, _
                "./*[local-name()='prod']/*[local-name()='vProd']"), _
            TOLERANCIA_MOEDA

        ValidarICMSPorCST formulario, itemNFe
        ValidarPISeCOFINSPorCST formulario, itemNFe

    End If

    ValidarTotaisXML xmlDoc
    ExibirRelatorioVisual arquivoXML

    ValidarXMLNFe = (mQuantidadeErros = 0)

Saida:

    Set itemNFe = Nothing
    Set infNFe = Nothing
    Set xmlDoc = Nothing

    Exit Function

TratarErro:

    MsgBox "Ocorreu um erro durante a validação do XML." & _
           vbCrLf & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, _
           vbCritical, "Validação do XML"

    ValidarXMLNFe = False
    Resume Saida

End Function

'===============================================================================
'ESTRUTURA, MODELO, AMBIENTE, AUTORIZAÇÃO E CHAVE
'===============================================================================
Private Sub ValidarEstruturaEAutorizacao( _
    ByVal xmlDoc As Object, _
    ByVal infNFe As Object)

    Dim modeloXML As String
    Dim ambienteXML As String
    Dim cStatXML As String
    Dim motivoXML As String
    Dim idInfNFe As String
    Dim chaveId As String
    Dim chaveProtocolo As String

    modeloXML = TextoNoXML( _
        xmlDoc, "//*[local-name()='ide']/*[local-name()='mod']")

    AdicionarResultado _
        "Modelo do documento - mod", "55 - NF-e", modeloXML, _
        Trim$(modeloXML) = "55", _
        "O documento deve ser uma NF-e modelo 55."

    ambienteXML = TextoNoXML( _
        xmlDoc, _
        "//*[local-name()='protNFe']/*[local-name()='infProt']/" & _
        "*[local-name()='tpAmb']")

    If ambienteXML = vbNullString Then
        ambienteXML = TextoNoXML( _
            xmlDoc, "//*[local-name()='ide']/*[local-name()='tpAmb']")
    End If

    AdicionarResultado _
        "Ambiente da NF-e - tpAmb", _
        "1 - Produção", DescricaoAmbiente(ambienteXML), _
        Trim$(ambienteXML) = "1", _
        "XML de homologação não deve ser utilizado para gravação."

    cStatXML = TextoNoXML( _
        xmlDoc, _
        "//*[local-name()='protNFe']/*[local-name()='infProt']/" & _
        "*[local-name()='cStat']")

    motivoXML = TextoNoXML( _
        xmlDoc, _
        "//*[local-name()='protNFe']/*[local-name()='infProt']/" & _
        "*[local-name()='xMotivo']")

    AdicionarResultado _
        "Autorização da NF-e - cStat", _
        "100 - Autorizado o uso", _
        cStatXML & IIfTexto(motivoXML <> vbNullString, _
                           " - " & motivoXML, vbNullString), _
        Trim$(cStatXML) = "100", _
        "A NF-e deve possuir protocolo de autorização com cStat 100."

    idInfNFe = TextoAtributoXML(infNFe, "Id")
    chaveId = RemoverPrefixoNFe(idInfNFe)

    chaveProtocolo = TextoNoXML( _
        xmlDoc, _
        "//*[local-name()='protNFe']/*[local-name()='infProt']/" & _
        "*[local-name()='chNFe']")

    AdicionarResultado _
        "Chave de acesso", chaveId, chaveProtocolo, _
        chaveId <> vbNullString And _
        StrComp(chaveId, chaveProtocolo, vbBinaryCompare) = 0, _
        "A chave do protocolo deve ser igual à chave do atributo Id da infNFe."

End Sub

'===============================================================================
'OPERAÇÃO E CFOP
'===============================================================================
Private Sub ValidarOperacaoECFOP( _
    ByVal operacaoFormulario As String, _
    ByVal cfopXML As String, _
    ByVal naturezaOperacaoXML As String)

    Dim operacaoNormalizada As String
    Dim naturezaNormalizada As String
    Dim resultadoCorreto As Boolean
    Dim esperado As String

    operacaoNormalizada = NormalizarTexto(operacaoFormulario)
    naturezaNormalizada = NormalizarTexto(naturezaOperacaoXML)
    cfopXML = Trim$(cfopXML)

    If InStr(1, operacaoNormalizada, "DEVOLUCAO", vbTextCompare) > 0 Then
        operacaoNormalizada = "DEVOLUCAO_COMPRA"
    ElseIf InStr(1, operacaoNormalizada, "TRANSFER", vbTextCompare) > 0 Then
        operacaoNormalizada = "TRANSFERENCIA"
    ElseIf InStr(1, operacaoNormalizada, "VENDA", vbTextCompare) > 0 Then
        operacaoNormalizada = "VENDA"
    ElseIf InStr(1, operacaoNormalizada, "CANCEL", vbTextCompare) > 0 Then
        operacaoNormalizada = "CANCELADO"
    ElseIf InStr(1, operacaoNormalizada, "DEMONSTR", vbTextCompare) > 0 Then
        operacaoNormalizada = "DEMONSTRACAO"
    End If

    Select Case operacaoNormalizada

        Case "VENDA"

            esperado = "5102 ou 6102"
            resultadoCorreto = cfopXML = "5102" Or cfopXML = "6102"

            AdicionarResultado _
                "CFOP da operação", esperado, cfopXML, resultadoCorreto, _
                "Operação selecionada: Venda."

            If InStr(1, naturezaNormalizada, "VENDA", vbTextCompare) = 0 Then

                AdicionarAviso _
                    "Natureza da operação - natOp", _
                    "Descrição relacionada a venda", naturezaOperacaoXML, _
                    "O CFOP é determinante, mas a natureza da operação " & _
                    "não contém a palavra Venda."

            Else

                AdicionarResultado _
                    "Natureza da operação - natOp", "Venda", _
                    naturezaOperacaoXML, True, vbNullString

            End If

        Case "TRANSFERENCIA"

            esperado = "5152"
            resultadoCorreto = (cfopXML = "5152")

            AdicionarResultado _
                "CFOP da operação", esperado, cfopXML, resultadoCorreto, _
                "Operação selecionada: Transferência."

            If InStr(1, naturezaNormalizada, "TRANSFER", vbTextCompare) = 0 Then

                AdicionarAviso _
                    "Natureza da operação - natOp", _
                    "Descrição relacionada a transferência", naturezaOperacaoXML, _
                    "O CFOP é determinante, mas a natureza da operação " & _
                    "não menciona transferência."

            Else

                AdicionarResultado _
                    "Natureza da operação - natOp", "Transferência", _
                    naturezaOperacaoXML, True, vbNullString

            End If

        Case "DEVOLUCAO_COMPRA"

            esperado = "5202"
            resultadoCorreto = (cfopXML = "5202")

            AdicionarResultado _
                "CFOP da operação", esperado, cfopXML, resultadoCorreto, _
                "Operação selecionada: Devolução de compra para comercialização."

            If InStr(1, naturezaNormalizada, "DEVOLUCAO", vbTextCompare) = 0 Then

                AdicionarAviso _
                    "Natureza da operação - natOp", _
                    "Descrição relacionada a devolução", naturezaOperacaoXML, _
                    "O CFOP é determinante, mas a natureza da operação " & _
                    "não menciona devolução."

            Else

                AdicionarResultado _
                    "Natureza da operação - natOp", "Devolução", _
                    naturezaOperacaoXML, True, vbNullString

            End If

        Case "CANCELADO"

            resultadoCorreto = _
                cfopXML <> vbNullString And IsNumeric(cfopXML)

            AdicionarResultado _
                "CFOP da operação", "Qualquer CFOP numérico", _
                cfopXML, resultadoCorreto, _
                "A operação foi informada como Cancelado."

            AdicionarAviso _
                "Comprovação do cancelamento", _
                "XML de evento de cancelamento", "XML normal da NF-e", _
                "O XML normal da NF-e não comprova o cancelamento. " & _
                "Para comprovação definitiva, deve ser analisado o XML do evento."

        Case "DEMONSTRACAO"

            AdicionarAviso _
                "CFOP da operação", "Regra não cadastrada", cfopXML, _
                "A operação Demonstração ainda não possui uma regra " & _
                "específica de CFOP."

        Case Else

            AdicionarResultado _
                "Operação do formulário", _
                "Venda, Transferência, Devolução de compra para " & _
                "comercialização, Cancelado ou Demonstração", _
                operacaoFormulario, False, _
                "A operação informada não possui regra cadastrada."

    End Select

End Sub

'===============================================================================
'BENEFÍCIO FISCAL
'===============================================================================
Private Sub ValidarBeneficioFiscal( _
    ByVal itemNFe As Object, _
    ByVal cfopXML As String)

    Dim beneficioEsperado As String
    Dim beneficioXML As String
    Dim resultadoCorreto As Boolean
    Dim observacao As String

    cfopXML = Trim$(cfopXML)

    beneficioXML = TextoNoXML( _
        itemNFe, "./*[local-name()='prod']/*[local-name()='cBenef']")

    beneficioXML = UCase$(Trim$(beneficioXML))

    Select Case cfopXML

        Case "5102", "6102"

            beneficioEsperado = "SC820019"

            resultadoCorreto = _
                StrComp(beneficioEsperado, beneficioXML, vbTextCompare) = 0

            If resultadoCorreto Then

                observacao = _
                    "Benefício fiscal correto para o CFOP " & cfopXML & "."

            ElseIf beneficioXML = vbNullString Then

                observacao = _
                    "A tag cBenef está ausente ou vazia. " & _
                    "Para o CFOP " & cfopXML & _
                    ", o benefício esperado é " & beneficioEsperado & "."

            Else

                observacao = _
                    "Benefício fiscal divergente. " & _
                    "Para o CFOP " & cfopXML & _
                    ", o benefício esperado é " & beneficioEsperado & "."

            End If

            AdicionarResultado _
                "Benefício fiscal - cBenef", beneficioEsperado, _
                beneficioXML, resultadoCorreto, observacao

        Case "5152"

            beneficioEsperado = "SC800006"

            resultadoCorreto = _
                StrComp(beneficioEsperado, beneficioXML, vbTextCompare) = 0

            If resultadoCorreto Then

                observacao = "Benefício fiscal correto para o CFOP 5152."

            ElseIf beneficioXML = vbNullString Then

                observacao = _
                    "A tag cBenef está ausente ou vazia. " & _
                    "Para o CFOP 5152, o benefício esperado é " & _
                    beneficioEsperado & "."

            Else

                observacao = _
                    "Benefício fiscal divergente. " & _
                    "Para o CFOP 5152, o benefício esperado é " & _
                    beneficioEsperado & "."

            End If

            AdicionarResultado _
                "Benefício fiscal - cBenef", beneficioEsperado, _
                beneficioXML, resultadoCorreto, observacao

        Case "5202"

            beneficioEsperado = "Não informado"
            resultadoCorreto = (beneficioXML = vbNullString)

            If resultadoCorreto Then

                observacao = _
                    "Benefício fiscal corretamente não informado " & _
                    "para o CFOP 5202."

            Else

                observacao = _
                    "O CFOP 5202 não deve possuir benefício fiscal. " & _
                    "O XML contém cBenef = " & beneficioXML & "."

            End If

            AdicionarResultado _
                "Benefício fiscal - cBenef", beneficioEsperado, _
                beneficioXML, resultadoCorreto, observacao

        Case Else

            AdicionarInformativo _
                "Benefício fiscal - cBenef", "Regra não cadastrada", _
                beneficioXML, "NÃO VALIDADO", _
                "Não existe regra de benefício fiscal cadastrada para o CFOP " & _
                IIfTexto(cfopXML = vbNullString, "(ausente)", cfopXML) & "."

    End Select

End Sub

'===============================================================================
'TRANSFERÊNCIA ENTRE ESTABELECIMENTOS
'===============================================================================
Private Sub ValidarTransferenciaEntreEstabelecimentos( _
    ByVal xmlDoc As Object, _
    ByVal operacaoFormulario As String, _
    ByVal cfopXML As String)

    Dim operacaoNormalizada As String
    Dim cnpjEmitente As String
    Dim cnpjDestinatario As String
    Dim raizEmitente As String
    Dim raizDestinatario As String
    Dim resultadoCorreto As Boolean

    operacaoNormalizada = NormalizarTexto(operacaoFormulario)

    If operacaoNormalizada <> "TRANSFERENCIA" And _
       Trim$(cfopXML) <> "5152" Then
        Exit Sub
    End If

    cnpjEmitente = SomenteDigitos(TextoNoXML( _
        xmlDoc, "//*[local-name()='emit']/*[local-name()='CNPJ']"))

    cnpjDestinatario = SomenteDigitos(TextoNoXML( _
        xmlDoc, "//*[local-name()='dest']/*[local-name()='CNPJ']"))

    If Len(cnpjEmitente) = 14 Then
        raizEmitente = Left$(cnpjEmitente, 8)
    End If

    If Len(cnpjDestinatario) = 14 Then
        raizDestinatario = Left$(cnpjDestinatario, 8)
    End If

    resultadoCorreto = _
        raizEmitente <> vbNullString And _
        raizDestinatario <> vbNullString And _
        raizEmitente = raizDestinatario

    AdicionarResultado _
        "Raiz dos CNPJs na transferência", _
        FormatarCNPJ(cnpjEmitente), FormatarCNPJ(cnpjDestinatario), _
        resultadoCorreto, _
        "Na transferência, emitente e destinatário devem pertencer à mesma raiz de CNPJ."

End Sub

'===============================================================================
'PLACA NA DESCRIÇÃO DO PRODUTO
'===============================================================================
Private Sub ValidarPlacaNoProduto( _
    ByVal formulario As Object, _
    ByVal itemNFe As Object)

    Dim placaFormulario As String
    Dim descricaoProduto As String
    Dim placaNormalizada As String
    Dim descricaoNormalizada As String

    placaFormulario = ObterValorControle(formulario, "TextBox_PLACA")

    If Trim$(placaFormulario) = vbNullString Then Exit Sub

    descricaoProduto = TextoNoXML( _
        itemNFe, "./*[local-name()='prod']/*[local-name()='xProd']")

    placaNormalizada = NormalizarAlfanumerico(placaFormulario)
    descricaoNormalizada = NormalizarAlfanumerico(descricaoProduto)

    AdicionarResultado _
        "Placa na descrição - xProd", placaFormulario, descricaoProduto, _
        placaNormalizada <> vbNullString And _
        InStr(1, descricaoNormalizada, placaNormalizada, vbTextCompare) > 0, _
        "A placa informada deve constar na descrição do produto da NF-e."

End Sub

'===============================================================================
'ICMS CONFORME O CST
'===============================================================================
Private Sub ValidarICMSPorCST( _
    ByVal formulario As Object, _
    ByVal itemNFe As Object)

    Dim cstOriginal As String
    Dim cstNormalizado As String
    Dim valorProdutoTexto As String
    Dim reducaoTexto As String
    Dim baseTexto As String
    Dim aliquotaTexto As String
    Dim icmsTexto As String
    Dim valorProduto As Double
    Dim reducao As Double
    Dim baseICMS As Double
    Dim aliquotaICMS As Double
    Dim esperadoBase As Double
    Dim esperadoICMS As Double

    cstOriginal = TextoNoXML( _
        itemNFe, _
        "./*[local-name()='imposto']/*[local-name()='ICMS']" & _
        "//*[local-name()='CST']")

    cstNormalizado = NormalizarCST(cstOriginal)

    Select Case cstNormalizado

        Case "20"

            AdicionarResultado _
                "ICMS - CST", "20", cstOriginal, True, _
                "CST com redução da base de cálculo."

            baseTexto = TextoNoXML( _
                itemNFe, _
                "./*[local-name()='imposto']/*[local-name()='ICMS']" & _
                "//*[local-name()='vBC']")

            aliquotaTexto = TextoNoXML( _
                itemNFe, _
                "./*[local-name()='imposto']/*[local-name()='ICMS']" & _
                "//*[local-name()='pICMS']")

            icmsTexto = TextoNoXML( _
                itemNFe, _
                "./*[local-name()='imposto']/*[local-name()='ICMS']" & _
                "//*[local-name()='vICMS']")

            reducaoTexto = TextoNoXML( _
                itemNFe, _
                "./*[local-name()='imposto']/*[local-name()='ICMS']" & _
                "//*[local-name()='pRedBC']")

            valorProdutoTexto = TextoNoXML( _
                itemNFe, "./*[local-name()='prod']/*[local-name()='vProd']")

            ValidarNumeroCampo _
                "Base de cálculo do ICMS - vBC", _
                ObterValorControle(formulario, "TextBoxBaseICMS"), _
                baseTexto, TOLERANCIA_MOEDA

            ValidarNumeroCampo _
                "Alíquota do ICMS - pICMS", _
                ObterValorControle(formulario, "ComboBox_Por"), _
                aliquotaTexto, TOLERANCIA_ALIQUOTA

            ValidarNumeroCampo _
                "Valor do ICMS - vICMS", _
                ObterValorControle(formulario, "TextBox_ICMS"), _
                icmsTexto, TOLERANCIA_MOEDA

            If TentarConverterNumeroXML(valorProdutoTexto, valorProduto) And _
               TentarConverterNumeroXML(reducaoTexto, reducao) Then

                esperadoBase = valorProduto * (1 - reducao / 100)

                ValidarNumeroCalculado _
                    "Cálculo da base reduzida do ICMS", _
                    esperadoBase, baseTexto, TOLERANCIA_MOEDA, _
                    "vProd × (1 - pRedBC ÷ 100)."

            Else

                AdicionarAviso _
                    "Cálculo da base reduzida do ICMS", _
                    "vProd e pRedBC disponíveis", _
                    valorProdutoTexto & " / " & reducaoTexto, _
                    "Não foi possível conferir matematicamente a redução da base."

            End If

            If TentarConverterNumeroXML(baseTexto, baseICMS) And _
               TentarConverterNumeroXML(aliquotaTexto, aliquotaICMS) Then

                esperadoICMS = baseICMS * aliquotaICMS / 100

                ValidarNumeroCalculado _
                    "Cálculo interno do ICMS", _
                    esperadoICMS, icmsTexto, TOLERANCIA_MOEDA, _
                    "vBC × pICMS ÷ 100."

            Else

                AdicionarAviso _
                    "Cálculo interno do ICMS", _
                    "vBC e pICMS disponíveis", _
                    baseTexto & " / " & aliquotaTexto, _
                    "Não foi possível conferir matematicamente o valor do ICMS."

            End If

        Case "40", "41"

            AdicionarResultado _
                "ICMS - CST", "40 ou 41", cstOriginal, True, _
                "Operação sem tributação do ICMS no item."

            AdicionarInformativo _
                "ICMS - valores do item", _
                "Conferência dispensada", "CST " & cstNormalizado, _
                "SEM TRIBUTAÇÃO", _
                "vBC, pICMS e vICMS não são exigidos no item para esse CST."

        Case Else

            AdicionarResultado _
                "ICMS - CST", "20, 40 ou 41", cstOriginal, False, _
                "O CST do ICMS está ausente ou ainda não possui regra cadastrada."

    End Select

End Sub

'===============================================================================
'PIS E COFINS CONFORME O CST
'===============================================================================
Private Sub ValidarPISeCOFINSPorCST( _
    ByVal formulario As Object, _
    ByVal itemNFe As Object)

    Dim cstPIS As String
    Dim cstCOFINS As String

    cstPIS = TextoNoXML( _
        itemNFe, _
        "./*[local-name()='imposto']/*[local-name()='PIS']" & _
        "//*[local-name()='CST']")

    cstCOFINS = TextoNoXML( _
        itemNFe, _
        "./*[local-name()='imposto']/*[local-name()='COFINS']" & _
        "//*[local-name()='CST']")

    ValidarPISPorCST formulario, itemNFe, cstPIS
    ValidarCOFINSPorCST formulario, itemNFe, cstCOFINS

End Sub

Private Sub ValidarPISPorCST( _
    ByVal formulario As Object, _
    ByVal itemNFe As Object, _
    ByVal cstXMLOriginal As String)

    Dim cstNormalizado As String
    Dim basePISXML As String
    Dim valorPISXML As String

    cstNormalizado = NormalizarCST(cstXMLOriginal)

    Select Case cstNormalizado

        Case "01"

            AdicionarResultado _
                "PIS - CST", "01", cstXMLOriginal, True, _
                "Operação tributável. A base e o valor do PIS serão conferidos."

            basePISXML = TextoNoXML( _
                itemNFe, _
                "./*[local-name()='imposto']/*[local-name()='PIS']" & _
                "//*[local-name()='vBC']")

            ValidarNumeroCampo _
                "Base de cálculo do PIS - vBC", _
                CStr(formulario.TextBoxBasePISeCOFINS.value), _
                basePISXML, 0.005

            valorPISXML = TextoNoXML( _
                itemNFe, _
                "./*[local-name()='imposto']/*[local-name()='PIS']" & _
                "//*[local-name()='vPIS']")

            ValidarNumeroCampo _
                "Valor do PIS - vPIS", _
                CStr(formulario.TextBox_PIS.value), _
                valorPISXML, 0.005

        Case "07", "08"

            AdicionarResultado _
                "PIS - CST", "07 ou 08", cstXMLOriginal, True, _
                DescricaoCSTSemIncidencia(cstNormalizado)

            AdicionarInformativo _
                "PIS - valores", "Conferência dispensada", _
                "CST " & cstNormalizado, "SEM INCIDÊNCIA", _
                "A base de cálculo e o valor do PIS não foram conferidos."

        Case Else

            AdicionarResultado _
                "PIS - CST", "01, 07 ou 08", cstXMLOriginal, False, _
                "O CST do PIS está ausente ou não está previsto na validação."

    End Select

End Sub

Private Function DescricaoCSTSemIncidencia( _
    ByVal cstNormalizado As String) As String

    Select Case Trim$(cstNormalizado)

        Case "07"
            DescricaoCSTSemIncidencia = _
                "Operação isenta da contribuição. " & _
                "A conferência dos valores foi dispensada."

        Case "08"
            DescricaoCSTSemIncidencia = _
                "Operação sem incidência da contribuição. " & _
                "A conferência dos valores foi dispensada."

        Case Else
            DescricaoCSTSemIncidencia = _
                "Conferência dos valores dispensada."

    End Select

End Function

Private Sub ValidarCOFINSPorCST( _
    ByVal formulario As Object, _
    ByVal itemNFe As Object, _
    ByVal cstOriginal As String)

    Dim cstNormalizado As String
    Dim baseTexto As String
    Dim aliquotaTexto As String
    Dim valorCOFINSTexto As String
    Dim baseCOFINS As Double
    Dim aliquotaCOFINS As Double
    Dim valorCOFINSEsperado As Double

    cstNormalizado = NormalizarCST(cstOriginal)

    Select Case cstNormalizado

        Case "01"

            AdicionarResultado _
                "COFINS - CST", "01", cstOriginal, True, _
                "Operação tributável. Os valores da COFINS serão conferidos."

            baseTexto = TextoNoXML( _
                itemNFe, _
                "./*[local-name()='imposto']/*[local-name()='COFINS']" & _
                "//*[local-name()='vBC']")

            aliquotaTexto = TextoNoXML( _
                itemNFe, _
                "./*[local-name()='imposto']/*[local-name()='COFINS']" & _
                "//*[local-name()='pCOFINS']")

            valorCOFINSTexto = TextoNoXML( _
                itemNFe, _
                "./*[local-name()='imposto']/*[local-name()='COFINS']" & _
                "//*[local-name()='vCOFINS']")

            ValidarNumeroCampo _
                "Base de cálculo da COFINS - vBC", _
                ObterValorControle(formulario, "TextBoxBasePISeCOFINS"), _
                baseTexto, TOLERANCIA_MOEDA

            ValidarNumeroCampo _
                "Valor da COFINS - vCOFINS", _
                ObterValorControle(formulario, "TextBox_COFINS"), _
                valorCOFINSTexto, TOLERANCIA_MOEDA

            If TentarConverterNumeroXML(baseTexto, baseCOFINS) And _
               TentarConverterNumeroXML(aliquotaTexto, aliquotaCOFINS) Then

                valorCOFINSEsperado = baseCOFINS * aliquotaCOFINS / 100

                ValidarNumeroCalculado _
                    "Cálculo interno da COFINS", valorCOFINSEsperado, _
                    valorCOFINSTexto, TOLERANCIA_MOEDA, _
                    "vBC × pCOFINS ÷ 100."

            End If

        Case "07", "08"

            AdicionarResultado _
                "COFINS - CST", "07 ou 08", cstOriginal, True, _
                DescricaoCSTPISCOFINS(cstNormalizado)

            AdicionarInformativo _
                "COFINS - valores", "Conferência dispensada", _
                "CST " & cstNormalizado, "SEM INCIDÊNCIA", _
                "A base, a alíquota e o valor da COFINS não foram exigidos."

        Case Else

            AdicionarResultado _
                "COFINS - CST", "01, 07 ou 08", cstOriginal, False, _
                "O CST da COFINS está ausente ou ainda não possui regra cadastrada."

    End Select

End Sub

'===============================================================================
'TOTAIS INTERNOS DA NF-e
'===============================================================================
Private Sub ValidarTotaisXML(ByVal xmlDoc As Object)

    Dim somaProdutos As Double
    Dim somaBaseICMS As Double
    Dim somaICMS As Double
    Dim somaPIS As Double
    Dim somaCOFINS As Double

    somaProdutos = SomarNosXML( _
        xmlDoc, _
        "//*[local-name()='det']/*[local-name()='prod']/" & _
        "*[local-name()='vProd']")

    somaBaseICMS = SomarNosXML( _
        xmlDoc, _
        "//*[local-name()='det']/*[local-name()='imposto']/" & _
        "*[local-name()='ICMS']//*[local-name()='vBC']")

    somaICMS = SomarNosXML( _
        xmlDoc, _
        "//*[local-name()='det']/*[local-name()='imposto']/" & _
        "*[local-name()='ICMS']//*[local-name()='vICMS']")

    somaPIS = SomarNosXML( _
        xmlDoc, _
        "//*[local-name()='det']/*[local-name()='imposto']/" & _
        "*[local-name()='PIS']//*[local-name()='vPIS']")

    somaCOFINS = SomarNosXML( _
        xmlDoc, _
        "//*[local-name()='det']/*[local-name()='imposto']/" & _
        "*[local-name()='COFINS']//*[local-name()='vCOFINS']")

    ValidarNumeroCalculado _
        "Total dos produtos", somaProdutos, _
        TextoNoXML(xmlDoc, _
            "//*[local-name()='total']/*[local-name()='ICMSTot']/" & _
            "*[local-name()='vProd']"), _
        TOLERANCIA_MOEDA, "Soma dos vProd de todos os itens."

    ValidarNumeroCalculado _
        "Total da base do ICMS", somaBaseICMS, _
        TextoNoXML(xmlDoc, _
            "//*[local-name()='total']/*[local-name()='ICMSTot']/" & _
            "*[local-name()='vBC']"), _
        TOLERANCIA_MOEDA, "Soma das bases de ICMS dos itens."

    ValidarNumeroCalculado _
        "Total do ICMS", somaICMS, _
        TextoNoXML(xmlDoc, _
            "//*[local-name()='total']/*[local-name()='ICMSTot']/" & _
            "*[local-name()='vICMS']"), _
        TOLERANCIA_MOEDA, "Soma dos valores de ICMS dos itens."

    ValidarNumeroCalculado _
        "Total do PIS", somaPIS, _
        TextoNoXML(xmlDoc, _
            "//*[local-name()='total']/*[local-name()='ICMSTot']/" & _
            "*[local-name()='vPIS']"), _
        TOLERANCIA_MOEDA, "Soma dos valores de PIS dos itens."

    ValidarNumeroCalculado _
        "Total da COFINS", somaCOFINS, _
        TextoNoXML(xmlDoc, _
            "//*[local-name()='total']/*[local-name()='ICMSTot']/" & _
            "*[local-name()='vCOFINS']"), _
        TOLERANCIA_MOEDA, "Soma dos valores de COFINS dos itens."

End Sub

Private Sub AdicionarDependenciasNaoValidadas()

    AdicionarInformativo _
        "CFOP", "Item localizado", "Item não localizado", _
        "NÃO VALIDADO", "O CFOP depende da localização correta do item."

    AdicionarInformativo _
        "Benefício fiscal - cBenef", "Item localizado", "Item não localizado", _
        "NÃO VALIDADO", "O benefício depende da localização correta do item."

    AdicionarInformativo _
        "ICMS", "Item localizado", "Item não localizado", _
        "NÃO VALIDADO", "A tributação do ICMS não foi conferida."

    AdicionarInformativo _
        "PIS", "Item localizado", "Item não localizado", _
        "NÃO VALIDADO", "A tributação do PIS não foi conferida."

    AdicionarInformativo _
        "COFINS", "Item localizado", "Item não localizado", _
        "NÃO VALIDADO", "A tributação da COFINS não foi conferida."

End Sub

'===============================================================================
'VALIDAÇÕES NUMÉRICAS
'===============================================================================
Private Sub ValidarNumeroCampo( _
    ByVal nomeCampo As String, _
    ByVal valorFormularioTexto As String, _
    ByVal valorXMLTexto As String, _
    ByVal tolerancia As Double)

    Dim valorFormulario As Double
    Dim valorXML As Double
    Dim formularioValido As Boolean
    Dim xmlValido As Boolean
    Dim resultadoCorreto As Boolean
    Dim observacao As String

    formularioValido = _
        TentarConverterNumeroFormulario(valorFormularioTexto, valorFormulario)

    xmlValido = TentarConverterNumeroXML(valorXMLTexto, valorXML)

    resultadoCorreto = False
    observacao = vbNullString

    If Not formularioValido Then

        observacao = _
            "O valor informado no formulário está vazio ou não é numérico."

    ElseIf Not xmlValido Then

        observacao = "A tag está ausente ou não contém um número válido."

    Else

        resultadoCorreto = Abs(valorFormulario - valorXML) <= tolerancia

        If Not resultadoCorreto Then
            observacao = "Diferença encontrada: " & _
                         Format$(valorFormulario - valorXML, "0.0000")
        End If

    End If

    AdicionarResultado _
        nomeCampo, valorFormularioTexto, valorXMLTexto, _
        resultadoCorreto, observacao

End Sub

Private Sub ValidarNumeroCalculado( _
    ByVal nomeCampo As String, _
    ByVal valorEsperado As Double, _
    ByVal valorXMLTexto As String, _
    ByVal tolerancia As Double, _
    ByVal observacaoBase As String)

    Dim valorXML As Double
    Dim xmlValido As Boolean
    Dim resultadoCorreto As Boolean
    Dim observacao As String

    xmlValido = TentarConverterNumeroXML(valorXMLTexto, valorXML)

    If Not xmlValido Then

        resultadoCorreto = False

        observacao = observacaoBase & _
                     " A tag está ausente ou não contém um número válido."

    Else

        resultadoCorreto = Abs(valorEsperado - valorXML) <= tolerancia

        If resultadoCorreto Then
            observacao = observacaoBase
        Else
            observacao = observacaoBase & " Diferença: " & _
                         Format$(valorEsperado - valorXML, "0.0000")
        End If

    End If

    AdicionarResultado _
        nomeCampo, FormatarNumeroRelatorio(valorEsperado), _
        valorXMLTexto, resultadoCorreto, observacao

End Sub

Private Function SomarNosXML( _
    ByVal xmlDoc As Object, _
    ByVal caminhoXPath As String) As Double

    Dim nos As Object
    Dim noAtual As Object
    Dim valor As Double
    Dim total As Double

    total = 0

    Set nos = xmlDoc.SelectNodes(caminhoXPath)

    For Each noAtual In nos

        If TentarConverterNumeroXML(CStr(noAtual.text), valor) Then
            total = total + valor
        End If

    Next noAtual

    SomarNosXML = total

    Set nos = Nothing

End Function

'===============================================================================
'SELEÇÃO E ACESSO AO XML
'===============================================================================
Private Function SelecionarArquivoXML() As String

    Dim seletor As Object

    SelecionarArquivoXML = vbNullString

    Set seletor = Application.FileDialog(3)

    With seletor

        .Title = "Selecione o arquivo XML da NF-e"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Arquivo XML da NF-e", "*.xml"

        If .Show = -1 Then
            SelecionarArquivoXML = CStr(.SelectedItems(1))
        End If

    End With

    Set seletor = Nothing

End Function

Private Function CriarDocumentoXML() As Object

    Dim documento As Object

    On Error Resume Next

    Set documento = CreateObject("MSXML2.DOMDocument.6.0")

    If documento Is Nothing Then
        Set documento = CreateObject("MSXML2.DOMDocument.3.0")
    End If

    On Error GoTo 0

    Set CriarDocumentoXML = documento

End Function

Private Function TextoNoXML( _
    ByVal noPai As Object, _
    ByVal caminhoXPath As String) As String

    Dim noEncontrado As Object

    TextoNoXML = vbNullString

    If noPai Is Nothing Then Exit Function

    On Error Resume Next
    Set noEncontrado = noPai.SelectSingleNode(caminhoXPath)
    On Error GoTo 0

    If Not noEncontrado Is Nothing Then
        TextoNoXML = Trim$(CStr(noEncontrado.text))
    End If

    Set noEncontrado = Nothing

End Function

Private Function TextoAtributoXML( _
    ByVal noXML As Object, _
    ByVal nomeAtributo As String) As String

    Dim atributo As Object

    TextoAtributoXML = vbNullString

    If noXML Is Nothing Then Exit Function

    On Error Resume Next
    Set atributo = noXML.Attributes.getNamedItem(nomeAtributo)
    On Error GoTo 0

    If Not atributo Is Nothing Then
        TextoAtributoXML = Trim$(CStr(atributo.text))
    End If

    Set atributo = Nothing

End Function

Private Function LocalizarItemPorCodigo( _
    ByVal xmlDoc As Object, _
    ByVal codigoProcurado As String) As Object

    Dim itens As Object
    Dim itemAtual As Object
    Dim codigoXML As String

    Set LocalizarItemPorCodigo = Nothing

    codigoProcurado = Trim$(codigoProcurado)

    If codigoProcurado = vbNullString Then Exit Function

    Set itens = xmlDoc.SelectNodes("//*[local-name()='det']")

    For Each itemAtual In itens

        codigoXML = TextoNoXML( _
            itemAtual, "./*[local-name()='prod']/*[local-name()='cProd']")

        If StrComp(Trim$(codigoXML), codigoProcurado, vbTextCompare) = 0 Then
            Set LocalizarItemPorCodigo = itemAtual
            Exit Function
        End If

    Next itemAtual

    Set itens = Nothing

End Function

Private Function ListaCodigosProdutos(ByVal xmlDoc As Object) As String

    Dim nos As Object
    Dim noAtual As Object
    Dim resultado As String

    Set nos = xmlDoc.SelectNodes( _
        "//*[local-name()='det']/*[local-name()='prod']/*[local-name()='cProd']")

    For Each noAtual In nos

        If resultado <> vbNullString Then
            resultado = resultado & ", "
        End If

        resultado = resultado & Trim$(CStr(noAtual.text))

    Next noAtual

    If resultado = vbNullString Then
        resultado = "(nenhum cProd encontrado)"
    End If

    ListaCodigosProdutos = resultado

    Set nos = Nothing

End Function

Private Function ObterValorControle( _
    ByVal formulario As Object, _
    ByVal nomeControle As String) As String

    ObterValorControle = vbNullString

    On Error Resume Next
    ObterValorControle = CStr(formulario.Controls(nomeControle).value)
    On Error GoTo 0

End Function

'===============================================================================
'CONVERSÃO DE NÚMEROS
'===============================================================================
Private Function TentarConverterNumeroXML( _
    ByVal texto As String, _
    ByRef resultado As Double) As Boolean

    Dim expressao As Object
    Dim textoLimpo As String

    TentarConverterNumeroXML = False

    textoLimpo = Trim$(texto)

    If textoLimpo = vbNullString Then Exit Function

    Set expressao = CreateObject("VBScript.RegExp")

    With expressao
        .Global = False
        .IgnoreCase = True
        .Pattern = "^-?[0-9]+(\.[0-9]+)?$"
    End With

    If Not expressao.Test(textoLimpo) Then
        Set expressao = Nothing
        Exit Function
    End If

    resultado = Val(textoLimpo)

    TentarConverterNumeroXML = True

    Set expressao = Nothing

End Function

Private Function TentarConverterNumeroFormulario( _
    ByVal texto As String, _
    ByRef resultado As Double) As Boolean

    Dim valorLimpo As String
    Dim posicaoPonto As Long
    Dim posicaoVirgula As Long
    Dim quantidadePontos As Long
    Dim casasDepoisPonto As Long
    Dim expressao As Object

    TentarConverterNumeroFormulario = False

    valorLimpo = Trim$(texto)
    valorLimpo = Replace(valorLimpo, "R$", vbNullString)
    valorLimpo = Replace(valorLimpo, "%", vbNullString)
    valorLimpo = Replace(valorLimpo, Chr$(160), vbNullString)
    valorLimpo = Replace(valorLimpo, " ", vbNullString)

    If valorLimpo = vbNullString Then Exit Function

    posicaoPonto = InStrRev(valorLimpo, ".")
    posicaoVirgula = InStrRev(valorLimpo, ",")

    If posicaoPonto > 0 And posicaoVirgula > 0 Then

        If posicaoVirgula > posicaoPonto Then
            valorLimpo = Replace(valorLimpo, ".", vbNullString)
            valorLimpo = Replace(valorLimpo, ",", ".")
        Else
            valorLimpo = Replace(valorLimpo, ",", vbNullString)
        End If

    ElseIf posicaoVirgula > 0 Then

        valorLimpo = Replace(valorLimpo, ",", ".")

    ElseIf posicaoPonto > 0 Then

        quantidadePontos = _
            Len(valorLimpo) - Len(Replace(valorLimpo, ".", vbNullString))

        casasDepoisPonto = Len(valorLimpo) - posicaoPonto

        If quantidadePontos > 1 Then
            valorLimpo = Replace(valorLimpo, ".", vbNullString)
        ElseIf casasDepoisPonto = 3 Then
            valorLimpo = Replace(valorLimpo, ".", vbNullString)
        End If

    End If

    Set expressao = CreateObject("VBScript.RegExp")

    With expressao
        .Global = False
        .IgnoreCase = True
        .Pattern = "^-?[0-9]+(\.[0-9]+)?$"
    End With

    If Not expressao.Test(valorLimpo) Then
        Set expressao = Nothing
        Exit Function
    End If

    resultado = Val(valorLimpo)

    TentarConverterNumeroFormulario = True

    Set expressao = Nothing

End Function

'===============================================================================
'DATA DE EMISSÃO
'===============================================================================
Private Sub ValidarDataEmissao( _
    ByVal dataFormularioTexto As String, _
    ByVal dataISO As String)

    Dim dataFormulario As Date
    Dim dataXML As Date
    Dim formularioValido As Boolean
    Dim xmlValido As Boolean
    Dim resultadoCorreto As Boolean
    Dim valorXMLVisual As String
    Dim observacao As String

    formularioValido = _
        TentarConverterDataFormulario(dataFormularioTexto, dataFormulario)

    xmlValido = ConverterDataISO(dataISO, dataXML)

    If xmlValido Then
        valorXMLVisual = Format$(dataXML, "dd/mm/yyyy") & " | " & dataISO
    Else
        valorXMLVisual = dataISO
    End If

    If Not formularioValido Then

        resultadoCorreto = False
        observacao = "A data informada no formulário está vazia ou é inválida."

    ElseIf Not xmlValido Then

        resultadoCorreto = False
        observacao = "A tag dhEmi está ausente ou contém uma data inválida."

    Else

        resultadoCorreto = DateValue(dataFormulario) = DateValue(dataXML)

        If Not resultadoCorreto Then
            observacao = _
                "A data do formulário é diferente da data de emissão do XML."
        End If

    End If

    AdicionarResultado _
        "Data de emissão - dhEmi", dataFormularioTexto, _
        valorXMLVisual, resultadoCorreto, observacao

End Sub

Private Function TentarConverterDataFormulario( _
    ByVal texto As String, _
    ByRef resultado As Date) As Boolean

    Dim partes() As String
    Dim textoLimpo As String
    Dim dia As Long
    Dim mes As Long
    Dim ano As Long

    TentarConverterDataFormulario = False

    textoLimpo = Trim$(texto)

    If textoLimpo = vbNullString Then Exit Function

    If InStr(textoLimpo, " ") > 0 Then
        textoLimpo = Split(textoLimpo, " ")(0)
    End If

    If InStr(textoLimpo, "/") > 0 Then

        partes = Split(textoLimpo, "/")

        If UBound(partes) <> 2 Then Exit Function
        If Not IsNumeric(partes(0)) Then Exit Function
        If Not IsNumeric(partes(1)) Then Exit Function
        If Not IsNumeric(partes(2)) Then Exit Function

        dia = CLng(partes(0))
        mes = CLng(partes(1))
        ano = CLng(partes(2))

    ElseIf InStr(textoLimpo, "-") > 0 Then

        partes = Split(textoLimpo, "-")

        If UBound(partes) <> 2 Then Exit Function
        If Not IsNumeric(partes(0)) Then Exit Function
        If Not IsNumeric(partes(1)) Then Exit Function
        If Not IsNumeric(partes(2)) Then Exit Function

        If Len(partes(0)) = 4 Then
            ano = CLng(partes(0))
            mes = CLng(partes(1))
            dia = CLng(partes(2))
        Else
            dia = CLng(partes(0))
            mes = CLng(partes(1))
            ano = CLng(partes(2))
        End If

    Else

        Exit Function

    End If

    TentarConverterDataFormulario = _
        CriarDataSegura(ano, mes, dia, resultado)

End Function

Private Function ConverterDataISO( _
    ByVal textoISO As String, _
    ByRef resultado As Date) As Boolean

    Dim ano As Long
    Dim mes As Long
    Dim dia As Long

    ConverterDataISO = False

    textoISO = Trim$(textoISO)

    If Len(textoISO) < 10 Then Exit Function
    If Mid$(textoISO, 5, 1) <> "-" Then Exit Function
    If Mid$(textoISO, 8, 1) <> "-" Then Exit Function
    If Not IsNumeric(Left$(textoISO, 4)) Then Exit Function
    If Not IsNumeric(Mid$(textoISO, 6, 2)) Then Exit Function
    If Not IsNumeric(Mid$(textoISO, 9, 2)) Then Exit Function

    ano = CLng(Left$(textoISO, 4))
    mes = CLng(Mid$(textoISO, 6, 2))
    dia = CLng(Mid$(textoISO, 9, 2))

    ConverterDataISO = CriarDataSegura(ano, mes, dia, resultado)

End Function

Private Function CriarDataSegura( _
    ByVal ano As Long, _
    ByVal mes As Long, _
    ByVal dia As Long, _
    ByRef resultado As Date) As Boolean

    CriarDataSegura = False

    On Error GoTo DataInvalida

    resultado = DateSerial(ano, mes, dia)

    If Year(resultado) <> ano Then GoTo DataInvalida
    If Month(resultado) <> mes Then GoTo DataInvalida
    If Day(resultado) <> dia Then GoTo DataInvalida

    CriarDataSegura = True
    Exit Function

DataInvalida:

    CriarDataSegura = False

End Function

'===============================================================================
'COMPARAÇÕES E DESCRIÇÕES
'===============================================================================
Private Function NumerosNFeIguais( _
    ByVal numeroFormulario As String, _
    ByVal numeroXML As String) As Boolean

    numeroFormulario = Trim$(numeroFormulario)
    numeroXML = Trim$(numeroXML)

    NumerosNFeIguais = False

    If numeroFormulario = vbNullString Then Exit Function
    If numeroXML = vbNullString Then Exit Function

    If IsNumeric(numeroFormulario) And IsNumeric(numeroXML) Then

        NumerosNFeIguais = CDec(numeroFormulario) = CDec(numeroXML)

    Else

        NumerosNFeIguais = _
            StrComp(numeroFormulario, numeroXML, vbTextCompare) = 0

    End If

End Function

Private Function NormalizarCST(ByVal cst As String) As String

    cst = Trim$(cst)

    If cst = vbNullString Then
        NormalizarCST = vbNullString
    ElseIf IsNumeric(cst) Then
        NormalizarCST = Format$(CLng(cst), "00")
    Else
        NormalizarCST = UCase$(cst)
    End If

End Function

Private Function DescricaoCSTPISCOFINS( _
    ByVal cstNormalizado As String) As String

    Select Case cstNormalizado

        Case "07"
            DescricaoCSTPISCOFINS = "Operação isenta da contribuição."

        Case "08"
            DescricaoCSTPISCOFINS = "Operação sem incidência da contribuição."

        Case Else
            DescricaoCSTPISCOFINS = "Conferência dispensada."

    End Select

End Function

Private Function DescricaoTipoNF(ByVal tpNF As String) As String

    Select Case Trim$(tpNF)

        Case "0"
            DescricaoTipoNF = "0 - Entrada"

        Case "1"
            DescricaoTipoNF = "1 - Saída"

        Case Else
            DescricaoTipoNF = tpNF

    End Select

End Function

Private Function DescricaoAmbiente(ByVal tpAmb As String) As String

    Select Case Trim$(tpAmb)

        Case "1"
            DescricaoAmbiente = "1 - Produção"

        Case "2"
            DescricaoAmbiente = "2 - Homologação"

        Case Else
            DescricaoAmbiente = tpAmb

    End Select

End Function

'===============================================================================
'NORMALIZAÇÃO E FORMATAÇÃO
'===============================================================================
Private Function NormalizarTexto(ByVal texto As String) As String

    texto = UCase$(Trim$(texto))

    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "À", "A")
    texto = Replace(texto, "Ã", "A")
    texto = Replace(texto, "Â", "A")
    texto = Replace(texto, "Ä", "A")

    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "È", "E")
    texto = Replace(texto, "Ê", "E")
    texto = Replace(texto, "Ë", "E")

    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ì", "I")
    texto = Replace(texto, "Î", "I")
    texto = Replace(texto, "Ï", "I")

    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ò", "O")
    texto = Replace(texto, "Õ", "O")
    texto = Replace(texto, "Ô", "O")
    texto = Replace(texto, "Ö", "O")

    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ù", "U")
    texto = Replace(texto, "Û", "U")
    texto = Replace(texto, "Ü", "U")

    texto = Replace(texto, "Ç", "C")

    NormalizarTexto = texto

End Function

Private Function NormalizarAlfanumerico(ByVal texto As String) As String

    Dim i As Long
    Dim caractere As String
    Dim resultado As String

    texto = UCase$(texto)

    For i = 1 To Len(texto)

        caractere = Mid$(texto, i, 1)

        If caractere Like "[A-Z0-9]" Then
            resultado = resultado & caractere
        End If

    Next i

    NormalizarAlfanumerico = resultado

End Function

Private Function SomenteDigitos(ByVal texto As String) As String

    Dim i As Long
    Dim caractere As String
    Dim resultado As String

    For i = 1 To Len(texto)

        caractere = Mid$(texto, i, 1)

        If caractere Like "#" Then
            resultado = resultado & caractere
        End If

    Next i

    SomenteDigitos = resultado

End Function

Private Function FormatarCNPJ(ByVal cnpj As String) As String

    cnpj = SomenteDigitos(cnpj)

    If Len(cnpj) = 14 Then

        FormatarCNPJ = _
            Left$(cnpj, 2) & "." & _
            Mid$(cnpj, 3, 3) & "." & _
            Mid$(cnpj, 6, 3) & "/" & _
            Mid$(cnpj, 9, 4) & "-" & _
            Right$(cnpj, 2)

    Else

        FormatarCNPJ = cnpj

    End If

End Function

Private Function RemoverPrefixoNFe(ByVal valor As String) As String

    valor = Trim$(valor)

    If UCase$(Left$(valor, 3)) = "NFE" Then
        valor = Mid$(valor, 4)
    End If

    RemoverPrefixoNFe = valor

End Function

Private Function FormatarNumeroRelatorio(ByVal valor As Double) As String

    FormatarNumeroRelatorio = Format$(valor, "0.00##")

End Function

Private Function IIfTexto( _
    ByVal condicao As Boolean, _
    ByVal valorVerdadeiro As String, _
    ByVal valorFalso As String) As String

    If condicao Then
        IIfTexto = valorVerdadeiro
    Else
        IIfTexto = valorFalso
    End If

End Function

'===============================================================================
'RESULTADOS E RELATÓRIO
'===============================================================================
Private Sub AdicionarResultado( _
    ByVal campo As String, _
    ByVal valorFormulario As String, _
    ByVal valorXML As String, _
    ByVal correto As Boolean, _
    ByVal observacao As String)

    Dim item As Variant
    Dim resultado As String

    If Trim$(valorFormulario) = vbNullString Then
        valorFormulario = "(vazio)"
    End If

    If Trim$(valorXML) = vbNullString Then
        valorXML = "(tag ausente ou vazia)"
    End If

    If correto Then

        resultado = "CORRETO"

    Else

        resultado = "INCORRETO"
        mQuantidadeErros = mQuantidadeErros + 1

    End If

    item = Array(campo, valorFormulario, valorXML, resultado, observacao)

    mResultados.Add item

End Sub

Private Sub AdicionarAviso( _
    ByVal campo As String, _
    ByVal valorFormulario As String, _
    ByVal valorXML As String, _
    ByVal observacao As String)

    Dim item As Variant

    If Trim$(valorFormulario) = vbNullString Then
        valorFormulario = "(vazio)"
    End If

    If Trim$(valorXML) = vbNullString Then
        valorXML = "(tag ausente ou vazia)"
    End If

    item = Array(campo, valorFormulario, valorXML, "AVISO", observacao)

    mResultados.Add item

    mQuantidadeAvisos = mQuantidadeAvisos + 1

End Sub

Private Sub AdicionarInformativo( _
    ByVal campo As String, _
    ByVal valorFormulario As String, _
    ByVal valorXML As String, _
    ByVal resultado As String, _
    ByVal observacao As String)

    Dim item As Variant

    If Trim$(valorFormulario) = vbNullString Then
        valorFormulario = "(vazio)"
    End If

    If Trim$(valorXML) = vbNullString Then
        valorXML = "(tag ausente ou vazia)"
    End If

    item = Array(campo, valorFormulario, valorXML, resultado, observacao)

    mResultados.Add item

End Sub

Private Sub ExibirRelatorioVisual(ByVal arquivoXML As String)

    Load frmRelatorioXML

    frmRelatorioXML.CarregarRelatorio _
        mResultados, _
        mQuantidadeErros, _
        mQuantidadeAvisos, _
        arquivoXML

    frmRelatorioXML.Show vbModal

End Sub

