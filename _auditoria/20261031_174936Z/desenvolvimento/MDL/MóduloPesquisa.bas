Attribute VB_Name = "MóduloPesquisa"

Option Explicit

Private Const LINHA_INICIAL_DADOS As Long = 2
Private Const COL_INICIO As Long = 1        ' Coluna A
Private Const COL_FIM As Long = 11         ' Coluna K - Entradas
Private Const COL_FIM_FILIAL As Long = 12  ' Coluna L - Entradas Filial

'==========================================================
' CONFIGURAR LISTBOX DE PESQUISA
'==========================================================
Public Sub ConfigurarListBoxPesquisa(ByVal Lst As Object)

    With Lst
        .Clear

        ' Coluna 0  = Nome da planilha
        ' Colunas 1 a 11 = Dados A até K
        ' Coluna 12 = Número da linha (oculta)
        ' Coluna 13 = Chave de acesso da NFe (Entradas Filial)

        .ColumnCount = 14

        .ColumnWidths = _
            "70 pt;20 pt;50 pt;20 pt;300 pt;" & _
            "60 pt;60 pt;50 pt;50 pt;50 pt;" & _
            "70 pt;50 pt;0 pt;280 pt"
    End With

End Sub

'==========================================================
' PESQUISAR ENTRADAS E ENTRADAS FILIAL
'==========================================================
Public Sub PesquisarEntradasRapido( _
    ByVal TextoPesquisa As String, _
    ByVal Lst As Object)

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim Abas As Variant
    Dim NomeAba As Variant

    Dim UltLinha As Long
    Dim TotalMaximo As Long
    Dim ColFimLeitura As Long

    Dim dados As Variant
    Dim resultado() As Variant

    Dim i As Long
    Dim j As Long

    Dim IndiceResultado As Long
    Dim TextoBusca As String
    Dim Encontrou As Boolean

    TextoBusca = Trim$(TextoPesquisa)

    ConfigurarListBoxPesquisa Lst

    If TextoBusca = vbNullString Then Exit Sub

    Set wb = ThisWorkbook

    Abas = Array("Entradas", "Entradas Filial")

    '======================================================
    ' CALCULAR TOTAL MÁXIMO DE REGISTROS
    '======================================================
    For Each NomeAba In Abas

        Set ws = wb.Worksheets(CStr(NomeAba))

        UltLinha = UltimaLinhaUsada(ws, "A")

        If UltLinha >= LINHA_INICIAL_DADOS Then

            TotalMaximo = TotalMaximo + _
                (UltLinha - LINHA_INICIAL_DADOS + 1)

        End If

    Next NomeAba

    If TotalMaximo = 0 Then Exit Sub

    '======================================================
    ' MATRIZ DE RESULTADOS
    '
    ' 0       = Nome da planilha
    ' 1 a 11  = Colunas A até K
    ' 12      = Número da linha (oculta)
    ' 13      = Coluna L - Chave de acesso da NFe
    '======================================================
    ReDim resultado(0 To 13, 0 To TotalMaximo - 1)

    IndiceResultado = -1

    '======================================================
    ' PERCORRER AS PLANILHAS
    '======================================================
    For Each NomeAba In Abas

        Set ws = wb.Worksheets(CStr(NomeAba))

        UltLinha = UltimaLinhaUsada(ws, "A")

        ' Determina a última coluna conforme a planilha
        If ws.Name = "Entradas Filial" Then

            ColFimLeitura = COL_FIM_FILIAL

        Else

            ColFimLeitura = COL_FIM

        End If

        If UltLinha >= LINHA_INICIAL_DADOS Then

            ' Carrega os dados em memória
            dados = ws.Range( _
                ws.Cells(LINHA_INICIAL_DADOS, COL_INICIO), _
                ws.Cells(UltLinha, ColFimLeitura) _
            ).value

            '==================================================
            ' PESQUISAR NAS COLUNAS DISPONÍVEIS
            '==================================================
            For i = 1 To UBound(dados, 1)

                Encontrou = False

                For j = 1 To UBound(dados, 2)

                    If Not IsError(dados(i, j)) Then

                        If InStr(1, _
                            CStr(dados(i, j)), _
                            TextoBusca, _
                            vbTextCompare) > 0 Then

                            Encontrou = True
                            Exit For

                        End If

                    End If

                Next j

                '==============================================
                ' ADICIONAR REGISTRO ENCONTRADO
                '==============================================
                If Encontrou Then

                    IndiceResultado = IndiceResultado + 1

                    ' Nome da planilha de origem
                    resultado(0, IndiceResultado) = ws.Name

                    ' Dados das colunas A até K
                    For j = 1 To COL_FIM

                        resultado(j, IndiceResultado) = dados(i, j)

                    Next j

                    ' Preserva o índice original da linha
                    resultado(12, IndiceResultado) = _
                        i + LINHA_INICIAL_DADOS - 1

                    ' Coluna L existe somente em Entradas Filial
                    If ws.Name = "Entradas Filial" Then

                        resultado(13, IndiceResultado) = _
                            dados(i, COL_FIM_FILIAL)

                    Else

                        resultado(13, IndiceResultado) = vbNullString

                    End If

                End If

            Next i

        End If

    Next NomeAba

    '======================================================
    ' TRANSFERIR RESULTADOS PARA A LISTBOX
    '======================================================
    If IndiceResultado >= 0 Then

        If IndiceResultado < TotalMaximo - 1 Then

            ReDim Preserve resultado( _
                0 To 13, _
                0 To IndiceResultado _
            )

        End If

        Lst.List = TransporParaListBox(resultado)

    End If

End Sub

'==========================================================
' ABRIR REGISTRO SELECIONADO NA LISTBOX
'==========================================================
Public Sub AbrirRegistroDaLista(ByVal Lst As Object)

    Dim NomeAba As String
    Dim NumeroLinha As Long
    Dim ColFimSelecao As Long
    Dim ws As Worksheet

    If Lst.ListIndex < 0 Then Exit Sub

    NomeAba = CStr(Lst.List(Lst.ListIndex, 0))

    ' Mantém a coluna 12 como referência da linha
    NumeroLinha = CLng(Val(Lst.List(Lst.ListIndex, 12)))

    If NomeAba = vbNullString Then Exit Sub
    If NumeroLinha <= 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(NomeAba)

    ' Define o intervalo conforme a planilha
    If ws.Name = "Entradas Filial" Then

        ColFimSelecao = COL_FIM_FILIAL

    Else

        ColFimSelecao = COL_FIM

    End If

    ' Navega até o registro selecionado
    Application.GoTo ws.Range( _
        ws.Cells(NumeroLinha, COL_INICIO), _
        ws.Cells(NumeroLinha, ColFimSelecao) _
    ), False

End Sub

'==========================================================
' IDENTIFICAR ÚLTIMA LINHA UTILIZADA
'==========================================================
Private Function UltimaLinhaUsada( _
    ByVal ws As Worksheet, _
    ByVal coluna As String) As Long

    UltimaLinhaUsada = ws.Cells( _
        ws.rows.count, coluna _
    ).End(xlUp).Row

End Function

'==========================================================
' TRANSPOR MATRIZ PARA LISTBOX
'==========================================================
Private Function TransporParaListBox( _
    ByVal Matriz As Variant) As Variant

    Dim i As Long
    Dim j As Long

    Dim Temp() As Variant

    ReDim Temp( _
        0 To UBound(Matriz, 2), _
        0 To UBound(Matriz, 1) _
    )

    For i = 0 To UBound(Matriz, 1)

        For j = 0 To UBound(Matriz, 2)

            Temp(j, i) = Matriz(i, j)

        Next j

    Next i

    TransporParaListBox = Temp

End Function

'==========================================================
' IR PARA CÉLULA DA NFe
'==========================================================
Public Sub IrParaCelulaNFe( _
    ByVal NomePlanilha As String, _
    ByVal NumeroNFe As String)

    Dim ws As Worksheet
    Dim celEncontrada As Range

    If Trim(NomePlanilha) = "" Then Exit Sub
    If Trim(NumeroNFe) = "" Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(NomePlanilha)

    Set celEncontrada = ws.Cells.Find( _
        What:=NumeroNFe, _
        After:=ws.Cells(1, 1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False _
    )

    If Not celEncontrada Is Nothing Then

        ws.Activate

        Application.GoTo celEncontrada, True

    Else

        MsgBox "A NFe nº " & NumeroNFe & _
               " não foi localizada na planilha " & _
               NomePlanilha & ".", _
               vbInformation, "NFe não encontrada"

    End If

End Sub

