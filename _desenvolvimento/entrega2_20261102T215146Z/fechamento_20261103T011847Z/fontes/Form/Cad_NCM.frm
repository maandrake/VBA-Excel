VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cad_NCM 
   Caption         =   "Pesquisa NCM"
   ClientHeight    =   7935
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   15760
   OleObjectBlob   =   "Cad_NCM.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Cad_NCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Autor: Tomás Vásquez
'       www.tomasvasquez.com.br
'       www.tomasvasquez.com.br/blog
'       www.tomasvasquez.com.br/forum
'       www.tomasvasquez.com.br/cursocsharp
Option Explicit
Private Const NomePlanilha As String = "NCM e CEST"
Private Const LinhaCabecalho As Integer = 1

Private Sub FramePesquisa_Enter()

    FramePesquisa.BorderColor = &H8000000D
    FramePesquisa.ForeColor = &H8000000D
End Sub

Private Sub FramePesquisa_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    FramePesquisa.BorderColor = &H8000000A
    FramePesquisa.ForeColor = &H80000011
    
    
    ComboBoxCampos.BorderColor = &H8000000A
    ComboBoxCampos.BackColor = &H80000005
    Label_Filtro.ForeColor = &H80000011
    
    
    TextBoxFiltro.BorderColor = &H8000000A
    TextBoxFiltro.BackColor = &H80000005
    Label_Pesquisa.ForeColor = &H80000011
End Sub

Private Sub ComboBoxCampos_Enter()

    ComboBoxCampos.BorderColor = &H8000000D
    ComboBoxCampos.BackColor = &H80000002
    Label_Filtro.ForeColor = &H80000012
End Sub

Private Sub ComboBoxCampos_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    ComboBoxCampos.BorderColor = &H8000000A
    ComboBoxCampos.BackColor = &H80000005
    Label_Filtro.ForeColor = &H80000011
End Sub

Private Sub TextBoxFiltro_Enter()

    TextBoxFiltro.BorderColor = &H8000000D
    TextBoxFiltro.BackColor = &H80000002
    Label_Pesquisa.ForeColor = &H80000012
End Sub

Private Sub TextBoxFiltro_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    TextBoxFiltro.BorderColor = &H8000000A
    TextBoxFiltro.BackColor = &H80000005
    Label_Pesquisa.ForeColor = &H80000011
End Sub

Private Sub ListBoxLista_Change()

    'Clear e a carga da lista podem disparar Change sem seleção.
    If ListBoxLista.ListIndex < 0 Then Exit Sub

    TextBoxMotor.text = ListBoxLista.list(ListBoxLista.ListIndex, 4)
    TextBoxCEST.text = ListBoxLista.list(ListBoxLista.ListIndex, 1)
    TextBoxNCM.text = ListBoxLista.list(ListBoxLista.ListIndex, 2)
End Sub

Private Sub FrameResultado_Enter()

    FrameResultado.BorderColor = &H8000000D
    FrameResultado.ForeColor = &H8000000D
End Sub

Private Sub FrameResultado_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    FrameResultado.BorderColor = &H8000000A
    FrameResultado.ForeColor = &H80000011
    
    TextBoxMotor.BorderColor = &H8000000A
    TextBoxMotor.BackColor = &H80000005
    LabelMotor.ForeColor = &H80000011
    
    TextBoxCEST.BorderColor = &H8000000A
    TextBoxCEST.BackColor = &H80000005
    LabelCEST.ForeColor = &H80000011
        
    TextBoxNCM.BorderColor = &H8000000A
    TextBoxNCM.BackColor = &H80000005
    LabelNCM.ForeColor = &H80000011
End Sub

Private Sub TextBoxMotor_Enter()

    TextBoxMotor.BorderColor = &H8000000D
    TextBoxMotor.BackColor = &H80000002
    LabelMotor.ForeColor = &H80000012
End Sub

Private Sub TextBoxMotor_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    TextBoxMotor.BorderColor = &H8000000A
    TextBoxMotor.BackColor = &H80000005
    LabelMotor.ForeColor = &H80000011
End Sub

Private Sub TextBoxCEST_Enter()

    TextBoxCEST.BorderColor = &H8000000D
    TextBoxCEST.BackColor = &H80000002
    LabelCEST.ForeColor = &H80000012
End Sub

Private Sub TextBoxCEST_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    TextBoxCEST.BorderColor = &H8000000A
    TextBoxCEST.BackColor = &H80000005
    LabelCEST.ForeColor = &H80000011
End Sub

Private Sub TextBoxNCM_Enter()

    TextBoxNCM.BorderColor = &H8000000D
    TextBoxNCM.BackColor = &H80000002
    LabelNCM.ForeColor = &H80000012
End Sub

Private Sub TextBoxNCM_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    TextBoxNCM.BorderColor = &H8000000A
    TextBoxNCM.BackColor = &H80000005
    LabelNCM.ForeColor = &H80000011
End Sub

Private Sub ListBoxLista_Enter()

    ListBoxLista.BorderColor = &H8000000D
    ListBoxLista.ForeColor = &H80000012
End Sub

Private Sub ListBoxLista_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    ListBoxLista.BorderColor = &H8000000A
    ListBoxLista.ForeColor = &H80000011
End Sub

Private Sub TextBoxFiltro_Change()
    If Me.ComboBoxCampos.ListIndex <> -1 Then
        Call PreencheLista(TextBoxFiltro.text)
    End If
End Sub

Private Sub UserForm_Initialize()
    
    Call PreencheCampos
    TextBoxFiltro.SetFocus
End Sub

Private Sub PreencheCampos()
    Dim WS As Worksheet
    Dim coluna As Long
    Dim linha As Long
    Set WS = ThisWorkbook.Worksheets(NomePlanilha)
    coluna = 1
    linha = LinhaCabecalho

    With WS
        While .Cells(linha, coluna).value <> Empty
            Me.ComboBoxCampos.AddItem .Cells(linha, coluna)
            coluna = coluna + 1
        Wend
    End With
End Sub

Private Sub PreencheCabecalho(ByRef lista())
    Dim WS As Worksheet
    Dim coluna As Long
    Dim linha As Long
    Set WS = ThisWorkbook.Worksheets(NomePlanilha)
    coluna = 1
    linha = LinhaCabecalho

    With WS
        While .Cells(linha, coluna).value <> Empty
            lista(coluna - 1, 0) = .Cells(linha, coluna)
            coluna = coluna + 1
        Wend
    End With
End Sub

Private Sub PreencheLista(ByVal TextoDigitado As String)
    Dim WS As Worksheet
    Dim i As Long
    Dim x As Long
    Dim indiceLista As Long
    Dim coluna As Long
    Dim TextoCelula As String
    Dim quantidadeColunas As Long
    Dim lista()

    If Me.ComboBoxCampos.ListIndex < 0 Then Exit Sub
    Set WS = ThisWorkbook.Worksheets(NomePlanilha)
    quantidadeColunas = WS.UsedRange.Columns.count
    ReDim lista(quantidadeColunas, 0)

    i = LinhaCabecalho + 1
    indiceLista = 1
    coluna = Me.ComboBoxCampos.ListIndex + 1
    Call PreencheCabecalho(lista)

    ListBoxLista.Clear
    With WS
        While .Cells(i, coluna).value <> Empty
            TextoCelula = .Cells(i, coluna).value
            If UCase(Left(TextoCelula, Len(TextoDigitado))) = UCase(TextoDigitado) Then

                'Uma expansão por registro, preservando as dimensões originais.
                ReDim Preserve lista(quantidadeColunas, indiceLista)
                For x = 0 To quantidadeColunas - 1
                    lista(x, indiceLista) = .Cells(i, x + 1)
                Next

                indiceLista = indiceLista + 1
            End If
            i = i + 1
        Wend
    End With
   
    lista = Array2DTranspose(lista)

    Me.ListBoxLista.list = lista
End Sub

Function Array2DTranspose(avValues As Variant) As Variant
    Dim lThisCol As Long, lThisRow As Long
    Dim lUb2 As Long, lLb2 As Long
    Dim lUb1 As Long, lLb1 As Long
    Dim avTransposed As Variant
    If IsArray(avValues) Then
        On Error GoTo ErrFailed
        lUb2 = UBound(avValues, 2)
        lLb2 = LBound(avValues, 2)
        lUb1 = UBound(avValues, 1)
        lLb1 = LBound(avValues, 1)
        ReDim avTransposed(lLb2 To lUb2, lLb1 To lUb1)
        For lThisCol = lLb1 To lUb1
            For lThisRow = lLb2 To lUb2
                avTransposed(lThisRow, lThisCol) = avValues(lThisCol, lThisRow)
            Next
        Next
    End If
    Array2DTranspose = avTransposed
    Exit Function
ErrFailed:
    Debug.Print Err.Description
    Debug.Assert False
    Array2DTranspose = Empty
    Exit Function
    Resume
End Function

Private Sub CommandButtonFechar_Click()

    Cad_NCM.Hide
End Sub

Private Sub CommandButtonFechar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

