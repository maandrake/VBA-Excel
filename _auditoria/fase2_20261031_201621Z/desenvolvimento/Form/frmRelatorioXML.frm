VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRelatorioXML 
   Caption         =   "Relatório"
   ClientHeight    =   11440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   19780
   OleObjectBlob   =   "frmRelatorioXML.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRelatorioXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CarregarRelatorio( _
    ByVal resultados As Collection, _
    ByVal quantidadeErros As Long, _
    ByVal quantidadeAvisos As Long, _
    ByVal arquivoXML As String)

    Dim i As Long
    Dim item As Variant
    Dim linha As Long

    Me.CAPTION = "Relatório de validação da NF-e"

    lblArquivo.CAPTION = _
        "Arquivo: " & arquivoXML

    With lstResultados

        .Clear
        .ColumnCount = 5
        .ColumnHeads = False

        .ColumnWidths = _
            "200 pt;150 pt;150 pt;100 pt;345 pt"

        .AddItem "CAMPO"
        .List(0, 1) = "ESPERADO / FORMULÁRIO"
        .List(0, 2) = "ENCONTRADO NO XML"
        .List(0, 3) = "RESULTADO"
        .List(0, 4) = "OBSERVAÇÃO"

        For i = 1 To resultados.count

            item = resultados(i)

            .AddItem CStr(item(0))

            linha = .ListCount - 1

            .List(linha, 1) = CStr(item(1))
            .List(linha, 2) = CStr(item(2))
            .List(linha, 3) = CStr(item(3))
            .List(linha, 4) = CStr(item(4))

        Next i

    End With

    If quantidadeErros > 0 Then

        lblResumo.CAPTION = _
            "XML COM DIVERGÊNCIAS | Erros: " & _
            quantidadeErros & _
            " | Avisos: " & quantidadeAvisos

        lblResumo.ForeColor = RGB(192, 0, 0)

    ElseIf quantidadeAvisos > 0 Then

        lblResumo.CAPTION = _
            "XML VALIDADO COM AVISOS | Avisos: " & _
            quantidadeAvisos

        lblResumo.ForeColor = RGB(192, 96, 0)

    Else

        lblResumo.CAPTION = _
            "XML VALIDADO | Todos os dados conferem"

        lblResumo.ForeColor = RGB(0, 128, 0)

    End If

End Sub

Private Sub CommandButton_Salvar_Click()

    Dim dlg As FileDialog
    Dim caminhoArquivo As String
    Dim numeroArquivo As Integer
    Dim linha As Long
    Dim coluna As Long
    Dim textoLinha As String

    On Error GoTo TrataErro

    'Verifica se existem resultados
    If Me.lstResultados.ListCount = 0 Then
        MsgBox "Não existem resultados para salvar.", _
               vbExclamation, _
               "Salvar relatório"
        Exit Sub
    End If

    'Abre janela para escolher onde salvar
    Set dlg = Application.FileDialog(msoFileDialogSaveAs)

    With dlg

        .Title = "Salvar relatório de validação"

        .InitialFileName = _
            "Relatorio_Validacao_NFe_" & _
            Format(Now, "yyyymmdd_hhnnss") & ".txt"

        If .Show <> -1 Then Exit Sub

        caminhoArquivo = .SelectedItems(1)

    End With

    'Garante extensão TXT
    If LCase$(Right$(caminhoArquivo, 4)) <> ".txt" Then
        caminhoArquivo = caminhoArquivo & ".txt"
    End If

    numeroArquivo = FreeFile

    Open caminhoArquivo For Output As #numeroArquivo

    '=========================================================
    ' CABEÇALHO DO RELATÓRIO
    '=========================================================

    Print #numeroArquivo, "RELATÓRIO DE VALIDAÇÃO DA NF-e"
    Print #numeroArquivo, String(100, "=")
    Print #numeroArquivo, Me.lblArquivo.CAPTION
    Print #numeroArquivo, Me.lblResumo.CAPTION
    Print #numeroArquivo, String(100, "=")
    Print #numeroArquivo, ""

    '=========================================================
    ' CONTEÚDO DA LISTBOX
    '=========================================================

    For linha = 0 To Me.lstResultados.ListCount - 1

        textoLinha = vbNullString

        For coluna = 0 To Me.lstResultados.ColumnCount - 1

            If coluna > 0 Then
                textoLinha = textoLinha & vbTab
            End If

            textoLinha = textoLinha & _
                         TextoSeguroListBox( _
                             Me.lstResultados.List(linha, coluna))

        Next coluna

        Print #numeroArquivo, textoLinha

    Next linha

    Close #numeroArquivo

    MsgBox _
        "Relatório salvo com sucesso:" & vbCrLf & vbCrLf & _
        caminhoArquivo, _
        vbInformation, _
        "Salvar relatório"

    Exit Sub

TrataErro:

    On Error Resume Next

    If numeroArquivo > 0 Then
        Close #numeroArquivo
    End If

    MsgBox _
        "Não foi possível salvar o relatório." & vbCrLf & vbCrLf & _
        "Erro " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "Salvar relatório"

End Sub

Private Sub CommandButton_Salvar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

Private Function TextoSeguroListBox(ByVal valor As Variant) As String

    If IsNull(valor) Or IsEmpty(valor) Then
        TextoSeguroListBox = vbNullString
    Else
        TextoSeguroListBox = CStr(valor)
    End If

End Function

Private Sub cmdFechar_Click()

    Unload Me
End Sub

Private Sub cmdFechar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

Private Sub cmdFechar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    cmdFechar.PicturePosition = fmPicturePositionRightTop
End Sub

Private Sub cmdFechar_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
        
    cmdFechar.PicturePosition = fmPicturePositionCenter
End Sub


