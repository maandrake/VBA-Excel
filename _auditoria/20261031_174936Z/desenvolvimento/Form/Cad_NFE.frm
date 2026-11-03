VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cad_NFE 
   Caption         =   "Cadastro de NFe de Entrada"
   ClientHeight    =   12440
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   18420
   OleObjectBlob   =   "Cad_NFE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cad_NFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const NomePlanilha As String = "Entradas"
Private Const NomePlanilha2 As String = "Entradas Filial"
Private Const LinhaCabecalho As Integer = 3

Private Sub UserForm_Initialize()

    Dim ws As Worksheet

    LabelStatus = "Bem vindo..."

    Me.Data.value = Format(Now, "dd/mm/yyyy")
    Me.DataFilial.value = Format(Now, "dd/mm/yyyy")

    Call RemoveBotao(Me)

    ' Posiciona a planilha no início da tabela Entradas
    Set ws = ThisWorkbook.Worksheets("Entradas")
    ws.Activate
    Application.GoTo Reference:=ws.Range("A3"), Scroll:=True

    ' Inicializa a LabelData e LabelHora
    LabelData.CAPTION = Format(Now, "dd/mm/yyyy")
    LabelHora.CAPTION = Format(Now, "hh:mm")

    UF_AtualizarPainel Me

    ConfigurarListBoxPesquisa Me.ListBoxLista
    ConfigurarListBoxPesquisa Me.ListBoxLista2

End Sub

Private Sub TextBoxPesquisa_Change()
    ' Para bases muito grandes, se quiser pesquisar só a partir de 2 caracteres,
    ' descomente o bloco abaixo:
    '
    'If Len(Trim$(Me.TextBoxPesquisa.Text)) < 2 Then
    '    ConfigurarListBoxPesquisa Me.ListBoxLista
    '    Exit Sub
    'End If
    
    PesquisarEntradasRapido Me.TextBoxPesquisa.text, Me.ListBoxLista
End Sub

Private Sub TextBoxPesquisaFilial_Change()
    ' Para bases muito grandes, se quiser pesquisar só a partir de 2 caracteres,
    ' descomente o bloco abaixo:
    '
    'If Len(Trim$(Me.TextBoxPesquisa.Text)) < 2 Then
    '    ConfigurarListBoxPesquisa Me.ListBoxLista
    '    Exit Sub
    'End If
    
    PesquisarEntradasRapido Me.TextBoxPesquisaFilial.text, Me.ListBoxLista2
End Sub

Private Sub CommandButtonCopiarEntradaPFilial_Click()
    
    On Error GoTo TrataErro

    'Copia Filial -> Campos principais
    UN.text = UNFilial.text
    COD.text = CODFilial.text
    
    ComboBox_VEÍCULO.value = ComboBox_VEÍCULOFilial.value
    
    PLACA.text = PLACAFilial.text
    ComboBox_COR.value = ComboBox_CORFilial.value
    Combo_ANO_MOD.value = Combo_ANO_MODFilial.value
    txt_Valor.text = txt_ValorFilial.text
    
    DataFilial = Format(Now, "dd/mm/yyyy")

    Exit Sub

TrataErro:
    MsgBox "Erro ao copiar a entrada: " & Err.Number & " - " & Err.Description, vbExclamation, "CommandButtonCopiarEntrada"
End Sub

Private Sub CommandButtonDetranNet_Click()
    CommandButtonConsultaDetran
End Sub

Private Sub CommandButtonDetranNet_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    CommandButtonDetranNet.PicturePosition = fmPicturePositionRightTop
End Sub

Private Sub CommandButtonDetranNet_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelConsulta.Visible = True
    
    Label_N_Nfe.ForeColor = &H80000012
    FrameMenu.BorderColor = &H8000000D
End Sub

Private Sub CommandButtonDetranNet_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
        
    CommandButtonDetranNet.PicturePosition = fmPicturePositionCenter
End Sub

Private Sub CommandButtonDetranNetBIN_Click()
    
    Call CommandButtonDetranNetBIN_Handler
End Sub

Private Sub CommandButton_Excluir_Click()

'Declarar a variável Resp para receber uma resposta
Dim resp As Integer
Dim c As Range
Dim msg, Style, Title, Help, Ctxt, Response, MyString

'Fazer a busca do registro digitado pelo usuário
LabelStatus = "Confirma a exclusão da NFe..."
With Worksheets("Entradas").Range("A:A")
Set c = .Find(N_Nfe.value, LookIn:=xlValues, LookAt:=xlWhole)

If Not c Is Nothing Then

        resp = MsgBox("Tem certeza que deseja excluir esta NFe?", vbYesNo + vbQuestion + vbDefaultButton2, "Exluir NFe")
    
        'Style = vbYesNo + vbInformation + vbDefaultButton2 ' Define os botões.
        'Title = "Exluir NFe" ' Define o título.
        'Response = MsgBox(Msg, Style, Title)
 
    If resp = vbYes Then
         c.Select
         ActiveCell.Resize(1, 14).Select
         Selection.Delete Shift:=xlUp
         Application.CalculateFull
         'Comando para deletar a linha toda
         'Selection.EntireRow.Delete
         'Limpar as caixas de texto
         N_Nfe.value = Empty
         ComboBox_Operação.value = Empty
         ComboBox_OutroEnvio.value = Empty
         UN.value = Empty
         LabelTabela1.CAPTION = Empty
            LabelTabela1.BorderStyle = fmBorderStyleNone
            LabelTabela1.BorderColor = &H80000006
            LabelTabela.ForeColor = &H80000006
         COD.value = Empty
         ComboBox_VEÍCULO.value = Empty
         PLACA.value = Empty
         Combo_ANO_MOD.value = Empty
         ComboBox_COR.value = Empty
         txt_Valor.value = Empty
         Data.value = Empty
         ListBoxLista.Clear
         'Colocar o foco na primeira caixa de texto
         ComboBox_Operação.SetFocus
         ActiveWorkbook.Save
         Selection.End(xlDown).Select
         
    Else
        
        msg = "Esta NFe não será excluído!" ' Define a mensagem.
        Style = vbOKOnly + vbInformation + vbDefaultButton2 ' Define os botões.
        Title = "Exluir NFe" ' Define o título.
        Response = MsgBox(msg, Style, Title)
    End If
    
        LabelStatus = "Nfe excluida com sucesso..."
    
Else
     MsgBox "NFe não encontrada!"
     LabelStatus = "NFe não encontrada..."
End If
End With
Exit Sub
End Sub

Private Sub CommandButton_Excluir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

Private Sub CommandButton_ExcluirFilial_Click()

    Dim resp As Integer
    Dim c As Range
    Dim cAnual As Range
    Dim msg, Style, Title, Response
    Dim nfeKey As String
    
    nfeKey = N_NfeFilial.value '<<< guarda antes de limpar os campos
    
    LabelStatus = "Confirma a exclusão da NFe..."

    With Worksheets("Entradas Filial").Range("A:A")
        Set c = .Find(What:=nfeKey, LookIn:=xlValues, LookAt:=xlWhole)

        If Not c Is Nothing Then

            resp = MsgBox("Tem certeza que deseja excluir esta NFe?", vbYesNo + vbQuestion + vbDefaultButton2, "Exluir NFe")

            If resp = vbYes Then

                'Exclui na planilha Entradas Filial (A:M)
                c.Resize(1, 13).Delete Shift:=xlUp

                'Exclui também na planilha Entradas Anual Filial (A:M)
                With Worksheets("Entradas Anual Filial").Range("A:A")
                    Set cAnual = .Find(What:=nfeKey, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not cAnual Is Nothing Then
                        cAnual.Resize(1, 16).Delete Shift:=xlUp
                    End If
                End With

                Application.CalculateFull

                'Limpar as caixas de texto
                N_NfeFilial.value = Empty
                ComboBox_OperaçãoFilial.value = Empty
                UNFilial.value = Empty
                LabelTabelaFilial1.CAPTION = Empty
                    LabelTabelaFilial1.BorderStyle = fmBorderStyleNone
                    LabelTabelaFilial1.BorderColor = &H80000006
                    LabelTabelaFilial.ForeColor = &H80000006
                CODFilial.value = Empty
                ComboBox_VEÍCULOFilial.value = Empty
                PLACAFilial.value = Empty
                Combo_ANO_MODFilial.value = Empty
                ComboBox_CORFilial.value = Empty
                txt_ValorFilial.value = Empty
                DataFilial.value = Empty

                ListBoxLista2.Clear

                ComboBox_OperaçãoFilial.SetFocus

                LabelStatus = "NFe excluída com sucesso..."

            Else
                msg = "Esta NFe não será excluída!"
                Style = vbOKOnly + vbInformation + vbDefaultButton2
                Title = "Exluir NFe"
                Response = MsgBox(msg, Style, Title)

                LabelStatus = "Exclusão cancelada."
            End If

        Else
            MsgBox "NFe não encontrada!"
            LabelStatus = "NFe não encontrada..."
        End If
    End With

End Sub

Private Sub CommandButton_ExcluirFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

Private Sub CommandButton_Gravar_Click()

    Dim atualizacaoAnterior As Boolean
    Dim numeroErro As Long, descricaoErro As String
    atualizacaoAnterior = Application.ScreenUpdating
    On Error GoTo TrataErroGravacao

'Verifica se existe células com o mesmo valor
Dim vernumeronfe As Double
Dim planilha As Worksheet
Set planilha = ThisWorkbook.Worksheets("Entradas")
vernumeronfe = WorksheetFunction.CountIf(planilha.Range("A:A"), N_Nfe.text)

If N_Nfe.text = "" Then
        Exit Sub
End If

If vernumeronfe > 0 Then
        MsgBox "Numero da NFe já cadastrado!", vbCritical, "Numero da NFe"
        LabelStatus = "Numero da NFe já cadastrado!"
        N_Nfe.BorderColor = &HFF&
        N_Nfe.BackColor = &H8080FF
        Label_N_Nfe.ForeColor = &HC0&
        Exit Sub
End If

Application.ScreenUpdating = False
'Ativar a primeira planilha
ThisWorkbook.Worksheets("Entradas").Activate
'Selecionar a célula
Range("A4").Select

'Procurar a primeira célula vazia
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
  End If
Loop Until IsEmpty(ActiveCell) = True

'Carregar os dados digitados nas caixas de texto para a planilha
ActiveCell.value = N_Nfe.value
ActiveCell.Offset(0, 1).value = ComboBox_Operação.value
ActiveCell.Offset(0, 2).value = UN.value
ActiveCell.Offset(0, 3).value = ComboBox_VEÍCULO.value
ActiveCell.Offset(0, 4).value = COD.value
ActiveCell.Offset(0, 5).value = PLACA.value
ActiveCell.Offset(0, 6).value = Combo_ANO_MOD.value
ActiveCell.Offset(0, 7).value = ComboBox_COR.value
ActiveCell.Offset(0, 8).value = CDbl(txt_Valor.value)
ActiveCell.Offset(0, 9).value = CDate(Data.value)
ActiveCell.Offset(0, 10).value = ComboBox_OutroEnvio.value

'Ativar a primeira planilha
ThisWorkbook.Worksheets("Entradas Anual").Activate
'Selecionar a célula
Range("A4").Select

'Procurar a primeira célula vazia
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
  End If
Loop Until IsEmpty(ActiveCell) = True

'Carregar os dados digitados nas caixas de texto para a planilha
ActiveCell.value = N_Nfe.value
ActiveCell.Offset(0, 1).value = ComboBox_Operação.value
ActiveCell.Offset(0, 2).value = UN.value
ActiveCell.Offset(0, 3).value = ComboBox_VEÍCULO.value
ActiveCell.Offset(0, 4).value = COD.value
ActiveCell.Offset(0, 5).value = PLACA.value
ActiveCell.Offset(0, 6).value = Combo_ANO_MOD.value
ActiveCell.Offset(0, 7).value = ComboBox_COR.value
ActiveCell.Offset(0, 8).value = CDbl(txt_Valor.value)
ActiveCell.Offset(0, 9).value = CDate(Data.value)

'Limpar as caixas de texto
N_Nfe.value = Empty
ComboBox_Operação.value = "Compra"
ComboBox_OutroEnvio.value = ""
UN.value = "1"
ComboBox_OutroEnvio.text = "NÃO"
Me.LabelTabela1.CAPTION = ""
    Me.LabelTabela1.BorderStyle = fmBorderStyleNone
    Me.LabelTabela1.BorderColor = &H80000006
    Me.LabelTabela.ForeColor = &H80000006
N_Nfe.value = ""
COD.value = Empty
ComboBox_VEÍCULO.value = Empty
PLACA.value = Empty
Combo_ANO_MOD.value = Empty
ComboBox_COR.value = Empty
txt_Valor.value = Empty
Data.value = Format(Now, "dd/mm/yyyy")

'Colocar o foco na primeira caixa de texto e retorna a planilha Entradas
ThisWorkbook.Worksheets("Entradas").Select
Range("A3").Select
ThisWorkbook.Save
ComboBox_Operação.SetFocus
Application.ScreenUpdating = atualizacaoAnterior

    LabelStatus = "Nfe inserida com sucesso..."
    Exit Sub
TrataErroGravacao:
    numeroErro = Err.Number
    descricaoErro = Err.Description
    Application.ScreenUpdating = atualizacaoAnterior
    MsgBox "Falha ao gravar NFe (" & numeroErro & "): " & descricaoErro, vbExclamation, "Gravação de NFe"

End Sub

Private Sub CommandButton_Gravar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

Private Sub CommandButton_GravarFilial_Click()

'Verifica se existe células com o mesmo valor
Dim vernumeronfe As Double
Dim planilha As Worksheet
Set planilha = ThisWorkbook.Worksheets("Entradas Filial")
vernumeronfe = WorksheetFunction.CountIf(planilha.Range("A:A"), N_NfeFilial.text)

If N_NfeFilial.text = "" Then
        Exit Sub
End If

If vernumeronfe > 0 Then
        MsgBox "Numero da NFe já cadastrado!", vbCritical, "Numero da NFe"
        LabelStatus = "Numero da NFe já cadastrado!"
        N_NfeFilial.BorderColor = &HFF&
        N_NfeFilial.BackColor = &H8080FF
        N_NfeFilial.ForeColor = &HC0&
        Exit Sub
End If

'Ativar a primeira planilha
ThisWorkbook.Worksheets("Entradas Filial").Activate
'Selecionar a célula
Range("A4").Select

'Procurar a primeira célula vazia
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
  End If
Loop Until IsEmpty(ActiveCell) = True

'Carregar os dados digitados nas caixas de texto para a planilha
ActiveCell.value = N_NfeFilial.value
ActiveCell.Offset(0, 1).value = ComboBox_OperaçãoFilial.value
ActiveCell.Offset(0, 2).value = UNFilial.value
ActiveCell.Offset(0, 3).value = ComboBox_VEÍCULOFilial.value
ActiveCell.Offset(0, 4).value = CODFilial.value
ActiveCell.Offset(0, 5).value = PLACAFilial.value
ActiveCell.Offset(0, 6).value = Combo_ANO_MODFilial.value
ActiveCell.Offset(0, 7).value = ComboBox_CORFilial.value
ActiveCell.Offset(0, 8).value = CDbl(txt_ValorFilial.value)
ActiveCell.Offset(0, 9).value = CDate(DataFilial.value)
ActiveCell.Offset(0, 10).value = ComboBox_OutroEnvioFilial.value
'Defina o formato antes do valor para preservar os 44 dígitos.
ActiveCell.Offset(0, 11).NumberFormat = "@"
ActiveCell.Offset(0, 11).value = CStr(TextBoxChaveDeAcessoFilial.value)

'Ativar a primeira planilha
ThisWorkbook.Worksheets("Entradas Anual Filial").Activate
'Selecionar a célula
Range("A4").Select

'Procurar a primeira célula vazia
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
  End If
Loop Until IsEmpty(ActiveCell) = True

'Carregar os dados digitados nas caixas de texto para a planilha
ActiveCell.value = N_NfeFilial.value
ActiveCell.Offset(0, 1).value = ComboBox_OperaçãoFilial.value
ActiveCell.Offset(0, 2).value = UNFilial.value
ActiveCell.Offset(0, 3).value = ComboBox_VEÍCULOFilial.value
ActiveCell.Offset(0, 4).value = CODFilial.value
ActiveCell.Offset(0, 5).value = PLACAFilial.value
ActiveCell.Offset(0, 6).value = Combo_ANO_MODFilial.value
ActiveCell.Offset(0, 7).value = ComboBox_CORFilial.value
ActiveCell.Offset(0, 8).value = CDbl(txt_ValorFilial.value)
ActiveCell.Offset(0, 9).value = CDate(DataFilial.value)

'Limpar as caixas de texto
Me.N_NfeFilial.value = ""
Me.ComboBox_OperaçãoFilial = "Compra"
Me.ComboBox_OutroEnvioFilial.value = ""
Me.TextBoxChaveDeAcessoFilial.value = ""
Me.UNFilial.value = "1"
Me.LabelTabelaFilial1.CAPTION = ""
    Me.LabelTabelaFilial1.BorderStyle = fmBorderStyleNone
    Me.LabelTabelaFilial1.BorderColor = &H80000006
    Me.LabelTabelaFilial.ForeColor = &H80000006
Me.CODFilial.value = ""
Me.ComboBox_VEÍCULOFilial.value = ""
Me.PLACAFilial.value = ""
Me.Combo_ANO_MODFilial.value = ""
Me.ComboBox_CORFilial.value = ""
Me.txt_ValorFilial.value = ""
Me.DataFilial.value = Format(Now, "dd/mm/yyyy")
Me.ListBoxLista2.Clear
Me.N_NfeFilial.SetFocus
Me.TextBoxPesquisaFilial.value = ""

'Colocar o foco na primeira caixa de texto e retorna a planilha Entradas
Sheets("Entradas Filial").Select
Range("A3").Select
ComboBox_OperaçãoFilial.SetFocus

    LabelStatus = "Nfe inserida com sucesso..."
End Sub

Private Sub CommandButton_GravarFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

Private Sub CommandButton_Limpar_Click()

Me.N_Nfe.value = ""
Me.ComboBox_Operação = "Compra"
Me.ComboBox_OutroEnvio.value = ""
Me.UN.value = "1"
Me.ComboBox_OutroEnvio.text = ""
Me.LabelTabela1.CAPTION = ""
    Me.LabelTabela1.BorderStyle = fmBorderStyleNone
    Me.LabelTabela1.BorderColor = &H80000006
    Me.LabelTabela.ForeColor = &H80000006
Me.COD.value = ""
Me.ComboBox_VEÍCULO.value = ""
Me.PLACA.value = ""
Me.Combo_ANO_MOD.value = ""
Me.ComboBox_COR.value = ""
Me.txt_Valor.value = ""
Me.Data.value = Format(Now, "dd/mm/yyyy")
ListBoxLista.Clear
Me.N_Nfe.SetFocus
Me.TextBoxPesquisa.value = ""


    Sheets("Entradas").Select
    Range("A3").Select
    LabelStatus = "Todos os campos limpos..."
End Sub

Private Sub CommandButton_Limpar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

Private Sub CommandButton_LimparFilial_Click()

Me.N_NfeFilial.value = ""
Me.ComboBox_OperaçãoFilial = "Compra"
Me.ComboBox_OutroEnvioFilial.value = ""
Me.TextBoxChaveDeAcessoFilial.value = ""
Me.UNFilial.value = "1"
Me.LabelTabelaFilial1.CAPTION = ""
    Me.LabelTabelaFilial1.BorderStyle = fmBorderStyleNone
    Me.LabelTabelaFilial1.BorderColor = &H80000006
    Me.LabelTabelaFilial.ForeColor = &H80000006
Me.CODFilial.value = ""
Me.ComboBox_VEÍCULOFilial.value = ""
Me.PLACAFilial.value = ""
Me.Combo_ANO_MODFilial.value = ""
Me.ComboBox_CORFilial.value = ""
Me.txt_ValorFilial.value = ""
Me.DataFilial.value = Format(Now, "dd/mm/yyyy")
Me.ListBoxLista2.Clear
Me.N_NfeFilial.SetFocus
Me.TextBoxPesquisaFilial.value = ""

    Sheets("Entradas Filial").Select
    Range("A3").Select
    LabelStatus = "Todos os campos limpos..."
End Sub

Private Sub CommandButton_LimparFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

Private Sub MultiPage1_Change()

    Dim ws As Worksheet

    If MultiPage1.SelectedItem.Index = 0 Then
        Set ws = ThisWorkbook.Worksheets("Entradas")
        ws.Activate
        Application.GoTo Reference:=ws.Range("A3"), Scroll:=True
    ElseIf MultiPage1.SelectedItem.Index = 1 Then
        Set ws = ThisWorkbook.Worksheets("Entradas Filial")
        ws.Activate
        Application.GoTo Reference:=ws.Range("A3"), Scroll:=True
    End If
End Sub

Private Sub CommandButtonCadNFESaida_Click()

    Cad_NFE.Hide
        Sheets("Saidas").Select
    Cad_NFE_Saida.Show
End Sub

Private Sub CommandButtonCadNFESaida_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
    LabelSaidas.Visible = True
End Sub

Private Sub CommandButtonCadNFESaida_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    CommandButtonCadNFESaida.PicturePosition = fmPicturePositionRightTop
End Sub

Private Sub CommandButtonCadNFESaida_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
        
    CommandButtonCadNFESaida.PicturePosition = fmPicturePositionCenter
End Sub

Private Sub CommandButtonFechar_Click()

    Cad_NFE.Hide
    EncerrarSessaoDetranNFE
End Sub

Private Sub CommandButtonFechar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

Private Sub CommandButtonFechar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    CommandButtonFechar.PicturePosition = fmPicturePositionRightTop
End Sub

Private Sub CommandButtonFechar_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
        
    CommandButtonFechar.PicturePosition = fmPicturePositionCenter
End Sub

Private Sub CommandButtonAlterar_Click()

'Declarar a variável Resp para receber uma resposta
Dim resp As Integer
Dim c As Range
Dim msg, Style, Title, Help, Ctxt, Response, MyString

'Fazer a busca do registro digitado pelo usuário
With Worksheets("Entradas").Range("A:A")
Set c = .Find(TextBoxPesquisa, LookIn:=xlValues, LookAt:=xlWhole)

If Not c Is Nothing Then
        
        LabelStatus = "Confirma a alteração da nota de entrada..."
        resp = MsgBox("Tem certeza que deseja alterar este Serviço Prestado?", vbYesNo + vbQuestion + vbDefaultButton2, "Alterar nota de entrada")
    
    If resp = vbYes Then
        LabelStatus = "Solicitação iniciada..."
        MouseCursor (32514)
        c.Select
        ActiveCell.Resize(1, 11).Select
        Selection.ClearContents 'Shift:=xlUp
        Application.CalculateFull
                       
                'Ativar a primeira planilha
                ThisWorkbook.Worksheets("Entradas").Activate
                'Selecionar a célula
                Range("A4").Select
                
                'Procurar a primeira célula vazia
                Do
                  If Not (IsEmpty(ActiveCell)) Then
                      ActiveCell.Offset(1, 0).Select
                  End If
                Loop Until IsEmpty(ActiveCell) = True
                
                'Carregar os dados digitados nas caixas de texto para a planilha
                ActiveCell.value = N_Nfe.value
                ActiveCell.Offset(0, 1).value = ComboBox_Operação.value
                ActiveCell.Offset(0, 2).value = UN.value
                ActiveCell.Offset(0, 3).value = ComboBox_VEÍCULO.value
                ActiveCell.Offset(0, 4).value = COD.value
                ActiveCell.Offset(0, 5).value = PLACA.value
                ActiveCell.Offset(0, 6).value = Combo_ANO_MOD.value
                ActiveCell.Offset(0, 7).value = ComboBox_COR.value
                ActiveCell.Offset(0, 8).value = CDbl(txt_Valor.value)
                ActiveCell.Offset(0, 9).value = CDate(Data.value)
                ActiveCell.Offset(0, 10).value = ComboBox_OutroEnvio.value
                
                'Ativar a primeira planilha
                ThisWorkbook.Worksheets("Entradas Anual").Activate
                
                        ActiveCell.Resize(1, 10).Select
                        Selection.ClearContents 'Shift:=xlUp
                        Application.CalculateFull
                
                'Selecionar a célula
                Range("A4").Select
                
                'Procurar a primeira célula vazia
                Do
                  If Not (IsEmpty(ActiveCell)) Then
                      ActiveCell.Offset(1, 0).Select
                  End If
                Loop Until IsEmpty(ActiveCell) = True
                
                'Carregar os dados digitados nas caixas de texto para a planilha
                ActiveCell.value = N_Nfe.value
                ActiveCell.Offset(0, 1).value = ComboBox_Operação.value
                ActiveCell.Offset(0, 2).value = UN.value
                ActiveCell.Offset(0, 3).value = ComboBox_VEÍCULO.value
                ActiveCell.Offset(0, 4).value = COD.value
                ActiveCell.Offset(0, 5).value = PLACA.value
                ActiveCell.Offset(0, 6).value = Combo_ANO_MOD.value
                ActiveCell.Offset(0, 7).value = ComboBox_COR.value
                ActiveCell.Offset(0, 8).value = CDbl(txt_Valor.value)
                ActiveCell.Offset(0, 9).value = CDate(Data.value)
                
                'Limpar as caixas de texto
                Me.N_Nfe.value = ""
                Me.ComboBox_Operação = "Compra"
                Me.ComboBox_OutroEnvio.value = ""
                Me.UN.value = "1"
                Me.COD.value = ""
                Me.ComboBox_VEÍCULO.value = ""
                Me.PLACA.value = ""
                Me.Combo_ANO_MOD.value = ""
                Me.ComboBox_COR.value = ""
                Me.txt_Valor.value = ""
                Me.Data.value = Format(Now, "dd/mm/yyyy")
                ListBoxLista.Clear
                LabelStatus = "Todos os campos limpos..."

        'Colocar o foco na primeira caixa de texto e retorna a planilha Entradas
        Sheets("Entradas").Select
        N_Nfe.SetFocus

    'Salva toda panilha do Excel
    LabelStatus = "Salvando panilha..."
    ActiveWorkbook.Save
    'Redimenciona automaticamente todas as células da planilha
    LabelStatus = "Redimencionando células..."
    ThisWorkbook.Worksheets("Entradas").Cells.EntireColumn.AutoFit
    Range("A3").End(xlDown).Offset(1, 0).Select
    LabelStatus = "Nota de entrada alterado com sucesso..."
         
    Else
                
        LabelStatus = "Nota de entrada não foi alterado..."
        Range("A3").End(xlDown).Offset(1, 0).Select
        msg = "Esta nota de entrada não será alterado!" ' Define a mensagem.
        Style = vbOKOnly + vbInformation + vbDefaultButton2 ' Define os botões.
        Title = "Alterar nota de entrada" ' Define o título.
        Response = MsgBox(msg, Style, Title)

    End If
    
    Else
        
        LabelStatus = "Nota de entrada não encontrado..."
        Range("A3").End(xlDown).Offset(1, 0).Select
        msg = "Nota de entrada não encontrado!" ' Define a mensagem.
        Style = vbOKOnly + vbInformation + vbDefaultButton2 ' Define os botões.
        Title = "Alterar nota de entrada" ' Define o título.
        Response = MsgBox(msg, Style, Title)
        
End If
End With
Exit Sub
End Sub

Private Sub CommandButtonCopiarEntrada_Click()
    
    Cad_NFE_Saida.MultiPage1.value = 0
    Cad_NFE_Saida.TextBox_ESTOQUE.value = Cad_NFE.UN.value
    Cad_NFE_Saida.TextBox_N_Entrada.value = Cad_NFE.N_Nfe.value
    Cad_NFE_Saida.ComboBox_OutroEnvio.value = Cad_NFE.ComboBox_OutroEnvio.value
    Cad_NFE_Saida.COD.value = Cad_NFE.COD.value
    Cad_NFE_Saida.ComboBox_VEÍCULO.value = Cad_NFE.ComboBox_VEÍCULO.value
    Cad_NFE_Saida.TextBox_PLACA.value = Cad_NFE.PLACA.value
    Cad_NFE_Saida.ComboBox_Ano.value = Cad_NFE.Combo_ANO_MOD.value
    Cad_NFE_Saida.ComboBox_COR.value = Cad_NFE.ComboBox_COR.value
    Cad_NFE_Saida.ComboBox_Por.value = "12"
    Cad_NFE_Saida.TextBox_Valor_Entrada.value = Cad_NFE.txt_Valor.value
    Cad_NFE_Saida.TextBox_Data.value = Format(Now, "dd/mm/yyyy")
    Cad_NFE.Hide
        Sheets("Saidas").Select
    Cad_NFE_Saida.Show
End Sub

Private Sub CommandButtonCopiarEntradaFilial_Click()
    
    Cad_NFE_Saida.MultiPage1.value = 1
    Cad_NFE_Saida.TextBox_ESTOQUEFilial.value = Cad_NFE.UNFilial.value
    Cad_NFE_Saida.TextBox_N_EntradaFilial.value = Cad_NFE.N_NfeFilial.value
    Cad_NFE_Saida.ComboBox_OutroEnvioFilial.value = Cad_NFE.ComboBox_OutroEnvioFilial.value
    Cad_NFE_Saida.TextBoxChaveDeAcessoFilial.value = Cad_NFE.TextBoxChaveDeAcessoFilial.value
    Cad_NFE_Saida.CODFilial.value = Cad_NFE.CODFilial.value
    Cad_NFE_Saida.ComboBox_VEÍCULOFilial.value = Cad_NFE.ComboBox_VEÍCULOFilial.value
    Cad_NFE_Saida.TextBox_PLACAFilial.value = Cad_NFE.PLACAFilial.value
    Cad_NFE_Saida.ComboBox_AnoFilial.value = Cad_NFE.Combo_ANO_MODFilial.value
    Cad_NFE_Saida.ComboBox_CORFilial.value = Cad_NFE.ComboBox_CORFilial.value
    Cad_NFE_Saida.TextBox_Valor_EntradaFilial.value = Cad_NFE.txt_ValorFilial.value
    Cad_NFE_Saida.TextBox_Valor_SaidaFilial.value = Cad_NFE.txt_ValorFilial.value
    Cad_NFE_Saida.TextBox_DataFilial.value = Format(Now, "dd/mm/yyyy")
    Cad_NFE_Saida.LabelStatus = "NFe copiada com sucesso..."
    Cad_NFE.Hide
        Sheets("Saidas Filial").Select
    Cad_NFE_Saida.Show
    
End Sub

Private Sub CommandButtonCopiarEntradaFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

Private Sub CommandButtonCopiarAutoEntradaFilial_Click()

    Dim ws As Worksheet
    Dim celula As Range
    Dim linha As Long
    Dim ValorMoeda As Currency
    Dim ValorFormatado As String

    On Error GoTo TratarErro

    Set ws = ThisWorkbook.Worksheets("Entradas Filial")

    'Pesquisa de cima para baixo, começando pela primeira linha.
    Set celula = ws.Columns("B").Find( _
        What:="Devolução", _
        After:=ws.Cells(ws.rows.count, "B"), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        MatchByte:=False, _
        SearchFormat:=False)

    If celula Is Nothing Then
        MsgBox "Nenhuma linha contendo 'Devolução' foi encontrada na coluna B.", _
               vbInformation
        Exit Sub
    End If

    linha = celula.Row

    '==========================================================
    ' VALIDA E FORMATA O VALOR MONETÁRIO
    '==========================================================
    If Not IsNumeric(ws.Cells(linha, "I").Value2) Then

        MsgBox "O valor informado na coluna I não é numérico.", _
               vbExclamation, "Valor inválido"

        Exit Sub

    End If

    ValorMoeda = CCur(ws.Cells(linha, "I").Value2)

    'Aplica a formatação monetária regional
    ValorFormatado = FormatCurrency(ValorMoeda, 2)

    '==========================================================
    ' CARREGA OS DADOS NO FORMULÁRIO Cad_NFE
    '==========================================================
    With Me

        .N_NfeFilial.value = ws.Cells(linha, "A").value
        .ComboBox_OperaçãoFilial.value = ws.Cells(linha, "B").value
        .UNFilial.value = ws.Cells(linha, "C").value
        .ComboBox_VEÍCULOFilial.value = ws.Cells(linha, "D").value
        .CODFilial.value = ws.Cells(linha, "E").value
        .PLACAFilial.value = ws.Cells(linha, "F").value
        .Combo_ANO_MODFilial.value = ws.Cells(linha, "G").value
        .ComboBox_CORFilial.value = ws.Cells(linha, "H").value

        'Valor formatado em moeda
        .txt_ValorFilial.value = ValorFormatado

        .DataFilial.value = Format( _
            CDate(ws.Cells(linha, "J").value), "dd/mm/yyyy")

        .ComboBox_OutroEnvioFilial.value = ws.Cells(linha, "K").value
        .TextBoxChaveDeAcessoFilial.value = ws.Cells(linha, "L").value

    End With

    '==========================================================
    ' CARREGA O FORMULÁRIO DE SAÍDA
    '==========================================================
    Load Cad_NFE_Saida

    With Cad_NFE_Saida

        .MultiPage1.value = 1

        .TextBox_ESTOQUEFilial.value = Me.UNFilial.value
        .TextBox_N_EntradaFilial.value = Me.N_NfeFilial.value

        .ComboBox_OutroEnvioFilial.value = Me.ComboBox_OutroEnvioFilial.value
        .TextBoxChaveDeAcessoFilial.value = Me.TextBoxChaveDeAcessoFilial.value

        .CODFilial.value = Me.CODFilial.value
        .ComboBox_VEÍCULOFilial.value = Me.ComboBox_VEÍCULOFilial.value
        .TextBox_PLACAFilial.value = Me.PLACAFilial.value
        .ComboBox_AnoFilial.value = Me.Combo_ANO_MODFilial.value
        .ComboBox_CORFilial.value = Me.ComboBox_CORFilial.value

        '======================================================
        ' VALORES MONETÁRIOS FORMATADOS
        '======================================================
        .TextBox_Valor_EntradaFilial.value = ValorFormatado

        .TextBox_Valor_SaidaFilial.value = ValorFormatado

        'Data atual
        .TextBox_DataFilial.value = Format(Date, "dd/mm/yyyy")

    End With

    '==========================================================
    ' FINALIZA A OPERAÇÃO
    '==========================================================
    Me.Hide

    Cad_NFE_Saida.LabelStatus.CAPTION = _
        "NFe copiada com sucesso..."

    Cad_NFE_Saida.Show

    Exit Sub

TratarErro:

    MsgBox "Não foi possível copiar os dados." & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, _
           vbExclamation

End Sub

Private Sub CommandButtonCopiarAutoEntradaFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    MouseCursor (32649)
End Sub

Private Sub CommandButtonFiltroCompra_Click()
    
    LabelStatus = "Filtrar NFe por operação..."
    UserFormFiltro.Show
End Sub

Private Sub CommandButtonFiltroCompra_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonAlterar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub
Private Sub CommandButtonLerXML_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonFiltroCompraFilial_Click()
    
    LabelStatus = "Filtrar NFe por operação..."
    UserFormFiltro.Show
End Sub

Private Sub CommandButtonFiltroCompraFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonNCM_Click()
    
    LabelStatus = "Pesquisa NCM..."
    Cad_NCM.Show
End Sub

Private Sub Command_EntradaFilial_Click()

    LabelStatus = "Planilha Entradas filial..."
    Sheets("Entradas Filial").Select
End Sub

Private Sub CommandButton_SaidasFilial_Click()
    
    LabelStatus = "Planilha Saidas filial..."
    Sheets("Saidas Filial").Select
End Sub

Private Sub Command_EntradaFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButton_SaidasFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonCopiarEntrada_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonNCM_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonNCMFilial_Click()

    Cad_NCM.Show
End Sub

Private Sub CommandButtonNCMFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonChaveR_Click()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim celula As Range
    Dim chaveOriginal As String
    Dim chaveCorrigida As String
    Dim totalCorrigido As Long
    Dim totalNumerico As Long
    Dim estadoEventos As Boolean
    Dim estadoTela As Boolean

    On Error GoTo TratarErro

    'Preserva as configurações atuais do Excel
    estadoEventos = Application.EnableEvents
    estadoTela = Application.ScreenUpdating

    Set ws = ThisWorkbook.Worksheets("Entradas Filial")

    'Localiza a última linha preenchida da coluna L
    ultimaLinha = ws.Cells(ws.rows.count, "L").End(xlUp).Row

    If ultimaLinha < 2 Then
        MsgBox "Nenhuma chave de acesso encontrada.", _
               vbInformation, "Chaves de Acesso"
        Exit Sub
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Define a coluna L como texto
    ws.Columns("L").NumberFormat = "@"

    totalCorrigido = 0
    totalNumerico = 0

    'Percorre todas as chaves da coluna L
    For Each celula In ws.Range("L2:L" & ultimaLinha).Cells

        If Not celula.HasFormula Then

            If Not IsError(celula.Value2) Then

                If VarType(celula.Value2) = vbString Then

                    chaveOriginal = CStr(celula.Value2)

                    If Len(chaveOriginal) > 0 Then

                        chaveCorrigida = chaveOriginal

                        'Remove espaços comuns e especiais
                        chaveCorrigida = Replace(chaveCorrigida, " ", "")
                        chaveCorrigida = Replace(chaveCorrigida, ChrW(160), "")
                        chaveCorrigida = Replace(chaveCorrigida, ChrW(8239), "")

                        'Remove tabulações e quebras de linha
                        chaveCorrigida = Replace(chaveCorrigida, vbTab, "")
                        chaveCorrigida = Replace(chaveCorrigida, vbCr, "")
                        chaveCorrigida = Replace(chaveCorrigida, vbLf, "")

                        'Atualiza somente quando necessário
                        If chaveOriginal <> chaveCorrigida Then

                            celula.Value2 = chaveCorrigida
                            totalCorrigido = totalCorrigido + 1

                        End If

                        'Ignora o aviso "Número armazenado como texto"
                        celula.Errors(xlNumberAsText).Ignore = True

                    End If

                ElseIf Len(CStr(celula.Value2)) > 0 Then

                    'Não altera números que possam ter perdido precisão
                    totalNumerico = totalNumerico + 1

                End If

            End If

        End If

    Next celula

    Application.EnableEvents = estadoEventos
    Application.ScreenUpdating = estadoTela

    MsgBox totalCorrigido & _
           " chave(s) de acesso corrigida(s)!" & _
           vbCrLf & vbCrLf & _
           "Avisos de números armazenados como texto ignorados." & _
           IIf(totalNumerico > 0, vbCrLf & _
           totalNumerico & " célula(s) numérica(s) não alterada(s).", ""), _
           vbInformation, "Chaves de Acesso NF-e"

    Exit Sub

TratarErro:

    Application.EnableEvents = estadoEventos
    Application.ScreenUpdating = estadoTela

    MsgBox "Erro ao corrigir as chaves de acesso: " & _
           Err.Description, vbCritical, "Erro"

End Sub

Private Sub CommandButtonChaveR_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub FrameMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    LabelBakupBakup.Visible = False
    LabelSaidas.Visible = False
    LabelConsulta.Visible = False
End Sub

Private Sub Command_Entrada_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButton_Saidas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub Command_Entrada_Click()
    
    LabelStatus = "Planilha Entradas matriz..."
    Sheets("Entradas").Select
End Sub

Private Sub CommandButton_Saidas_Click()
    
    LabelStatus = "Planilha Saidas matriz..."
    Sheets("Saidas").Select
End Sub

Private Sub CommandButtonSalvarTudo_Click()

    LabelStatus = "Deseja iniciar o Backup?"

    Range("u1") = Range("u1") + 1
    Dim msg, Style, Title, Help, Ctxt, Response, MyString
    msg = "Deseja Salvar e efetuar bakup do Excel?" ' Define a mensagem.
    Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define os botões.
    Title = "Salvar Arquivo" ' Define o título.
    Response = MsgBox(msg, Style, Title)
    If Response = vbYes Then ' O usuário escolheu Sim.
    LabelStatus = "Backup iniciado"
        MyString = "Sim" ' Executa alguma ação.
            Sheets("Menu Principal").Select
            ThisWorkbook.Save
            Call SaveWorkbookBackup
    Else
    MyString = "Não" ' Nada é executado.
    'não salva
    End If
    
    LabelStatus = "Backup finalizado..."
    
End Sub

Private Sub CommandButtonSalvarTudo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelBakupBakup.Visible = True
End Sub

Private Sub CommandButtonSalvarTudo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    CommandButtonSalvarTudo.PicturePosition = fmPicturePositionRightTop
End Sub

Private Sub CommandButtonSalvarTudo_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
        
    CommandButtonSalvarTudo.PicturePosition = fmPicturePositionCenter
End Sub

Private Sub Data_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
    
Data.MaxLength = 10 '10/10/2014
    Select Case KeyAscii
      Case 8       'Aceita o BACK SPACE
      Case 13: SendKeys "{TAB}"    'Emula o TAB
      Case 48 To 57
         If Data.SelStart = 2 Then Data.SelText = "/"
         If Data.SelStart = 5 Then Data.SelText = "/"
      Case Else: KeyAscii = 0     'Ignora os outros caracteres
   
   End Select
End Sub

Private Sub DataFilial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
    
Data.MaxLength = 10 '10/10/2014
    Select Case KeyAscii
      Case 8       'Aceita o BACK SPACE
      Case 13: SendKeys "{TAB}"    'Emula o TAB
      Case 48 To 57
         If DataFilial.SelStart = 2 Then DataFilial.SelText = "/"
         If DataFilial.SelStart = 5 Then DataFilial.SelText = "/"
      Case Else: KeyAscii = 0     'Ignora os outros caracteres
   
   End Select
End Sub

Private Sub txt_Valor_AfterUpdate()
    
    txt_Valor.text = Format(txt_Valor.text, "Currency")
    txt_Valor.ForeColor = &HFF
End Sub

Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)

  If CloseMode = vbFormControlMenu Then
    cancel = True
    Range("u1") = Range("u1") + 1
    Dim msg, Style, Title, Help, Ctxt, Response, MyString
    msg = "Por favor, use o botão Fechar em Ações" ' Define a mensagem.
    Style = vbOKOnly + vbInformation + vbDefaultButton2 ' Define os botões.
    Title = "Fechar Cadastro" ' Define o título.
    Response = MsgBox(msg, Style, Title)
  End If
End Sub

Private Sub CommandButtonLerXML_Click()

    Dim xmlFile As String
    Dim xmlDoc As Object 'MSXML2.DOMDocument60 ou versão correspondente
    
        LabelStatus = "Selecionar o arquivo XML..."
    
    'Selecionar o arquivo XML usando o diálogo "Abrir"
    xmlFile = Application.GetOpenFilename("Arquivos XML (*.xml), *.xml")
    
    If xmlFile = "False" Then
        'Usuário cancelou a seleção do arquivo
            LabelStatus = "Operação cancelada..."
        Exit Sub
    End If
    
    'Criar o objeto DOMDocument e carregar o arquivo XML
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0") 'ou versão correspondente
    xmlDoc.async = False
    xmlDoc.Load xmlFile
    
    'Verificar se ocorreu algum erro ao carregar o arquivo XML
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "Erro ao carregar o arquivo XML."
            LabelStatus = "Erro ao carregar o XML..."
        Exit Sub
    End If
    
    'Definir o namespace padrão para facilitar as consultas XPath
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:ns='http://www.portalfiscal.inf.br/nfe'"
    
    'Preencher as TextBoxes com os dados do XML
    Me.N_Nfe.value = xmlDoc.SelectSingleNode("//ns:nNF").text
    Me.COD.value = xmlDoc.SelectSingleNode("//ns:cProd").text
    
    ' Converter o valor para um número (tipo Double) e substituir o ponto por vírgula
    Dim valor As Double
    valor = CDbl(Replace(xmlDoc.SelectSingleNode("//ns:vNF").text, ".", ","))
    
    ' Aplicar a formatação e inserir na TextBox
    Me.txt_Valor.value = "R$ " & Format(valor, "#,##0.00")
    
    'Extrair a parte "Marca/Modelo Placa Renavam Chassi Ano/Mod Cor" do elemento <xProd>
    Dim xProdText As String
    xProdText = xmlDoc.SelectSingleNode("//ns:xProd").text
    
    Dim marcaModelo As String
    Dim PLACA As String
    Dim AnoMod As String
    Dim tudoAposCor As String
    Dim startPosition As Long
    Dim endPosition As Long
    Dim startPlaca As Long
    Dim endPlaca As Long
    Dim startChassi As Long
    Dim endChassi As Long
    Dim startAnoMod As Long
    Dim endAnoMod As Long
    Dim startCor As Long
    
    startPosition = InStr(xProdText, "Marca/Modelo ") + Len("Marca/Modelo ")
    endPosition = InStr(startPosition, xProdText, " Placa") - 1
    startPlaca = endPosition + Len(" Placa") + 2
    endPlaca = InStr(startPlaca, xProdText, " Renavam") - 1
    startChassi = endPlaca + Len(" Renavam") + 1
    endChassi = InStr(startChassi, xProdText, " Ano/Mod") - 1
    startAnoMod = endChassi + Len(" Ano/Mod ") + 4
    startCor = InStr(startAnoMod, xProdText, " Cor") + Len(" Cor ")
    
    If startPosition > 0 And endPosition > startPosition And startPlaca > endPosition And endPlaca > startPlaca _
    And startChassi > endPlaca And endChassi > startChassi And startAnoMod > endChassi And startCor > startAnoMod Then
        marcaModelo = Mid(xProdText, startPosition, endPosition - startPosition + 1)
        PLACA = Mid(xProdText, startPlaca, endPlaca - startPlaca + 1)
        AnoMod = Mid(xProdText, startAnoMod, startCor - startAnoMod - 5)
        tudoAposCor = Mid(xProdText, startCor)
    Else
        marcaModelo = "Não encontrado"
        PLACA = "Não encontrado"
        AnoMod = "Não encontrado"
        tudoAposCor = "Não encontrado"
    End If
    
    Me.ComboBox_VEÍCULO.value = marcaModelo
    Me.PLACA.value = PLACA
    Me.Combo_ANO_MOD.value = AnoMod
    Me.ComboBox_COR.value = tudoAposCor

    Dim anoDigitado As String
    Dim anoConvertido As Integer

    ' Obtém o ano digitado pelo usuário da TextBox Combo_ANO_MOD
    anoDigitado = Combo_ANO_MOD.value
    
    ' Verifica se o ano digitado possui o formato "yy"
    If Len(anoDigitado) = 2 And IsNumeric(anoDigitado) Then
        ' Converte o ano de "yy" para um valor inteiro
        anoConvertido = CInt(anoDigitado)
        
        ' Determina o século correto (1900's ou 2000's) com base no valor do ano
        If anoConvertido >= 30 And anoConvertido <= 99 Then
            anoConvertido = 1900 + anoConvertido
        ElseIf anoConvertido >= 0 And anoConvertido <= 29 Then
            anoConvertido = 2000 + anoConvertido
        Else
            ' Ano inválido, exibe uma mensagem de erro e limpa o valor da TextBox
            MsgBox "Digite um ano válido no formato 'yy'.", vbExclamation, "Ano Inválido"
            Combo_ANO_MOD.text = ""
            Exit Sub
        End If
        
        ' Atribui o ano convertido à propriedade "Text" da TextBox
        Combo_ANO_MOD.text = CStr(anoConvertido)
    Else
        ' Ano inválido, exibe uma mensagem de erro e limpa o valor da TextBox
        MsgBox "Digite um ano válido no formato 'yy'.", vbExclamation, "Ano Inválido"
        Combo_ANO_MOD.text = ""
    End If
    
    Dim dhEmiText As String
    dhEmiText = xmlDoc.SelectSingleNode("//ns:dhEmi").text
    Dim dataEmissao As String
    dataEmissao = Mid(dhEmiText, 1, 10)

    ' Converter a data para o formato dd/mm/yyyy
    Dim dataFormatada As Date
    dataFormatada = CDate(dataEmissao)

    ' Inserir a data na TextBox chamada "Data" no formato dd/mm/yyyy
    Me.Data.value = Format(dataFormatada, "dd/mm/yyyy")
    
    'Fechar o objeto DOMDocument
    Set xmlDoc = Nothing
    
        LabelStatus = "Dados do XML foram importados com sucesso!"
End Sub

Private Sub txt_ValorFilial_AfterUpdate()
    
    txt_ValorFilial.text = Format(txt_ValorFilial.text, "Currency")
    txt_ValorFilial.ForeColor = &HFF
End Sub

Private Sub ComboBox_Operação_Enter()

    ComboBox_Operação.BorderColor = &H8000000D
    ComboBox_Operação.BackColor = &H80000002
    Label_Operação.ForeColor = &H80000012
    Frame_Dados_da_NFe.BorderColor = &H8000000D
End Sub

Private Sub ComboBox_Operação_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_Operação.BorderColor = &H8000000A
    ComboBox_Operação.BackColor = &H80000005
    Label_Operação.ForeColor = &H80000011
End Sub

Private Sub ComboBox_OutroEnvio_Enter()

    ComboBox_OutroEnvio.BorderColor = &H8000000D
    ComboBox_OutroEnvio.BackColor = &H80000002
    Label_OutroEnvio.ForeColor = &H80000012
    Frame_Dados_da_NFe.BorderColor = &H8000000D
End Sub

Private Sub ComboBox_OutroEnvio_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_OutroEnvio.BorderColor = &H8000000A
    ComboBox_OutroEnvio.BackColor = &H80000005
    Label_OutroEnvio.ForeColor = &H80000011
End Sub

Private Sub ComboBox_OutroEnvioFilial_Enter()

    ComboBox_OutroEnvioFilial.BorderColor = &H8000000D
    ComboBox_OutroEnvioFilial.BackColor = &H80000002
    Label_OutroEnvioFilial.ForeColor = &H80000012
    Frame_Dados_da_NFeFilial.BorderColor = &H8000000D
End Sub

Private Sub ComboBox_OutroEnvioFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_OutroEnvioFilial.BorderColor = &H8000000A
    ComboBox_OutroEnvioFilial.BackColor = &H80000005
    Label_OutroEnvioFilial.ForeColor = &H80000011
End Sub

Private Sub TextBoxChaveDeAcessoFilial_Enter()

    TextBoxChaveDeAcessoFilial.BorderColor = &H8000000D
    TextBoxChaveDeAcessoFilial.BackColor = &H80000002
    LabelChaveDeAcessoFilial.ForeColor = &H80000012
End Sub

Private Sub TextBoxChaveDeAcessoFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBoxChaveDeAcessoFilial.BorderColor = &H8000000A
    TextBoxChaveDeAcessoFilial.BackColor = &H80000005
    LabelChaveDeAcessoFilial.ForeColor = &H80000011
End Sub

Private Sub N_Nfe_Enter()

    N_Nfe.BorderColor = &H8000000D
    N_Nfe.BackColor = &H80000002
    Label_N_Nfe.ForeColor = &H80000012
    Frame_Dados_da_NFe.BorderColor = &H8000000D
End Sub

Private Sub N_Nfe_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
'Verifica se existe células com o mesmo valor
Dim vernumeronfe As Double
Dim planilha As Worksheet
Set planilha = ThisWorkbook.Worksheets("Entradas")
vernumeronfe = WorksheetFunction.CountIf(planilha.Range("A:A"), N_Nfe.text)

If N_Nfe.text = "" Then
        Exit Sub
End If

If vernumeronfe > 0 Then
        MsgBox "Numero da NFe já cadastrado!", vbCritical, "Numero da NFe"
        LabelStatus = "Numero da NFe já cadastrado!"
        N_Nfe.BorderColor = &HFF&
        N_Nfe.BackColor = &H8080FF
        Label_N_Nfe.ForeColor = &HC0&
        Exit Sub
End If
    
    N_Nfe.BorderColor = &H8000000A
    N_Nfe.BackColor = &H80000005
    Label_N_Nfe.ForeColor = &H80000011
End Sub

Private Sub UN_Enter()

    UN.BorderColor = &H8000000D
    UN.BackColor = &H80000002
    Label_UN.ForeColor = &H80000012
    Frame_Dados_da_NFe.BorderColor = &H8000000D
End Sub

Private Sub UN_Exit(ByVal cancel As MSForms.ReturnBoolean)

    UN.BorderColor = &H8000000A
    UN.BackColor = &H80000005
    Label_UN.ForeColor = &H80000011
End Sub

Private Sub Frame_Dados_da_NFe_Enter()

    Frame_Dados_da_NFe.ForeColor = &H8000000D
    Frame_Dados_da_NFe.BorderColor = &H8000000D
End Sub

Private Sub Frame_Dados_da_NFe_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_Operação.BorderColor = &H8000000A
    ComboBox_Operação.BackColor = &H80000005
    Label_Operação.ForeColor = &H80000011
    
    ComboBox_OutroEnvio.BorderColor = &H8000000A
    ComboBox_OutroEnvio.BackColor = &H80000005
    Label_OutroEnvio.ForeColor = &H80000011

    N_Nfe.BorderColor = &H8000000A
    N_Nfe.BackColor = &H80000005
    Label_N_Nfe.ForeColor = &H80000011
    
    UN.BorderColor = &H8000000A
    UN.BackColor = &H80000005
    Label_UN.ForeColor = &H80000011
    
    Frame_Dados_da_NFe.BorderColor = &H8000000A
    Frame_Dados_da_NFe.ForeColor = &H8000000A
End Sub

Private Sub PLACA_Enter()

    PLACA.BorderColor = &H8000000D
    PLACA.BackColor = &H80000002
    Label_PLACA.ForeColor = &H80000012
    Frame_Veículo.BorderColor = &H8000000D
End Sub

Private Sub PLACA_Exit(ByVal cancel As MSForms.ReturnBoolean)

    PLACA.BorderColor = &H8000000A
    PLACA.BackColor = &H80000005
    Label_PLACA.ForeColor = &H80000011
End Sub

Private Sub COD_Enter()

    COD.BorderColor = &H8000000D
    COD.BackColor = &H80000002
    Label_COD.ForeColor = &H80000012
    Frame_Veículo.BorderColor = &H8000000D
End Sub

Private Sub COD_Exit(ByVal cancel As MSForms.ReturnBoolean)

    COD.BorderColor = &H8000000A
    COD.BackColor = &H80000005
    Label_COD.ForeColor = &H80000011
End Sub

Private Sub Combo_ANO_MOD_Enter()

    Combo_ANO_MOD.BorderColor = &H8000000D
    Combo_ANO_MOD.BackColor = &H80000002
    Label_ANO_MOD.ForeColor = &H80000012
    Frame_Veículo.BorderColor = &H8000000D
End Sub

Private Sub Combo_ANO_MOD_Exit(ByVal cancel As MSForms.ReturnBoolean)

    Combo_ANO_MOD.BorderColor = &H8000000A
    Combo_ANO_MOD.BackColor = &H80000005
    Label_ANO_MOD.ForeColor = &H80000011
End Sub

Private Sub ComboBox_COR_Enter()

    ComboBox_COR.BorderColor = &H8000000D
    ComboBox_COR.BackColor = &H80000002
    Label_COR.ForeColor = &H80000012
    Frame_Veículo.BorderColor = &H8000000D
End Sub

Private Sub ComboBox_COR_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_COR.BorderColor = &H8000000A
    ComboBox_COR.BackColor = &H80000005
    Label_COR.ForeColor = &H80000011
End Sub

Private Sub ComboBox_VEÍCULO_Enter()

    ComboBox_VEÍCULO.BorderColor = &H8000000D
    ComboBox_VEÍCULO.BackColor = &H80000002
    Label_VEÍCULO.ForeColor = &H80000012
    Frame_Veículo.BorderColor = &H8000000D
End Sub

Private Sub ComboBox_VEÍCULO_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_VEÍCULO.BorderColor = &H8000000A
    ComboBox_VEÍCULO.BackColor = &H80000005
    Label_VEÍCULO.ForeColor = &H80000011
End Sub

Private Sub ListBoxLista_Change()

    Dim linha As Long
    Dim NumeroLinhaPlanilha As Long
    Dim TabelaOrigem As String
    Dim ws As Worksheet

    On Error GoTo TrataErro

    '==========================================================
    ' VALIDAÇÕES
    '==========================================================
    If Me.ListBoxLista.ListCount = 0 Then Exit Sub
    If Me.ListBoxLista.ListIndex < 0 Then Exit Sub

    linha = Me.ListBoxLista.ListIndex

    ' Coluna 0 = Nome da planilha de origem
    TabelaOrigem = Trim$(CStr(Me.ListBoxLista.List(linha, 0)))

    ' Coluna 12 = Número real da linha na planilha
    NumeroLinhaPlanilha = CLng(Val(Me.ListBoxLista.List(linha, 12)))

    If NumeroLinhaPlanilha < 2 Then Exit Sub

    '==========================================================
    ' IDENTIFICA A TABELA E PREENCHE OS CAMPOS
    '==========================================================
    Select Case UCase$(TabelaOrigem)

        '------------------------------------------------------
        ' ENTRADAS
        '------------------------------------------------------
        Case "ENTRADAS"

            ' Seleciona automaticamente a aba PageEntradas
            Me.MultiPage1.value = _
                Me.MultiPage1.Pages("PageEntradas").Index

            ' Preenche os campos da aba Entradas
            PreencherDadosEntradas _
                Me, _
                Me.ListBoxLista, _
                linha

            ' A planilha Entradas não possui chave na coluna L
            Me.TextBoxChaveDeAcessoFilial.value = vbNullString

        '------------------------------------------------------
        ' ENTRADAS FILIAL
        '------------------------------------------------------
        Case "ENTRADAS FILIAL"

            ' Seleciona automaticamente PageEntradasFilial
            Me.MultiPage1.value = _
                Me.MultiPage1.Pages("PageEntradasFilial").Index

            ' Preenche os campos da aba Entradas Filial
            PreencherDadosEntradasFilial _
                Me, _
                Me.ListBoxLista, _
                linha

            '==================================================
            ' CHAVE DE ACESSO DA NFe - COLUNA L
            '
            ' Índice 13 da ListBox corresponde à coluna L
            ' da planilha Entradas Filial
            '==================================================
            Me.TextBoxChaveDeAcessoFilial.value = _
                CStr(Me.ListBoxLista.List(linha, 13))

        '------------------------------------------------------
        ' ORIGEM NÃO IDENTIFICADA
        '------------------------------------------------------
        Case Else

            MsgBox _
                "A tabela de origem não foi identificada." & _
                vbCrLf & vbCrLf & _
                "Origem encontrada: " & TabelaOrigem, _
                vbExclamation, _
                "Origem do registro"

            Exit Sub

    End Select

    '==========================================================
    ' SELECIONAR A CÉLULA DA COLUNA A NA LINHA DO REGISTRO
    '==========================================================
    Set ws = ThisWorkbook.Worksheets(TabelaOrigem)

    ' Ativa a planilha correspondente
    ws.Activate

    ' Seleciona somente a célula da coluna A
    ' na linha original do registro pesquisado
    Application.GoTo _
        Reference:=ws.Cells(NumeroLinhaPlanilha, 1), _
        Scroll:=True

    Exit Sub

TrataErro:

    MsgBox _
        "Erro ao carregar os dados selecionados." & _
        vbCrLf & vbCrLf & _
        "Tabela: " & TabelaOrigem & vbCrLf & _
        "Erro " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "Pesquisa"

End Sub

Private Sub ListBoxLista_Enter()

    ListBoxLista.BorderColor = &H8000000D
End Sub

Private Sub ListBoxLista_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ListBoxLista.BorderColor = &H8000000A
End Sub

Private Sub ListBoxLista2_Change()

    Dim linha As Long
    Dim NumeroLinhaPlanilha As Long
    Dim TabelaOrigem As String
    Dim ws As Worksheet

    On Error GoTo TrataErro

    '==========================================================
    ' VALIDAÇÕES
    '==========================================================
    If Me.ListBoxLista2.ListCount = 0 Then Exit Sub
    If Me.ListBoxLista2.ListIndex < 0 Then Exit Sub

    linha = Me.ListBoxLista2.ListIndex

    ' Coluna 0 = Nome da planilha de origem
    TabelaOrigem = Trim$(CStr(Me.ListBoxLista2.List(linha, 0)))

    ' Coluna 12 = Número real da linha na planilha
    NumeroLinhaPlanilha = CLng(Val(Me.ListBoxLista2.List(linha, 12)))

    If NumeroLinhaPlanilha < 2 Then Exit Sub

    '==========================================================
    ' IDENTIFICA A TABELA E PREENCHE OS CAMPOS
    '==========================================================
    Select Case UCase$(TabelaOrigem)

        '------------------------------------------------------
        ' ENTRADAS
        '------------------------------------------------------
        Case "ENTRADAS"

            ' Seleciona automaticamente PageEntradas
            Me.MultiPage1.value = _
                Me.MultiPage1.Pages("PageEntradas").Index

            ' Preenche os campos da PageEntradas
            PreencherDadosEntradas _
                Me, _
                Me.ListBoxLista2, _
                linha

            ' A planilha Entradas não possui coluna L
            Me.TextBoxChaveDeAcessoFilial.value = vbNullString

        '------------------------------------------------------
        ' ENTRADAS FILIAL
        '------------------------------------------------------
        Case "ENTRADAS FILIAL"

            ' Seleciona automaticamente PageEntradasFilial
            Me.MultiPage1.value = _
                Me.MultiPage1.Pages("PageEntradasFilial").Index

            ' Preenche os campos da PageEntradasFilial
            PreencherDadosEntradasFilial _
                Me, _
                Me.ListBoxLista2, _
                linha

            '==================================================
            ' CHAVE DE ACESSO DA NFe - COLUNA L
            '
            ' Índice 13 da ListBox corresponde à coluna L
            ' da planilha Entradas Filial
            '==================================================
            Me.TextBoxChaveDeAcessoFilial.value = _
                CStr(Me.ListBoxLista2.List(linha, 13))

        '------------------------------------------------------
        ' ORIGEM NÃO IDENTIFICADA
        '------------------------------------------------------
        Case Else

            MsgBox _
                "A tabela de origem não foi identificada." & _
                vbCrLf & vbCrLf & _
                "Origem encontrada: " & TabelaOrigem, _
                vbExclamation, _
                "Origem do registro"

            Exit Sub

    End Select

    '==========================================================
    ' SELECIONAR A CÉLULA DA COLUNA A NA LINHA DO REGISTRO
    '==========================================================
    Set ws = ThisWorkbook.Worksheets(TabelaOrigem)

    ' Ativa a planilha correspondente
    ws.Activate

    ' Seleciona somente a célula da coluna A
    ' na linha original do registro pesquisado
    Application.GoTo _
        Reference:=ws.Cells(NumeroLinhaPlanilha, 1), _
        Scroll:=True

    Exit Sub

TrataErro:

    MsgBox _
        "Erro ao carregar os dados selecionados." & _
        vbCrLf & vbCrLf & _
        "Tabela: " & TabelaOrigem & vbCrLf & _
        "Erro " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "Pesquisa"

End Sub

Private Sub Frame_Veículo_Enter()

    Frame_Veículo.ForeColor = &H8000000D
    Frame_Veículo.BorderColor = &H8000000D
End Sub

Private Sub Frame_Veículo_Exit(ByVal cancel As MSForms.ReturnBoolean)

    COD.BorderColor = &H8000000A
    COD.BackColor = &H80000005
    Label_COD.ForeColor = &H80000011
        
    ComboBox_VEÍCULO.BorderColor = &H8000000A
    ComboBox_VEÍCULO.BackColor = &H80000005
    Label_VEÍCULO.ForeColor = &H80000011
    
    PLACA.BorderColor = &H8000000A
    PLACA.BackColor = &H80000005
    Label_PLACA.ForeColor = &H80000011
    
    ComboBox_COR.BorderColor = &H8000000A
    ComboBox_COR.BackColor = &H80000005
    Label_COR.ForeColor = &H80000011
        
    Combo_ANO_MOD.BorderColor = &H8000000A
    Combo_ANO_MOD.BackColor = &H80000005
    Label_ANO_MOD.ForeColor = &H80000011
    
    Frame_Veículo.BorderColor = &H8000000A
    Frame_Veículo.ForeColor = &H8000000A
End Sub

Private Sub Data_Enter()

    Data.BorderColor = &H8000000D
    Data.BackColor = &H80000002
    Label_Data.ForeColor = &H80000012
    
    Frame_Data.ForeColor = &H8000000D
    Frame_Data.BorderColor = &H8000000D
End Sub

Private Sub Data_Exit(ByVal cancel As MSForms.ReturnBoolean)

    Data.BorderColor = &H8000000A
    Data.BackColor = &H80000005
    Label_Data.ForeColor = &H80000011
End Sub

Private Sub Frame_Data_Enter()

    Data.BorderColor = &H8000000D
    Data.BackColor = &H80000002
    Label_Data.ForeColor = &H80000012
    
    Frame_Data.ForeColor = &H8000000D
    Frame_Data.BorderColor = &H8000000D
End Sub

Private Sub Frame_Data_Exit(ByVal cancel As MSForms.ReturnBoolean)

    Data.BorderColor = &H8000000A
    Data.BackColor = &H80000005
    Label_Data.ForeColor = &H80000011
    
    Frame_Data.ForeColor = &H8000000A
    Frame_Data.BorderColor = &H8000000A
End Sub

Private Sub txt_Valor_Enter()

    txt_Valor.BorderColor = &H8000000D
    txt_Valor.BackColor = &H80000002
    Label_Valores.ForeColor = &H80000012
    Frame_Valores.BorderColor = &H8000000D
End Sub

Private Sub txt_Valor_Exit(ByVal cancel As MSForms.ReturnBoolean)

    txt_Valor.BorderColor = &H8000000A
    txt_Valor.BackColor = &H80000005
    Label_Valores.ForeColor = &H80000011
End Sub

Private Sub Frame_Valores_Enter()

    txt_Valor.BorderColor = &H8000000D
    txt_Valor.BackColor = &H80000002
    Label_Valores.ForeColor = &H80000012
    
    Frame_Valores.ForeColor = &H8000000D
    Frame_Valores.BorderColor = &H8000000D
End Sub

Private Sub Frame_Valores_Exit(ByVal cancel As MSForms.ReturnBoolean)

    txt_Valor.BorderColor = &H8000000A
    txt_Valor.BackColor = &H80000005
    Label_Valores.ForeColor = &H80000011
    
    Frame_Valores.ForeColor = &H8000000A
    Frame_Valores.BorderColor = &H8000000A
End Sub

Private Sub TextBoxPesquisa_Enter()

    TextBoxPesquisa.BorderColor = &H8000000D
    TextBoxPesquisa.BackColor = &H80000002
    Frame_Pesquisa.BorderColor = &H8000000D
End Sub

Private Sub TextBoxPesquisa_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBoxPesquisa.BorderColor = &H8000000A
    TextBoxPesquisa.BackColor = &H80000005
End Sub

Private Sub TextBoxPesquisaFilial_Enter()

    TextBoxPesquisaFilial.BorderColor = &H8000000D
    TextBoxPesquisaFilial.BackColor = &H80000002
    Fame_PesquisaFilial.BorderColor = &H8000000D
End Sub

Private Sub TextBoxPesquisaFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBoxPesquisaFilial.BorderColor = &H8000000A
    TextBoxPesquisaFilial.BackColor = &H80000005
End Sub

Private Sub ListBoxLista2_Enter()

    ListBoxLista2.BorderColor = &H8000000D
End Sub

Private Sub ListBoxLista2_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ListBoxLista2.BorderColor = &H8000000A
End Sub

Private Sub Fame_PesquisaFilial_Enter()

    Fame_PesquisaFilial.ForeColor = &H8000000D
    Fame_PesquisaFilial.BorderColor = &H8000000D
End Sub

Private Sub Fame_PesquisaFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)
  
    Fame_PesquisaFilial.ForeColor = &H80000011
    Fame_PesquisaFilial.BorderColor = &H8000000A
End Sub

Private Sub Frame_Pesquisa_Enter()

    Frame_Pesquisa.ForeColor = &H8000000D
    Frame_Pesquisa.BorderColor = &H8000000D
End Sub

Private Sub Frame_Pesquisa_Exit(ByVal cancel As MSForms.ReturnBoolean)
  
    Frame_Pesquisa.ForeColor = &H80000011
    Frame_Pesquisa.BorderColor = &H8000000A
End Sub

Private Sub Frame_Ações_Enter()

    Frame_Ações.ForeColor = &H8000000D
    Frame_Ações.BorderColor = &H8000000D
End Sub

Private Sub Frame_Ações_Exit(ByVal cancel As MSForms.ReturnBoolean)

    Frame_Ações.ForeColor = &H8000000A
    Frame_Ações.BorderColor = &H8000000A
End Sub

Private Sub Frame_AçõesFilial_Enter()

    Frame_AçõesFilial.ForeColor = &H8000000D
    Frame_AçõesFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_AçõesFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    Frame_AçõesFilial.ForeColor = &H8000000A
    Frame_AçõesFilial.BorderColor = &H8000000A
End Sub

Private Sub FrameMenu_Enter()

    FrameMenu.ForeColor = &H8000000D
    FrameMenu.BorderColor = &H8000000D
        
    ListBoxLista.BorderColor = &H8000000A
    ListBoxLista.BackColor = &H80000005
      
    ListBoxLista2.BorderColor = &H8000000A
    ListBoxLista2.BackColor = &H80000005
          
    Frame_Pesquisa.ForeColor = &H80000011
    Frame_Pesquisa.BorderColor = &H8000000A
      
    Fame_PesquisaFilial.ForeColor = &H80000011
    Fame_PesquisaFilial.BorderColor = &H8000000A
End Sub

Private Sub FrameMenu_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    FrameMenu.ForeColor = &H8000000A
    FrameMenu.BorderColor = &H8000000A
        
    ListBoxLista2.BorderColor = &H8000000A
    ListBoxLista2.BackColor = &H80000005
End Sub

Private Sub ComboBox_OperaçãoFilial_Enter()

    ComboBox_OperaçãoFilial.BorderColor = &H8000000D
    ComboBox_OperaçãoFilial.BackColor = &H80000002
    Label_OperaçãoFilial.ForeColor = &H80000012
    Frame_Dados_da_NFeFilial.BorderColor = &H8000000D
End Sub

Private Sub ComboBox_OperaçãoFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_OperaçãoFilial.BorderColor = &H8000000A
    ComboBox_OperaçãoFilial.BackColor = &H80000005
    Label_OperaçãoFilial.ForeColor = &H80000011
End Sub

Private Sub N_NfeFilial_Enter()

    N_NfeFilial.BorderColor = &H8000000D
    N_NfeFilial.BackColor = &H80000002
    Label_N_NfeFilial.ForeColor = &H80000012
    Frame_Dados_da_NFeFilial.BorderColor = &H8000000D
End Sub

Private Sub N_NfeFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

'Verifica se existe células com o mesmo valor
Dim vernumeronfe As Double
Dim planilha As Worksheet
Set planilha = ThisWorkbook.Worksheets("Entradas Filial")
vernumeronfe = WorksheetFunction.CountIf(planilha.Range("A:A"), N_NfeFilial.text)

If N_NfeFilial.text = "" Then
        Exit Sub
End If

If vernumeronfe > 0 Then
        MsgBox "Numero da NFe já cadastrado!", vbCritical, "Numero da NFe"
        LabelStatus = "Numero da NFe já cadastrado!"
        N_NfeFilial.BorderColor = &HFF&
        N_NfeFilial.BackColor = &H8080FF
        N_NfeFilial.ForeColor = &HC0&
        Exit Sub
End If

    N_NfeFilial.BorderColor = &H8000000A
    N_NfeFilial.BackColor = &H80000005
    Label_N_NfeFilial.ForeColor = &H80000011
End Sub

Private Sub UNFilial_Enter()

    UNFilial.BorderColor = &H8000000D
    UNFilial.BackColor = &H80000002
    Label_UNFilial.ForeColor = &H80000012
    Frame_Dados_da_NFeFilial.BorderColor = &H8000000D
End Sub

Private Sub UNFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    UNFilial.BorderColor = &H8000000A
    UNFilial.BackColor = &H80000005
    Label_UNFilial.ForeColor = &H80000011
End Sub

Private Sub Frame_Dados_da_NFeFilial_Enter()

    Frame_Dados_da_NFeFilial.ForeColor = &H8000000D
    Frame_Dados_da_NFeFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_Dados_da_NFeFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_OperaçãoFilial.BorderColor = &H8000000A
    ComboBox_OperaçãoFilial.BackColor = &H80000005
    Label_OperaçãoFilial.ForeColor = &H80000011

    N_NfeFilial.BorderColor = &H8000000A
    N_NfeFilial.BackColor = &H80000005
    Label_N_NfeFilial.ForeColor = &H80000011
    
    UNFilial.BorderColor = &H8000000A
    UNFilial.BackColor = &H80000005
    Label_UNFilial.ForeColor = &H80000011
    
    Frame_Dados_da_NFeFilial.BorderColor = &H8000000A
    Frame_Dados_da_NFeFilial.ForeColor = &H8000000A
End Sub

Private Sub PLACAFilial_Enter()

    PLACAFilial.BorderColor = &H8000000D
    PLACAFilial.BackColor = &H80000002
    Label_PLACAFilial.ForeColor = &H80000012
    Frame_VeículoFilial.BorderColor = &H8000000D
End Sub

Private Sub PLACAFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    PLACAFilial.BorderColor = &H8000000A
    PLACAFilial.BackColor = &H80000005
    Label_PLACAFilial.ForeColor = &H80000011
End Sub

Private Sub CODFilial_Enter()

    CODFilial.BorderColor = &H8000000D
    CODFilial.BackColor = &H80000002
    Label_CODFilial.ForeColor = &H80000012
    Frame_VeículoFilial.BorderColor = &H8000000D
End Sub

Private Sub CODFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    CODFilial.BorderColor = &H8000000A
    CODFilial.BackColor = &H80000005
    Label_CODFilial.ForeColor = &H80000011
End Sub

Private Sub Combo_ANO_MODFilial_Enter()

    Combo_ANO_MODFilial.BorderColor = &H8000000D
    Combo_ANO_MODFilial.BackColor = &H80000002
    Label_ANO_MODFilial.ForeColor = &H80000012
    Frame_VeículoFilial.BorderColor = &H8000000D
End Sub

Private Sub Combo_ANO_MODFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    Combo_ANO_MODFilial.BorderColor = &H8000000A
    Combo_ANO_MODFilial.BackColor = &H80000005
    Label_ANO_MODFilial.ForeColor = &H80000011
End Sub

Private Sub ComboBox_CORFilial_Enter()

    ComboBox_CORFilial.BorderColor = &H8000000D
    ComboBox_CORFilial.BackColor = &H80000002
    Label_CORFilial.ForeColor = &H80000012
    Frame_VeículoFilial.BorderColor = &H8000000D
End Sub

Private Sub ComboBox_CORFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_CORFilial.BorderColor = &H8000000A
    ComboBox_CORFilial.BackColor = &H80000005
    Label_CORFilial.ForeColor = &H80000011
End Sub

Private Sub ComboBox_VEÍCULOFilial_Enter()

    ComboBox_VEÍCULOFilial.BorderColor = &H8000000D
    ComboBox_VEÍCULOFilial.BackColor = &H80000002
    Label_VEÍCULOFilial.ForeColor = &H80000012
    Frame_VeículoFilial.BorderColor = &H8000000D
End Sub

Private Sub ComboBox_VEÍCULOFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_VEÍCULOFilial.BorderColor = &H8000000A
    ComboBox_VEÍCULOFilial.BackColor = &H80000005
    Label_VEÍCULOFilial.ForeColor = &H80000011
End Sub

Private Sub Frame_VeículoFilial_Enter()

    Frame_VeículoFilial.ForeColor = &H8000000D
    Frame_VeículoFilial.BorderColor = &H8000000D
    
    TextBoxChaveDeAcessoFilial.BorderColor = &H8000000A
    TextBoxChaveDeAcessoFilial.BackColor = &H80000005
    LabelChaveDeAcessoFilial.ForeColor = &H80000011
End Sub

Private Sub Frame_VeículoFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    CODFilial.BorderColor = &H8000000A
    CODFilial.BackColor = &H80000005
    Label_CODFilial.ForeColor = &H80000011
        
    ComboBox_VEÍCULOFilial.BorderColor = &H8000000A
    ComboBox_VEÍCULOFilial.BackColor = &H80000005
    Label_VEÍCULOFilial.ForeColor = &H80000011
    
    PLACAFilial.BorderColor = &H8000000A
    PLACAFilial.BackColor = &H80000005
    Label_PLACAFilial.ForeColor = &H80000011
    
    ComboBox_CORFilial.BorderColor = &H8000000A
    ComboBox_CORFilial.BackColor = &H80000005
    Label_CORFilial.ForeColor = &H80000011
        
    Combo_ANO_MODFilial.BorderColor = &H8000000A
    Combo_ANO_MODFilial.BackColor = &H80000005
    Label_ANO_MODFilial.ForeColor = &H80000011
    
    Frame_VeículoFilial.BorderColor = &H8000000A
    Frame_VeículoFilial.ForeColor = &H8000000A
End Sub

Private Sub DataFilial_Enter()

    DataFilial.BorderColor = &H8000000D
    DataFilial.BackColor = &H80000002
    Label_DataFilial.ForeColor = &H80000012
    
    Frame_DataFilial.ForeColor = &H8000000D
    Frame_DataFilial.BorderColor = &H8000000D
End Sub

Private Sub DataFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    DataFilial.BorderColor = &H8000000A
    DataFilial.BackColor = &H80000005
    Label_DataFilial.ForeColor = &H80000011
End Sub

Private Sub Frame_DataFilial_Enter()

    DataFilial.BorderColor = &H8000000D
    DataFilial.BackColor = &H80000002
    Label_DataFilial.ForeColor = &H80000012
    
    Frame_DataFilial.ForeColor = &H8000000D
    Frame_DataFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_DataFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    DataFilial.BorderColor = &H8000000A
    DataFilial.BackColor = &H80000005
    Label_DataFilial.ForeColor = &H80000011
    
    Frame_DataFilial.ForeColor = &H8000000A
    Frame_DataFilial.BorderColor = &H8000000A
End Sub

Private Sub txt_ValorFilial_Enter()

    txt_ValorFilial.BorderColor = &H8000000D
    txt_ValorFilial.BackColor = &H80000002
    Label_ValoresFilial.ForeColor = &H80000012
    Frame_ValoresFilial.BorderColor = &H8000000D
End Sub

Private Sub txt_ValorFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    txt_ValorFilial.BorderColor = &H8000000A
    txt_ValorFilial.BackColor = &H80000005
    Label_ValoresFilial.ForeColor = &H80000011
End Sub

Private Sub Frame_ValoresFilial_Enter()

    txt_ValorFilial.BorderColor = &H8000000D
    txt_ValorFilial.BackColor = &H80000002
    Label_ValoresFilial.ForeColor = &H80000012
    
    Frame_ValoresFilial.ForeColor = &H8000000D
    Frame_ValoresFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_ValoresFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    txt_ValorFilial.BorderColor = &H8000000A
    txt_ValorFilial.BackColor = &H80000005
    Label_ValoresFilial.ForeColor = &H80000011
    
    Frame_ValoresFilial.ForeColor = &H8000000A
    Frame_ValoresFilial.BorderColor = &H8000000A
End Sub

Private Sub Frame_PesquisaFilial_Enter()

    Fame_PesquisaFilial.ForeColor = &H8000000D
    Fame_PesquisaFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_PesquisaFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBoxPesquisaFilial.BorderColor = &H8000000A
    TextBoxPesquisaFilial.BackColor = &H80000005
  
    Fame_PesquisaFilial.ForeColor = &H80000011
    Fame_PesquisaFilial.BorderColor = &H8000000A
End Sub

Private Sub PLACA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub PLACA_AfterUpdate()
    
    Me.PLACA = UCase(Me.PLACA)
End Sub

Private Function SóNúmeros(l As IReturnInteger)
    Select Case l
        Case Asc("0") To Asc("9")
            SóNúmeros = l
        Case Else
            SóNúmeros = 0
            MsgBox "Favor inserir apenas números!", vbExclamation, "CAMPO TIPO NÚMERO"
    End Select
End Function

Private Sub N_Nfe_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub UN_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub COD_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub txt_Valor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub N_NfeFilial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub UNFilial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub CODFilial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub txt_ValorFilial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub CommandButtonCelEsquerda_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonCelEsquerda_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    CommandButtonCelEsquerda.PicturePosition = fmPicturePositionRightTop
End Sub

Private Sub CommandButtonCelEsquerda_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
        
    CommandButtonCelEsquerda.PicturePosition = fmPicturePositionCenter
End Sub

Private Sub CommandButtonCelDireita_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonCelDireita_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    CommandButtonCelDireita.PicturePosition = fmPicturePositionRightTop
End Sub

Private Sub CommandButtonCelDireita_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
        
    CommandButtonCelDireita.PicturePosition = fmPicturePositionCenter
End Sub

Private Sub CommandButtonCelAcima_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonCelAcima_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    CommandButtonCelAcima.PicturePosition = fmPicturePositionRightTop
End Sub

Private Sub CommandButtonCelAcima_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
        
    CommandButtonCelAcima.PicturePosition = fmPicturePositionCenter
End Sub

Private Sub CommandButtonCelAbaixo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonCelAbaixo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    CommandButtonCelAbaixo.PicturePosition = fmPicturePositionRightTop
End Sub

Private Sub CommandButtonCelAbaixo_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
        
    CommandButtonCelAbaixo.PicturePosition = fmPicturePositionCenter
End Sub

Private Sub CommandButtonCelCopiar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonCelCopiar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    CommandButtonCelCopiar.PicturePosition = fmPicturePositionRightTop
End Sub

Private Sub CommandButtonCelCopiar_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
        
    CommandButtonCelCopiar.PicturePosition = fmPicturePositionCenter
End Sub

Private Sub CommandButtonCelEsquerda_Click()
    UF_NavegarCelula Me, 0, -1
    Me.LabelConteudoCélula.BorderStyle = fmBorderStyleSingle
    Me.LabelConteudoCélula.BorderColor = &HFF&
End Sub

Private Sub CommandButtonCelDireita_Click()
    UF_NavegarCelula Me, 0, 1
    Me.LabelConteudoCélula.BorderStyle = fmBorderStyleSingle
    Me.LabelConteudoCélula.BorderColor = &HFF&
End Sub

Private Sub CommandButtonCelAcima_Click()
    UF_NavegarCelula Me, -1, 0
    Me.LabelConteudoCélula.BorderStyle = fmBorderStyleSingle
    Me.LabelConteudoCélula.BorderColor = &HFF&
End Sub

Private Sub CommandButtonCelAbaixo_Click()
    UF_NavegarCelula Me, 1, 0
    Me.LabelConteudoCélula.BorderStyle = fmBorderStyleSingle
    Me.LabelConteudoCélula.BorderColor = &HFF&
End Sub

Private Sub CommandButtonCelCopiar_Click()
    UF_CopiarCelula Me, True 'True = copia como aparece (ActiveCell.Text)
End Sub

Private Sub AtualizaPainelCelula()
    On Error Resume Next
    If Not ActiveCell Is Nothing Then
        If ControleExiste(Me, "LabelCelAtual") Then Me.Controls("LabelCelAtual").CAPTION = ActiveCell.Address(False, False)
        If ControleExiste(Me, "LabelCelValor") Then Me.Controls("LabelCelValor").CAPTION = CStr(ActiveCell.text)
    End If
End Sub

Private Function ControleExiste(Frm As Object, ByVal nome As String) As Boolean
    On Error GoTo FIM
    Dim tmp As Object
    Set tmp = Frm.Controls(nome)
    ControleExiste = True
    Exit Function
FIM:
    ControleExiste = False
End Function



