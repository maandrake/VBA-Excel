VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cad_NFE_Saida 
   Caption         =   "Cadastro de NFe Saída"
   ClientHeight    =   11880
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   19380
   OleObjectBlob   =   "Cad_NFE_Saida.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cad_NFE_Saida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const NomePlanilha As String = "Saidas"
Private Const NomePlanilha2 As String = "Saidas Filial"
Private Const LinhaCabecalho As Integer = 1

Private Sub UserForm_Initialize()

    LabelStatus = "Bem vindo..."
    
    Call PreencheCampos
    Call PreencheCampos2
    
    ' Inicializa a LabelData e  LabelHora com a data e hora atual
    LabelData.CAPTION = Format(Now, "dd/mm/yyyy")
    LabelHora.CAPTION = Format(Now, "hh:mm")
        
    TextBox_Valor_Entrada.text = Format(TextBox_Valor_Entrada.text, "Currency")
    TextBox_Data.value = Format(Now, "dd/mm/yyyy")
    TextBox_DataFilial.value = Format(Now, "dd/mm/yyyy")
    Sheets("Saidas").Select
    
End Sub

Private Sub Command_EntradaFilial_Click()

    LabelStatus = "Planilha Entradas filial..."
    Sheets("Entradas Filial").Select
End Sub

Private Sub Command_Saida_Click()

    Sheets("Saidas Filial").Select
End Sub

Private Sub CommandButton_ExcluirFilial_Click()

'Declarar a variável Resp para receber uma resposta
Dim resp As Integer
Dim c As Range
Dim msg, Style, Title, Help, Ctxt, Response, MyString

'Fazer a busca do registro digitado pelo usuário
LabelStatus = "Confirma a exclusão da NFe..."
With Worksheets("Saidas Filial").Range("A:A")
Set c = .Find(TextBox_N_SaidaFilial.value, LookIn:=xlValues, LookAt:=xlWhole)

If Not c Is Nothing Then
    resp = MsgBox("Tem certeza que deseja excluir Esta NFe?", vbYesNo, "Confirmação")
    If resp = vbYes Then
         c.Select
         Selection.EntireRow.Delete
         'Limpar as caixas de texto
            ComboBox_OperaçãoFilial.value = Empty
            TextBox_N_SaidaFilial.value = Empty
            TextBox_N_EntradaFilial.value = Empty
            TextBox_ESTOQUEFilial.value = Empty
            ComboBox_VEÍCULOFilial.value = Empty
            TextBox_PLACAFilial.value = Empty
            ComboBox_AnoFilial.value = Empty
            ComboBox_CORFilial.value = Empty
            TextBox_Valor_EntradaFilial.value = Empty
            TextBox_Valor_SaidaFilial.value = Empty
            TextBox_DataFilial.value = Empty
            ComboBox_PorFilial.value = Empty
            TextBoxBaseICMSFilial.value = Empty
            TextBox_ICMSFilial.value = Empty
            TextBox_PISFilial.value = Empty
            TextBox_COFINSFilial.value = Empty
            ComboBoxCamposFilial.value = "Nº Nfe Saida"
            TextBoxFiltroFilial.value = Empty
            ListBoxListaFilial.Clear
         'Colocar o foco na primeira caixa de texto
         ComboBox_OperaçãoFilial.SetFocus
    Else
         MsgBox "Esta NFe não será excluído!"
    End If
Else
     MsgBox "NFe não encontrada!"
End If
End With
Exit Sub
End Sub

Private Sub CommandButton_GravarFilial_Click()

'Verifica se existe células com o mesmo valor
Dim vernumeronfe As Double
Dim planilha As Worksheet
Set planilha = ThisWorkbook.Worksheets("Saidas Filial")
vernumeronfe = WorksheetFunction.CountIf(planilha.Range("A:A"), TextBox_N_SaidaFilial.text)

If TextBox_N_SaidaFilial.text = "" Then
        Exit Sub
End If

If vernumeronfe > 0 Then
        MsgBox "Numero da NFe já cadastrado!", vbCritical, "Numero da NFe"
        LabelStatus = "Numero da NFe já cadastrado!"
        TextBox_N_SaidaFilial.BorderColor = &HFF&
        TextBox_N_SaidaFilial.BackColor = &H8080FF
        TextBox_N_SaidaFilial.ForeColor = &HC0&
        Exit Sub
End If

    'Ativar a primeira planilha
    ThisWorkbook.Worksheets("Saidas Filial").Activate
    'Selecionar a célula
    Range("A2").Select

    'Procurar a primeira célula vazia
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
    End If
Loop Until IsEmpty(ActiveCell) = True

    'Carregar os dados digitados nas caixas de texto para a planilha
    ActiveCell.value = TextBox_N_SaidaFilial.value
    ActiveCell.Offset(0, 1).value = TextBox_N_EntradaFilial.value
    ActiveCell.Offset(0, 2).value = ComboBox_OperaçãoFilial.value
    ActiveCell.Offset(0, 3).value = TextBox_ESTOQUEFilial.value
    ActiveCell.Offset(0, 4).value = ComboBox_VEÍCULOFilial.value
    ActiveCell.Offset(0, 5).value = TextBox_PLACAFilial.value
    ActiveCell.Offset(0, 6).value = ComboBox_AnoFilial.value
    ActiveCell.Offset(0, 7).value = ComboBox_CORFilial.value
    ActiveCell.Offset(0, 8).value = CDbl(TextBox_Valor_EntradaFilial.value)
    ActiveCell.Offset(0, 9).value = CDbl(TextBox_Valor_SaidaFilial.value)
    ActiveCell.Offset(0, 11).value = CDate(TextBox_DataFilial.value)
    ActiveCell.Offset(0, 12).value = ComboBox_PorFilial.value
    ActiveCell.Offset(0, 13).value = CDbl(TextBox_ICMSFilial.value)
    ActiveCell.Offset(0, 14).value = CDbl(TextBox_PISFilial.value)
    ActiveCell.Offset(0, 15).value = CDbl(TextBox_COFINSFilial.value)
    
    'Ativar a primeira planilha
    ThisWorkbook.Worksheets("Todas Saidas Filial").Activate
    'Selecionar a célula
    Range("A2").Select

    'Procurar a primeira célula vazia
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
    End If
Loop Until IsEmpty(ActiveCell) = True

    'Carregar os dados digitados nas caixas de texto para a planilha
    ActiveCell.value = TextBox_N_SaidaFilial.value
    ActiveCell.Offset(0, 1).value = TextBox_N_EntradaFilial.value
    ActiveCell.Offset(0, 2).value = ComboBox_OperaçãoFilial.value
    ActiveCell.Offset(0, 3).value = TextBox_ESTOQUEFilial.value
    ActiveCell.Offset(0, 4).value = ComboBox_VEÍCULOFilial.value
    ActiveCell.Offset(0, 5).value = TextBox_PLACAFilial.value
    ActiveCell.Offset(0, 6).value = ComboBox_AnoFilial.value
    ActiveCell.Offset(0, 7).value = ComboBox_CORFilial.value
    ActiveCell.Offset(0, 8).value = CDbl(TextBox_Valor_EntradaFilial.value)
    ActiveCell.Offset(0, 9).value = CDbl(TextBox_Valor_SaidaFilial.value)
    ActiveCell.Offset(0, 11).value = CDate(TextBox_DataFilial.value)
    ActiveCell.Offset(0, 12).value = ComboBox_PorFilial.value
    ActiveCell.Offset(0, 13).value = CDbl(TextBox_ICMSFilial.value)
    ActiveCell.Offset(0, 14).value = CDbl(TextBox_PISFilial.value)
    ActiveCell.Offset(0, 15).value = CDbl(TextBox_COFINSFilial.value)
    
ThisWorkbook.Worksheets("Entradas Filial").Activate

Dim wsEntradas As Worksheet
Dim proxLinha As Long

Set wsEntradas = ThisWorkbook.Worksheets("Entradas")

With wsEntradas
    'próxima linha vazia pela coluna A (considerando cabeçalho na linha 1)
    proxLinha = .Cells(.rows.count, "A").End(xlUp).Row
    If proxLinha < 2 Then
        proxLinha = 2
    Else
        proxLinha = proxLinha + 1
    End If

    .Cells(proxLinha, "A").value = TextBox_N_SaidaFilial.value          'Col A
    .Cells(proxLinha, "B").value = "Compra"                            'Col B (fixo)
    .Cells(proxLinha, "C").value = TextBox_ESTOQUEFilial.value          'Col C
    .Cells(proxLinha, "D").value = ComboBox_VEÍCULOFilial.value         'Col D
    .Cells(proxLinha, "E").value = CODFilial.value                      'Col E
    .Cells(proxLinha, "F").value = TextBox_PLACAFilial.value            'Col F
    .Cells(proxLinha, "G").value = ComboBox_AnoFilial.value             'Col G
    .Cells(proxLinha, "H").value = ComboBox_CORFilial.value             'Col H
    .Cells(proxLinha, "I").value = CDbl(TextBox_Valor_EntradaFilial.value) 'Col I
    .Cells(proxLinha, "J").value = CDate(TextBox_DataFilial.value)      'Col J
    .Cells(proxLinha, "K").value = ComboBox_OutroEnvioFilial.value      'Col K (fixo)
End With

Set wsEntradas = ThisWorkbook.Worksheets("Entradas Anual")

With wsEntradas
    'próxima linha vazia pela coluna A (considerando cabeçalho na linha 1)
    proxLinha = .Cells(.rows.count, "A").End(xlUp).Row
    If proxLinha < 2 Then
        proxLinha = 2
    Else
        proxLinha = proxLinha + 1
    End If

    .Cells(proxLinha, "A").value = TextBox_N_SaidaFilial.value          'Col A
    .Cells(proxLinha, "B").value = "Compra"                            'Col B (fixo)
    .Cells(proxLinha, "C").value = TextBox_ESTOQUEFilial.value          'Col C
    .Cells(proxLinha, "D").value = ComboBox_VEÍCULOFilial.value         'Col D
    .Cells(proxLinha, "E").value = CODFilial.value                      'Col E
    .Cells(proxLinha, "F").value = TextBox_PLACAFilial.value            'Col F
    .Cells(proxLinha, "G").value = ComboBox_AnoFilial.value             'Col G
    .Cells(proxLinha, "H").value = ComboBox_CORFilial.value             'Col H
    .Cells(proxLinha, "I").value = CDbl(TextBox_Valor_EntradaFilial.value) 'Col I
    .Cells(proxLinha, "J").value = CDate(TextBox_DataFilial.value)      'Col J
End With


Dim c As Range

LabelStatus = "Excluindo a NFe de Entrada..."

With Worksheets("Entradas Filial").Range("A:A")
    Set c = .Find(TextBox_N_EntradaFilial.value, LookIn:=xlValues, LookAt:=xlWhole)

    If Not c Is Nothing Then
        'Exclui direto (sem confirmação) — equivalente ao seu Resize(1,13)
        c.Resize(1, 13).Delete Shift:=xlUp

        Application.CalculateFull
        ThisWorkbook.Save

        'Limpar as caixas de texto
        ComboBox_OperaçãoFilial.value = "Devolução"
        TextBox_N_SaidaFilial.value = Empty
        TextBox_N_EntradaFilial.value = Empty
        ComboBox_OutroEnvioFilial.value = ""
        TextBoxChaveDeAcessoFilial.value = ""
        TextBox_ESTOQUEFilial.value = "1"
        CODFilial.value = Empty
        ComboBox_VEÍCULOFilial.value = Empty
        TextBox_PLACAFilial.value = Empty
        ComboBox_AnoFilial.value = Empty
        ComboBox_CORFilial.value = Empty
        TextBox_Valor_EntradaFilial.value = Empty
        TextBox_Valor_SaidaFilial.value = Empty
        TextBox_DataFilial.value = Empty
        TextBox_Valor_LucroFilial.value = "R$ 0,00"
        ComboBox_PorFilial.value = "12"
        TextBoxBaseICMSFilial.value = "R$ 0,00"
        TextBox_ICMSFilial.value = "R$ 0,00"
        TextBox_PISFilial.value = "R$ 0,00"
        TextBox_COFINSFilial.value = "R$ 0,00"
        ListBoxListaFilial.Clear
        Me.TextBox_ESTOQUEFilial.SetFocus

        Sheets("Saidas Filial").Select
        ComboBox_OperaçãoFilial.SetFocus

        LabelStatus = "NFe inserida com sucesso..."

        Cad_NFE.MultiPage1.value = 1
        Cad_NFE.N_NfeFilial.value = ""
        Cad_NFE.ComboBox_OutroEnvioFilial.value = ""
        Cad_NFE.TextBoxChaveDeAcessoFilial.value = ""
        Cad_NFE.ComboBox_OperaçãoFilial.value = "Compra"
        Cad_NFE.UNFilial.value = "1"
        Cad_NFE.CODFilial.value = ""
        Cad_NFE.ComboBox_VEÍCULOFilial.value = ""
        Cad_NFE.PLACAFilial.value = ""
        Cad_NFE.Combo_ANO_MODFilial.value = ""
        Cad_NFE.ComboBox_CORFilial.value = ""
        Cad_NFE.txt_ValorFilial.value = ""
        Cad_NFE.DataFilial.value = Format(Now, "dd/mm/yyyy")
        Cad_NFE.LabelTabelaFilial1.CAPTION = ""
        Cad_NFE.TextBoxPesquisaFilial.value = ""
        Cad_NFE.ListBoxLista2.Clear

        Cad_NFE_Saida.Hide

        ' Definir o índice da página desejada
        Dim indicePagina As Integer
        indicePagina = 1

        ' Definir a página ativa no MultiPage
        Cad_NFE.MultiPage1.value = indicePagina

        'Salva a pasta de trabalho
        ThisWorkbook.Save
        
        'Seleciona a célula A3 da planilha Entradas Filial
        With ThisWorkbook.Worksheets("Entradas Filial")
            .Activate
            .Range("A3").Select
        End With
        
        'Mostra novamente o formulário
        Cad_NFE.Show

    Else
        MsgBox "NFe não encontrada!", vbExclamation, "Excluir NFe"
        LabelStatus = "NFe não encontrada..."
    End If
End With

Exit Sub
End Sub

Private Sub CommandButton_LimparFilial_Click()

    ComboBox_OperaçãoFilial.value = "Devolução"
    TextBox_N_SaidaFilial.value = ""
    TextBox_N_EntradaFilial.value = ""
    TextBox_ESTOQUEFilial.value = "1"
    ComboBox_OutroEnvioFilial.value = ""
    TextBoxChaveDeAcessoFilial.value = ""
    CODFilial.value = ""
    ComboBox_VEÍCULOFilial.value = ""
    TextBox_PLACAFilial.value = ""
    ComboBox_AnoFilial.value = ""
    ComboBox_CORFilial.value = ""
    TextBox_Valor_EntradaFilial.value = ""
    TextBox_Valor_SaidaFilial.value = ""
    TextBox_DataFilial.value = Format(Now, "dd/mm/yyyy")
    TextBox_Valor_LucroFilial.value = "R$ 0,00"
    ComboBox_PorFilial.value = "12"
    TextBoxBaseICMSFilial.value = "R$ 0,00"
    TextBox_ICMSFilial.value = "R$ 0,00"
    TextBox_PISFilial.value = "R$ 0,00"
    TextBox_COFINSFilial.value = "R$ 0,00"
    TextBoxFiltroFilial.value = ""
    ComboBoxCamposFilial.value = "Nº Nfe Saida"
    ListBoxListaFilial.Clear
    Me.TextBox_ESTOQUEFilial.SetFocus

    Sheets("Saidas Filial").Select
    LabelStatus = "Todos os campos limpos..."
End Sub

Private Sub CommandButton_SaidasFilial_Click()
    
    LabelStatus = "Planilha Saidas filial..."
    Sheets("Saidas Filial").Select
End Sub

Private Sub CommandButtonNCM_Click()
    
    LabelStatus = "Pesquisa NCM..."
    Cad_NCM.Show
End Sub

Private Sub CommandButtonNCMFilial_Click()
    
    LabelStatus = "Pesquisa NCM..."
    Cad_NCM.Show
End Sub

Private Sub CommandButtonCadNFE_Click()

    Cad_NFE_Saida.Hide
        Sheets("Entradas").Select
    Cad_NFE.Show
End Sub

Private Sub CommandButtonFechar_Click()

    Cad_NFE_Saida.Hide
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

Private Sub Command_Entrada_Click()
    
    LabelStatus = "Planilha Entradas matriz..."
    Sheets("Entradas").Select
End Sub

Private Sub CommandButton_Saidas_Click()
    
    LabelStatus = "Planilha Saidas matriz..."
    Sheets("Saidas").Select
End Sub

Private Sub MultiPage1_Change()

    If MultiPage1.SelectedItem.Index = 0 Then
        Sheets("Saidas").Select
    ElseIf MultiPage1.SelectedItem.Index = 1 Then
        Sheets("Saidas Filial").Select
    End If
End Sub

Private Sub ComboBox_PorFilial_AfterUpdate()

    TextBoxBaseICMSFilial = (TextBox_Valor_SaidaFilial / 100) * (5)
        TextBoxBaseICMSFilial.text = Format(TextBoxBaseICMSFilial.text, "Currency")

    TextBox_ICMSFilial = (TextBox_Valor_SaidaFilial / 100) * (5) / 100 * (ComboBox_PorFilial)
        TextBox_ICMSFilial.text = Format(TextBox_ICMSFilial.text, "Currency")
    
    TextBox_PISFilial = (TextBox_Valor_LucroFilial / 100) * (0.65)
        TextBox_PISFilial.text = Format(TextBox_PISFilial.text, "Currency")
        
    TextBox_COFINSFilial = (TextBox_Valor_LucroFilial / 100) * (3)
        TextBox_COFINSFilial.text = Format(TextBox_COFINSFilial.text, "Currency")

End Sub

Private Sub CommandButton_Excluir_Click()

'Declarar a variável Resp para receber uma resposta
Dim resp As Integer
Dim c As Range
Dim msg, Style, Title, Help, Ctxt, Response, MyString

'Fazer a busca do registro digitado pelo usuário
LabelStatus = "Confirma a exclusão da NFe..."
With Worksheets("Saidas").Range("A:A")
Set c = .Find(TextBox_N_Saida.value, LookIn:=xlValues, LookAt:=xlWhole)

If Not c Is Nothing Then
    resp = MsgBox("Tem certeza que deseja excluir Esta NFe?", vbYesNo, "Confirmação")
    If resp = vbYes Then
         c.Select
         Selection.EntireRow.Delete
         'Limpar as caixas de texto
            ComboBox_Operação.value = Empty
            TextBox_N_Saida.value = Empty
            TextBox_N_Entrada.value = Empty
            TextBox_ESTOQUE.value = Empty
            ComboBox_VEÍCULO.value = Empty
            TextBox_PLACA.value = Empty
            ComboBox_Ano.value = Empty
            ComboBox_COR.value = Empty
            TextBox_Valor_Entrada.value = Empty
            TextBox_Valor_Saida.value = Empty
            TextBox_Data.value = Empty
            ComboBox_Por.value = "12"
            TextBox_Valor_Lucro.value = Empty
            ComboBox_Por.value = "12"
            TextBoxBaseICMS.value = Empty
            TextBox_ICMS.value = Empty
            TextBox_PIS.value = Empty
            TextBox_COFINS.value = Empty
            ComboBoxCampos.value = "Nº Nfe Saida"
            TextBoxFiltro.value = Empty
            ListBoxLista.Clear
            ActiveWorkbook.Save
         'Colocar o foco na primeira caixa de texto
         ComboBox_Operação.SetFocus
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

Private Sub CommandButtonCopiarSaida_Click()
    
    Cad_NFE.MultiPage1.value = 0
    Cad_NFE.UN.value = Cad_NFE_Saida.TextBox_ESTOQUE.value
    'Cad_NFE.N_Nfe.Value = Cad_NFE_Saida.TextBox_N_Entrada.Value
    Cad_NFE.COD.value = Cad_NFE_Saida.COD.value
    Cad_NFE.ComboBox_VEÍCULO.value = Cad_NFE_Saida.ComboBox_VEÍCULO.value
    Cad_NFE.PLACA.value = Cad_NFE_Saida.TextBox_PLACA.value
    Cad_NFE.Combo_ANO_MOD.value = Cad_NFE_Saida.ComboBox_Ano.value
    Cad_NFE.ComboBox_COR.value = Cad_NFE_Saida.ComboBox_COR.value
    Cad_NFE.txt_Valor.value = Cad_NFE_Saida.TextBox_Valor_Entrada.value
    Cad_NFE.Data.value = Format(Now, "dd/mm/yyyy")
    Cad_NFE_Saida.Hide
        Sheets("Entradas").Select
    Cad_NFE.Show
End Sub

Private Sub CommandButtonCopiarSaida_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButton_Fechar_Click()

    Cad_NFE_Saida.Hide
End Sub

Private Sub CommandButton_Gravar_Click()

    Dim vernumeronfe As Double
    Dim planilha As Worksheet
    Dim wsDemo As Worksheet
    Dim linhaDemo As Long
    Dim c As Range
    Dim indicePagina As Integer
    Dim operacaoCancelada As Boolean
    Dim xmlDoc As Object
    Dim arquivoXML As String

    'Controle dos eventos do Excel
    Dim estadoEventos As Boolean
    Dim numeroErro As Long
    Dim descricaoErro As String

    estadoEventos = Application.EnableEvents

    On Error GoTo TratarErro

    Set planilha = ThisWorkbook.Worksheets("Saidas")

    operacaoCancelada = _
        (StrComp(Trim$(ComboBox_Operação.value), _
                 "Cancelado", vbTextCompare) = 0)

    '==========================================================
    'IMPORTAR XML
    '==========================================================
    If Not operacaoCancelada Then

        LabelStatus = "Lendo o arquivo XML..."

        If Not CarregarXMLNFe(Me, xmlDoc, arquivoXML) Then
            LabelStatus = "Gravação cancelada: o XML não foi importado."
            Exit Sub
        End If

    End If

    If Trim$(CStr(TextBox_N_Saida.value)) = vbNullString Then

        MsgBox "Informe o número da NF-e.", _
               vbExclamation, "Número da NFe"

        LabelStatus = "Gravação cancelada: número não informado."
        Exit Sub

    End If

    '==========================================================
    'VERIFICAR DUPLICIDADE
    '==========================================================
    vernumeronfe = WorksheetFunction.CountIf( _
        planilha.Range("A:A"), _
        Trim$(CStr(TextBox_N_Saida.value)))

    If vernumeronfe > 0 Then

        MsgBox "Número da NFe já cadastrado!", _
               vbCritical, "Número da NFe"

        LabelStatus = "Número da NFe já cadastrado!"

        TextBox_N_Saida.BorderColor = &HFF&
        TextBox_N_Saida.BackColor = &H8080FF
        TextBox_N_Saida.ForeColor = &HC0&

        Exit Sub

    End If

    '==========================================================
    'VALIDAR XML
    '==========================================================
    If operacaoCancelada Then

        LabelStatus = "Gravando a NF-e cancelada..."

    Else

        LabelStatus = "Validando o arquivo XML..."

        If Not ValidarXMLNFe(Me, xmlDoc, arquivoXML) Then

            LabelStatus = _
                "Gravação cancelada: existem divergências no XML."

            Exit Sub

        End If

        'Localizar a NF-e em Entradas
        With ThisWorkbook.Worksheets("Entradas").Range("A:A")

            Set c = .Find( _
                What:=TextBox_N_Entrada.value, _
                After:=.Cells(.Cells.count), _
                LookIn:=xlValues, _
                LookAt:=xlWhole, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False, _
                SearchFormat:=False)

        End With

        If c Is Nothing Then

            MsgBox "NFe não encontrada!"
            LabelStatus = "NFe não encontrada..."
            Exit Sub

        End If

        LabelStatus = "XML validado. Gravando a NF-e..."

    End If

    Set xmlDoc = Nothing

    '==========================================================
    'GRAVAR EM SAIDAS
    '==========================================================
    ThisWorkbook.Worksheets("Saidas").Activate
    Range("A2").Select

    Do While Not IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
    Loop

    ActiveCell.value = TextBox_N_Saida.value
    ActiveCell.Offset(0, 1).value = TextBox_N_Entrada.value
    ActiveCell.Offset(0, 2).value = ComboBox_Operação.value
    ActiveCell.Offset(0, 3).value = TextBox_ESTOQUE.value
    ActiveCell.Offset(0, 4).value = ComboBox_VEÍCULO.value
    ActiveCell.Offset(0, 5).value = TextBox_PLACA.value
    ActiveCell.Offset(0, 6).value = ComboBox_Ano.value
    ActiveCell.Offset(0, 7).value = ComboBox_COR.value
    ActiveCell.Offset(0, 8).value = CDbl(TextBox_Valor_Entrada.value)
    ActiveCell.Offset(0, 9).value = CDbl(TextBox_Valor_Saida.value)
    ActiveCell.Offset(0, 11).value = CDate(TextBox_Data.value)
    ActiveCell.Offset(0, 12).value = CDbl(ComboBox_Por.value) / 100
    ActiveCell.Offset(0, 13).value = CDbl(TextBox_ICMS.value)
    ActiveCell.Offset(0, 14).value = CDbl(TextBox_PIS.value)
    ActiveCell.Offset(0, 15).value = CDbl(TextBox_COFINS.value)

    '==========================================================
    'DEMONSTRAÇÃO
    '==========================================================
    If StrComp(Trim$(ComboBox_Operação.value), _
               "Demonstração", vbTextCompare) = 0 Then

        Set wsDemo = ThisWorkbook.Worksheets("Demonstração")

        linhaDemo = _
            wsDemo.Cells(wsDemo.rows.count, "A").End(xlUp).Row + 1

        If linhaDemo < 2 Then linhaDemo = 2

        With wsDemo

            .Cells(linhaDemo, 1).value = TextBox_N_Saida.value
            .Cells(linhaDemo, 2).value = ComboBox_Operação.value
            .Cells(linhaDemo, 3).value = TextBox_ESTOQUE.value
            .Cells(linhaDemo, 4).value = ComboBox_VEÍCULO.value
            .Cells(linhaDemo, 5).value = COD.value
            .Cells(linhaDemo, 6).value = TextBox_PLACA.value
            .Cells(linhaDemo, 7).value = ComboBox_Ano.value
            .Cells(linhaDemo, 8).value = ComboBox_COR.value
            .Cells(linhaDemo, 9).value = CDbl(TextBox_Valor_Saida.value)
            .Cells(linhaDemo, 10).value = CDate(TextBox_Data.value)

        End With

    End If

    '==========================================================
    'TRANSFERÊNCIA - ROTINA REVISADA
    '==========================================================
    If StrComp(Trim$(ComboBox_Operação.value), _
               "Transferência", vbTextCompare) = 0 Then

        LabelStatus = "Gravando transferência para a filial..."

        'Impede a execução de eventos Worksheet_Change
        'durante a gravação nas planilhas da filial.
        Application.EnableEvents = False

        '------------------------------------------------------
        'ENTRADAS FILIAL
        '------------------------------------------------------
        Set wsDemo = ThisWorkbook.Worksheets("Entradas Filial")

        'Próxima linha disponível com base na coluna A.
        'Não insere nem exclui linhas.
        linhaDemo = _
            wsDemo.Cells(wsDemo.rows.count, "A").End(xlUp).Row + 1

        If linhaDemo < 2 Then linhaDemo = 2

        With wsDemo

            'Gravação limitada às colunas A até L.
            'As colunas M, N e O não são modificadas.

            .Cells(linhaDemo, 1).value = TextBox_N_Saida.value
            .Cells(linhaDemo, 2).value = "Compra"
            .Cells(linhaDemo, 3).value = TextBox_ESTOQUE.value
            .Cells(linhaDemo, 4).value = ComboBox_VEÍCULO.value
            .Cells(linhaDemo, 5).value = COD.value
            .Cells(linhaDemo, 6).value = TextBox_PLACA.value
            .Cells(linhaDemo, 7).value = ComboBox_Ano.value
            .Cells(linhaDemo, 8).value = ComboBox_COR.value
            .Cells(linhaDemo, 9).value = CDbl(TextBox_Valor_Saida.value)
            .Cells(linhaDemo, 10).value = CDate(TextBox_Data.value)
            .Cells(linhaDemo, 11).value = ComboBox_OutroEnvio.value
            .Cells(linhaDemo, 12).value = TextBoxChaveDeAcesso.value
            
            celula.Errors(xlNumberAsText).Ignore = True

        End With

        '------------------------------------------------------
        'ENTRADAS ANUAL FILIAL
        '------------------------------------------------------
        Set wsDemo = ThisWorkbook.Worksheets("Entradas Anual Filial")

        linhaDemo = _
            wsDemo.Cells(wsDemo.rows.count, "A").End(xlUp).Row + 1

        If linhaDemo < 2 Then linhaDemo = 2

        With wsDemo

            .Cells(linhaDemo, 1).value = TextBox_N_Saida.value
            .Cells(linhaDemo, 2).value = "Compra"
            .Cells(linhaDemo, 3).value = TextBox_ESTOQUE.value
            .Cells(linhaDemo, 4).value = ComboBox_VEÍCULO.value
            .Cells(linhaDemo, 5).value = COD.value
            .Cells(linhaDemo, 6).value = TextBox_PLACA.value
            .Cells(linhaDemo, 7).value = ComboBox_Ano.value
            .Cells(linhaDemo, 8).value = ComboBox_COR.value
            .Cells(linhaDemo, 9).value = CDbl(TextBox_Valor_Saida.value)
            .Cells(linhaDemo, 10).value = CDate(TextBox_Data.value)

        End With

        'Restaura a configuração original dos eventos.
        Application.EnableEvents = estadoEventos

        LabelStatus = "Transferência registrada na filial."

    End If

    '==========================================================
    'GRAVAR EM TODAS SAIDAS
    '==========================================================
    ThisWorkbook.Worksheets("Todas Saidas").Activate
    Range("A2").Select

    Do While Not IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
    Loop

    ActiveCell.value = TextBox_N_Saida.value
    ActiveCell.Offset(0, 1).value = TextBox_N_Entrada.value
    ActiveCell.Offset(0, 2).value = ComboBox_Operação.value
    ActiveCell.Offset(0, 3).value = TextBox_ESTOQUE.value
    ActiveCell.Offset(0, 4).value = ComboBox_VEÍCULO.value
    ActiveCell.Offset(0, 5).value = TextBox_PLACA.value
    ActiveCell.Offset(0, 6).value = ComboBox_Ano.value
    ActiveCell.Offset(0, 7).value = ComboBox_COR.value
    ActiveCell.Offset(0, 8).value = CDbl(TextBox_Valor_Entrada.value)
    ActiveCell.Offset(0, 9).value = CDbl(TextBox_Valor_Saida.value)
    ActiveCell.Offset(0, 11).value = CDate(TextBox_Data.value)
    ActiveCell.Offset(0, 12).value = CDbl(ComboBox_Por.value) / 100
    ActiveCell.Offset(0, 13).value = CDbl(TextBox_ICMS.value)
    ActiveCell.Offset(0, 14).value = CDbl(TextBox_PIS.value)
    ActiveCell.Offset(0, 15).value = CDbl(TextBox_COFINS.value)

    '==========================================================
    'EXCLUIR DE ENTRADAS SOMENTE SE NÃO FOR CANCELADO
    '==========================================================
    If Not operacaoCancelada Then

        LabelStatus = "Localizando a NF-e em Entradas..."

        'Mantido o comportamento original.
        'ATENÇÃO: desloca A:N para cima na planilha Entradas.
        c.Resize(1, 14).Delete Shift:=xlUp

    End If

    Application.CalculateFull
    ThisWorkbook.Save

    '==========================================================
    'LIMPAR O FORMULÁRIO DE SAÍDA
    '==========================================================
    ComboBox_Operação.value = "Venda"
    TextBox_N_Saida.value = ""
    TextBox_N_Entrada.value = ""
    ComboBox_OutroEnvio.value = ""
    TextBoxChaveDeAcesso.value = ""
    TextBox_ESTOQUE.value = "1"
    COD.value = ""
    ComboBox_VEÍCULO.value = ""
    TextBox_PLACA.value = ""
    ComboBox_Ano.value = ""
    ComboBox_COR.value = ""
    TextBox_Valor_Entrada.value = ""
    TextBox_Valor_Saida.value = ""
    TextBox_Data.value = Format(Now, "dd/mm/yyyy")
    TextBox_Valor_Lucro.value = ""
    ComboBox_Por.value = "12"
    TextBoxBaseICMS.value = ""
    TextBox_ICMS.value = ""
    TextBoxBasePISeCOFINS.value = ""
    TextBox_PIS.value = ""
    TextBox_COFINS.value = ""
    TextBoxFiltro.value = ""
    ComboBoxCampos.value = "Nº Nfe Saida"

    ListBoxLista.Clear

    Me.TextBox_ESTOQUE.SetFocus

    VisibilidadeOutroEnvio Me

    ThisWorkbook.Worksheets("Saidas").Activate
    ComboBox_Operação.SetFocus

    If operacaoCancelada Then
        LabelStatus = "NF-e cancelada gravada com sucesso."
    Else
        LabelStatus = "NFe inserida com sucesso..."
    End If

    '==========================================================
    'LIMPAR O FORMULÁRIO CAD_NFE
    '==========================================================
    With Cad_NFE

        .MultiPage1.value = 1
        .N_Nfe.value = ""
        .ComboBox_Operação.value = "Compra"
        .ComboBox_OutroEnvio.value = ""
        .UN.value = "1"
        .COD.value = ""
        .ComboBox_VEÍCULO.value = ""
        .PLACA.value = ""
        .Combo_ANO_MOD.value = ""
        .ComboBox_COR.value = ""
        .txt_Valor.value = ""
        .Data.value = Format(Now, "dd/mm/yyyy")
        .LabelTabela1.CAPTION = ""
        .TextBoxPesquisa.value = ""
        .ListBoxLista.Clear
        .LabelStatus = "NFe deinserida com sucesso..."

    End With

    Cad_NFE_Saida.Hide

    indicePagina = 0
    Cad_NFE.MultiPage1.value = indicePagina

    ThisWorkbook.Save

    With ThisWorkbook.Worksheets("Entradas")

        .Activate
        .Range("A3").Select

    End With

    Cad_NFE.Show

    Exit Sub

'==============================================================
'TRATAMENTO DE ERROS
'==============================================================
TratarErro:

    numeroErro = Err.Number
    descricaoErro = Err.Description

    'Garante a restauração dos eventos mesmo em caso de erro.
    On Error Resume Next
    Application.EnableEvents = estadoEventos
    Set xmlDoc = Nothing

    LabelStatus = "A gravação foi interrompida por um erro."

    MsgBox "Erro " & numeroErro & ": " & descricaoErro & vbCrLf & _
           "Confira os registros nas planilhas antes de tentar gravar novamente.", _
           vbCritical, "Gravação da NF-e"

End Sub

Private Sub CommandButton_Limpar_Click()

    ComboBox_Operação.value = "Venda"
    TextBox_N_Saida.value = ""
    TextBox_N_Entrada.value = ""
    TextBox_ESTOQUE.value = "1"
    ComboBox_OutroEnvio.value = ""
    TextBoxChaveDeAcesso.value = ""
    COD.value = ""
    ComboBox_VEÍCULO.value = ""
    TextBox_PLACA.value = ""
    ComboBox_Ano.value = ""
    ComboBox_COR.value = ""
    TextBox_Valor_Entrada.value = ""
    TextBox_Valor_Saida.value = ""
    TextBox_Data.value = Format(Now, "dd/mm/yyyy")
    TextBox_Valor_Lucro.value = ""
    ComboBox_Por.value = "12"
    TextBoxBaseICMS.value = ""
    TextBox_ICMS.value = ""
    TextBoxBasePISeCOFINS = ""
    TextBox_PIS.value = ""
    TextBox_COFINS.value = ""
    TextBoxFiltro.value = ""
    ComboBoxCampos.value = "Nº Nfe Saida"
    ListBoxLista.Clear
    Me.TextBox_ESTOQUE.SetFocus
    
    VisibilidadeOutroEnvio Me
    
    Sheets("Saidas").Select
    LabelStatus = "Todos os campos limpos..."
End Sub

Private Sub TextBox_Valor_EntradaFilial_AfterUpdate()

    Dim valor1 As Double
    Dim valor2 As Double
        valor1 = 0
        valor2 = 0

    If Me.TextBox_Valor_SaidaFilial = Empty Then
        Else
    valor1 = Me.TextBox_Valor_SaidaFilial.value
        End If
    If Me.TextBox_Valor_EntradaFilial = Empty Then
        Else
    valor2 = Me.TextBox_Valor_EntradaFilial.value
        End If

    Me.TextBox_Valor_LucroFilial = valor1 - valor2
    
    TextBox_Valor_EntradaFilial.text = Format(TextBox_Valor_EntradaFilial.text, "Currency")

End Sub

Private Sub TextBox_Valor_LucroFilial_AfterUpdate()

Dim valor1 As Double
    Dim valor2 As Double
        valor1 = 0
        valor2 = 0

    If Me.TextBox_Valor_SaidaFilial = Empty Then
        Else
    valor1 = Me.TextBox_Valor_SaidaFilial.value
        End If
    If Me.TextBox_Valor_EntradaFilial = Empty Then
        Else
    valor2 = Me.TextBox_Valor_EntradaFilial.value
        End If

    Me.TextBox_Valor_LucroFilial = valor1 - valor2
    
    TextBox_Valor_LucroFilial.text = Format(TextBox_Valor_LucroFilial.text, "Currency")

End Sub

Private Sub TextBox_Valor_Saida_AfterUpdate()
    
    LabelStatus = "Calculando..."
    Dim valor1 As Double
    Dim valor2 As Double
    Dim valor3 As Double
    Dim valor4 As Double
    Dim valor5 As Double
        valor1 = 0
        valor2 = 0
        valor3 = 0
        valor4 = 0
        valor5 = 0

    If Me.TextBox_Valor_Saida = Empty Then
        Else
    valor1 = Me.TextBox_Valor_Saida.value
        End If
    If Me.TextBox_Valor_Entrada = Empty Then
        Else
    valor2 = Me.TextBox_Valor_Entrada.value
        End If
    If Me.ComboBox_Por = Empty Then
        Else
    valor3 = Me.ComboBox_Por.value
        End If
    If Me.TextBox_Valor_Lucro = Empty Then
        Else
    valor4 = Me.TextBox_Valor_Lucro.value
        End If
    If Me.TextBox_ICMS = Empty Then
        Else
    valor5 = Me.TextBox_ICMS.value
        End If

    Me.TextBox_Valor_Lucro = (valor1 - valor2)
    Me.TextBoxBaseICMS = (valor1 / 100) * (5)
    Me.TextBox_ICMS = (valor1 / 100) * (5) / 100 * (valor3)
    Me.TextBoxBasePISeCOFINS = (valor1 - valor2) - (valor1 / 100) * (5) / 100 * (valor3)
    Me.TextBox_PIS = ((valor1 - valor2) - (valor1 / 100) * (5) / 100 * (valor3)) * (0.65) / 100
    Me.TextBox_COFINS = ((valor1 - valor2) - (valor1 / 100) * (5) / 100 * (valor3)) * (3) / 100
       
    TextBox_Valor_Lucro.text = Format(TextBox_Valor_Lucro.text, "Currency")
    TextBoxBaseICMS.text = Format(TextBoxBaseICMS.text, "Currency")
    TextBox_ICMS.text = Format(TextBox_ICMS.text, "Currency")
    TextBoxBasePISeCOFINS.text = Format(TextBoxBasePISeCOFINS.text, "Currency")
    TextBox_PIS.text = Format(TextBox_PIS.text, "Currency")
    TextBox_COFINS.text = Format(TextBox_COFINS.text, "Currency")
    TextBox_Valor_Saida.text = Format(TextBox_Valor_Saida.text, "Currency")
    TextBox_Valor_Lucro.text = Format(TextBox_Valor_Lucro.text, "Currency")
    
    LabelStatus = "Calculado..."
End Sub

Private Sub TextBox_Valor_SaidaFilial_AfterUpdate()

    Dim valor1 As Double
    Dim valor2 As Double
        valor1 = 0
        valor2 = 0

    If Me.TextBox_Valor_SaidaFilial = Empty Then
        Else
    valor1 = Me.TextBox_Valor_SaidaFilial.value
        End If
    If Me.TextBox_Valor_EntradaFilial = Empty Then
        Else
    valor2 = Me.TextBox_Valor_EntradaFilial.value
        End If

    Me.TextBox_Valor_LucroFilial = valor1 - valor2
    
    TextBox_Valor_SaidaFilial.text = Format(TextBox_Valor_SaidaFilial.text, "Currency")

End Sub

Private Sub TextBoxBaseICMSFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    TextBoxBaseICMSFilial.BorderColor = &H8000000D
    TextBoxBaseICMSFilial.BackColor = &H80000002
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

Private Sub CommandButtonSalvarTudo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelBakupBakup.Visible = True
End Sub

Private Sub CommandButtonCadNFE_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelEntradas.Visible = True
End Sub

Private Sub CommandButtonFechar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButton_Gravar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButton_Limpar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButton_Excluir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub Command_Entrada_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButton_Saidas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonNCM_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButton_GravarFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButton_LimparFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButton_ExcluirFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub Command_EntradaFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButton_SaidasFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub CommandButtonNCMFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
End Sub

Private Sub FrameMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    LabelBakupBakup.Visible = False
    LabelEntradas.Visible = False
End Sub

Private Sub FrameMenu_Enter()
    
    FrameMenu.ForeColor = &H8000000D
    FrameMenu.BorderColor = &H8000000D
End Sub

Private Sub FrameMenu_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    FrameMenu.ForeColor = &H8000000A
    FrameMenu.BorderColor = &H8000000A
End Sub

Private Sub Frame_Ações_Enter()

    TextBox_Data.BorderColor = &H8000000A
    TextBox_Data.BackColor = &H80000005
    Label_Data.ForeColor = &H80000011
    
    Frame_Ações.ForeColor = &H8000000D
    Frame_Ações.BorderColor = &H8000000D
End Sub

Private Sub Frame_Ações_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Frame_Ações.ForeColor = &H8000000A
    Frame_Ações.BorderColor = &H8000000A
End Sub

Private Sub Frame_Dados_da_NFe_Enter()
    
    Frame_Dados_da_NFe.ForeColor = &H8000000D
    Frame_Dados_da_NFe.BorderColor = &H8000000D
End Sub

Private Sub Frame_Dados_da_NFe_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Frame_Dados_da_NFe.ForeColor = &H8000000A
    Frame_Dados_da_NFe.BorderColor = &H8000000A
End Sub

Private Sub Frame_Veículo_Enter()
        
    TextBoxChaveDeAcesso.BorderColor = &H8000000A
    TextBoxChaveDeAcesso.BackColor = &H80000005
    LabelChaveDeAcesso.ForeColor = &H80000011
        
    TextBox_N_Entrada.BorderColor = &H8000000A
    TextBox_N_Entrada.BackColor = &H80000005
    Label_NFe_Entrada.ForeColor = &H80000011
    
    Frame_Veículo.ForeColor = &H8000000D
    Frame_Veículo.BorderColor = &H8000000D
End Sub

Private Sub Frame_Veículo_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Frame_Veículo.ForeColor = &H8000000A
    Frame_Veículo.BorderColor = &H8000000A
End Sub

Private Sub Frame_Valores_Enter()

    ComboBox_COR.BorderColor = &H8000000A
    ComboBox_COR.BackColor = &H80000005
    Label_COR.ForeColor = &H80000011
    
    Frame_Valores.ForeColor = &H8000000D
    Frame_Valores.BorderColor = &H8000000D
End Sub

Private Sub Frame_Valores_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Frame_Valores.ForeColor = &H8000000A
    Frame_Valores.BorderColor = &H8000000A
End Sub

Private Sub Frame_Data_Enter()

    TextBox_Data.BorderColor = &H8000000D
    TextBox_Data.BackColor = &H80000002
    Label_Data.ForeColor = &H80000012
    
    TextBox_COFINS.BorderColor = &H8000000A
    TextBox_COFINS.BackColor = &H80000005
    Label_COFINS.ForeColor = &H80000011
    
    Frame_Data.ForeColor = &H8000000D
    Frame_Data.BorderColor = &H8000000D
End Sub

Private Sub Frame_Data_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Frame_Data.ForeColor = &H8000000A
    Frame_Data.BorderColor = &H8000000A
End Sub

Private Sub Frame_Pesquisa_Enter()

    Frame_Pesquisa.ForeColor = &H8000000D
    Frame_Pesquisa.BorderColor = &H8000000D
End Sub

Private Sub Frame_Pesquisa_Exit(ByVal cancel As MSForms.ReturnBoolean)
  
    Frame_Pesquisa.ForeColor = &H80000011
    Frame_Pesquisa.BorderColor = &H8000000A
End Sub

Private Sub Frame_AçõesFilial_Enter()

    TextBox_DataFilial.BorderColor = &H8000000A
    TextBox_DataFilial.BackColor = &H80000005
    Label_DataFilial.ForeColor = &H80000011
    
    Frame_AçõesFilial.ForeColor = &H8000000D
    Frame_AçõesFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_AçõesFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Frame_AçõesFilial.ForeColor = &H8000000A
    Frame_AçõesFilial.BorderColor = &H8000000A
End Sub

Private Sub Frame_PesquisaFilial_Enter()
    
    Fame_PesquisaFilial.ForeColor = &H8000000D
    Fame_PesquisaFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_PesquisaFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Fame_PesquisaFilial.ForeColor = &H8000000A
    Fame_PesquisaFilial.BorderColor = &H8000000A
End Sub

Private Sub Frame_Dados_da_NFeFilial_Enter()
    
    Frame_Dados_da_NFeFilial.ForeColor = &H8000000D
    Frame_Dados_da_NFeFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_Dados_da_NFeFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Frame_Dados_da_NFeFilial.ForeColor = &H8000000A
    Frame_Dados_da_NFeFilial.BorderColor = &H8000000A
End Sub

Private Sub Frame_VeículoFilial_Enter()

    TextBoxChaveDeAcessoFilial.BorderColor = &H8000000A
    TextBoxChaveDeAcessoFilial.BackColor = &H80000005
    LabelChaveDeAcessoFilial.ForeColor = &H80000011
    
    Frame_VeículoFilial.ForeColor = &H8000000D
    Frame_VeículoFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_VeículoFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Frame_VeículoFilial.ForeColor = &H8000000A
    Frame_VeículoFilial.BorderColor = &H8000000A
End Sub

Private Sub Frame_ValoresFilial_Enter()

    ComboBox_CORFilial.BorderColor = &H8000000A
    ComboBox_CORFilial.BackColor = &H80000005
    Label_CORFilial.ForeColor = &H80000011
    
    Frame_ValoresFilial.ForeColor = &H8000000D
    Frame_ValoresFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_ValoresFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Frame_ValoresFilial.ForeColor = &H8000000A
    Frame_ValoresFilial.BorderColor = &H8000000A
End Sub

Private Sub Frame_ImpostoFilial_Enter()

    TextBox_Valor_LucroFilial.BorderColor = &H8000000A
    TextBox_Valor_LucroFilial.BackColor = &H80000005
    Label_LucroFilial.ForeColor = &H80000011
    
    Frame_ImpostoFilial.ForeColor = &H8000000D
    Frame_ImpostoFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_ImpostoFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Frame_ImpostoFilial.ForeColor = &H8000000A
    Frame_ImpostoFilial.BorderColor = &H8000000A
End Sub

Private Sub Frame_DataFilial_Enter()

    TextBox_DataFilial.BorderColor = &H8000000D
    TextBox_DataFilial.BackColor = &H80000002
    Label_DataFilial.ForeColor = &H80000012
    
    TextBox_COFINSFilial.BorderColor = &H8000000A
    TextBox_COFINSFilial.BackColor = &H80000005
    Label_COFINSFilial.ForeColor = &H80000011
    
    Frame_DataFilial.ForeColor = &H8000000D
    Frame_DataFilial.BorderColor = &H8000000D
End Sub

Private Sub Frame_DataFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)
    
    Frame_DataFilial.ForeColor = &H8000000A
    Frame_DataFilial.BorderColor = &H8000000A
End Sub

Private Sub Fame_PesquisaFilial_Enter()

    Fame_PesquisaFilial.ForeColor = &H8000000D
    Fame_PesquisaFilial.BorderColor = &H8000000D
End Sub

Private Sub Fame_PesquisaFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)
  
    Fame_PesquisaFilial.ForeColor = &H80000011
    Fame_PesquisaFilial.BorderColor = &H8000000A
End Sub

Private Sub TextBox_ESTOQUE_Enter()

    TextBox_ESTOQUE.BorderColor = &H8000000D
    TextBox_ESTOQUE.BackColor = &H80000002
    Label_Quant.ForeColor = &H80000012
End Sub

Private Sub TextBox_ESTOQUE_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_ESTOQUE.BorderColor = &H8000000A
    TextBox_ESTOQUE.BackColor = &H80000005
    Label_Quant.ForeColor = &H80000011
End Sub

Private Sub ComboBox_Operação_Change()

    If Trim$(Me.ComboBox_Operação.value) = "Aguardando" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If

    If Trim$(Me.ComboBox_Operação.value) = "Ajuste" Then
        
        VisibilidadeOutroEnvio Me

        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If

    If Trim$(Me.ComboBox_Operação.value) = "Cancelado" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_N_Saida = ""
        Me.TextBox_N_Entrada = ""
        ZerarCampo Me.TextBox_Valor_Entrada
        ZerarCampo Me.TextBox_Valor_Saida
        ZerarCampo Me.TextBox_Valor_Lucro
        ZerarCampo Me.TextBoxBaseICMS
        ZerarCampo Me.TextBox_ICMS
        ZerarCampo Me.TextBoxBasePISeCOFINS
        ZerarCampo Me.TextBox_PIS
        ZerarCampo Me.TextBox_COFINS
        
    End If
    
    If Trim$(Me.ComboBox_Operação.value) = "Complemento de Nfe" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If
    
    If Trim$(Me.ComboBox_Operação.value) = "Compra" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If
    
    If Trim$(Me.ComboBox_Operação.value) = "Demonstração" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If

    If Trim$(Me.ComboBox_Operação.value) = "Devolução" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_Valor_Saida.value = Me.TextBox_Valor_Entrada.value
        ZerarCampo Me.TextBox_Valor_Lucro
        ZerarCampo Me.TextBoxBaseICMS
        ZerarCampo Me.TextBox_ICMS
        ZerarCampo Me.TextBoxBasePISeCOFINS
        ZerarCampo Me.TextBox_PIS
        ZerarCampo Me.TextBox_COFINS
        TextBox_Valor_Entrada.value = TextBox_Valor_Saida
        
    End If
    
    If Trim$(Me.ComboBox_Operação.value) = "Estorno" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If
    
    If Trim$(Me.ComboBox_Operação.value) = "Exportação" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If
    
    If Trim$(Me.ComboBox_Operação.value) = "Importação" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If
    
    If Trim$(Me.ComboBox_Operação.value) = "Inutilizada" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If
    
    If Trim$(Me.ComboBox_Operação.value) = "Remessa" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If
    
    If Trim$(Me.ComboBox_Operação.value) = "Repasse" Then
        
        VisibilidadeOutroEnvio Me
        
        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If
    
    If Trim$(Me.ComboBox_Operação.value) = "Transferência" Then
        
        VisibilidadeOutroEnvio Me, True
        
        Me.TextBox_Valor_Saida.value = Me.TextBox_Valor_Entrada.value
        
        ZerarCampo Me.TextBox_Valor_Lucro
        ZerarCampo Me.TextBoxBaseICMS
        ZerarCampo Me.TextBox_ICMS
        ZerarCampo Me.TextBoxBasePISeCOFINS
        ZerarCampo Me.TextBox_PIS
        ZerarCampo Me.TextBox_COFINS
        TextBox_Valor_Entrada.value = TextBox_Valor_Saida
        
    End If
    
    If Trim$(Me.ComboBox_Operação.value) = "Venda" Then
        
        VisibilidadeOutroEnvio Me

        Me.TextBox_Valor_Lucro.value = ""
        Me.TextBoxBaseICMS.value = ""
        Me.TextBox_ICMS.value = ""
        Me.TextBoxBasePISeCOFINS.value = ""
        Me.TextBox_PIS.value = ""
        Me.TextBox_COFINS.value = ""
        Me.TextBox_Valor_Entrada.value = ""
        Me.TextBox_Valor_Saida.value = ""
    End If
End Sub

Public Sub VisibilidadeOutroEnvio(ByVal Frm As Object, Optional ByVal Exibir As Boolean = False)

    With Frm
        .Label_OutroEnvio.Visible = Exibir
        .ComboBox_OutroEnvio.Visible = Exibir
        .LabelChaveDeAcesso.Visible = Exibir
        .TextBoxChaveDeAcesso.Visible = Exibir
    End With

End Sub

Private Sub ZerarCampo(ByVal txt As MSForms.TextBox)
    txt.value = FormatCurrency(0, 2)
End Sub

Private Sub ComboBox_Operação_Enter()

    ComboBox_Operação.BorderColor = &H8000000D
    ComboBox_Operação.BackColor = &H80000002
    Label_Operação.ForeColor = &H80000012
End Sub

Private Sub ComboBox_Operação_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_Operação.BorderColor = &H8000000A
    ComboBox_Operação.BackColor = &H80000005
    Label_Operação.ForeColor = &H80000011
End Sub

Private Sub TextBox_N_Saida_Enter()

    TextBox_N_Saida.BorderColor = &H8000000D
    TextBox_N_Saida.BackColor = &H80000002
    Label_NFe_Saida.ForeColor = &H80000012
End Sub

Private Sub TextBox_N_Saida_Exit(ByVal cancel As MSForms.ReturnBoolean)

'Verifica se existe células com o mesmo valor
Dim vernumeronfe As Double
Dim planilha As Worksheet
Set planilha = ThisWorkbook.Worksheets("Saidas")
vernumeronfe = WorksheetFunction.CountIf(planilha.Range("A:A"), TextBox_N_Saida.text)

If TextBox_N_Saida.text = "" Then
        Exit Sub
End If

If vernumeronfe > 0 Then
        MsgBox "Numero da NFe já cadastrado!", vbCritical, "Numero da NFe"
        LabelStatus = "Numero da NFe já cadastrado!"
        TextBox_N_Saida.BorderColor = &HFF&
        TextBox_N_Saida.BackColor = &H8080FF
        TextBox_N_Saida.ForeColor = &HC0&
        Exit Sub
End If

    TextBox_N_Saida.BorderColor = &H8000000A
    TextBox_N_Saida.BackColor = &H80000005
    Label_NFe_Saida.ForeColor = &H80000011
End Sub

Private Sub TextBox_N_Entrada_Enter()

    TextBox_N_Entrada.BorderColor = &H8000000D
    TextBox_N_Entrada.BackColor = &H80000002
    Label_NFe_Entrada.ForeColor = &H80000012
End Sub

Private Sub TextBox_N_Entrada_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_N_Entrada.BorderColor = &H8000000A
    TextBox_N_Entrada.BackColor = &H80000005
    Label_NFe_Entrada.ForeColor = &H80000011
End Sub

Private Sub ComboBox_OutroEnvio_Enter()

    ComboBox_OutroEnvio.BorderColor = &H8000000D
    ComboBox_OutroEnvio.BackColor = &H80000002
    Label_OutroEnvio.ForeColor = &H80000012
End Sub

Private Sub ComboBox_OutroEnvio_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_OutroEnvio.BorderColor = &H8000000A
    ComboBox_OutroEnvio.BackColor = &H80000005
    Label_OutroEnvio.ForeColor = &H80000011
End Sub

Private Sub TextBoxChaveDeAcesso_Enter()

    TextBoxChaveDeAcesso.BorderColor = &H8000000D
    TextBoxChaveDeAcesso.BackColor = &H80000002
    LabelChaveDeAcesso.ForeColor = &H80000012
End Sub

Private Sub TextBoxChaveDeAcesso_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBoxChaveDeAcesso.BorderColor = &H8000000A
    TextBoxChaveDeAcesso.BackColor = &H80000005
    LabelChaveDeAcesso.ForeColor = &H80000011
End Sub

Private Sub ComboBox_OutroEnvioFilial_Enter()

    ComboBox_OutroEnvioFilial.BorderColor = &H8000000D
    ComboBox_OutroEnvioFilial.BackColor = &H80000002
    Label_OutroEnvioFilial.ForeColor = &H80000012
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

Private Sub COD_Enter()

    COD.BorderColor = &H8000000D
    COD.BackColor = &H80000002
    Label_COD.ForeColor = &H80000012
End Sub

Private Sub COD_Exit(ByVal cancel As MSForms.ReturnBoolean)

    COD.BorderColor = &H8000000A
    COD.BackColor = &H80000005
    Label_COD.ForeColor = &H80000011
End Sub

Private Sub ComboBox_VEÍCULO_Enter()

    ComboBox_VEÍCULO.BorderColor = &H8000000D
    ComboBox_VEÍCULO.BackColor = &H80000002
    Label_VEÍCULO.ForeColor = &H80000012
End Sub

Private Sub ComboBox_VEÍCULO_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_VEÍCULO.BorderColor = &H8000000A
    ComboBox_VEÍCULO.BackColor = &H80000005
    Label_VEÍCULO.ForeColor = &H80000011
End Sub

Private Sub TextBox_PLACA_Enter()

    TextBox_PLACA.BorderColor = &H8000000D
    TextBox_PLACA.BackColor = &H80000002
    Label_PLACA.ForeColor = &H80000012
End Sub

Private Sub TextBox_PLACA_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_PLACA.BorderColor = &H8000000A
    TextBox_PLACA.BackColor = &H80000005
    Label_PLACA.ForeColor = &H80000011
End Sub

Private Sub ComboBox_Ano_Enter()

    ComboBox_Ano.BorderColor = &H8000000D
    ComboBox_Ano.BackColor = &H80000002
    Label_ANO_MOD.ForeColor = &H80000012
End Sub

Private Sub ComboBox_Ano_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_Ano.BorderColor = &H8000000A
    ComboBox_Ano.BackColor = &H80000005
    Label_ANO_MOD.ForeColor = &H80000011
End Sub

Private Sub ComboBox_COR_Enter()

    ComboBox_COR.BorderColor = &H8000000D
    ComboBox_COR.BackColor = &H80000002
    Label_COR.ForeColor = &H80000012
End Sub

Private Sub ComboBox_COR_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_COR.BorderColor = &H8000000A
    ComboBox_COR.BackColor = &H80000005
    Label_COR.ForeColor = &H80000011
End Sub

Private Sub TextBox_Valor_Entrada_Enter()

    TextBox_Valor_Entrada.BorderColor = &H8000000D
    TextBox_Valor_Entrada.BackColor = &H80000002
    Label_Valor_Entrada.ForeColor = &H80000012
End Sub

Private Sub TextBox_Valor_Entrada_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_Valor_Entrada.BorderColor = &H8000000A
    TextBox_Valor_Entrada.BackColor = &H80000005
    Label_Valor_Entrada.ForeColor = &H80000011
    
    TextBox_Valor_Entrada.text = Format(TextBox_Valor_Entrada.text, "Currency")
End Sub

Private Sub TextBox_Valor_Saida_Enter()

    TextBox_Valor_Saida.BorderColor = &H8000000D
    TextBox_Valor_Saida.BackColor = &H80000002
    Label_Valor_Saida.ForeColor = &H80000012
End Sub

Private Sub TextBox_Valor_Saida_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_Valor_Saida.BorderColor = &H8000000A
    TextBox_Valor_Saida.BackColor = &H80000005
    Label_Valor_Saida.ForeColor = &H80000011
End Sub

Private Sub TextBox_Valor_Lucro_Enter()

    TextBox_Valor_Lucro.BorderColor = &H8000000D
    TextBox_Valor_Lucro.BackColor = &H80000002
    Label_Valor_Lucro.ForeColor = &H80000012
End Sub

Private Sub TextBox_Valor_Lucro_Exit(ByVal cancel As MSForms.ReturnBoolean)

    Dim valorLucro As Double
    
    ' Tenta converter o valor do TextBox em um número
    On Error Resume Next
    valorLucro = CDbl(TextBox_Valor_Lucro.value)
    On Error GoTo 0
    
    ' Verifica se o valor é negativo
    If valorLucro <= 0 Then
        ' Define os campos como R$ 0,00
        TextBox_Valor_Lucro.value = "R$ 0,00"
        TextBoxBasePISeCOFINS.value = "R$ 0,00"
        TextBox_PIS.value = "R$ 0,00"
        TextBox_COFINS.value = "R$ 0,00"
    End If
    
    ' Configura as propriedades de cor dos campos
    TextBox_Valor_Lucro.BorderColor = &H8000000A
    TextBox_Valor_Lucro.BackColor = &H80000005
    Label_Valor_Lucro.ForeColor = &H80000011
End Sub

Private Sub TextBox_Data_Enter()

    TextBox_Data.BorderColor = &H8000000D
    TextBox_Data.BackColor = &H80000002
    Label_Data.ForeColor = &H80000012
End Sub

Private Sub TextBox_Data_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_Data.BorderColor = &H8000000A
    TextBox_Data.BackColor = &H80000005
    Label_Data.ForeColor = &H80000011
End Sub

Private Sub ComboBox_Por_Enter()

    ComboBox_Por.BorderColor = &H8000000D
    ComboBox_Por.BackColor = &H80000002
    Label_Aliq.ForeColor = &H80000012
End Sub

Private Sub ComboBox_Por_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_Por.BorderColor = &H8000000A
    ComboBox_Por.BackColor = &H80000005
    Label_Aliq.ForeColor = &H80000011
End Sub

Private Sub TextBoxBaseICMS_Enter()

    TextBoxBaseICMS.BorderColor = &H8000000D
    TextBoxBaseICMS.BackColor = &H80000002
    LabelBaseICMS.ForeColor = &H80000012
End Sub

Private Sub TextBoxBaseICMS_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBoxBaseICMS.BorderColor = &H8000000A
    TextBoxBaseICMS.BackColor = &H80000005
    LabelBaseICMS.ForeColor = &H80000011
End Sub

Private Sub TextBox_ICMS_Enter()

    TextBox_ICMS.BorderColor = &H8000000D
    TextBox_ICMS.BackColor = &H80000002
    Label_ICMS.ForeColor = &H80000012
End Sub

Private Sub TextBox_ICMS_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_ICMS.BorderColor = &H8000000A
    TextBox_ICMS.BackColor = &H80000005
    Label_ICMS.ForeColor = &H80000011
End Sub

Private Sub TextBoxBasePISeCOFINS_Enter()

    TextBoxBasePISeCOFINS.BorderColor = &H8000000D
    TextBoxBasePISeCOFINS.BackColor = &H80000002
    LabelBasePISeCOFINS.ForeColor = &H80000012
End Sub

Private Sub TextBoxBasePISeCOFINS_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBoxBasePISeCOFINS.BorderColor = &H8000000A
    TextBoxBasePISeCOFINS.BackColor = &H80000005
    LabelBasePISeCOFINS.ForeColor = &H80000011
End Sub

Private Sub TextBox_PIS_Enter()

    TextBox_PIS.BorderColor = &H8000000D
    TextBox_PIS.BackColor = &H80000002
    Label_PIS.ForeColor = &H80000012
End Sub

Private Sub TextBox_PIS_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_PIS.BorderColor = &H8000000A
    TextBox_PIS.BackColor = &H80000005
    Label_PIS.ForeColor = &H80000011
End Sub

Private Sub TextBox_COFINS_Enter()

    TextBox_COFINS.BorderColor = &H8000000D
    TextBox_COFINS.BackColor = &H80000002
    Label_COFINS.ForeColor = &H80000012
End Sub

Private Sub TextBox_COFINS_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_COFINS.BorderColor = &H8000000A
    TextBox_COFINS.BackColor = &H80000005
    Label_COFINS.ForeColor = &H80000011
End Sub

Private Sub TextBox_ESTOQUEFilial_Enter()

    TextBox_ESTOQUEFilial.BorderColor = &H8000000D
    TextBox_ESTOQUEFilial.BackColor = &H80000002
    Label_QuantFilial.ForeColor = &H80000012
End Sub

Private Sub TextBox_ESTOQUEFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_ESTOQUEFilial.BorderColor = &H8000000A
    TextBox_ESTOQUEFilial.BackColor = &H80000005
    Label_QuantFilial.ForeColor = &H80000011
End Sub

Private Sub ComboBox_OperaçãoFilial_Enter()

    ComboBox_OperaçãoFilial.BorderColor = &H8000000D
    ComboBox_OperaçãoFilial.BackColor = &H80000002
    Label_OperaçãoFilial.ForeColor = &H80000012
End Sub

Private Sub ComboBox_OperaçãoFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_OperaçãoFilial.BorderColor = &H8000000A
    ComboBox_OperaçãoFilial.BackColor = &H80000005
    Label_OperaçãoFilial.ForeColor = &H80000011
End Sub

Private Sub TextBox_N_SaidaFilial_Enter()

    TextBox_N_SaidaFilial.BorderColor = &H8000000D
    TextBox_N_SaidaFilial.BackColor = &H80000002
    Label_Quant.ForeColor = &H80000012
End Sub

Private Sub TextBox_N_SaidaFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

'Verifica se existe células com o mesmo valor
Dim vernumeronfe As Double
Dim planilha As Worksheet
Set planilha = ThisWorkbook.Worksheets("Saidas Filial")
vernumeronfe = WorksheetFunction.CountIf(planilha.Range("A:A"), TextBox_N_SaidaFilial.text)

If TextBox_N_SaidaFilial.text = "" Then
        Exit Sub
End If

If vernumeronfe > 0 Then
        MsgBox "Numero da NFe já cadastrado!", vbCritical, "Numero da NFe"
        LabelStatus = "Numero da NFe já cadastrado!"
        TextBox_N_SaidaFilial.BorderColor = &HFF&
        TextBox_N_SaidaFilial.BackColor = &H8080FF
        TextBox_N_SaidaFilial.ForeColor = &HC0&
        Exit Sub
End If

    TextBox_N_SaidaFilial.BorderColor = &H8000000A
    TextBox_N_SaidaFilial.BackColor = &H80000005
    Label_NFe_SaidaFilial.ForeColor = &H80000011
End Sub

Private Sub TextBox_N_EntradaFilial_Enter()

    TextBox_N_EntradaFilial.BorderColor = &H8000000D
    TextBox_N_EntradaFilial.BackColor = &H80000002
    Label_NFe_EntradaFilial.ForeColor = &H80000012
End Sub

Private Sub TextBox_N_EntradaFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_N_EntradaFilial.BorderColor = &H8000000A
    TextBox_N_EntradaFilial.BackColor = &H80000005
    Label_NFe_EntradaFilial.ForeColor = &H80000011
End Sub

Private Sub CODFilial_Enter()

    CODFilial.BorderColor = &H8000000D
    CODFilial.BackColor = &H80000002
    Label_CODFilial.ForeColor = &H80000012
End Sub

Private Sub CODFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    CODFilial.BorderColor = &H8000000A
    CODFilial.BackColor = &H80000005
    Label_CODFilial.ForeColor = &H80000011
End Sub

Private Sub ComboBox_VEÍCULOFilial_Enter()

    ComboBox_VEÍCULOFilial.BorderColor = &H8000000D
    ComboBox_VEÍCULOFilial.BackColor = &H80000002
    Label_VEÍCULOFilial.ForeColor = &H80000012
End Sub

Private Sub ComboBox_VEÍCULOFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_VEÍCULOFilial.BorderColor = &H8000000A
    ComboBox_VEÍCULOFilial.BackColor = &H80000005
    Label_VEÍCULOFilial.ForeColor = &H80000011
End Sub

Private Sub TextBox_PLACAFilial_Enter()

    TextBox_PLACAFilial.BorderColor = &H8000000D
    TextBox_PLACAFilial.BackColor = &H80000002
    Label_PLACAFilial.ForeColor = &H80000012
End Sub

Private Sub TextBox_PLACAFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_PLACAFilial.BorderColor = &H8000000A
    TextBox_PLACAFilial.BackColor = &H80000005
    Label_PLACAFilial.ForeColor = &H80000011
End Sub

Private Sub ComboBox_AnoFilial_Enter()

    ComboBox_AnoFilial.BorderColor = &H8000000D
    ComboBox_AnoFilial.BackColor = &H80000002
    Label_ANO_MODFilial.ForeColor = &H80000012
End Sub

Private Sub ComboBox_AnoFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_AnoFilial.BorderColor = &H8000000A
    ComboBox_AnoFilial.BackColor = &H80000005
    Label_ANO_MODFilial.ForeColor = &H80000011
End Sub

Private Sub ComboBox_CORFilial_Enter()

    ComboBox_CORFilial.BorderColor = &H8000000D
    ComboBox_CORFilial.BackColor = &H80000002
    Label_CORFilial.ForeColor = &H80000012
End Sub

Private Sub ComboBox_CORFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_CORFilial.BorderColor = &H8000000A
    ComboBox_CORFilial.BackColor = &H80000005
    Label_CORFilial.ForeColor = &H80000011
End Sub

Private Sub TextBox_Valor_EntradaFilial_Enter()

    TextBox_Valor_EntradaFilial.BorderColor = &H8000000D
    TextBox_Valor_EntradaFilial.BackColor = &H80000002
    Label_Valor_EntradaFilial.ForeColor = &H80000012
End Sub

Private Sub TextBox_Valor_EntradaFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_Valor_EntradaFilial.BorderColor = &H8000000A
    TextBox_Valor_EntradaFilial.BackColor = &H80000005
    Label_Valor_EntradaFilial.ForeColor = &H80000011
End Sub

Private Sub TextBox_Valor_SaidaFilial_Enter()

    TextBox_Valor_SaidaFilial.BorderColor = &H8000000D
    TextBox_Valor_SaidaFilial.BackColor = &H80000002
    Label_Valor_SaidaFilial.ForeColor = &H80000012
End Sub

Private Sub TextBox_Valor_SaidaFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_Valor_SaidaFilial.BorderColor = &H8000000A
    TextBox_Valor_SaidaFilial.BackColor = &H80000005
    Label_Valor_SaidaFilial.ForeColor = &H80000011
End Sub

Private Sub TextBox_Valor_LucroFilial_Enter()

    TextBox_Valor_LucroFilial.BorderColor = &H8000000D
    TextBox_Valor_LucroFilial.BackColor = &H80000002
    Label_LucroFilial.ForeColor = &H80000012
End Sub

Private Sub TextBox_Valor_LucroFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_Valor_LucroFilial.BorderColor = &H8000000A
    TextBox_Valor_LucroFilial.BackColor = &H80000005
    Label_LucroFilial.ForeColor = &H80000011
End Sub

Private Sub TextBox_DataFilial_Enter()

    TextBox_DataFilial.BorderColor = &H8000000D
    TextBox_DataFilial.BackColor = &H80000002
    Label_DataFilial.ForeColor = &H80000012
End Sub

Private Sub TextBox_DataFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_DataFilial.BorderColor = &H8000000A
    TextBox_DataFilial.BackColor = &H80000005
    Label_DataFilial.ForeColor = &H80000011
End Sub

Private Sub ComboBox_PorFilial_Enter()

    ComboBox_PorFilial.BorderColor = &H8000000D
    ComboBox_PorFilial.BackColor = &H80000002
    Label_AliqFilial.ForeColor = &H80000012
End Sub

Private Sub ComboBox_PorFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBox_PorFilial.BorderColor = &H8000000A
    ComboBox_PorFilial.BackColor = &H80000005
    Label_AliqFilial.ForeColor = &H80000011
End Sub

Private Sub TextBoxBaseICMSFilial_Enter()

    TextBoxBaseICMSFilial.BorderColor = &H8000000D
    TextBoxBaseICMSFilial.BackColor = &H80000002
    LabelBaseICMSFilial.ForeColor = &H80000012
End Sub

Private Sub TextBoxBaseICMSFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBoxBaseICMSFilial.BorderColor = &H8000000A
    TextBoxBaseICMSFilial.BackColor = &H80000005
    LabelBaseICMSFilial.ForeColor = &H80000011
End Sub

Private Sub TextBox_ICMSFilial_Enter()

    TextBox_ICMSFilial.BorderColor = &H8000000D
    TextBox_ICMSFilial.BackColor = &H80000002
    Label_ICMSFilial.ForeColor = &H80000012
End Sub

Private Sub TextBox_ICMSFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_ICMSFilial.BorderColor = &H8000000A
    TextBox_ICMSFilial.BackColor = &H80000005
    Label_ICMSFilial.ForeColor = &H80000011
End Sub

Private Sub TextBoxBasePISeCOFINSFilial_Enter()

    TextBoxBasePISeCOFINSFilial.BorderColor = &H8000000D
    TextBoxBasePISeCOFINSFilial.BackColor = &H80000002
    LabelBasePISeCOFINSFilial.ForeColor = &H80000012
End Sub

Private Sub TextBoxBasePISeCOFINSFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBoxBasePISeCOFINSFilial.BorderColor = &H8000000A
    TextBoxBasePISeCOFINSFilial.BackColor = &H80000005
    LabelBasePISeCOFINSFilial.ForeColor = &H80000011
End Sub

Private Sub TextBox_PISFilial_Enter()

    TextBox_PISFilial.BorderColor = &H8000000D
    TextBox_PISFilial.BackColor = &H80000002
    Label_PISFilial.ForeColor = &H80000012
End Sub

Private Sub TextBox_PISFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_PISFilial.BorderColor = &H8000000A
    TextBox_PISFilial.BackColor = &H80000005
    Label_PISFilial.ForeColor = &H80000011
End Sub

Private Sub TextBox_COFINSFilial_Enter()

    TextBox_COFINSFilial.BorderColor = &H8000000D
    TextBox_COFINSFilial.BackColor = &H80000002
    Label_COFINSFilial.ForeColor = &H80000012
End Sub

Private Sub TextBox_COFINSFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBox_COFINSFilial.BorderColor = &H8000000A
    TextBox_COFINSFilial.BackColor = &H80000005
    Label_COFINSFilial.ForeColor = &H80000011
End Sub
Private Sub ComboBoxCampos_Enter()

    ComboBoxCampos.BorderColor = &H8000000D
    ComboBoxCampos.BackColor = &H80000002
    Label_Filtro.ForeColor = &H80000012
End Sub

Private Sub ComboBoxCampos_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBoxCampos.BorderColor = &H8000000A
    ComboBoxCampos.BackColor = &H80000005
    Label_Filtro.ForeColor = &H80000011
End Sub

Private Sub ComboBoxCamposFilial_Enter()

    ComboBoxCamposFilial.BorderColor = &H8000000D
    ComboBoxCamposFilial.BackColor = &H80000002
    Label_FiltroFilial.ForeColor = &H80000012
End Sub

Private Sub ComboBoxCamposFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ComboBoxCamposFilial.BorderColor = &H8000000A
    ComboBoxCamposFilial.BackColor = &H80000005
    Label_FiltroFilial.ForeColor = &H80000011
End Sub

Private Sub TextBoxFiltroFilial_Enter()

    TextBoxFiltroFilial.BorderColor = &H8000000D
    TextBoxFiltroFilial.BackColor = &H80000002
    Label_PesquisaFilial.ForeColor = &H80000012
End Sub

Private Sub TextBoxFiltroFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBoxFiltroFilial.BorderColor = &H8000000A
    TextBoxFiltroFilial.BackColor = &H80000005
    Label_PesquisaFilial.ForeColor = &H80000011
End Sub

Private Sub TextBoxFiltro_Enter()

    TextBoxFiltro.BorderColor = &H8000000D
    TextBoxFiltro.BackColor = &H80000002
    Label_Pesquisa.ForeColor = &H80000012
End Sub

Private Sub TextBoxFiltro_Exit(ByVal cancel As MSForms.ReturnBoolean)

    TextBoxFiltro.BorderColor = &H8000000A
    TextBoxFiltro.BackColor = &H80000005
    Label_Pesquisa.ForeColor = &H80000011
End Sub

Private Sub ListBoxLista_Enter()

    ListBoxLista.BorderColor = &H8000000D
End Sub

Private Sub ListBoxLista_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ListBoxLista.BorderColor = &H8000000A
End Sub

Private Sub ListBoxListaFilial_Enter()

    ListBoxListaFilial.BorderColor = &H8000000D
End Sub

Private Sub ListBoxListaFilial_Exit(ByVal cancel As MSForms.ReturnBoolean)

    ListBoxListaFilial.BorderColor = &H8000000A
End Sub

'Inicio da pesquisa saidas na matriz
Private Sub TextBoxFiltro_Change()
    If Me.ComboBoxCampos.ListIndex <> -1 Then
        Call PreencheLista(TextBoxFiltro.text)
    End If
End Sub

Private Sub PreencheCampos()
    Dim ws As Worksheet
    Dim coluna As Long
    Dim linha As Long
    Set ws = ThisWorkbook.Worksheets(NomePlanilha)
    coluna = 1
    linha = LinhaCabecalho

    With ws
        While .Cells(linha, coluna).value <> Empty
            Me.ComboBoxCampos.AddItem .Cells(linha, coluna)
            coluna = coluna + 1
        Wend
    End With
End Sub

Private Sub PreencheCabecalho(ByRef Lista())
    Dim ws As Worksheet
    Dim coluna As Long
    Dim linha As Long
    Set ws = ThisWorkbook.Worksheets(NomePlanilha)
    coluna = 1
    linha = LinhaCabecalho

    With ws
        While .Cells(linha, coluna).value <> Empty
            Lista(coluna - 1, 0) = .Cells(linha, coluna)
            coluna = coluna + 1
        Wend
    End With
End Sub

Private Sub PreencheLista(ByVal TextoDigitado As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim x As Long
    Dim indiceLista As Long
    Dim coluna As Long
    Dim TextoCelula As String
    Set ws = ThisWorkbook.Worksheets(NomePlanilha)
    Dim Lista()

    ReDim Lista(ws.UsedRange.Columns.count, 0)

    i = LinhaCabecalho + 1
    indiceLista = 1
    coluna = Me.ComboBoxCampos.ListIndex + 1
    Call PreencheCabecalho(Lista)

    ListBoxLista.Clear
    With ws
        While .Cells(i, coluna).value <> Empty
            TextoCelula = .Cells(i, coluna).value
            If UCase(Left(TextoCelula, Len(TextoDigitado))) = UCase(TextoDigitado) Then

                For x = 0 To ws.UsedRange.Columns.count - 1
                    ReDim Preserve Lista(ws.UsedRange.Columns.count, indiceLista)
                    Lista(x, indiceLista) = .Cells(i, x + 1)
                Next

                indiceLista = indiceLista + 1
            End If
            i = i + 1
        Wend
    End With
   
    Lista = Array2DTranspose(Lista)

    Me.ListBoxLista.List = Lista
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

Private Sub ListBoxLista_Change()
    
    If Me.ListBoxLista.ListCount = 0 Then
    
    Else
     
    TextBox_ESTOQUE.value = ListBoxLista.List(ListBoxLista.ListIndex, 3)
    ComboBox_Operação.value = ListBoxLista.List(ListBoxLista.ListIndex, 2)
    TextBox_N_Saida.value = ListBoxLista.List(ListBoxLista.ListIndex, 0)
    TextBox_N_Entrada.value = ListBoxLista.List(ListBoxLista.ListIndex, 1)
    ComboBox_VEÍCULO.value = ListBoxLista.List(ListBoxLista.ListIndex, 4)
    TextBox_PLACA.value = ListBoxLista.List(ListBoxLista.ListIndex, 5)
    ComboBox_Ano.value = ListBoxLista.List(ListBoxLista.ListIndex, 6)
    ComboBox_COR.value = ListBoxLista.List(ListBoxLista.ListIndex, 7)
    TextBox_Valor_Entrada.value = ListBoxLista.List(ListBoxLista.ListIndex, 8)
    TextBox_Valor_Entrada.value = Format(TextBox_Valor_Entrada.value, "Currency")
    TextBox_Valor_Saida.value = ListBoxLista.List(ListBoxLista.ListIndex, 9)
    TextBox_Valor_Saida.value = Format(TextBox_Valor_Saida.value, "Currency")
    ComboBox_Por.value = ListBoxLista.List(ListBoxLista.ListIndex, 12)
    ComboBox_Por.value = (ComboBox_Por.value * 100)
    TextBox_Data.value = ListBoxLista.List(ListBoxLista.ListIndex, 11)
    
        'Calculo do lucro
        Dim valor1 As Double
            Dim valor2 As Double
                valor1 = 0
                valor2 = 0

            If Me.TextBox_Valor_Saida = Empty Then
                Else
            valor1 = Me.TextBox_Valor_Saida.value
                End If
            If Me.TextBox_Valor_Entrada = Empty Then
                Else
            valor2 = Me.TextBox_Valor_Entrada.value
                End If

            Me.TextBox_Valor_Lucro = valor1 - valor2
    
            TextBox_Valor_Lucro.value = Format(TextBox_Valor_Lucro.value, "Currency")
        
        'Calculo do Imposto
        TextBoxBaseICMS = (TextBox_Valor_Saida / 100) * (5)
            TextBoxBaseICMS.value = Format(TextBoxBaseICMS.value, "Currency")

        TextBox_ICMS = (TextBox_Valor_Saida / 100) * (5) / 100 * (ComboBox_Por)
            TextBox_ICMS.value = Format(TextBox_ICMS.value, "Currency")
    
        TextBox_PIS = (TextBox_Valor_Lucro / 100) * (0.65)
            TextBox_PIS.value = Format(TextBox_PIS.value, "Currency")
        
        TextBox_COFINS = (TextBox_Valor_Lucro / 100) * (3)
            TextBox_COFINS.value = Format(TextBox_COFINS.value, "Currency")
    
    End If
End Sub
'Fim da pesquisa saidas na matriz

'Inicio da pesquisa saidas na filial
Private Sub TextBoxFiltroFilial_Change()
    If Me.ComboBoxCamposFilial.ListIndex <> -1 Then
        Call PreencheLista2(TextBoxFiltroFilial.text)
    End If
End Sub

Private Sub PreencheCampos2()
    Dim ws As Worksheet
    Dim coluna As Long
    Dim linha As Long
    Set ws = ThisWorkbook.Worksheets(NomePlanilha2)
    coluna = 1
    linha = LinhaCabecalho

    With ws
        While .Cells(linha, coluna).value <> Empty
            Me.ComboBoxCamposFilial.AddItem .Cells(linha, coluna)
            coluna = coluna + 1
        Wend
    End With
End Sub

Private Sub PreencheCabecalho2(ByRef Lista())
    Dim ws As Worksheet
    Dim coluna As Long
    Dim linha As Long
    Set ws = ThisWorkbook.Worksheets(NomePlanilha2)
    coluna = 1
    linha = LinhaCabecalho

    With ws
        While .Cells(linha, coluna).value <> Empty
            Lista(coluna - 1, 0) = .Cells(linha, coluna)
            coluna = coluna + 1
        Wend
    End With
End Sub

Private Sub PreencheLista2(ByVal TextoDigitado As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim x As Long
    Dim indiceLista As Long
    Dim coluna As Long
    Dim TextoCelula As String
    Set ws = ThisWorkbook.Worksheets(NomePlanilha2)
    Dim Lista()

    ReDim Lista(ws.UsedRange.Columns.count, 0)

    i = LinhaCabecalho + 1
    indiceLista = 1
    coluna = Me.ComboBoxCamposFilial.ListIndex + 1
    Call PreencheCabecalho2(Lista)

    ListBoxListaFilial.Clear
    With ws
        While .Cells(i, coluna).value <> Empty
            TextoCelula = .Cells(i, coluna).value
            If UCase(Left(TextoCelula, Len(TextoDigitado))) = UCase(TextoDigitado) Then

                For x = 0 To ws.UsedRange.Columns.count - 1
                    ReDim Preserve Lista(ws.UsedRange.Columns.count, indiceLista)
                    Lista(x, indiceLista) = .Cells(i, x + 1)
                Next

                indiceLista = indiceLista + 1
            End If
            i = i + 1
        Wend
    End With
   
    Lista = Array2DTranspose(Lista)

    Me.ListBoxListaFilial.List = Lista
End Sub

Private Sub ListBoxListaFilial_Change()
    
    If Me.ListBoxListaFilial.ListCount = 0 Then
    
    Else
     
    TextBox_ESTOQUEFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 3)
    ComboBox_OperaçãoFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 2)
    TextBox_N_SaidaFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 0)
    TextBox_N_EntradaFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 1)
    ComboBox_VEÍCULOFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 4)
    TextBox_PLACAFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 5)
    ComboBox_AnoFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 6)
    ComboBox_CORFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 7)
    TextBox_Valor_EntradaFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 8)
    TextBox_Valor_EntradaFilial.value = Format(TextBox_Valor_EntradaFilial.value, "Currency")
    TextBox_Valor_SaidaFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 9)
    TextBox_Valor_SaidaFilial.value = Format(TextBox_Valor_SaidaFilial.value, "Currency")
    ComboBox_PorFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 12)
    ComboBox_PorFilial.value = (ComboBox_PorFilial.value * 100)
    TextBox_DataFilial.value = ListBoxListaFilial.List(ListBoxListaFilial.ListIndex, 11)
    
        'Calculo do lucro
        
        Dim valor1 As Double
            Dim valor2 As Double
                valor1 = 0
                valor2 = 0

            If Me.TextBox_Valor_SaidaFilial = Empty Then
                Else
            valor1 = Me.TextBox_Valor_SaidaFilial.value
                End If
            If Me.TextBox_Valor_EntradaFilial = Empty Then
                Else
            valor2 = Me.TextBox_Valor_EntradaFilial.value
                End If

            Me.TextBox_Valor_LucroFilial = valor1 - valor2
    
            TextBox_Valor_LucroFilial.value = Format(TextBox_Valor_LucroFilial.value, "Currency")

    TextBoxBaseICMSFilial = (TextBox_Valor_SaidaFilial / 100) * (5)
        TextBoxBaseICMSFilial.text = Format(TextBoxBaseICMSFilial.text, "Currency")

    TextBox_ICMSFilial = (TextBox_Valor_SaidaFilial / 100) * (5) / 100 * (ComboBox_PorFilial)
        TextBox_ICMSFilial.text = Format(TextBox_ICMSFilial.text, "Currency")
    
    TextBox_PISFilial = (TextBox_Valor_LucroFilial / 100) * (0.65)
        TextBox_PISFilial.text = Format(TextBox_PISFilial.text, "Currency")
        
    TextBox_COFINSFilial = (TextBox_Valor_LucroFilial / 100) * (3)
        TextBox_COFINSFilial.text = Format(TextBox_COFINSFilial.text, "Currency")
        
    End If
End Sub
'Fim da pesquisa saidas na filial

Private Function SóNúmeros(l As IReturnInteger)
    Select Case l
        Case Asc("0") To Asc("9")
            SóNúmeros = l
        Case Else
            SóNúmeros = 0
            MsgBox "Favor inserir apenas números!", vbExclamation, "CAMPO TIPO NÚMERO"
    End Select
End Function

Private Sub TextBox_ESTOQUE_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub TextBox_N_Entrada_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub TextBox_N_Saida_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub COD_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub TextBox_Valor_Entrada_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub TextBox_Valor_Saida_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub TextBox_ESTOQUEFilial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub TextBox_N_SaidaFilial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub TextBox_N_EntradaFilial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub CODFilial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub TextBox_Valor_EntradaFilial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub

Private Sub TextBox_Valor_SaidaFilial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    KeyAscii = SóNúmeros(KeyAscii)
End Sub



