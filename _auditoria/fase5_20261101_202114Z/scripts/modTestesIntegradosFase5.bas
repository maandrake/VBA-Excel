Attribute VB_Name = "modTestesIntegradosFase5"
Option Explicit

' Exclusivamente para fixtures ficticias e descartaveis da Fase 5.
Private Const CHAVE_F5 As String = "01234567890123456789012345678901234567890123"
Private mCenarioF5 As String

Private Sub F5_Guardar()
    If InStr(1, ThisWorkbook.Name, "Fixture_Auditoria_Fase5_", vbTextCompare) <> 1 Then Err.Raise vbObjectError + 2501, , "Fixture Fase 5 obrigatoria."
    If InStr(1, ThisWorkbook.Path, "VBA_Fase5_20261101_202114Z", vbTextCompare) = 0 Then Err.Raise vbObjectError + 2502, , "Pasta isolada obrigatoria."
    If InStr(1, ThisWorkbook.Path, "OneDrive", vbTextCompare) > 0 Then Err.Raise vbObjectError + 2503, , "Proibido testar em OneDrive."
    If ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then Err.Raise vbObjectError + 2504, , "Marcador ficticio ausente."
End Sub

Public Sub F5_Preparar(ByVal cenario As String)
    Dim ws As Worksheet, nomes As Variant, nome As Variant
    Dim r As Long, j As Long, i As Long, filial As Boolean, op As String
    Dim ev As Boolean, su As Boolean
    F5_Guardar
    If cenario Like "*[!A-Za-z0-9_]*" Or Len(cenario) = 0 Then Err.Raise 5
    If InStr(1, ThisWorkbook.Name, cenario, vbTextCompare) = 0 Then Err.Raise 5, , "Nome da copia difere do cenario."
    If Dir(ThisWorkbook.Path & "\" & cenario & "_antes.tsv") <> "" Then Err.Raise 5, , "Cenario ja preparado; use copia nova."
    mCenarioF5 = cenario
    ev = Application.EnableEvents: su = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    nomes = Array("Entradas", "Entradas Anual", "Entradas Filial", "Entradas Anual Filial", "Saidas", "Todas Saidas", "Saidas Filial", "Todas Saidas Filial", "Demonstração")
    For Each nome In nomes
        Set ws = ThisWorkbook.Worksheets(CStr(nome))
        ws.Range("A1:U40").Clear
        For j = 1 To 21: ws.Cells(1, j).Value2 = "F5_" & CStr(j): Next j
        If InStr(1, CStr(nome), "Saidas", vbTextCompare) > 0 Then ws.Range("A1").Value2 = "Nº Nfe Saida"
    Next nome
    For i = 0 To 2
        r = i + 4
        For Each nome In Array("Entradas", "Entradas Anual", "Entradas Filial", "Entradas Anual Filial")
            Set ws = ThisWorkbook.Worksheets(CStr(nome))
            filial = (InStr(1, CStr(nome), "Filial", vbTextCompare) > 0)
            ws.Cells(r, 1).Value2 = IIf(filial, 21001, 11001) + i
            ws.Cells(r, 2).Value2 = "Compra"
            ws.Cells(r, 3).Value2 = "F5_ESTOQUE_" & CStr(i)
            ws.Cells(r, 4).Value2 = "FICTICIO " & Chr(65 + i)
            ws.Cells(r, 5).Value2 = "F5" & Chr(65 + i)
            ws.Cells(r, 6).Value2 = "FIC5" & Chr(65 + i) & "01"
            ws.Cells(r, 7).Value2 = "2026"
            ws.Cells(r, 8).Value2 = "BRANCO"
            ws.Cells(r, 9).Value2 = 900 + i * 100
            ws.Cells(r, 10).Value2 = CDbl(DateSerial(2026, 11, 1))
            ws.Cells(r, 10).NumberFormat = "dd/mm/yyyy"
        Next nome
        Set ws = ThisWorkbook.Worksheets("Entradas")
        ws.Cells(r, 11).Value2 = "TESTE"
        ws.Cells(r, 12).NumberFormat = "@": ws.Cells(r, 12).Value2 = CHAVE_F5
        ws.Cells(r, 13).Formula = "=I" & r & "*2"
        ws.Cells(r, 14).Formula = "=A" & r
        ws.Cells(r, 15).Value2 = "MATRIZ_" & Chr(65 + i)
        Set ws = ThisWorkbook.Worksheets("Entradas Filial")
        ws.Cells(r, 11).Value2 = "TESTE"
        ws.Cells(r, 12).NumberFormat = "@": ws.Cells(r, 12).Value2 = CHAVE_F5
        ws.Cells(r, 13).Formula = "=I" & r & "*2"
        ws.Cells(r, 14).Formula = "=A" & r
        ws.Cells(r, 15).Value2 = "FILIAL_" & Chr(65 + i)
    Next i
    Set ws = ThisWorkbook.Worksheets("Demonstração")
    ws.Range("Q4").Formula = "='Entradas'!A4"
    ws.Range("R4").Formula = "='Entradas'!F4"
    ws.Range("S4").Formula = "=INDIRECT(""'Entradas'!F4"")"
    ws.Range("Q5").Formula = "='Entradas Filial'!A4"
    ws.Range("R5").Formula = "='Entradas Filial'!F4"
    ws.Range("S5").Formula = "=INDIRECT(""'Entradas Filial'!F4"")"
    With ThisWorkbook.Worksheets("Entradas Filial")
        .Range("M7").Formula = "=I7*2": .Range("N7").Formula = "=A7": .Range("O7").Value2 = "DESTINO_SENTINELA"
    End With
    If InStr(cenario, "DUP") > 0 Then
        If Left$(cenario, 1) = "D" Then
            ThisWorkbook.Worksheets("Saidas Filial").Range("A2").Value2 = 22001
        ElseIf cenario = "T_DUP_DESTINO" Then
            ThisWorkbook.Worksheets("Entradas Filial").Range("A4").Value2 = 22001
        Else
            ThisWorkbook.Worksheets("Saidas").Range("A2").Value2 = 22001
        End If
    End If
    Application.CalculateFull
    Application.EnableEvents = ev: Application.ScreenUpdating = su
    F5_Preencher cenario
    ThisWorkbook.Save
    F5_Snapshot "antes"
    Debug.Print "F5 PREPARADO: " & cenario
End Sub

Private Sub F5_Preencher(ByVal cenario As String)
    Dim f As Object, sufixo As String, op As String
    Load Cad_NFE_Saida
    Set f = Cad_NFE_Saida
    If Left$(cenario, 1) = "D" Then sufixo = "Filial": op = "Devolução" Else sufixo = "": op = "Venda"
    If Left$(cenario, 1) = "T" Then op = "Transferência"
    If Left$(cenario, 1) = "C" Then op = "Cancelado"
    f.MultiPage1.Value = IIf(sufixo = "", 0, 1)
    f.Controls("ComboBox_Operação" & sufixo).Value = op
    f.Controls("TextBox_N_Saida" & sufixo).Value = "22001"
    f.Controls("TextBox_N_Entrada" & sufixo).Value = IIf(sufixo = "", "11001", "21001")
    f.Controls("TextBox_ESTOQUE" & sufixo).Value = "F5_ESTOQUE_0"
    f.Controls("ComboBox_VEÍCULO" & sufixo).Value = "FICTICIO A"
    f.Controls("COD" & sufixo).Value = "F5A"
    f.Controls("TextBox_PLACA" & sufixo).Value = "FIC5A01"
    f.Controls("ComboBox_Ano" & sufixo).Value = "2026"
    f.Controls("ComboBox_COR" & sufixo).Value = "BRANCO"
    f.Controls("TextBox_Valor_Entrada" & sufixo).Value = "900"
    f.Controls("TextBox_Valor_Saida" & sufixo).Value = "1000"
    f.Controls("TextBox_Data" & sufixo).Value = "01/11/2026"
    f.Controls("ComboBox_OutroEnvio" & sufixo).Value = "TESTE"
    f.Controls("ComboBox_Por" & sufixo).Value = "12"
    f.Controls("TextBox_ICMS" & sufixo).Value = "0"
    f.Controls("TextBox_PIS" & sufixo).Value = "0"
    f.Controls("TextBox_COFINS" & sufixo).Value = "0"
    If sufixo = "" Then
        f.TextBoxChaveDeAcesso.Value = CHAVE_F5
        f.TextBoxBaseICMS.Value = "0"
        f.TextBoxBasePISeCOFINS.Value = "0"
    End If
    If InStr(cenario, "VAZIO") > 0 Then f.Controls("TextBox_N_Saida" & sufixo).Value = ""
    If InStr(cenario, "VALOR") > 0 Then f.Controls("TextBox_Valor_Entrada" & sufixo).Value = "ABC"
    If InStr(cenario, "DATA") > 0 Then f.Controls("TextBox_Data" & sufixo).Value = "31/02/2026"
    If InStr(cenario, "SEMORIGEM") > 0 Then f.Controls("TextBox_N_Entrada" & sufixo).Value = "99999"
End Sub

Public Sub F5_Mostrar()
    F5_Guardar
    Cad_NFE_Saida.Show vbModeless
End Sub

Public Sub F5_Concluir()
    Dim arq As Object, fso As Object
    F5_Guardar
    If mCenarioF5 = "" Then Err.Raise 5, , "Contexto perdido; nao concluir sem identificar cenario."
    Application.CalculateFull
    F5_Snapshot "depois"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arq = fso.CreateTextFile(ThisWorkbook.Path & "\" & mCenarioF5 & "_resultado.txt", False, True)
    arq.WriteLine "Cenario=" & mCenarioF5
    arq.WriteLine "Workbook=" & ThisWorkbook.FullName
    arq.WriteLine "Status=" & Cad_NFE_Saida.LabelStatus.Caption
    arq.WriteLine "EnableEvents=" & Application.EnableEvents
    arq.WriteLine "ScreenUpdating=" & Application.ScreenUpdating
    arq.WriteLine "DisplayAlerts=" & Application.DisplayAlerts
    arq.WriteLine "Calculation=" & Application.Calculation
    arq.Close
    ThisWorkbook.Save
    Debug.Print "F5 EVIDENCIA CONCLUIDA: " & mCenarioF5
End Sub

Private Function F5_Texto(ByVal v As Variant) As String
    If IsError(v) Then F5_Texto = "#ERRO" Else F5_Texto = Replace(Replace(Replace(CStr(v), vbTab, "{TAB}"), vbCr, "{CR}"), vbLf, "{LF}")
End Function

Private Sub F5_Snapshot(ByVal etapa As String)
    Dim arq As Object, fso As Object, nome As Variant, ws As Worksheet, cel As Range
    Dim ignorar As Variant, conteudo As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arq = fso.CreateTextFile(ThisWorkbook.Path & "\" & mCenarioF5 & "_" & etapa & ".tsv", False, True)
    arq.WriteLine "Planilha" & vbTab & "Celula" & vbTab & "Tipo" & vbTab & "Valor" & vbTab & "Formula" & vbTab & "NumberFormat" & vbTab & "Ignore" & vbTab & "Locked" & vbTab & "Protegida"
    For Each nome In Array("Entradas", "Entradas Anual", "Entradas Filial", "Entradas Anual Filial", "Saidas", "Todas Saidas", "Saidas Filial", "Todas Saidas Filial", "Demonstração")
        Set ws = ThisWorkbook.Worksheets(CStr(nome))
        For Each cel In ws.Range("A1:U40").Cells
            ignorar = "N/A"
            On Error Resume Next
            ignorar = cel.Errors(xlNumberAsText).Ignore
            On Error GoTo 0
            If IsError(cel.Value2) Then conteudo = cel.Text Else conteudo = F5_Texto(cel.Value2)
            arq.WriteLine ws.Name & vbTab & cel.Address(False, False) & vbTab & TypeName(cel.Value2) & vbTab & conteudo & vbTab & F5_Texto(cel.Formula) & vbTab & F5_Texto(cel.NumberFormat) & vbTab & F5_Texto(ignorar) & vbTab & cel.Locked & vbTab & ws.ProtectContents
        Next cel
    Next nome
    arq.Close
End Sub

Public Sub F5_FalhaEntreGravacoes()
    F5_Guardar
    If Left$(mCenarioF5, 1) = "D" Then Cad_NFE_Saida.TextBox_Valor_EntradaFilial.Value = "ABC" Else Cad_NFE_Saida.TextBox_Valor_Entrada.Value = "ABC"
    Debug.Print "INJECAO TESTE: conversao invalida entre gravacoes"
End Sub

Public Sub F5_ProtegerOrigem()
    F5_Guardar
    If Left$(mCenarioF5, 1) = "D" Then ThisWorkbook.Worksheets("Entradas Filial").Protect "F5QA" Else ThisWorkbook.Worksheets("Entradas").Protect "F5QA"
    Debug.Print "INJECAO TESTE: origem protegida antes de Delete"
End Sub

Public Sub F5_XML(ByVal caso As String)
    Dim dom As Object, resultado As Boolean, arq As Object, fso As Object, caminho As String
    F5_Guardar
    F5_Preencher "V_VALIDO"
    caminho = ThisWorkbook.Path & "\xml\" & caso & ".xml"
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False: dom.validateOnParse = False: dom.resolveExternals = False
    dom.setProperty "ProhibitDTD", True
    dom.Load caminho
    resultado = ValidarXMLNFe(Cad_NFE_Saida, dom, caminho)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arq = fso.CreateTextFile(ThisWorkbook.Path & "\XML_" & caso & "_resultado.txt", False, True)
    arq.WriteLine "Tipo=Execucao funcional VBA: modValidacaoXML.ValidarXMLNFe"
    arq.WriteLine "Arquivo=" & caminho
    arq.WriteLine "Retorno=" & resultado
    arq.WriteLine "ParseError=" & dom.parseError.errorCode
    arq.Close
    Debug.Print "F5 XML " & caso & " = " & resultado
End Sub
