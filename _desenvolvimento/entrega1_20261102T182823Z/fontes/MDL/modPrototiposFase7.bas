Attribute VB_Name = "modPrototiposFase7"
Option Explicit
' Ambiente sintetico; nao alterar gravadores, baixa ou fiscalidade do produto.
Private Const MARCA As String = "VBA_Fase7_20261102_025115Z"
Private gLog7 As Object
Private Function FS7() As Object
    Set FS7 = CreateObject("Scripting.FileSystemObject")
End Function
Private Sub Guard7()
    Dim raiz As String, atual As String
    If ThisWorkbook.name <> "Fixture_Arquitetura_Fase7_Mestre.xlsm" Then Err.Raise 5, , "Mestre F7 obrigatorio"
    If InStr(1, ThisWorkbook.Path, MARCA, vbTextCompare) = 0 Then Err.Raise 5, , "Pasta F7 obrigatoria"
    If InStr(1, ThisWorkbook.Path, "OneDrive", vbTextCompare) > 0 Then Err.Raise 5, , "Proibido OneDrive"
    raiz = FS7.GetFolder(Environ$("TEMP")).ShortPath & "\"
    atual = FS7.GetFolder(ThisWorkbook.Path).ShortPath & "\"
    If StrComp(Left$(atual, Len(raiz)), raiz, vbTextCompare) <> 0 Then Err.Raise 5, , "Somente TEMP"
    If ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then Err.Raise 5, , "Marcador ausente"
End Sub
Private Function Txt7(ByVal v As Variant) As String
    If IsError(v) Then
        Txt7 = CStr(v)
    ElseIf IsEmpty(v) Then
        Txt7 = "<EMPTY>"
    Else
        Txt7 = Replace(Replace(Replace(CStr(v), vbTab, "<TAB>"), vbCr, "<CR>"), vbLf, "<LF>")
    End If
End Function
Private Function Ig7(ByVal c As Range) As String
    On Error GoTo Falhou
    Ig7 = CStr(c.Errors(xlNumberAsText).Ignore): Exit Function
Falhou:
    Ig7 = "NAO_LIDO:" & CStr(Err.Number)
End Function
Private Function Selo7(ByVal sh As Worksheet, ByVal r As Long) As String
    Dim c As Range, v As String, s As String
    For Each c In sh.Cells(r, 1).Resize(1, 12).Cells
        v = TypeName(c.Value2) & ":" & Txt7(c.Value2) & ":" & Txt7(c.Formula) & ":" & c.NumberFormat & ":" & Ig7(c)
        s = s & Len(v) & ":" & v
    Next c
    Selo7 = s
End Function
Private Function Last7(ByVal sh As Worksheet, Optional ByVal col As Long = 1) As Long
    Last7 = sh.Cells(sh.rows.count, col).End(xlUp).row
End Function
Private Function Unique7(ByVal sh As Worksheet, ByVal col As Long, ByVal id As String) As Long
    Dim i As Long, n As Long, r As Long
    If id = "" Then Exit Function
    For i = 2 To Last7(sh, col)
        If CStr(sh.Cells(i, col).Value2) = id Then n = n + 1: r = i
    Next i
    If n = 1 Then Unique7 = r
End Function
Private Function Resolve7(ByVal wb As Workbook, ByVal pid As String, ByVal ep As String, ByVal vid As String, ByRef sh As Worksheet) As Long
    Dim links As Worksheet, l As Long, r As Long
    On Error GoTo Invalido
    Set links = wb.Worksheets("_F6_Vinculos")
    l = Unique7(links, 1, pid): If l = 0 Then Exit Function
    If CStr(links.Cells(l, 2).Value2) <> ep Or CStr(links.Cells(l, 3).Value2) <> vid Then Exit Function
    Set sh = wb.Worksheets(CStr(links.Cells(l, 6).Value2))
    r = Unique7(sh, 24, pid): If r = 0 Then Exit Function
    If CStr(sh.Cells(r, 21).Value2) <> ep Or CStr(sh.Cells(r, 22).Value2) <> vid Then Exit Function
    If CStr(links.Cells(l, 9).Value2) <> Selo7(sh, r) Then Exit Function
    Resolve7 = r
Invalido:
End Function
Private Function Pending7(ByVal wb As Workbook, ByVal ep As String) As Boolean
    Dim sh As Worksheet, i As Long, estado As String
    Set sh = wb.Worksheets("_F6_Operacoes")
    For i = 2 To Last7(sh)
        If CStr(sh.Cells(i, 3).Value2) = ep Then
            estado = CStr(sh.Cells(i, 5).Value2)
            If estado <> "CONCLUIDA" And estado <> "COMPENSADA" Then Pending7 = True: Exit Function
        End If
    Next i
End Function
Private Function Eligible7(ByVal wb As Workbook, ByVal e As Long, ByRef sh As Worksheet) As Long
    Dim eps As Worksheet, vid As String, ep As String, pid As String, i As Long, n As Long, r As Long
    Set eps = wb.Worksheets("_F6_Episodios")
    ep = CStr(eps.Cells(e, 1).Value2): vid = CStr(eps.Cells(e, 2).Value2)
    If Unique7(eps, 1, ep) <> e Then Err.Raise 5, , "EpisodeID ambiguo"
    If CStr(eps.Cells(e, 3).Value2) <> "ATIVO" Then Exit Function
    If Pending7(wb, ep) Then Err.Raise 5, , "Episodio pendente"
    For i = 2 To Last7(eps)
        If CStr(eps.Cells(i, 2).Value2) = vid And CStr(eps.Cells(i, 3).Value2) = "ATIVO" Then n = n + 1
    Next i
    If n <> 1 Then Err.Raise 5, , "Veiculo com custodia ambigua"
    pid = CStr(eps.Cells(e, 5).Value2): r = Resolve7(wb, pid, ep, vid, sh)
    If r = 0 Then Err.Raise 5, , "Projecao ausente ou divergente"
    If sh.name <> IIf(CStr(eps.Cells(e, 4).Value2) = "MATRIZ", "Entradas", "Entradas Filial") Then Err.Raise 5, , "Custodia divergente"
    Eligible7 = r
End Function
Private Function Search7(ByVal wb As Workbook, ByVal query As String, ByVal lista As AuditoriaListaFake) As Long
    Dim eps As Worksheet, sh As Worksheet, e As Long, r As Long, n As Long, a() As Variant, i As Long
    Set eps = wb.Worksheets("_F6_Episodios")
    lista.Clear: lista.ColumnCount = 7
    ReDim a(0 To Last7(eps), 0 To 6)
    For e = 2 To Last7(eps)
        r = Eligible7(wb, e, sh)
        If r > 0 Then
            If query = "" Or InStr(1, CStr(sh.Cells(r, 6).Value2) & ";" & CStr(sh.Cells(r, 1).Value2), query, vbTextCompare) > 0 Then
                a(n, 0) = sh.name: a(n, 1) = sh.Cells(r, 1).Value2: a(n, 2) = sh.Cells(r, 6).Value2
                a(n, 3) = eps.Cells(e, 5).Value2: a(n, 4) = eps.Cells(e, 1).Value2
                a(n, 5) = eps.Cells(e, 2).Value2: a(n, 6) = eps.Cells(e, 6).Value2: n = n + 1
            End If
        End If
    Next e
    If n > 0 Then
        Dim final() As Variant
        ReDim final(0 To n - 1, 0 To 6)
        For e = 0 To n - 1: For i = 0 To 6: final(e, i) = a(e, i): Next i: Next e
        lista.list = final
    End If
    Search7 = n
End Function
Private Function Item7(ByVal lista As AuditoriaListaFake, ByVal r As Long, ByVal c As Long) As Variant
    Dim a As Variant
    a = lista.list: Item7 = a(r, c)
End Function
Private Function Navigate7(ByVal wb As Workbook, ByVal lista As AuditoriaListaFake, ByRef sh As Worksheet) As Long
    Dim i As Long, e As Long, r As Long, eps As Worksheet, token As Variant
    i = lista.ListIndex: If i < 0 Or i >= lista.ListCount Then Exit Function
    token = lista.list: Set eps = wb.Worksheets("_F6_Episodios")
    e = Unique7(eps, 1, CStr(token(i, 4))): If e = 0 Then Exit Function
    If eps.Cells(e, 6).Value2 <> token(i, 6) Or eps.Cells(e, 5).Value2 <> token(i, 3) Then Exit Function
    r = Eligible7(wb, e, sh)
    Navigate7 = r
End Function
Public Function F7_Campo(ByVal pid As String, ByVal ep As String, ByVal vid As String, ByVal dados As Range, ByVal ids As Range, ByVal campo As Long) As Variant
    Dim i As Long, n As Long, r As Long
    On Error GoTo Invalido
    If pid = "" Or ep = "" Or vid = "" Or campo < 1 Or campo > dados.Columns.count Then GoTo Invalido
    If dados.rows.count <> ids.rows.count Or ids.Columns.count <> 4 Then GoTo Invalido
    If Not (dados.Parent Is ids.Parent) Then GoTo Invalido
    For i = 1 To ids.rows.count
        If CStr(ids.Cells(i, 4).Value2) = pid Then n = n + 1: r = i
    Next i
    If n = 0 Then F7_Campo = CVErr(xlErrNA): Exit Function
    If n <> 1 Then GoTo Invalido
    If CStr(ids.Cells(r, 1).Value2) <> ep Or CStr(ids.Cells(r, 2).Value2) <> vid Then GoTo Invalido
    F7_Campo = dados.Cells(r, campo).Value2: Exit Function
Invalido:
    F7_Campo = CVErr(xlErrValue)
End Function
Public Function F7_Vendido(ByVal ep As String, ByVal operacoes As Range) As Variant
    Dim i As Long, n As Long, ids As Object, op As String
    On Error GoTo Invalido
    If ep = "" Or operacoes.Columns.count < 9 Then GoTo Invalido
    Set ids = CreateObject("Scripting.Dictionary")
    For i = 1 To operacoes.rows.count
        op = CStr(operacoes.Cells(i, 1).Value2)
        If op <> "" Then
            If ids.Exists(op) Then GoTo Invalido
            ids.Add op, True
            If CStr(operacoes.Cells(i, 3).Value2) = ep Then
                If CStr(operacoes.Cells(i, 5).Value2) <> "CONCLUIDA" And CStr(operacoes.Cells(i, 5).Value2) <> "COMPENSADA" Then GoTo Invalido
                If CStr(operacoes.Cells(i, 9).Value2) = "Venda" And CStr(operacoes.Cells(i, 5).Value2) = "CONCLUIDA" Then n = n + 1
            End If
        End If
    Next i
    If n > 1 Then GoTo Invalido
    F7_Vendido = IIf(n = 1, "SIM", "NAO"): Exit Function
Invalido:
    F7_Vendido = CVErr(xlErrValue)
End Function
Private Sub Check7(ByVal id As String, ByVal ok As Boolean, ByVal esperado As String, ByVal Observado As String)
    gLog7.WriteLine id & vbTab & IIf(ok, "PASS", "FAIL") & vbTab & Txt7(esperado) & vbTab & Txt7(Observado)
End Sub
Private Sub Snapshot7(ByVal wb As Workbook, ByVal nome As String)
    Dim arq As Object, sh As Worksheet, c As Range
    Set arq = FS7.CreateTextFile(ThisWorkbook.Path & "\" & nome & ".tsv", False, True)
    arq.WriteLine "Planilha" & vbTab & "Celula" & vbTab & "Tipo" & vbTab & "Valor" & vbTab & "Formula" & vbTab & "Formato" & vbTab & "Ignore"
    For Each sh In wb.Worksheets
        For Each c In sh.Range("A1:X40").Cells
            arq.WriteLine sh.name & vbTab & c.Address(False, False) & vbTab & TypeName(c.Value2) & vbTab & Txt7(c.Value2) & vbTab & Txt7(c.Formula) & vbTab & c.NumberFormat & vbTab & Ig7(c)
        Next c
    Next sh
    arq.Close
End Sub
Private Function OpenInput7(ByVal nome As String) As Workbook
    Dim origem As String, destino As String
    origem = ThisWorkbook.Path & "\insumos\" & nome & ".xlsm"
    destino = ThisWorkbook.Path & "\exec_" & nome & ".xlsm"
    FS7.CopyFile origem, destino, False
    Set OpenInput7 = Application.Workbooks.Open(destino, UpdateLinks:=0): Call RestaurarFormatoSintetico7(OpenInput7)
    
End Function
Private Sub CopyCell7(ByVal a As Range, ByVal b As Range)
    b.NumberFormat = a.NumberFormat
    If a.HasFormula Then b.Formula = a.Formula Else b.Value2 = a.Value2
    b.Errors(xlNumberAsText).Ignore = a.Errors(xlNumberAsText).Ignore
End Sub
Private Function Sig7(ByVal wb As Workbook) As String
    Dim nome As Variant, sh As Worksheet, c As Range, s As String
    For Each nome In Array("Entradas", "Entradas Filial", "Entradas Anual", "Entradas Anual Filial", "Saidas", "Todas Saidas", "Saidas Filial", "Todas Saidas Filial", "_F6_Episodios", "_F6_Vinculos", "_F6_Operacoes")
        Set sh = wb.Worksheets(CStr(nome))
        For Each c In sh.Range("A1:X20").Cells
            s = s & TypeName(c.Value2) & ":" & Txt7(c.Value2) & ":" & Txt7(c.Formula) & ":" & c.NumberFormat & ":" & Ig7(c) & vbLf
        Next c
    Next nome
    Sig7 = s
End Function
Public Sub F7_Leitores()
    Dim wb As Workbook, sh As Worksheet, eps As Worksheet, demo As Worksheet, links As Worksheet
    Dim lista As AuditoriaListaFake, n As Long, r As Long, j As Long, pid As String, ep As String, vid As String, selo As String, form As String, antes As String
    Dim eventos As Boolean, tela As Boolean, alertas As Boolean, calculo As Long
    Call Guard7
    eventos = Application.EnableEvents: tela = Application.ScreenUpdating: alertas = Application.DisplayAlerts: calculo = Application.Calculation
    On Error GoTo Falhou
    Set gLog7 = FS7.CreateTextFile(ThisWorkbook.Path & "\leitores_resultados.tsv", False, True)
    gLog7.WriteLine "Cenario" & vbTab & "Resultado" & vbTab & "Esperado" & vbTab & "Observado"
    Set wb = OpenInput7("leitor"): Snapshot7 wb, "leitor_antes"
    Set lista = New AuditoriaListaFake: Set eps = wb.Worksheets("_F6_Episodios")
    n = Search7(wb, "FIC6A01", lista)
    Check7 "R01_PESQUISA_ATIVA_REENTRADA", n = 1 And Item7(lista, 0, 4) <> eps.Cells(2, 1).Value2, "somente novo episodio", CStr(n) & ";" & CStr(Item7(lista, 0, 4))
    n = Search7(wb, "", lista)
    Check7 "R02_CONTAGEM_ATIVA", n = 6, "6 ativos, excluindo antigo A", CStr(n)
    With wb.Worksheets("Entradas Anual")
        ep = CStr(.Range("U8").Value2)
        .Range("Q8").Formula = "='" & ThisWorkbook.name & "'!F7_Vendido(""" & ep & """,'_F6_Operacoes'!A2:J100)"
        .Range("Q4").Formula = "='" & ThisWorkbook.name & "'!F7_Vendido(""" & CStr(.Range("U4").Value2) & """,'_F6_Operacoes'!A2:J100)"
        Application.Calculate
        Check7 "R03_ANUAL_POR_EPISODIO", .Range("Q8").Value2 = "NAO" And .Range("Q4").Value2 = "SIM", "novo NAO; antigo SIM", Txt7(.Range("Q8").Value2) & ";" & Txt7(.Range("Q4").Value2)
        Check7 "BASE_ANUAL_PLACA", .Range("P8").Value2 = "NAO", "NAO para nova entrada", Txt7(.Range("P8").Value2)
    End With
    Snapshot7 wb, "leitor_depois": wb.save: wb.Close False: Set wb = Nothing
    Set wb = OpenInput7("semente"): Snapshot7 wb, "semente_antes"
    Set sh = wb.Worksheets("Entradas"): Set eps = wb.Worksheets("_F6_Episodios"): Set demo = wb.Worksheets("Demonstração")
    pid = CStr(sh.Range("X4").Value2): ep = CStr(sh.Range("U4").Value2): vid = CStr(sh.Range("V4").Value2): selo = Selo7(sh, 4)
    n = Search7(wb, "FIC6A01", lista): lista.ListIndex = 0
    form = "='" & ThisWorkbook.name & "'!F7_Campo(""" & pid & """,""" & ep & """,""" & vid & """,'Entradas'!A4:L20,'Entradas'!U4:X20,"
    demo.Range("U4").Formula = form & "6)": demo.Range("V4").NumberFormat = "@": demo.Range("V4").Formula = form & "12)"
    demo.Range("W4").Formula = form & "9)*2": demo.Range("X4").Formula = form & "1)"
    For j = 1 To 24: CopyCell7 sh.Cells(4, j), sh.Cells(10, j): Next j
    sh.Range("A4:X4").ClearContents
    For j = 1 To 24: CopyCell7 sh.Cells(6, j), sh.Cells(4, j): Next j
    sh.Range("A6:X6").ClearContents
    Application.Calculate
    r = Navigate7(wb, lista, sh)
    Check7 "R04_LISTA_APOS_REORDENAR", r = 10 And sh.Cells(r, 6).Value2 = "FIC6A01", "token resolve linha10 A", CStr(r)
    Check7 "R05_TIPOS_FORMATOS_IGNORE", Selo7(sh, 10) = selo And TypeName(sh.Range("L10").Value2) = "String" And Len(sh.Range("L10").Value2) = 44 And sh.Range("L10").NumberFormat = "@" And sh.Range("L10").Errors(xlNumberAsText).Ignore, "selo A:L exato e String44/@/Ignore", Selo7(sh, 10)
    Check7 "R06_FORMULAS_ID_REORDENACAO", demo.Range("U4").Value2 = "FIC6A01" And demo.Range("V4").Value2 = sh.Range("L10").Value2 And demo.Range("W4").Value2 = 1800 And demo.Range("X4").Value2 = 11001, "A/chave/1800/11001", Txt7(demo.Range("U4").Value2) & ";" & Txt7(demo.Range("W4").Value2)
    Check7 "BASE_REFERENCIA_FISICA", demo.Range("S4").Value2 = "FIC6A01", "A", Txt7(demo.Range("S4").Value2)
    Check7 "BASE_AUXILIAR_FISICO", sh.Range("M10").Value2 = 1800, "1800", Txt7(sh.Range("M10").Value2)
    n = Search7(wb, "", lista)
    Check7 "R07_LACUNAS_CONTAGEM", n = 6, "6 atraves das lacunas", CStr(n)
    lista.ListIndex = -1: r = Navigate7(wb, lista, sh)
    Check7 "R08_LISTINDEX_NEGATIVO", r = 0, "sem selecao", CStr(r)
    n = Search7(wb, "FIC6A01", lista): lista.ListIndex = 0
    eps.Cells(2, 6).Value2 = 2: r = Navigate7(wb, lista, sh)
    Check7 "R09_SELECAO_OBSOLETA", r = 0, "versao divergente recusa", CStr(r)
    eps.Cells(2, 6).Value2 = 1
    sh.Range("A12:X12").Value2 = sh.Range("A10:X10").Value2
    r = Resolve7(wb, pid, ep, vid, sh)
    Check7 "R10_PROJECAO_DUPLICADA", r = 0 And IsError(F7_Campo(pid, ep, vid, sh.Range("A4:L20"), sh.Range("U4:X20"), 6)), "recusa, sem primeiro candidato", CStr(r)
    sh.Range("A12:X12").ClearContents
    r = Resolve7(wb, "ID_AUSENTE", ep, vid, sh)
    Check7 "R11_ID_AUSENTE", r = 0, "nao usa placa/NFe/linha", CStr(r)
    Snapshot7 wb, "semente_depois": wb.save: wb.Close False: Set wb = Nothing
    Set wb = OpenInput7("tipos_perdidos")
    Set eps = wb.Worksheets("_F6_Episodios")
    r = Resolve7(wb, CStr(eps.Cells(2, 5).Value2), CStr(eps.Cells(2, 1).Value2), CStr(eps.Cells(2, 2).Value2), sh)
    Check7 "R12_COERCAO_DETECTADA", r = 0, "conteudo tipado divergente bloqueia", CStr(r)
    wb.Close False: Set wb = Nothing
    Set wb = OpenInput7("historico_ambiguo"): Set sh = wb.Worksheets("Entradas Anual")
    r = Unique7(sh, 24, CStr(sh.Range("X4").Value2))
    Check7 "R13_HISTORICO_AMBIGUO", r = 0, "nenhum primeiro historico", CStr(r)
    wb.Close False: Set wb = Nothing
    Set wb = OpenInput7("transferencias"): Set eps = wb.Worksheets("_F6_Episodios")
    n = Search7(wb, "FIC6A01", lista): lista.ListIndex = 0: r = Navigate7(wb, lista, sh)
    Check7 "R14_CUSTODIA_SUCESSIVA", n = 1 And r > 0 And sh.name = "Entradas Filial" And Item7(lista, 0, 6) = 4, "um atual filial, versao4 apos T/D/T", CStr(n) & ";" & sh.name & ";" & CStr(Item7(lista, 0, 6))
    ThisWorkbook.Activate: n = Search7(wb, "", lista)
    Check7 "R15_OUTRA_PASTA_ATIVA", n = 6 And (ActiveWorkbook Is ThisWorkbook), "leitor usa proprietario explicito", CStr(n)
    wb.Close False: Set wb = Nothing
    Check7 "R16_ESTADOS_APPLICATION", eventos = Application.EnableEvents And tela = Application.ScreenUpdating And alertas = Application.DisplayAlerts And calculo = Application.Calculation, "estados preservados", CStr(Application.EnableEvents) & ";" & CStr(Application.Calculation)
    gLog7.Close: Set gLog7 = Nothing: Debug.Print "F7_LEITORES_CONCLUIDO": Exit Sub
Falhou:
    If Not gLog7 Is Nothing Then gLog7.WriteLine "ERRO_EXECUCAO" & vbTab & "FAIL" & vbTab & "sem erro" & vbTab & Err.Number & ":" & Err.Description: gLog7.Close
    Debug.Print "F7_ERRO=" & Err.Number & ":" & Err.Description
    If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = eventos: Application.ScreenUpdating = tela: Application.DisplayAlerts = alertas: Application.Calculation = calculo
End Sub
Private Function ColdNames7() As Variant
    ColdNames7 = Array("OK", "ORFAO", "TRUNCADO", "PENDENTE", "DIVERGENTE", "POS_SAVE", "HISTORICO", "AUSENTE", "PAYLOAD", "VERSAO", "BAIXA", "COMPENSACAO")
End Function
Public Sub F7_PrepararFrio()
    Dim nomes As Variant, nome As Variant, pasta As String, origem As String, wb As Workbook, ctl As Worksheet, arq As Object, linha As String, estado As String, base As String, ep As String
    Call Guard7: nomes = ColdNames7()
    For Each nome In nomes
        pasta = ThisWorkbook.Path & "\frio_" & CStr(nome): FS7.CreateFolder pasta
        origem = "semente"
        If nome = "BAIXA" Then origem = "falha_baixa"
        If nome = "COMPENSACAO" Then origem = "falha_compensacao"
        FS7.CopyFile ThisWorkbook.Path & "\insumos\" & origem & ".xlsm", pasta & "\Fixture_Fria.xlsm", False
        Set wb = Application.Workbooks.Open(pasta & "\Fixture_Fria.xlsm", UpdateLinks:=0)
        Set ctl = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.count)): ctl.name = "_F7_Contrato"
        estado = "CONCLUIDA"
        If nome = "PENDENTE" Or nome = "BAIXA" Or nome = "COMPENSACAO" Then estado = "PENDENTE_CONCILIACAO"
        ep = CStr(wb.Worksheets("_F6_Episodios").Cells(2, 1).Value2)
        ctl.Range("A1:A9").value = Application.Transpose(Array("BaseID", "Revision", "OperationID", "EpisodeID", "Payload", "Estado", "Version", "SeloMinimo", "Schema"))
        ctl.Range("B1").Value2 = "BASE_SINTETICA_F7": ctl.Range("B2").Value2 = 1: ctl.Range("B3").Value2 = "OP_SINTETICA_F7"
        ctl.Range("B4").Value2 = ep: ctl.Range("B5").Value2 = "ENVELOPE_SINTETICO_V1": ctl.Range("B6").Value2 = estado
        ctl.Range("B7").Value2 = 2: ctl.Range("B8").Value2 = Selo7(wb.Worksheets("Entradas"), 4) & Selo7(wb.Worksheets("Todas Saidas"), 2): ctl.Range("B9").Value2 = 1
        base = "BASE_SINTETICA_F7"
        If nome = "DIVERGENTE" Then base = "OUTRA_BASE"
        linha = "1" & vbTab & base & vbTab & "1" & vbTab & "OP_SINTETICA_F7" & vbTab & ep & vbTab & "ENVELOPE_SINTETICO_V1" & vbTab & "2" & vbTab & estado & vbTab & "SCHEMA1"
        If nome = "ORFAO" Then linha = Replace(linha, "OP_SINTETICA_F7", "OP_ORFA_SINTETICA")
        If nome = "TRUNCADO" Then linha = "1" & vbTab & "INCOMPLETO"
        If nome = "POS_SAVE" Then linha = Replace(linha, "CONCLUIDA", "SAVE_INICIADO")
        If nome <> "AUSENTE" Then
            Set arq = FS7.CreateTextFile(pasta & "\diario-f7.tsv", False, True): arq.WriteLine linha: arq.Close
        End If
        If nome = "HISTORICO" Then wb.Worksheets("Todas Saidas").Range("F2").Value2 = "FIC7MUTACAO"
        wb.save: wb.Close False
    Next nome
    Debug.Print "F7_FRIO_PREPARADO_FECHAR_SESSAO"
End Sub
Private Function GuardCold7(ByVal wb As Workbook, ByVal payload As String, ByVal versao As Long, ByVal novaOperacao As Boolean) As String
    Dim ctl As Worksheet, arq As Object, linha As String, campos As Variant, pasta As String, n As Long, seq As Long
    On Error GoTo Invalido
    Set ctl = wb.Worksheets("_F7_Contrato"): pasta = wb.Path
    If Not FS7.FileExists(pasta & "\diario-f7.tsv") Then GuardCold7 = "CONCILIAR_AUSENTE": Exit Function
    Set arq = FS7.OpenTextFile(pasta & "\diario-f7.tsv", 1, False, -1)
    Do Until arq.AtEndOfStream
        linha = arq.ReadLine: campos = Split(linha, vbTab)
        If UBound(campos) <> 8 Then arq.Close: GuardCold7 = "CONCILIAR_TRUNCADO": Exit Function
        If Not IsNumeric(campos(0)) Then arq.Close: GuardCold7 = "CONCILIAR_TRUNCADO": Exit Function
        If CLng(campos(0)) <= seq Then arq.Close: GuardCold7 = "CONCILIAR_TRUNCADO": Exit Function
        seq = CLng(campos(0)): n = n + 1
    Loop
    arq.Close
    If n <> 1 Then GuardCold7 = "CONCILIAR_TRUNCADO": Exit Function
    If campos(1) <> ctl.Range("B1").Value2 Or campos(2) <> CStr(ctl.Range("B2").Value2) Or campos(8) <> "SCHEMA1" Then GuardCold7 = "CONCILIAR_BASE": Exit Function
    If campos(3) <> ctl.Range("B3").Value2 Or campos(4) <> ctl.Range("B4").Value2 Then GuardCold7 = "CONCILIAR_ORFAO": Exit Function
    If campos(5) <> ctl.Range("B5").Value2 Or campos(6) <> CStr(ctl.Range("B7").Value2) Or campos(7) <> ctl.Range("B6").Value2 Then GuardCold7 = "CONCILIAR_ESTADO": Exit Function
    If campos(7) <> "CONCLUIDA" Then GuardCold7 = "CONCILIAR_PENDENTE": Exit Function
    If ctl.Range("B8").Value2 <> Selo7(wb.Worksheets("Entradas"), 4) & Selo7(wb.Worksheets("Todas Saidas"), 2) Then GuardCold7 = "CONCILIAR_EFEITOS": Exit Function
    If payload <> ctl.Range("B5").Value2 Then GuardCold7 = "RECUSADA_PAYLOAD": Exit Function
    If novaOperacao And versao <> ctl.Range("B7").Value2 Then GuardCold7 = "RECUSADA_VERSAO": Exit Function
    GuardCold7 = "JA_CONCLUIDA": Exit Function
Invalido:
    GuardCold7 = "CONCILIAR_ERRO:" & CStr(Err.Number)
End Function
Public Sub F7_VerificarFrio()
    Dim nomes As Variant, nome As Variant, wb As Workbook, antes As String, resultado As String, esperado As String, payload As String, versao As Long, nova As Boolean
    Call Guard7: nomes = ColdNames7()
    Set gLog7 = FS7.CreateTextFile(ThisWorkbook.Path & "\frio_resultados.tsv", False, True)
    gLog7.WriteLine "Cenario" & vbTab & "Resultado" & vbTab & "Esperado" & vbTab & "Observado"
    On Error GoTo Falhou
    For Each nome In nomes
        Set wb = Application.Workbooks.Open(ThisWorkbook.Path & "\frio_" & CStr(nome) & "\Fixture_Fria.xlsm", UpdateLinks:=0)
        antes = Sig7(wb): payload = "ENVELOPE_SINTETICO_V1": versao = 2: nova = False
        esperado = "JA_CONCLUIDA"
        Select Case CStr(nome)
            Case "ORFAO": esperado = "CONCILIAR_ORFAO"
            Case "TRUNCADO": esperado = "CONCILIAR_TRUNCADO"
            Case "PENDENTE", "BAIXA", "COMPENSACAO": esperado = "CONCILIAR_PENDENTE": nova = True
            Case "DIVERGENTE": esperado = "CONCILIAR_BASE"
            Case "POS_SAVE": esperado = "CONCILIAR_ESTADO"
            Case "HISTORICO": esperado = "CONCILIAR_EFEITOS"
            Case "AUSENTE": esperado = "CONCILIAR_AUSENTE"
            Case "PAYLOAD": esperado = "RECUSADA_PAYLOAD": payload = "OUTRO_PAYLOAD"
            Case "VERSAO": esperado = "RECUSADA_VERSAO": nova = True: versao = 1
        End Select
        resultado = GuardCold7(wb, payload, versao, nova)
        Check7 "FRIO_" & CStr(nome), resultado = esperado And antes = Sig7(wb), esperado & "; sem escrita de negocio", resultado & ";IGUAL=" & CStr(antes = Sig7(wb))
        If nome = "PENDENTE" Or nome = "BAIXA" Or nome = "COMPENSACAO" Then
            resultado = GuardCold7(wb, payload, versao, False)
            Check7 "FRIO_" & CStr(nome) & "_MESMA_OP", resultado = esperado And antes = Sig7(wb), esperado, resultado
        End If
        wb.Close False: Set wb = Nothing
    Next nome
    gLog7.Close: Set gLog7 = Nothing: Debug.Print "F7_FRIO_CONCLUIDO": Exit Sub
Falhou:
    If Not gLog7 Is Nothing Then gLog7.WriteLine "ERRO_FRIO" & vbTab & "FAIL" & vbTab & "sem erro" & vbTab & Err.Number & ":" & Err.Description: gLog7.Close
    Debug.Print "F7_FRIO_ERRO=" & Err.Number & ":" & Err.Description
    If Not wb Is Nothing Then wb.Close False
End Sub

Private Sub RestaurarFormatoSintetico7(ByVal wb As Workbook)
    Dim links As Worksheet, sh As Worksheet, l As Long, r As Long, pid As String, mudou As Long, estritos As Boolean
    Set links = wb.Worksheets("_F6_Vinculos"): estritos = True
    For l = 2 To Last7(links)
        pid = CStr(links.Cells(l, 1).Value2): Set sh = wb.Worksheets(CStr(links.Cells(l, 6).Value2))
        r = Unique7(sh, 24, pid)
        If r > 0 Then
            If CStr(links.Cells(l, 9).Value2) <> Selo7(sh, r) Then estritos = False
            If sh.Cells(r, 10).NumberFormat = "m/d/yyyy" And InStr(1, CStr(links.Cells(l, 9).Value2), ":dd/mm/yyyy:", vbBinaryCompare) > 0 Then
                sh.Cells(r, 10).NumberFormat = "dd/mm/yyyy": mudou = mudou + 1
            End If
        End If
    Next l
    If wb.name = "exec_leitor.xlsm" Then Check7 "BASE_FORMATO_REABERTURA", estritos, "selo estrito preservado ao reabrir", "ANTES=" & CStr(estritos) & ";formatos restaurados na fixture=" & CStr(mudou)
End Sub


Public Sub F7_RetestarFormulas()
    Dim wb As Workbook, demo As Worksheet, sh As Worksheet, f As String, n As Long, lista As AuditoriaListaFake
    Call Guard7
    Set gLog7 = FS7.CreateTextFile(ThisWorkbook.Path & "\formulas_reteste.tsv", False, True)
    gLog7.WriteLine "Cenario" & vbTab & "Resultado" & vbTab & "Esperado" & vbTab & "Observado"
    Set wb = Application.Workbooks.Open(ThisWorkbook.Path & "\exec_semente.xlsm", UpdateLinks:=0)
    Snapshot7 wb, "formula_reteste_antes"
    Set demo = wb.Worksheets("Demonstração"): Set sh = wb.Worksheets("Entradas")
    f = CStr(demo.Range("U4").Formula): f = Left$(f, Len(f) - 2) & "12)"
    demo.Range("V4").NumberFormat = "General": demo.Range("V4").Formula = f
    demo.Range("V4").NumberFormat = "@": demo.Range("V4").Errors(xlNumberAsText).Ignore = True
    sh.Range("I4").Value2 = 1200
    Application.Calculate
    Check7 "R06_FORMULAS_ID_REORDENACAO", demo.Range("U4").Value2 = "FIC6A01" And demo.Range("V4").Value2 = sh.Range("L10").Value2 And demo.Range("W4").Value2 = 1800 And demo.Range("X4").Value2 = 11001, "A/chave/1800/11001", Txt7(demo.Range("U4").Value2) & ";" & Txt7(demo.Range("V4").Value2) & ";" & Txt7(demo.Range("W4").Value2)
    Check7 "R17_CHAVE_FORMULA_STRING", TypeName(demo.Range("V4").Value2) = "String" And Len(demo.Range("V4").Value2) = 44 And demo.Range("V4").NumberFormat = "@" And demo.Range("V4").Errors(xlNumberAsText).Ignore, "String44/@/Ignore", TypeName(demo.Range("V4").Value2) & ";" & CStr(Len(demo.Range("V4").Value2))
    Check7 "BASE_AUXILIAR_VALOR_DISTINTO", sh.Range("M10").Value2 = 1800, "1800 do episodio A", Txt7(sh.Range("M10").Value2)
    Snapshot7 wb, "formula_reteste_depois"
    wb.SaveCopyAs ThisWorkbook.Path & "\formula_reteste_nao_promover.xlsm": wb.Close False
    Set wb = Application.Workbooks.Open(ThisWorkbook.Path & "\exec_leitor.xlsm", UpdateLinks:=0)
    Set demo = wb.Worksheets("Demonstração")
    demo.Range("U8").Formula = "=COUNTIFS('_F6_Episodios'!C2:C100,""ATIVO"",'_F6_Episodios'!D2:D100,""MATRIZ"")"
    Application.Calculate
    Set lista = New AuditoriaListaFake: RestaurarFormatoSintetico7 wb: n = Search7(wb, "", lista)
    Check7 "R18_FORMULA_CONTAGEM_ATIVA", demo.Range("U8").Value2 = 3 And n = 6, "3 matriz e 6 globais validados", Txt7(demo.Range("U8").Value2) & ";" & CStr(n)
    wb.SaveCopyAs ThisWorkbook.Path & "\contagem_reteste_nao_promover.xlsm": wb.Close False
    gLog7.Close: Set gLog7 = Nothing: Debug.Print "F7_FORMULAS_RETESTADAS"
End Sub

