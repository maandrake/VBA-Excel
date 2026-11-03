Attribute VB_Name = "modExperimentosFase8"
Option Explicit
' SOMENTE FIXTURES FICTICIAS. Nao e gravador fiscal ou operacional aprovado.
Private Const MARCADOR8 As String = "VBA_Fase8_20261102_121422Z"
Private Const CHAVE8 As String = "01234567890123456789012345678901234567890123"
Private Type Dbl8
    Valor As Double
End Type
Private Type Bytes8
    B(0 To 7) As Byte
End Type
Private Type Mem8
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As LongPtr
    WorkingSetSize As LongPtr
    QuotaPeakPagedPoolUsage As LongPtr
    QuotaPagedPoolUsage As LongPtr
    QuotaPeakNonPagedPoolUsage As LongPtr
    QuotaNonPagedPoolUsage As LongPtr
    PagefileUsage As LongPtr
    PeakPagefileUsage As LongPtr
End Type
#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef cont As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef freq As Currency) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function K32GetProcessMemoryInfo Lib "kernel32" (ByVal processo As LongPtr, ByRef pmc As Mem8, ByVal tamanho As Long) As Long
#End If
Private log8 As Object
Private Function FSO8() As Object
    Set FSO8 = CreateObject("Scripting.FileSystemObject")
End Function
Private Sub Guard8()
    Dim raiz As String, atual As String
    If ThisWorkbook.Name <> "Fixture_Integridade_Fase8_Mestre.xlsm" Then Err.Raise 5, , "Mestre F8 obrigatorio"
    If InStr(1, ThisWorkbook.Path, MARCADOR8, vbTextCompare) = 0 Or InStr(1, ThisWorkbook.Path, "OneDrive", vbTextCompare) > 0 Then Err.Raise 5, , "Pasta F8 TEMP obrigatoria"
    raiz = FSO8.GetFolder(Environ$("TEMP")).ShortPath & "\": atual = FSO8.GetFolder(ThisWorkbook.Path).ShortPath & "\"
    If StrComp(Left$(atual, Len(raiz)), raiz, vbTextCompare) <> 0 Then Err.Raise 5, , "Somente TEMP"
    If ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then Err.Raise 5, , "Marcador ausente"
End Sub
Private Function Limpo8(ByVal v As Variant) As String
    If IsEmpty(v) Then Limpo8 = "<EMPTY>": Exit Function
    Limpo8 = Replace(Replace(Replace(CStr(v), vbTab, "<TAB>"), vbCr, "<CR>"), vbLf, "<LF>")
End Function
Private Function Ignore8(ByVal c As Range) As String
    On Error GoTo Falhou
    Ignore8 = CStr(c.Errors(xlNumberAsText).Ignore): Exit Function
Falhou:
    Ignore8 = "NAO_LIDO:" & CStr(Err.Number)
End Function
Private Sub Check8(ByVal id As String, ByVal ok As Boolean, ByVal esperado As String, ByVal observado As String)
    log8.WriteLine id & vbTab & IIf(ok, "PASS", "FAIL") & vbTab & Limpo8(esperado) & vbTab & Limpo8(observado)
End Sub
Private Sub IniciarLog8(ByVal nome As String)
    Set log8 = FSO8.CreateTextFile(ThisWorkbook.Path & "\" & nome & ".tsv", False, True)
    log8.WriteLine "Cenario" & vbTab & "Resultado" & vbTab & "Esperado" & vbTab & "Observado"
End Sub
Public Function F8_HexDouble(ByVal valor As Double) As String
    Dim origem As Dbl8, destino As Bytes8, i As Long, s As String
    origem.Valor = valor: LSet destino = origem
    For i = 7 To 0 Step -1: s = s & Right$("0" & Hex$(destino.B(i)), 2): Next i
    F8_HexDouble = s
End Function
Public Function F8_CanonData(ByVal celula As Range, ByVal politica As String) As String
    Dim v As Variant, f As String, grupo As String, serial As Double
    v = celula.Value2
    If TypeName(v) <> "Double" Then F8_CanonData = "ERRO:DATA_TIPO_" & TypeName(v): Exit Function
    serial = CDbl(v)
    If celula.Parent.Parent.Date1904 Then F8_CanonData = "ERRO:SISTEMA_1904": Exit Function
    If serial < 61 Or serial > 2958465 Or serial <> Fix(serial) Then F8_CanonData = "ERRO:DATA_INVALIDA": Exit Function
    f = LCase$(CStr(celula.NumberFormat))
    Select Case politica
        Case "ISO_LITERAL"
            If f = "yyyy""-""mm""-""dd" Or f = "yyyy\-mm\-dd" Or f = "yyyy-mm-dd" Then grupo = "ISO_YMD"
        Case "BR_LITERAL"
            If f = "dd""/""mm""/""yyyy" Or f = "dd\/mm\/yyyy" Then grupo = "BR_DMY"
        Case "LEGACY_D14_LOCAL"
            ' Aceitacao explicita de apresentacao local; nao exige exibicao brasileira.
            If f = "dd/mm/yyyy" Or f = "m/d/yyyy" Then grupo = "D14_LOCAL"
    End Select
    If grupo = "" Then F8_CanonData = "ERRO:FORMATO_" & f: Exit Function
    F8_CanonData = "DATE_ONLY|EXCEL1900|DOUBLE_HEX=" & F8_HexDouble(serial) & "|POLICY=" & politica & "|FMT=" & grupo
End Function
Public Function F8_CanonRegistro(ByVal wb As Workbook, ByVal r As Long) As String
    Dim sh As Worksheet, a As Variant, c As Long, v As String, parte As String, f As String
    Set sh = wb.Worksheets("_F8_Dados"): a = sh.Cells(r, 1).Resize(1, 18).Value2
    If CStr(a(1, 18)) <> "F8_V1" Then F8_CanonRegistro = "ERRO:SCHEMA": Exit Function
    For c = 1 To 18
        If sh.Cells(r, c).HasFormula Then F8_CanonRegistro = "ERRO:FORMULA_NA_FONTE": Exit Function
        If c = 15 Then
            v = F8_CanonData(sh.Cells(r, c), CStr(a(1, 17)))
            If Left$(v, 5) = "ERRO:" Then F8_CanonRegistro = v: Exit Function
            f = "POLITICA_DATA"
        Else
            If c = 7 Or c = 13 Or c = 14 Or c = 16 Then
                If TypeName(a(1, c)) <> "Double" Then F8_CanonRegistro = "ERRO:TIPO_NUMERICO": Exit Function
                v = "DOUBLE_HEX:" & F8_HexDouble(CDbl(a(1, c)))
            Else
                If TypeName(a(1, c)) <> "String" Then F8_CanonRegistro = "ERRO:TIPO_TEXTO": Exit Function
                v = "STRING:" & CStr(a(1, c))
            End If
            f = CStr(sh.Cells(r, c).NumberFormat)
        End If
        If c = 12 Then
            If Len(CStr(a(1, c))) <> 44 Or CStr(a(1, c)) Like "*[!0-9]*" Or f <> "@" Then F8_CanonRegistro = "ERRO:CHAVE": Exit Function
        End If
        parte = v & ";FMT=" & f & ";IGNORE=" & Ignore8(sh.Cells(r, c))
        F8_CanonRegistro = F8_CanonRegistro & Len(parte) & ":" & parte
    Next c
End Function
Private Function Tabela8(ByVal wb As Workbook, ByVal nome As String, ByVal nomes As Variant) As Worksheet
    Dim sh As Worksheet, i As Long
    Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): sh.Name = nome
    For i = LBound(nomes) To UBound(nomes): sh.Cells(1, i + 1).Value2 = nomes(i): Next i
    Set Tabela8 = sh
End Function
Private Function Last8(ByVal sh As Worksheet) As Long
    Last8 = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
End Function
Private Function Unica8(ByVal sh As Worksheet, ByVal coluna As Long, ByVal id As String) As Long
    Dim r As Long, qtd As Long, pos As Long, n As Long
    If id = "" Then Exit Function
    n = sh.Cells(sh.Rows.Count, coluna).End(xlUp).Row
    For r = 2 To n
        If CStr(sh.Cells(r, coluna).Value2) = id Then qtd = qtd + 1: pos = r
    Next r
    If qtd = 1 Then Unica8 = pos
End Function
Private Sub Estrutura8(ByVal wb As Workbook)
    Tabela8 wb, "_F8_Dados", Array("ProjectionID", "EpisodeID", "VehicleID", "MovementID", "OperationID", "AttemptID", "Version", "Estado", "Local", "Placa", "NFe", "Chave", "Entrada", "Saida", "Data", "Fator", "DatePolicy", "Schema")
    Tabela8 wb, "_F8_Episodios", Array("EpisodeID", "VehicleID", "Estado", "ProjectionIDAtual", "Version")
    Tabela8 wb, "_F8_Movimentos", Array("MovementID", "EpisodeID", "Tipo", "Estado", "ProjectionIDOrigem", "OperationID")
    Tabela8 wb, "_F8_Operacoes", Array("OperationID", "EpisodeID", "Payload", "Estado", "AttemptID", "ExpectedVersion", "EffectSeal", "Revision")
    Tabela8 wb, "_F8_Selos", Array("ProjectionID", "Canonico")
    Tabela8 wb, "_F8_Controle", Array("Campo", "Valor")
    Tabela8 wb, "_F8_Leitura", Array("ProjectionID", "EpisodeID", "VehicleID", "Version", "Placa", "Entrada", "Saida", "Margem", "Auxiliar", "NFe", "Chave", "Data", "Anual", "Elegivel")
    With wb.Worksheets("_F8_Controle")
        .Range("A2").Value2 = "Revision": .Range("B2").Value2 = 1
        .Range("A1").Value2 = "BaseID": .Range("B1").Value2 = "BASE_F8_SINTETICA"
        .Range("A3").Value2 = "Schema": .Range("B3").Value2 = "F8_SCHEMA_1"
    End With
End Sub
Private Sub GravarLinha8(ByVal wb As Workbook, ByVal r As Long, ByVal pid As String, ByVal ep As String, ByVal vid As String, ByVal placa As String, ByVal entrada As Double, ByVal saida As Double, ByVal fator As Double, ByVal localEstoque As String, ByVal estado As String, ByVal versao As Long)
    Dim a(1 To 1, 1 To 18) As Variant, c As Long, sh As Worksheet
    a(1, 1) = pid: a(1, 2) = ep: a(1, 3) = vid: a(1, 4) = "MOV_ENT_" & pid: a(1, 5) = "OP_ENT_" & pid: a(1, 6) = "ATT_ENT_" & pid
    a(1, 7) = versao: a(1, 8) = estado: a(1, 9) = localEstoque: a(1, 10) = placa: a(1, 11) = "88001": a(1, 12) = CHAVE8
    a(1, 13) = entrada: a(1, 14) = saida: a(1, 15) = CDbl(DateSerial(2026, 11, 2)): a(1, 16) = fator: a(1, 17) = "ISO_LITERAL": a(1, 18) = "F8_V1"
    Set sh = wb.Worksheets("_F8_Dados")
    For c = 1 To 18
        If c <> 7 And c <> 13 And c <> 14 And c <> 15 And c <> 16 Then sh.Cells(r, c).NumberFormat = "@"
    Next c
    sh.Cells(r, 15).NumberFormat = "yyyy""-""mm""-""dd"
    sh.Cells(r, 1).Resize(1, 18).Value2 = a: sh.Cells(r, 12).Errors(xlNumberAsText).Ignore = True
End Sub
Private Sub Selar8(ByVal wb As Workbook, ByVal r As Long)
    Dim sh As Worksheet, pos As Long, pid As String
    pid = CStr(wb.Worksheets("_F8_Dados").Cells(r, 1).Value2): Set sh = wb.Worksheets("_F8_Selos")
    pos = Unica8(sh, 1, pid): If pos = 0 Then pos = Last8(sh) + 1
    sh.Cells(pos, 1).Value2 = pid: sh.Cells(pos, 2).Value2 = F8_CanonRegistro(wb, r)
End Sub
Private Sub PrepararPequena8(ByVal wb As Workbook)
    Dim sh As Worksheet, i As Long, nomes As Variant, eps As Variant, vids As Variant, estados As Variant, locais As Variant, valores As Variant, saidas As Variant, fatores As Variant, a As Variant
    Estrutura8 wb
    nomes = Array("PID_A1", "PID_B1", "PID_A2", "PID_C1"): eps = Array("EP_A1", "EP_B1", "EP_A2", "EP_C1")
    vids = Array("VID_A", "VID_B", "VID_A", "VID_C"): estados = Array("BAIXADO", "ATIVO", "ATIVO", "ATIVO"): locais = Array("MATRIZ", "MATRIZ", "MATRIZ", "FILIAL")
    valores = Array(1000#, 2200#, 1500#, 3100#): saidas = Array(1370#, 2790#, 1940#, 4110#): fatores = Array(2#, 3#, 4#, 5#)
    For i = 0 To 3
        GravarLinha8 wb, i + 2, CStr(nomes(i)), CStr(eps(i)), CStr(vids(i)), IIf(i = 0 Or i = 2, "FIC8A01", "FIC8" & Chr$(65 + i) & "01"), CDbl(valores(i)), CDbl(saidas(i)), CDbl(fatores(i)), CStr(locais(i)), CStr(estados(i)), IIf(i = 0, 2, 1)
        Set sh = wb.Worksheets("_F8_Episodios")
        sh.Cells(i + 2, 1).Value2 = eps(i): sh.Cells(i + 2, 2).Value2 = vids(i): sh.Cells(i + 2, 3).Value2 = IIf(i = 0, "ENCERRADO", "ATIVO"): sh.Cells(i + 2, 4).Value2 = nomes(i): sh.Cells(i + 2, 5).Value2 = IIf(i = 0, 2, 1)
        Set sh = wb.Worksheets("_F8_Movimentos")
        sh.Cells(i + 2, 1).Value2 = "MOV_ENT_" & nomes(i): sh.Cells(i + 2, 2).Value2 = eps(i): sh.Cells(i + 2, 3).Value2 = "Entrada": sh.Cells(i + 2, 4).Value2 = "CONCLUIDA": sh.Cells(i + 2, 5).Value2 = nomes(i): sh.Cells(i + 2, 6).Value2 = "OP_ENT_" & nomes(i)
        Selar8 wb, i + 2
    Next i
    With wb.Worksheets("_F8_Movimentos")
        .Range("A6:F6").Value2 = Array("MOV_VENDA_A1", "EP_A1", "Venda", "CONCLUIDA", "PID_A1", "OP_VENDA_A1")
    End With
End Sub
Public Function F8_Campo(ByVal pid As String, ByVal ep As String, ByVal vid As String, ByVal dados As Range, ByVal campo As Long) As Variant
    Dim a As Variant, r As Long, n As Long, pos As Long
    On Error GoTo Invalido
    If pid = "" Or ep = "" Or vid = "" Or dados.Columns.Count <> 18 Or campo < 1 Or campo > 18 Then GoTo Invalido
    a = dados.Value2
    For r = 1 To UBound(a, 1)
        If CStr(a(r, 1)) = pid Then n = n + 1: pos = r
    Next r
    If n = 0 Then F8_Campo = CVErr(xlErrNA): Exit Function
    If n <> 1 Then GoTo Invalido
    If CStr(a(pos, 2)) <> ep Or CStr(a(pos, 3)) <> vid Then GoTo Invalido
    F8_Campo = a(pos, campo): Exit Function
Invalido:
    F8_Campo = CVErr(xlErrValue)
End Function
Public Function F8_Anual(ByVal ep As String, ByVal episodios As Range, ByVal movimentos As Range) As Variant
    Dim e As Variant, m As Variant, i As Long, n As Long, qtd As Long, vistos As Object, mov As String
    On Error GoTo Invalido
    If ep = "" Then F8_Anual = CVErr(xlErrNA): Exit Function
    e = episodios.Value2: m = movimentos.Value2
    For i = 1 To UBound(e, 1): If CStr(e(i, 1)) = ep Then n = n + 1
    Next i
    If n = 0 Then F8_Anual = CVErr(xlErrNA): Exit Function
    If n <> 1 Then GoTo Invalido
    Set vistos = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(m, 1)
        mov = CStr(m(i, 1))
        If mov <> "" Then
            If vistos.Exists(mov) Then GoTo Invalido
            vistos.Add mov, True
            If CStr(m(i, 2)) = ep Then
                If CStr(m(i, 4)) <> "CONCLUIDA" Then GoTo Invalido
                If CStr(m(i, 3)) = "Venda" Then qtd = qtd + 1
            End If
        End If
    Next i
    If qtd > 1 Then GoTo Invalido
    F8_Anual = IIf(qtd = 1, "SIM", "NAO"): Exit Function
Invalido:
    F8_Anual = CVErr(xlErrValue)
End Function
Private Function NovaFixture8(ByVal nome As String) As Workbook
    Dim destino As String
    Call Guard8: destino = ThisWorkbook.Path & "\" & nome & ".xlsm"
    If FSO8.FileExists(destino) Then Err.Raise 58, , "Fixture ja existe"
    ThisWorkbook.SaveCopyAs destino
    Set NovaFixture8 = Application.Workbooks.Open(destino, UpdateLinks:=0)
End Function
Private Sub CopiarCelula8(ByVal origem As Range, ByVal destino As Range)
    destino.NumberFormat = origem.NumberFormat
    If origem.HasFormula Then destino.Formula = origem.Formula Else destino.Value2 = origem.Value2
    destino.Errors(xlNumberAsText).Ignore = origem.Errors(xlNumberAsText).Ignore
End Sub
Private Sub Projecao8(ByVal wb As Workbook)
    Dim sh As Worksheet, src As Worksheet, r As Long, n As Long, pid As String, ep As String, vid As String, f As String
    Set sh = wb.Worksheets("_F8_Leitura"): Set src = wb.Worksheets("_F8_Dados")
    n = Last8(src): sh.Range("A2:N40").ClearContents
    For r = 2 To n
        If CStr(src.Cells(r, 1).Value2) <> "" Then
            pid = CStr(src.Cells(r, 1).Value2): ep = CStr(src.Cells(r, 2).Value2): vid = CStr(src.Cells(r, 3).Value2)
            sh.Cells(r, 1).Value2 = pid: sh.Cells(r, 2).Value2 = ep: sh.Cells(r, 3).Value2 = vid: sh.Cells(r, 4).Value2 = src.Cells(r, 7).Value2
            f = "='" & ThisWorkbook.Name & "'!F8_Campo(""" & pid & """,""" & ep & """,""" & vid & """,'_F8_Dados'!A2:R40,"
            sh.Cells(r, 5).Formula = f & "10)": sh.Cells(r, 6).Formula = f & "13)": sh.Cells(r, 7).Formula = f & "14)"
            sh.Cells(r, 8).Formula = f & "14)-" & Mid$(f, 2) & "13)": sh.Cells(r, 9).Formula = f & "13)*" & Mid$(f, 2) & "16)"
            sh.Cells(r, 10).Formula = f & "11)"
            sh.Cells(r, 11).NumberFormat = "General": sh.Cells(r, 11).Formula = f & "12)": sh.Cells(r, 11).NumberFormat = "@": sh.Cells(r, 11).Errors(xlNumberAsText).Ignore = True
            sh.Cells(r, 13).Formula = "='" & ThisWorkbook.Name & "'!F8_Anual(""" & ep & """,'_F8_Episodios'!A2:E40,'_F8_Movimentos'!A2:F40)"
        End If
    Next r
    sh.Calculate
End Sub
Private Sub Snapshot8(ByVal wb As Workbook, ByVal nome As String)
    Dim arq As Object, sh As Worksheet, c As Range
    Set arq = FSO8.CreateTextFile(ThisWorkbook.Path & "\" & nome & ".tsv", False, True)
    arq.WriteLine "Aba" & vbTab & "Celula" & vbTab & "Tipo" & vbTab & "Valor" & vbTab & "Formula" & vbTab & "NumberFormat" & vbTab & "Ignore"
    For Each sh In wb.Worksheets
        If Left$(sh.Name, 4) = "_F8_" Then
            For Each c In sh.Range("A1:R40").Cells
                arq.WriteLine sh.Name & vbTab & c.Address(False, False) & vbTab & TypeName(c.Value2) & vbTab & Limpo8(c.Value2) & vbTab & Limpo8(c.Formula) & vbTab & c.NumberFormat & vbTab & Ignore8(c)
            Next c
        End If
    Next sh
    arq.Close
End Sub

Public Function F8_Admitir(ByVal wb As Workbook, ByVal pid As String, ByVal ep As String, ByVal vid As String, ByVal versao As Long, ByVal op As String, ByVal payload As String) As String
    Dim sh As Worksheet, r As Long, i As Long, repo As clsRepositorioFase8, dado As Object
    Set sh = wb.Worksheets("_F8_Operacoes")
    If op = "" Or payload = "" Then F8_Admitir = "INTENCAO_INVALIDA": Exit Function
    If Application.CountIf(sh.Range("A2:A" & Application.Max(2, Last8(sh))), op) > 1 Then F8_Admitir = "OPERACAO_DUPLICADA": Exit Function
    r = Unica8(sh, 1, op)
    If r > 0 Then
        If CStr(sh.Cells(r, 2).Value2) <> ep Or CStr(sh.Cells(r, 3).Value2) <> payload Then F8_Admitir = "RECUSADA_PAYLOAD": Exit Function
        If CStr(sh.Cells(r, 4).Value2) = "CONCLUIDA" Then
            i = Unica8(wb.Worksheets("_F8_Dados"), 1, pid)
            If i = 0 Then F8_Admitir = "AGUARDANDO_CONCILIACAO": Exit Function
            If CStr(wb.Worksheets("_F8_Dados").Cells(i, 2).Value2) <> ep Or CStr(wb.Worksheets("_F8_Dados").Cells(i, 3).Value2) <> vid Or CLng(sh.Cells(r, 6).Value2) <> versao Then F8_Admitir = "RECUSADA_IDENTIDADE_VERSAO": Exit Function
            If CStr(sh.Cells(r, 7).Value2) <> F8_CanonRegistro(wb, i) Then F8_Admitir = "AGUARDANDO_CONCILIACAO": Exit Function
            F8_Admitir = "JA_CONCLUIDA": Exit Function
        End If
        F8_Admitir = "OPERACAO_PENDENTE": Exit Function
    End If
    Set repo = New clsRepositorioFase8: repo.Iniciar wb
    Set dado = repo.Resolver(pid, ep, vid, versao)
    F8_Admitir = CStr(dado("Status")): repo.Finalizar
    If F8_Admitir <> "ENCONTRADO" Then Exit Function
    r = Last8(sh) + 1
    sh.Cells(r, 1).Value2 = op: sh.Cells(r, 2).Value2 = ep: sh.Cells(r, 3).Value2 = payload: sh.Cells(r, 4).Value2 = "PREPARADA"
    sh.Cells(r, 5).Value2 = op & "_TENTATIVA1": sh.Cells(r, 6).Value2 = versao
    sh.Cells(r, 7).Value2 = F8_CanonRegistro(wb, CLng(dado("Row"))): sh.Cells(r, 8).Value2 = wb.Worksheets("_F8_Controle").Range("B2").Value2
    F8_Admitir = "PREPARADA"
End Function
Private Function MudarCustodia8(ByVal wb As Workbook, ByVal pid As String, ByVal ep As String, ByVal vid As String, ByVal versao As Long, ByVal novoLocal As String, ByVal novoPID As String) As String
    Dim repo As clsRepositorioFase8, d As Object, sh As Worksheet, r As Long, nr As Long, e As Long, c As Long, mov As Long
    ' Mutacao de fixture, nao gravador de transferencia/devolucao do produto.
    Set repo = New clsRepositorioFase8: repo.Iniciar wb
    Set d = repo.Resolver(pid, ep, vid, versao): MudarCustodia8 = CStr(d("Status"))
    If MudarCustodia8 <> "ENCONTRADO" Then repo.Finalizar: Exit Function
    Set sh = wb.Worksheets("_F8_Dados")
    If Application.CountIf(sh.Range("A2:A" & Application.Max(2, Last8(sh))), novoPID) <> 0 Then MudarCustodia8 = "DUPLICADO": repo.Finalizar: Exit Function
    r = CLng(d("Row")): nr = Last8(sh) + 1
    For c = 1 To 18: CopiarCelula8 sh.Cells(r, c), sh.Cells(nr, c): Next c
    sh.Cells(r, 8).Value2 = "ARQUIVADO"
    sh.Cells(nr, 1).Value2 = novoPID: sh.Cells(nr, 7).Value2 = versao + 1: sh.Cells(nr, 9).Value2 = novoLocal
    sh.Cells(nr, 4).Value2 = "MOV_" & novoPID: sh.Cells(nr, 5).Value2 = "OP_" & novoPID: sh.Cells(nr, 6).Value2 = "ATT_" & novoPID
    e = Unica8(wb.Worksheets("_F8_Episodios"), 1, ep)
    wb.Worksheets("_F8_Episodios").Cells(e, 4).Value2 = novoPID: wb.Worksheets("_F8_Episodios").Cells(e, 5).Value2 = versao + 1
    mov = Last8(wb.Worksheets("_F8_Movimentos")) + 1
    wb.Worksheets("_F8_Movimentos").Cells(mov, 1).Resize(1, 6).Value2 = Array("MOV_" & novoPID, ep, IIf(novoLocal = "FILIAL", "Transferencia", "Devolucao"), "CONCLUIDA", pid, "OP_" & novoPID)
    Selar8 wb, r: Selar8 wb, nr
    wb.Worksheets("_F8_Controle").Range("B2").Value2 = wb.Worksheets("_F8_Controle").Range("B2").Value2 + 1
    MudarCustodia8 = "CONCLUIDA_FIXTURE": repo.Finalizar
End Function
Public Sub F8_ReproduzirBase7()
    Dim wb As Workbook, canon As String, n As Variant
    Call Guard8: IniciarLog8 "reproducao_f7"
    Set wb = Application.Workbooks.Open(ThisWorkbook.Path & "\baseline\leitor.xlsm", UpdateLinks:=0)
    canon = CStr(wb.Worksheets("_F6_Vinculos").Range("I2").Value2)
    Check8 "BASE_F8_01_FORMATO", InStr(canon, ":" & CStr(wb.Worksheets("Entradas").Range("J4").NumberFormat) & ":") > 0, "formato efetivo compativel com selo F7", CStr(wb.Worksheets("Entradas").Range("J4").NumberFormat)
    Check8 "BASE_F8_02_ANUAL", wb.Worksheets("Entradas Anual").Range("P8").Value2 = "NAO", "NAO para reentrada", Limpo8(wb.Worksheets("Entradas Anual").Range("P8").Value2)
    wb.Close False
    Set wb = Application.Workbooks.Open(ThisWorkbook.Path & "\baseline\formula.xlsm", UpdateLinks:=0)
    Check8 "BASE_F8_03_FISICA", wb.Worksheets("Demonstração").Range("S4").Value2 = "FIC6A01", "FIC6A01", Limpo8(wb.Worksheets("Demonstração").Range("S4").Value2)
    Check8 "BASE_F8_04_AUXILIAR", wb.Worksheets("Entradas").Range("M10").Value2 = 1800, "1800 do veiculo A", Limpo8(wb.Worksheets("Entradas").Range("M10").Value2)
    wb.Close False: log8.Close: Set log8 = Nothing: Debug.Print "F8_BASE_REPRODUZIDA"
End Sub
Public Sub F8_Integrado()
    Dim wb As Workbook, sh As Worksheet, leitura As Worksheet, repo As clsRepositorioFase8, d As Object, lista As AuditoriaListaFake
    Dim a As Variant, token As Variant, r As Long, c As Long, e As Long, estado As String, p As String, eps As Worksheet, frm As Cad_NFE, epoca As Long, eventos As Boolean
    Call Guard8: IniciarLog8 "integrado_f8_r2": eventos = Application.EnableEvents
    On Error GoTo Falhou
    Set wb = NovaFixture8("Fixture_F8_Integrada_R2"): PrepararPequena8 wb: Set sh = wb.Worksheets("_F8_Dados"): Set eps = wb.Worksheets("_F8_Episodios")
    Projecao8 wb: Snapshot8 wb, "integrado_antes_r2"
    Set repo = New clsRepositorioFase8: repo.Iniciar wb: Set lista = New AuditoriaListaFake
    estado = repo.Listar("FIC8A01", "OPERACIONAL", "TODOS", a): lista.ColumnCount = 8: lista.List = a: lista.ListIndex = 0: token = lista.List
    Check8 "F8_02_REENTRADA_PESQUISA", estado = "ENCONTRADO" And lista.ListCount = 1 And CStr(token(0, 1)) = "EP_A2", "somente EP_A2 ativo", estado & ";" & CStr(token(0, 1))
    Set d = repo.Resolver(CStr(token(0, 0)), CStr(token(0, 1)), CStr(token(0, 2)), CLng(token(0, 3)), "OPERACIONAL", CStr(token(0, 7)))
    a = d("Data")
    Set frm = New Cad_NFE: frm.N_Nfe.Value = CStr(a(1, 11)): frm.PLACA.Value = CStr(a(1, 10)): frm.txt_Valor.Value = CStr(a(1, 13))
    Check8 "F8_FORMULARIO_REAL", frm.N_Nfe.Value = "88001" And frm.PLACA.Value = "FIC8A01" And CDbl(frm.txt_Valor.Value) = 1500, "controles reais com EP_A2", frm.N_Nfe.Name & ";" & frm.PLACA.Name & ";" & frm.txt_Valor.Name & ";" & CStr(frm.txt_Valor.Value)
    Unload frm: Set frm = Nothing
    Set leitura = wb.Worksheets("_F8_Leitura")
    Check8 "F8_02_ANUAL_EPISODIOS", leitura.Range("M2").Value2 = "SIM" And leitura.Range("M4").Value2 = "NAO", "EP_A1 SIM; EP_A2 NAO", Limpo8(leitura.Range("M2").Value2) & ";" & Limpo8(leitura.Range("M4").Value2)
    Check8 "F8_02_ANUAL_AUSENTE", IsError(F8_Anual("INEXISTENTE", eps.Range("A2:E40"), wb.Worksheets("_F8_Movimentos").Range("A2:F40"))), "erro explicito", Limpo8(F8_Anual("INEXISTENTE", eps.Range("A2:E40"), wb.Worksheets("_F8_Movimentos").Range("A2:F40")))
    epoca = repo.EpocaIndice
    OrdenarComMetadados8 sh, 5
    Set d = repo.Resolver("PID_A2", "EP_A2", "VID_A", 1): r = CLng(d("Row"))
    For c = 1 To 18: CopiarCelula8 sh.Cells(r, c), sh.Cells(12, c): Next c
    sh.Cells(r, 1).Resize(1, 18).ClearContents
    GravarLinha8 wb, r, "PID_NOVO", "EP_NOVO", "VID_NOVO", "FIC8Z99", 9900, 12900, 9, "MATRIZ", "ATIVO", 1
    eps.Range("A6:E6").Value2 = Array("EP_NOVO", "VID_NOVO", "ATIVO", "PID_NOVO", 1)
    wb.Worksheets("_F8_Movimentos").Range("A7:F7").Value2 = Array("MOV_ENT_PID_NOVO", "EP_NOVO", "Entrada", "CONCLUIDA", "PID_NOVO", "OP_ENT_PID_NOVO")
    Selar8 wb, r
    leitura.Calculate
    Set d = repo.Resolver(CStr(token(0, 0)), CStr(token(0, 1)), CStr(token(0, 2)), CLng(token(0, 3)))
    Check8 "F8_03_TOKEN_APOS_SORT_REUSO", CStr(d("Status")) = "ENCONTRADO" And CLng(d("Row")) = 12 And repo.EpocaIndice > epoca, "PID_A2 linha12, indice invalidado", CStr(d("Status")) & ";" & CStr(d("Row"))
    Check8 "F8_03_FORMULA_ID", leitura.Range("E4").Value2 = "FIC8A01" And leitura.Range("F4").Value2 = 1500, "mesmo veiculo/episodio apos sort e reuso", Limpo8(leitura.Range("E4").Value2) & ";" & Limpo8(leitura.Range("F4").Value2)
    Check8 "F8_04_AUX_DISCRIMINANTE", leitura.Range("H4").Value2 = 440 And leitura.Range("I4").Value2 = 6000 And leitura.Range("I3").Value2 = 6600 And leitura.Range("I5").Value2 = 15500, "A2 margem440/aux6000; B6600; C15500", Limpo8(leitura.Range("H4").Value2) & ";" & Limpo8(leitura.Range("I4").Value2) & ";" & Limpo8(leitura.Range("I3").Value2) & ";" & Limpo8(leitura.Range("I5").Value2)
    Check8 "F8_CHAVE_TIPOS_FORMATOS", TypeName(sh.Range("L12").Value2) = "String" And sh.Range("L12").Value2 = CHAVE8 And sh.Range("L12").NumberFormat = "@" And sh.Range("L12").Errors(xlNumberAsText).Ignore, "String44/@/Ignore", TypeName(sh.Range("L12").Value2) & ";" & CStr(Len(sh.Range("L12").Value2)) & ";" & sh.Range("L12").NumberFormat
    lista.Clear: Check8 "F8_LISTINDEX_NEGATIVO", lista.ListIndex = -1, "nenhuma selecao autoriza escrita", CStr(lista.ListIndex)
    estado = repo.Listar("", "OPERACIONAL", "TODOS", a)
    Check8 "F8_LACUNAS_CONTAGEM", estado = "ENCONTRADO" And UBound(a, 1) + 1 = 4, "4 ativos incluindo novo ocupante e lacuna", estado & ";" & CStr(UBound(a, 1) + 1)
    Set d = repo.Resolver("PID_A1", "EP_A1", "VID_A", 2): Check8 "F8_RETORNO_BAIXADO", d("Status") = "BAIXADO", "BAIXADO", CStr(d("Status"))
    Set d = repo.Resolver("PID_A1", "EP_A1", "VID_A", 2, "HISTORICO"): Check8 "F8_HISTORICO_ENTRADA", d("Status") = "ENCONTRADO", "historico preservado", CStr(d("Status"))
    sh.Cells(12, 13).Value2 = 1555: sh.Cells(12, 7).Value2 = 2: eps.Range("E4").Value2 = 2: Selar8 wb, 12: wb.Worksheets("_F8_Controle").Range("B2").Value2 = 2
    Set d = repo.Resolver(CStr(token(0, 0)), CStr(token(0, 1)), CStr(token(0, 2)), CLng(token(0, 3)), "OPERACIONAL", CStr(token(0, 7)))
    Check8 "F8_TOKEN_OBSOLETO_SEM_MUTACAO", d("Status") = "VERSAO_OBSOLETA", "VERSAO_OBSOLETA", CStr(d("Status"))
    sh.Cells(12, 13).Value2 = 1400
    Set d = repo.Resolver("PID_A2", "EP_A2", "VID_A", 2): Check8 "F8_EDICAO_FORA_CONTRATO", d("Status") = "CONTEUDO_DIVERGENTE", "CONTEUDO_DIVERGENTE", CStr(d("Status"))
    sh.Cells(12, 13).Value2 = 1555
    eps.Range("A8:E8").Value2 = eps.Range("A4:E4").Value2
    Set d = repo.Resolver("PID_A2", "EP_A2", "VID_A", 2)
    Check8 "F8_02_EPISODIO_DUPLICADO", d("Status") = "IDENTIDADE_AMBIGUA" And IsError(F8_Anual("EP_A2", eps.Range("A2:E40"), wb.Worksheets("_F8_Movimentos").Range("A2:F40"))), "erro explicito sem primeiro resultado", CStr(d("Status"))
    eps.Range("A8:E8").ClearContents
    Application.EnableEvents = False: p = CStr(sh.Cells(r, 1).Value2): sh.Cells(r, 1).Value2 = "PID_A2": Application.EnableEvents = eventos
    Set d = repo.Resolver("PID_A2", "EP_A2", "VID_A", 2): Check8 "F8_DUPLICADO_SEM_EVENTO", d("Status") = "DUPLICADO", "indice nao oculta duplicado", CStr(d("Status"))
    sh.Cells(r, 1).Value2 = p
    Set d = repo.Resolver("AUSENTE", "EP_A2", "VID_A", 2): Check8 "F8_RETORNO_INEXISTENTE", d("Status") = "INEXISTENTE", "INEXISTENTE", CStr(d("Status"))
    Set d = repo.Resolver("PID_A2", "EP_A2", "OUTRO_VID", 2): Check8 "F8_RETORNO_AMBIGUO", d("Status") = "IDENTIDADE_AMBIGUA", "IDENTIDADE_AMBIGUA", CStr(d("Status"))
    estado = MudarCustodia8(wb, "PID_A2", "EP_A2", "VID_A", 2, "FILIAL", "PID_A2_T1")
    estado = estado & ";" & MudarCustodia8(wb, "PID_A2_T1", "EP_A2", "VID_A", 3, "MATRIZ", "PID_A2_D1")
    estado = estado & ";" & MudarCustodia8(wb, "PID_A2_D1", "EP_A2", "VID_A", 4, "FILIAL", "PID_A2_T2")
    Set d = repo.Resolver("PID_A2_T2", "EP_A2", "VID_A", 5)
    Check8 "F8_TRANSFERENCIAS_DEVOLUCAO", CStr(d("Status")) = "ENCONTRADO" And InStr(estado, "CONCLUIDA_FIXTURE;CONCLUIDA_FIXTURE;CONCLUIDA_FIXTURE") = 1, "T/D/T mesmo EP, versao5", estado & ";" & CStr(d("Status"))
    Check8 "F8_02_ANUAL_TRANSFERIDO", F8_Anual("EP_A2", eps.Range("A2:E40"), wb.Worksheets("_F8_Movimentos").Range("A2:F40")) = "NAO", "transferencia/devolucao nao herdam Venda A1", CStr(F8_Anual("EP_A2", eps.Range("A2:E40"), wb.Worksheets("_F8_Movimentos").Range("A2:F40")))
    estado = F8_Admitir(wb, "PID_A2_T2", "EP_A2", "VID_A", 5, "OP_EXPERIMENTO8", "PAYLOAD8_V1")
    wb.Worksheets("_F8_Operacoes").Range("D2").Value2 = "CONCLUIDA"
    p = F8_Admitir(wb, "PID_A2_T2", "EP_A2", "VID_A", 5, "OP_EXPERIMENTO8", "PAYLOAD8_V1")
    Check8 "F8_OPERATIONID_REPETIDO", estado = "PREPARADA" And p = "JA_CONCLUIDA" And Last8(wb.Worksheets("_F8_Operacoes")) = 2, "mesmo recibo sem duplicar intencao", estado & ";" & p
    p = F8_Admitir(wb, "PID_A2_T2", "EP_A2", "VID_A", 5, "OP_EXPERIMENTO8", "OUTRO_PAYLOAD")
    Check8 "F8_PAYLOAD_DIVERGENTE", p = "RECUSADA_PAYLOAD", "RECUSADA_PAYLOAD", p
    wb.Worksheets("_F8_Operacoes").Range("D2").Value2 = "PREPARADA"
    Set d = repo.Resolver("PID_A2_T2", "EP_A2", "VID_A", 5): Check8 "F8_RETORNO_PENDENTE", d("Status") = "OPERACAO_PENDENTE", "OPERACAO_PENDENTE", CStr(d("Status"))
    p = F8_Admitir(wb, "PID_A2_T2", "EP_A2", "VID_A", 5, "OP_OUTRA_TENTATIVA", "PAYLOAD8_V1")
    Check8 "F8_NOVO_ID_NAO_CONTORNA", p = "OPERACAO_PENDENTE", "OPERACAO_PENDENTE", p
    wb.Worksheets("_F8_Operacoes").Range("D2").Value2 = "AGUARDANDO_CONCILIACAO"
    Set d = repo.Resolver("PID_A2_T2", "EP_A2", "VID_A", 5): Check8 "F8_RETORNO_CONCILIACAO", d("Status") = "AGUARDANDO_CONCILIACAO", "AGUARDANDO_CONCILIACAO", CStr(d("Status"))
    Snapshot8 wb, "integrado_depois_r2": wb.Save: repo.Finalizar: Set repo = Nothing
    wb.Close False: Set wb = Nothing
    Check8 "F8_ESTADOS_APPLICATION", Application.EnableEvents = eventos, "estado de eventos anterior", CStr(Application.EnableEvents)
    log8.Close: Set log8 = Nothing: Debug.Print "F8_INTEGRADO_CONCLUIDO": Exit Sub
Falhou:
    If Not log8 Is Nothing Then log8.WriteLine "ERRO_INTEGRADO" & vbTab & "FAIL" & vbTab & "sem erro inesperado" & vbTab & Err.Number & ":" & Err.Description: log8.Close
    Debug.Print "F8_ERRO=" & Err.Number & ":" & Err.Description
    Application.EnableEvents = eventos
    If Not frm Is Nothing Then Unload frm
    If Not repo Is Nothing Then repo.Finalizar
    If Not wb Is Nothing Then wb.SaveCopyAs ThisWorkbook.Path & "\integrado_falha_memoria_r2.xlsm": wb.Close False
End Sub

Private Sub CasoData8(ByVal sh As Worksheet, ByVal r As Long, ByVal nome As String, ByVal valor As Variant, ByVal formato As String, ByVal politica As String)
    sh.Cells(r, 1).Value2 = nome: sh.Cells(r, 2).NumberFormat = formato: sh.Cells(r, 2).Value2 = valor
    sh.Cells(r, 3).Value2 = politica: sh.Cells(r, 4).Value2 = F8_CanonData(sh.Cells(r, 2), politica): sh.Cells(r, 5).Value2 = sh.Cells(r, 2).NumberFormat
End Sub
Public Sub F8_DatasPreparar()
    Dim wb As Workbook, sh As Worksheet, v As Double
    Call Guard8: IniciarLog8 "datas_sessao1"
    Set wb = NovaFixture8("Fixture_F8_Datas"): Set sh = Tabela8(wb, "_F8_Datas", Array("Caso", "Valor", "Policy", "CanonEsperado", "FormatoAntes"))
    v = CDbl(DateSerial(2026, 11, 2))
    CasoData8 sh, 2, "LEGADO_BR", v, "dd/mm/yyyy", "LEGACY_D14_LOCAL"
    CasoData8 sh, 3, "LEGADO_LOCAL", v, "m/d/yyyy", "LEGACY_D14_LOCAL"
    CasoData8 sh, 4, "ISO_LITERAL", v, "yyyy""-""mm""-""dd", "ISO_LITERAL"
    CasoData8 sh, 5, "ISO_ESCAPE", v, "yyyy\-mm\-dd", "ISO_LITERAL"
    CasoData8 sh, 6, "OUTRA_DATA", CDbl(DateSerial(2030, 3, 19)), "yyyy""-""mm""-""dd", "ISO_LITERAL"
    CasoData8 sh, 7, "FORMATO_MATERIAL", v, "0.00", "ISO_LITERAL"
    CasoData8 sh, 8, "DATA_INVALIDA_TEXTO", "31/02/2026", "@", "ISO_LITERAL"
    CasoData8 sh, 9, "TEXTO_PARECE_DATA", "2026-11-02", "@", "ISO_LITERAL"
    CasoData8 sh, 10, "DATA_COM_HORA", v + 0.5, "yyyy""-""mm""-""dd", "ISO_LITERAL"
    CasoData8 sh, 11, "DATA_OVERFLOW", 1E+20, "yyyy""-""mm""-""dd", "ISO_LITERAL"
    CasoData8 sh, 12, "BR_LITERAL", v, "dd""/""mm""/""yyyy", "BR_LITERAL"
    CasoData8 sh, 13, "BR_ESTRITO_NAO_LOCAL", v, "m/d/yyyy", "BR_LITERAL"
    Check8 "F8_01_FORMATOS_EQUIVALENTES", sh.Range("D2").Value2 = sh.Range("D3").Value2 And sh.Range("D4").Value2 = sh.Range("D5").Value2, "equivalencia somente pela politica explicita", CStr(sh.Range("D2").Value2) & ";" & CStr(sh.Range("D4").Value2)
    Check8 "F8_01_DATAS_DISTINTAS", sh.Range("D4").Value2 <> sh.Range("D6").Value2, "datas distintas possuem canonico distinto", CStr(sh.Range("D6").Value2)
    Check8 "F8_01_FORMATO_MATERIAL_RECUSADO", Left$(CStr(sh.Range("D7").Value2), 5) = "ERRO:", "erro de formato", CStr(sh.Range("D7").Value2)
    Check8 "F8_01_INVALIDAS_RECUSADAS", Left$(CStr(sh.Range("D8").Value2), 5) = "ERRO:" And Left$(CStr(sh.Range("D9").Value2), 5) = "ERRO:" And Left$(CStr(sh.Range("D10").Value2), 5) = "ERRO:" And Left$(CStr(sh.Range("D11").Value2), 5) = "ERRO:", "texto, hora e overflow recusados", CStr(sh.Range("D8").Value2) & ";" & CStr(sh.Range("D10").Value2)
    Check8 "F8_01_BR_ESTRITO", Left$(CStr(sh.Range("D12").Value2), 5) <> "ERRO:" And Left$(CStr(sh.Range("D13").Value2), 5) = "ERRO:", "BR literal aceito; local recusado sob BR estrito", CStr(sh.Range("D12").Value2) & ";" & CStr(sh.Range("D13").Value2)
    wb.Save: wb.Close False: log8.Close: Set log8 = Nothing: Debug.Print "F8_DATAS_PREPARADAS"
End Sub
Public Sub F8_DatasFrio()
    Dim wb As Workbook, sh As Worksheet, r As Long, atual As String, esperado As String
    Call Guard8: IniciarLog8 "datas_sessao_fria"
    Set wb = Application.Workbooks.Open(ThisWorkbook.Path & "\Fixture_F8_Datas.xlsm", UpdateLinks:=0): Set sh = wb.Worksheets("_F8_Datas")
    For r = 2 To 13
        esperado = CStr(sh.Cells(r, 4).Value2): atual = F8_CanonData(sh.Cells(r, 2), CStr(sh.Cells(r, 3).Value2))
        Check8 "F8_01_FRIO_" & CStr(sh.Cells(r, 1).Value2), atual = esperado, esperado, atual & ";FMT_ANTES=" & CStr(sh.Cells(r, 5).Value2) & ";FMT_ATUAL=" & CStr(sh.Cells(r, 2).NumberFormat)
    Next r
    wb.Close False: log8.Close: Set log8 = Nothing: Debug.Print "F8_DATAS_FRIO_CONCLUIDO"
End Sub
Public Sub F8_SaveRecusado()
    Dim wb As Workbook, repo As clsRepositorioFase8, estado As String
    Call Guard8: IniciarLog8 "save_recusado_f8"
    Set wb = NovaFixture8("Fixture_F8_SaveRecusado"): PrepararPequena8 wb: wb.Save
    estado = F8_Admitir(wb, "PID_A2", "EP_A2", "VID_A", 1, "OP_SAVE_RECUSADO", "PAYLOAD_SAVE8")
    Set repo = New clsRepositorioFase8: repo.Iniciar wb: repo.CancelarProximoSave
    wb.Save
    Check8 "F8_SAVE_REAL_CANCELADO", estado = "PREPARADA" And repo.SaveInterceptado And Not wb.Saved, "Save chamado, cancelado por evento e sem recibo", estado & ";EVENTO=" & CStr(repo.SaveInterceptado) & ";SAVED=" & CStr(wb.Saved)
    estado = F8_Admitir(wb, "PID_A2", "EP_A2", "VID_A", 1, "OP_NOVA_APOS_SAVE_INCERTO", "PAYLOAD_SAVE8")
    Check8 "F8_SAVE_PENDENCIA_NOVA_OP", estado = "OPERACAO_PENDENTE", "conciliare antes de repetir", estado
    wb.SaveCopyAs ThisWorkbook.Path & "\save_recusado_memoria_nao_promover.xlsm"
    repo.Finalizar: wb.Close False: log8.Close: Set log8 = Nothing: Debug.Print "F8_SAVE_RECUSADO_CONCLUIDO"
End Sub
Private Function Tempo8() As Double
    Dim cont As Currency, freq As Currency
    QueryPerformanceCounter cont: QueryPerformanceFrequency freq
    Tempo8 = CDbl(cont / freq)
End Function
Private Function Memoria8() As Double
    Dim m As Mem8
    m.cb = LenB(m)
    If K32GetProcessMemoryInfo(GetCurrentProcess(), m, m.cb) <> 0 Then Memoria8 = CDbl(m.WorkingSetSize) / 1048576# Else Memoria8 = -1
End Function
Private Sub Massa8(ByVal wb As Workbook, ByVal n As Long)
    Dim a() As Variant, e() As Variant, mov() As Variant, selos() As Variant, i As Long, c As Long, sh As Worksheet
    Estrutura8 wb
    ReDim a(1 To n, 1 To 18): ReDim e(1 To n, 1 To 5): ReDim mov(1 To n, 1 To 6): ReDim selos(1 To n, 1 To 2)
    For i = 1 To n
        a(i, 1) = "BENCH_PID_" & Format$(i, "00000"): a(i, 2) = "BENCH_EP_" & Format$(i, "00000"): a(i, 3) = "BENCH_VID_" & Format$(i, "00000")
        a(i, 4) = "BENCH_MOV_" & Format$(i, "00000"): a(i, 5) = "BENCH_OP_" & Format$(i, "00000"): a(i, 6) = "BENCH_ATT_" & Format$(i, "00000")
        a(i, 7) = 1: a(i, 8) = "ATIVO": a(i, 9) = IIf(i Mod 2 = 0, "FILIAL", "MATRIZ"): a(i, 10) = "FIC8" & Format$(i, "00000"): a(i, 11) = "88001": a(i, 12) = CHAVE8
        a(i, 13) = 1000# + i: a(i, 14) = 1500# + 2 * i: a(i, 15) = CDbl(DateSerial(2026, 11, 2)): a(i, 16) = CDbl(2 + i Mod 5): a(i, 17) = "ISO_LITERAL": a(i, 18) = "F8_V1"
        e(i, 1) = a(i, 2): e(i, 2) = a(i, 3): e(i, 3) = "ATIVO": e(i, 4) = a(i, 1): e(i, 5) = 1
        mov(i, 1) = a(i, 4): mov(i, 2) = a(i, 2): mov(i, 3) = "Entrada": mov(i, 4) = "CONCLUIDA": mov(i, 5) = a(i, 1): mov(i, 6) = a(i, 5)
    Next i
    Set sh = wb.Worksheets("_F8_Dados")
    For c = 1 To 18
        If c <> 7 And c <> 13 And c <> 14 And c <> 15 And c <> 16 Then sh.Cells(2, c).Resize(n, 1).NumberFormat = "@"
    Next c
    sh.Range("O2").Resize(n, 1).NumberFormat = "yyyy""-""mm""-""dd"
    sh.Range("A2").Resize(n, 18).Value2 = a: sh.Range("L2").Resize(n, 1).Errors(xlNumberAsText).Ignore = True
    wb.Worksheets("_F8_Episodios").Range("A2").Resize(n, 5).Value2 = e
    wb.Worksheets("_F8_Movimentos").Range("A2").Resize(n, 6).Value2 = mov
    For i = 1 To n: selos(i, 1) = a(i, 1): selos(i, 2) = F8_CanonRegistro(wb, i + 1): Next i
    wb.Worksheets("_F8_Selos").Range("A2").Resize(n, 2).Value2 = selos
End Sub
Public Sub F8_BenchmarkUm(ByVal n As Long)
    Dim wb As Workbook, repo As clsRepositorioFase8, d As Object, a As Variant, lista As AuditoriaListaFake, frm As Cad_NFE
    Dim i As Long, j As Long, alvo As Long, pos As Long, consultas As Long, pid As String, estado As String, tempo As Double, inicio As Double, linha As String
    Dim tSeed As Double, tIndex As Double, tLinear As Double, tProbe As Double, tVerified As Double, tLista As Double, tBind As Double, tMSForms As Double, tProjection As Double, tCalc As Double, tSave As Double, memAntes As Double, memDepois As Double
    Dim arrUI() As Variant, arrProj() As Variant, sh As Worksheet, dados As Variant, formulas As Long, faixaID As String, faixaEP As String, faixaVID As String, faixaValor As String, calcAnterior As Long
    Call Guard8
    If n <> 100 And n <> 1000 And n <> 5000 And n <> 10000 Then Err.Raise 5, , "Volume nao previsto"
    calcAnterior = Application.Calculation: memAntes = Memoria8(): inicio = Tempo8()
    On Error GoTo Falhou
    Set wb = NovaFixture8("Fixture_F8_Bench_" & n)
    tempo = Tempo8(): Massa8 wb, n: tSeed = Tempo8() - tempo
    Set repo = New clsRepositorioFase8: tempo = Tempo8(): repo.Iniciar wb: tIndex = Tempo8() - tempo
    dados = wb.Worksheets("_F8_Dados").Range("A2:R" & n + 1).Value2: consultas = 200
    tempo = Tempo8()
    For i = 1 To consultas
        alvo = 1 + ((i * 97) Mod n): pid = "BENCH_PID_" & Format$(alvo, "00000"): pos = 0
        For j = 1 To n
            If CStr(dados(j, 1)) = pid Then pos = j: Exit For
        Next j
        If pos <> alvo Then Err.Raise 5, , "Linear divergente"
    Next i
    tLinear = Tempo8() - tempo: tempo = Tempo8()
    For i = 1 To consultas
        alvo = 1 + ((i * 97) Mod n): pid = "BENCH_PID_" & Format$(alvo, "00000")
        If repo.LinhaIndice(pid) <> alvo + 1 Then Err.Raise 5, , "Indice divergente"
    Next i
    tProbe = Tempo8() - tempo: tempo = Tempo8()
    For i = 1 To consultas
        alvo = 1 + ((i * 97) Mod n): pid = "BENCH_PID_" & Format$(alvo, "00000")
        Set d = repo.Resolver(pid, "BENCH_EP_" & Format$(alvo, "00000"), "BENCH_VID_" & Format$(alvo, "00000"), 1)
        If d("Status") <> "ENCONTRADO" Or CLng(d("Row")) <> alvo + 1 Then Err.Raise 5, , "Resolucao verificada divergente"
    Next i
    tVerified = Tempo8() - tempo: tempo = Tempo8()
    estado = repo.Listar("", "OPERACIONAL", "TODOS", a): tLista = Tempo8() - tempo
    If estado <> "ENCONTRADO" Or UBound(a, 1) + 1 <> n Then Err.Raise 5, , "Lista divergente"
    Set lista = New AuditoriaListaFake: lista.ColumnCount = 8: tempo = Tempo8(): lista.List = a: tBind = Tempo8() - tempo
    ReDim arrUI(0 To n - 1, 0 To 13): ReDim arrProj(1 To n, 1 To 14)
    For i = 1 To n
        arrUI(i - 1, 0) = "_F8_Dados": arrUI(i - 1, 1) = dados(i, 11): arrUI(i - 1, 5) = dados(i, 10): arrUI(i - 1, 12) = i + 1: arrUI(i - 1, 13) = dados(i, 12)
        arrProj(i, 1) = dados(i, 1): arrProj(i, 2) = dados(i, 2): arrProj(i, 3) = dados(i, 3): arrProj(i, 4) = dados(i, 7): arrProj(i, 5) = dados(i, 10): arrProj(i, 6) = dados(i, 13): arrProj(i, 7) = dados(i, 14)
    Next i
    Set frm = New Cad_NFE: frm.ListBoxLista.ListIndex = -1: tempo = Tempo8(): frm.ListBoxLista.List = arrUI: tMSForms = Tempo8() - tempo
    If frm.ListBoxLista.ListCount <> n Or frm.ListBoxLista.ListIndex <> -1 Then Err.Raise 5, , "MSForms divergente"
    Unload frm: Set frm = Nothing
    Set sh = wb.Worksheets("_F8_Leitura"): tempo = Tempo8(): sh.Range("A2").Resize(n, 14).Value2 = arrProj: tProjection = Tempo8() - tempo
    formulas = Application.Min(n, 200): faixaID = "'_F8_Dados'!$A$2:$A$" & n + 1: faixaEP = "'_F8_Dados'!$B$2:$B$" & n + 1: faixaVID = "'_F8_Dados'!$C$2:$C$" & n + 1: faixaValor = "'_F8_Dados'!$M$2:$M$" & n + 1
    Application.Calculation = xlCalculationManual
    For i = 2 To formulas + 1
        sh.Cells(i, 6).Formula = "=IF(COUNTIF(" & faixaID & ",A" & i & ")<>1,NA(),IF(AND(INDEX(" & faixaEP & ",MATCH(A" & i & "," & faixaID & ",0))=B" & i & ",INDEX(" & faixaVID & ",MATCH(A" & i & "," & faixaID & ",0))=C" & i & "),INDEX(" & faixaValor & ",MATCH(A" & i & "," & faixaID & ",0)),NA()))"
    Next i
    tempo = Tempo8(): sh.Calculate: tCalc = Tempo8() - tempo
    If sh.Range("F2").Value2 <> 1001 Or sh.Cells(formulas + 1, 6).Value2 <> 1000 + formulas Then Err.Raise 5, , "Formula nativa divergente"
    Application.Calculation = calcAnterior
    tempo = Tempo8(): wb.Save: tSave = Tempo8() - tempo: memDepois = Memoria8()
    Set log8 = FSO8.CreateTextFile(ThisWorkbook.Path & "\benchmark_" & n & ".tsv", False, True)
    log8.WriteLine "Registros" & vbTab & "Consultas" & vbTab & "FormulasNativas" & vbTab & "Seed_s" & vbTab & "IndexBuild_s" & vbTab & "Linear200_s" & vbTab & "IndexProbe200_s" & vbTab & "Verified200_s" & vbTab & "ListaValidada_s" & vbTab & "FakeBind_s" & vbTab & "MSFormsBind_s" & vbTab & "ProjectionWrite_s" & vbTab & "RecalcSample_s" & vbTab & "Save_s" & vbTab & "WorkingSetAntes_MiB" & vbTab & "WorkingSetDepois_MiB" & vbTab & "Total_s" & vbTab & "Resultado"
    log8.WriteLine n & vbTab & consultas & vbTab & formulas & vbTab & tSeed & vbTab & tIndex & vbTab & tLinear & vbTab & tProbe & vbTab & tVerified & vbTab & tLista & vbTab & tBind & vbTab & tMSForms & vbTab & tProjection & vbTab & tCalc & vbTab & tSave & vbTab & memAntes & vbTab & memDepois & vbTab & Tempo8() - inicio & vbTab & "PASS"
    log8.Close: Set log8 = Nothing: repo.Finalizar: wb.Close False: Set wb = Nothing: Debug.Print "F8_BENCH_CONCLUIDO=" & n
    Exit Sub
Falhou:
    Debug.Print "F8_BENCH_ERRO=" & n & ";" & Err.Number & ";" & Err.Description
    Set log8 = FSO8.CreateTextFile(ThisWorkbook.Path & "\benchmark_falha_" & n & ".tsv", False, True)
    log8.WriteLine "Cenario" & vbTab & "Resultado" & vbTab & "Erro"
    log8.WriteLine "BENCH_" & n & vbTab & "FAIL" & vbTab & Err.Number & ":" & Err.Description
    log8.Close: Set log8 = Nothing
    Application.Calculation = calcAnterior
    If Not frm Is Nothing Then Unload frm
    If Not repo Is Nothing Then repo.Finalizar
    If Not wb Is Nothing Then wb.SaveCopyAs ThisWorkbook.Path & "\bench_erro_" & n & ".xlsm": wb.Close False
End Sub
Public Sub F8_BenchFrio()
    Dim n As Variant, wb As Workbook, repo As clsRepositorioFase8, d As Object, sh As Worksheet
    Call Guard8: IniciarLog8 "benchmark_frio"
    For Each n In Array(100, 1000, 5000, 10000)
        Set wb = Application.Workbooks.Open(ThisWorkbook.Path & "\Fixture_F8_Bench_" & n & ".xlsm", UpdateLinks:=0)
        Set repo = New clsRepositorioFase8: repo.Iniciar wb: Set sh = wb.Worksheets("_F8_Dados")
        Set d = repo.Resolver("BENCH_PID_" & Format$(n, "00000"), "BENCH_EP_" & Format$(n, "00000"), "BENCH_VID_" & Format$(n, "00000"), 1)
        Check8 "F8_BENCH_FRIO_" & n, d("Status") = "ENCONTRADO" And Last8(sh) - 1 = n And TypeName(sh.Cells(n + 1, 12).Value2) = "String" And sh.Cells(n + 1, 12).Value2 = CHAVE8 And sh.Cells(n + 1, 12).Errors(xlNumberAsText).Ignore, "identidade/quantidade/String44/formato/Ignore persistentes", CStr(d("Status")) & ";" & CStr(Last8(sh) - 1) & ";FMT=" & sh.Cells(n + 1, 15).NumberFormat
        repo.Finalizar: wb.Close False
    Next n
    log8.Close: Set log8 = Nothing: Debug.Print "F8_BENCH_FRIO_CONCLUIDO"
End Sub

Public Sub F8_IntegradoFrio()
    Dim wb As Workbook, repo As clsRepositorioFase8, d As Object
    Call Guard8: IniciarLog8 "integrado_frio"
    Set wb = Application.Workbooks.Open(ThisWorkbook.Path & "\Fixture_F8_Integrada_R2.xlsm", UpdateLinks:=0)
    Set repo = New clsRepositorioFase8: repo.Iniciar wb
    Set d = repo.Resolver("PID_A2_T2", "EP_A2", "VID_A", 5)
    Check8 "F8_FRIO_PENDENCIA", d("Status") = "AGUARDANDO_CONCILIACAO", "AGUARDANDO_CONCILIACAO", d("Status")
    Set d = repo.Resolver("PID_B1", "EP_B1", "VID_B", 1)
    Check8 "F8_FRIO_CONTEUDO", d("Status") = "ENCONTRADO", "ENCONTRADO sem restaurar formatos", d("Status")
    Check8 "F8_FRIO_ANUAL", F8_Anual("EP_A1", wb.Worksheets("_F8_Episodios").Range("A2:E40"), wb.Worksheets("_F8_Movimentos").Range("A2:F40")) = "SIM" And F8_Anual("EP_A2", wb.Worksheets("_F8_Episodios").Range("A2:E40"), wb.Worksheets("_F8_Movimentos").Range("A2:F40")) = "NAO", "SIM/NAO por episodio", "anual por identidade"
    repo.Finalizar: wb.Close False: log8.Close: Set log8 = Nothing: Debug.Print "F8_INTEGRADO_FRIO_CONCLUIDO"
End Sub

Private Sub OrdenarComMetadados8(ByVal sh As Worksheet, ByVal ultima As Long)
    Dim mapa As Object, r As Long, c As Long, id As String, erros() As Boolean, a As Variant
    Set mapa = CreateObject("Scripting.Dictionary")
    For r = 2 To ultima
        id = CStr(sh.Cells(r, 1).Value2)
        If id = "" Or mapa.Exists(id) Then Err.Raise 5, , "Sort exige identidade unica"
        ReDim erros(1 To 18)
        For c = 1 To 18: erros(c) = sh.Cells(r, c).Errors(xlNumberAsText).Ignore: Next c
        mapa.Add id, erros
    Next r
    sh.Range("A2:R" & ultima).Sort Key1:=sh.Range("M2"), Order1:=xlDescending, Header:=xlNo
    For r = 2 To ultima
        id = CStr(sh.Cells(r, 1).Value2)
        If Not mapa.Exists(id) Then Err.Raise 5, , "Sort alterou identidade"
        a = mapa(id)
        For c = 1 To 18: sh.Cells(r, c).Errors(xlNumberAsText).Ignore = a(c): Next c
    Next r
End Sub
