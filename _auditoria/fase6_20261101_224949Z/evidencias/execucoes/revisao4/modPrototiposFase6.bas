Attribute VB_Name = "modPrototiposFase6"
Option Explicit

' SOMENTE PROTOTIPOS SINTETICOS. Nao e gravador do produto nem validacao fiscal.
Private Type F6GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef id As F6GUID) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (ByRef id As F6GUID, ByVal texto As LongPtr, ByVal tamanho As Long) As Long
#Else
Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef id As F6GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByRef id As F6GUID, ByVal texto As Long, ByVal tamanho As Long) As Long
#End If
Private Const MARCADOR As String = "VBA_Fase6_20261101_224949Z"
Private Const CHAVE As String = "01234567890123456789012345678901234567890123"
Private gCaso As String, gPasta As String, gLog As Object, gIds As Object
Private gSeq As Long, gFalhas As Long, gRecuperou As Boolean
Private gEpM As String, gEpF As String, gOp As String
Private gEventos As Boolean, gTela As Boolean, gAlertas As Boolean, gCalculo As Long

Private Function WS(ByVal nome As String) As Worksheet
    Set WS = ThisWorkbook.Worksheets(nome)
End Function
Private Function FSO() As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
End Function
Private Sub Guardar()
    Dim raiz As String, atual As String
    If InStr(1, ThisWorkbook.Name, "Fixture_Arquitetura_Fase6_", vbTextCompare) <> 1 Then Err.Raise 5, , "Fixture Fase6 obrigatoria"
    If InStr(1, ThisWorkbook.Path, MARCADOR, vbTextCompare) = 0 Then Err.Raise 5, , "Pasta Fase6 obrigatoria"
    If InStr(1, ThisWorkbook.Path, "OneDrive", vbTextCompare) > 0 Then Err.Raise 5, , "Proibido OneDrive"
    raiz = FSO.GetFolder(Environ$("TEMP")).ShortPath & "\"
    atual = FSO.GetFolder(ThisWorkbook.Path).ShortPath & "\"
    If StrComp(Left$(atual, Len(raiz)), raiz, vbTextCompare) <> 0 Then Err.Raise 5, , "Somente TEMP"
    If WS("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then Err.Raise 5, , "Marcador sintetico ausente"
End Sub
Private Function Limpo(ByVal v As Variant) As String
    If IsError(v) Then
        Limpo = "#ERRO:" & CStr(v)
    ElseIf IsEmpty(v) Then
        Limpo = "<EMPTY>"
    Else
        Limpo = Replace(Replace(Replace(CStr(v), vbTab, "<TAB>"), vbCr, "<CR>"), vbLf, "<LF>")
    End If
End Function
Private Sub Teste(ByVal nome As String, ByVal ok As Boolean, ByVal esperado As String, ByVal observado As String)
    If Not ok Then gFalhas = gFalhas + 1
    gLog.WriteLine gCaso & vbTab & nome & vbTab & IIf(ok, "PASS", "FAIL") & vbTab & Limpo(esperado) & vbTab & Limpo(observado)
End Sub
Private Sub Exportar(ByVal etapa As String)
    Dim arq As Object, sh As Worksheet, c As Range, limite As Long
    Set arq = FSO.CreateTextFile(gPasta & "\" & gCaso & "_" & etapa & ".tsv", False, True)
    arq.WriteLine "Planilha" & vbTab & "Celula" & vbTab & "Tipo" & vbTab & "Valor" & vbTab & "Formula" & vbTab & "NumberFormat" & vbTab & "Ignore" & vbTab & "Protegida"
    Application.Calculate
    For Each sh In ThisWorkbook.Worksheets
        limite = 40
        If Left$(sh.Name, 4) = "_F6_" Then limite = 100
        For Each c In sh.Range("A1:X" & limite).Cells
            arq.WriteLine sh.Name & vbTab & c.Address(False, False) & vbTab & TypeName(c.Value2) & vbTab & Limpo(c.Value2) & vbTab & Limpo(c.Formula) & vbTab & c.NumberFormat & vbTab & IgnoreSeguro(c) & vbTab & CStr(sh.ProtectContents)
        Next c
    Next sh
    arq.Close
End Sub
Private Function AssinaturaDados() As String
    Dim nome As Variant, c As Range, s As String
    For Each nome In Array("Entradas", "Entradas Anual", "Entradas Filial", "Entradas Anual Filial", "Saidas", "Todas Saidas", "Saidas Filial", "Todas Saidas Filial", "Demonstração")
        For Each c In WS(CStr(nome)).Range("A1:X40").Cells
            s = s & TypeName(c.Value2) & ":" & Limpo(c.Value2) & ":" & Limpo(c.Formula) & ":" & c.NumberFormat & ":" & IgnoreSeguro(c) & vbLf
        Next c
    Next nome
    AssinaturaDados = s
End Function
Private Function NovaID(Optional ByVal forcada As String = "") As String
    Dim id As F6GUID, b As String, s As String, n As Variant
    If gIds Is Nothing Then Set gIds = CreateObject("Scripting.Dictionary")
    If forcada <> "" Then
        s = forcada
    Else
        b = String$(39, vbNullChar)
        If CoCreateGuid(id) <> 0 Then Err.Raise 5, , "Geracao GUID falhou"
        If StringFromGUID2(id, StrPtr(b), 39) <> 39 Then Err.Raise 5, , "Conversao GUID falhou"
        s = UCase$(Left$(b, 38))
    End If
    If gIds.Exists(s) Then Exit Function
    For Each n In Array("_F6_Episodios", "_F6_Veiculos", "_F6_Operacoes", "_F6_Vinculos")
        If Application.CountIf(WS(CStr(n)).UsedRange, s) > 0 Then Exit Function
    Next n
    gIds.Add s, True
    NovaID = s
End Function
Private Sub CriarTabela(ByVal nome As String, ByVal cabecalho As Variant)
    Dim sh As Worksheet, i As Long
    Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    sh.Name = nome
    For i = LBound(cabecalho) To UBound(cabecalho): sh.Cells(1, i + 1).Value2 = cabecalho(i): Next i
End Sub
Private Function Ultima(ByVal sh As Worksheet, Optional ByVal coluna As Long = 1) As Long
    Ultima = sh.Cells(sh.Rows.Count, coluna).End(xlUp).Row
End Function
Private Function Unica(ByVal sh As Worksheet, ByVal coluna As Long, ByVal id As String) As Long
    Dim i As Long, qtd As Long, linha As Long
    If id = "" Then Exit Function
    For i = 2 To Ultima(sh, coluna)
        If CStr(sh.Cells(i, coluna).Value2) = id Then qtd = qtd + 1: linha = i
    Next i
    If qtd = 1 Then Unica = linha
End Function
Private Sub Campos(ByVal sh As Worksheet, ByVal r As Long, ByVal nf As Long, ByVal placa As String, ByVal valor As Double)
    sh.Cells(r, 1).Value2 = nf: sh.Cells(r, 2).Value2 = "FICTICIO"
    sh.Cells(r, 3).Value2 = "SINTETICO": sh.Cells(r, 4).Value2 = "VEICULO " & placa
    sh.Cells(r, 5).Value2 = "F6_" & placa: sh.Cells(r, 6).Value2 = placa
    sh.Cells(r, 7).Value2 = "2026": sh.Cells(r, 8).Value2 = "BRANCO"
    sh.Cells(r, 9).Value2 = valor: sh.Cells(r, 10).Value2 = CDbl(DateSerial(2026, 11, 1))
    sh.Cells(r, 10).NumberFormat = "dd/mm/yyyy"
End Sub
Private Function Vincular(ByVal sh As Worksheet, ByVal r As Long, ByVal ep As String, ByVal vid As String, ByVal mov As String, ByVal op As String, ByVal papel As String) As String
    Dim pid As String, l As Long
    pid = NovaID(): l = Ultima(WS("_F6_Vinculos")) + 1
    sh.Cells(r, 21).Value2 = ep: sh.Cells(r, 22).Value2 = vid
    sh.Cells(r, 23).Value2 = mov: sh.Cells(r, 24).Value2 = pid
    With WS("_F6_Vinculos")
        .Cells(l, 1).Value2 = pid: .Cells(l, 2).Value2 = ep: .Cells(l, 3).Value2 = vid
        .Cells(l, 4).Value2 = mov: .Cells(l, 5).Value2 = op: .Cells(l, 6).Value2 = sh.Name
        .Cells(l, 7).Value2 = r: .Cells(l, 8).Value2 = papel
        .Cells(l, 9).Value2 = SeloConteudo6(sh, r)
    End With
    Vincular = pid
End Function
Private Function CriarEpisodio(ByVal localEstoque As String, ByVal r As Long, ByVal placa As String, ByVal nf As Long, Optional ByVal vid As String = "", Optional ByVal anterior As String = "") As String
    Dim ep As String, mov As String, origem As String, anual As String, pid As String, i As Long
    If vid = "" Then
        vid = NovaID(): i = Ultima(WS("_F6_Veiculos")) + 1
        WS("_F6_Veiculos").Cells(i, 1).Value2 = vid: WS("_F6_Veiculos").Cells(i, 2).Value2 = placa
        WS("_F6_Veiculos").Cells(i, 3).Value2 = "IDENTIDADE DECLARADA SOMENTE NO MODELO SINTETICO"
    End If
    ep = NovaID(): mov = NovaID()
    origem = IIf(localEstoque = "MATRIZ", "Entradas", "Entradas Filial")
    anual = IIf(localEstoque = "MATRIZ", "Entradas Anual", "Entradas Anual Filial")
    Campos WS(origem), r, nf, placa, 900
    With WS(origem)
        .Cells(r, 12).NumberFormat = "@": .Cells(r, 12).Value2 = CHAVE
        .Cells(r, 12).Errors(xlNumberAsText).Ignore = True
        .Cells(r, 13).Formula = "=I" & r & "*2"
        .Cells(r, 14).Formula = "=A" & r
        .Cells(r, 15).Value2 = placa & "_AUX"
    End With
    Campos WS(anual), r, nf, placa, 900
    pid = Vincular(WS(origem), r, ep, vid, mov, "SEMENTE", "CUSTODIA")
    Vincular WS(anual), r, ep, vid, mov, "SEMENTE", "HISTORICO"
    i = Ultima(WS("_F6_Episodios")) + 1
    With WS("_F6_Episodios")
        .Cells(i, 1).Value2 = ep: .Cells(i, 2).Value2 = vid: .Cells(i, 3).Value2 = "ATIVO"
        .Cells(i, 4).Value2 = localEstoque: .Cells(i, 5).Value2 = pid: .Cells(i, 6).Value2 = 1
        .Cells(i, 7).Value2 = anterior: .Cells(i, 8).Value2 = "SEMENTE"
    End With
    CriarEpisodio = ep
End Function
Private Sub Preparar()
    Dim n As Variant, i As Long, p As String, pid As String
    Guardar
    For Each n In Array("Entradas", "Entradas Anual", "Entradas Filial", "Entradas Anual Filial", "Saidas", "Todas Saidas", "Saidas Filial", "Todas Saidas Filial", "Demonstração")
        WS(CStr(n)).Range("A1:X40").Clear
    Next n
    CriarTabela "_F6_Episodios", Array("EpisodeID", "VehicleID", "Estado", "Local", "ProjectionIDAtual", "Versao", "EpisodioAnterior", "CriadoPor")
    CriarTabela "_F6_Veiculos", Array("VehicleID", "PlacaSintetica", "Proveniencia")
    CriarTabela "_F6_Operacoes", Array("OperationID", "MovementID", "EpisodeID", "PayloadCanonico", "Estado", "VersaoEsperada", "AttemptID", "Erro", "Tipo", "IntentID")
    CriarTabela "_F6_Vinculos", Array("ProjectionID", "EpisodeID", "VehicleID", "MovementID", "OperationID", "Planilha", "LinhaDica", "Papel", "SeloConteudo")
    CriarTabela "_F6_Diario", Array("Seq", "OperationID", "AttemptID", "Etapa", "Resultado", "HoraLocal", "Erro")
    Set gIds = CreateObject("Scripting.Dictionary")
    For i = 0 To 2
        p = CriarEpisodio("MATRIZ", 4 + i, "FIC6" & Chr(65 + i) & "01", 11001 + i)
        If i = 0 Then gEpM = p
        p = CriarEpisodio("FILIAL", 4 + i, "FIC6" & Chr(68 + i) & "01", 21001 + i)
        If i = 0 Then gEpF = p
    Next i
    pid = CStr(WS("_F6_Episodios").Cells(Unica(WS("_F6_Episodios"), 1, gEpM), 5).Value2)
    With WS("Demonstração")
        .Range("Q4").Formula = "='Entradas'!A4": .Range("R4").Formula = "='Entradas'!F4"
        .Range("S4").Formula = "=INDIRECT(""'Entradas'!F4"")"
        .Range("T4").Formula = "=INDEX('Entradas'!F:F,MATCH(""" & pid & """,'Entradas'!X:X,0))"
        .Range("Q5").Formula = "='Entradas Filial'!A4": .Range("R5").Formula = "='Entradas Filial'!F4"
        .Range("S5").Formula = "=INDIRECT(""'Entradas Filial'!F4"")"
        pid = CStr(WS("_F6_Episodios").Cells(Unica(WS("_F6_Episodios"), 1, gEpF), 5).Value2)
        .Range("T5").Formula = "=INDEX('Entradas Filial'!F:F,MATCH(""" & pid & """,'Entradas Filial'!X:X,0))"
    End With
    WS("Entradas Filial").Range("M7").Formula = "=I7*2"
    WS("Entradas Filial").Range("N7").Formula = "=A7"
    WS("Entradas Filial").Range("O7").Value2 = "DESTINO_SENTINELA"
End Sub
Private Sub Diario(ByVal op As String, ByVal tentativa As String, ByVal etapa As String, ByVal resultado As String, Optional ByVal erro As Long = 0)
    Dim i As Long, t As Object
    gSeq = gSeq + 1: i = Ultima(WS("_F6_Diario")) + 1
    With WS("_F6_Diario")
        .Cells(i, 1).Value2 = gSeq: .Cells(i, 2).Value2 = op: .Cells(i, 3).Value2 = tentativa
        .Cells(i, 4).Value2 = etapa: .Cells(i, 5).Value2 = resultado: .Cells(i, 6).Value2 = CStr(Now): .Cells(i, 7).Value2 = erro
    End With
    Set t = FSO.OpenTextFile(ThisWorkbook.Path & "\diario-operacoes.tsv", 8, True, -1)
    t.WriteLine gSeq & vbTab & op & vbTab & tentativa & vbTab & etapa & vbTab & resultado & vbTab & CStr(erro) & vbTab & CStr(WS("_F6_Operacoes").Cells(Unica(WS("_F6_Operacoes"), 1, op), 3).Value2)
    t.Close
End Sub
Private Function EstadoDuravel(ByVal op As String) As String
    Dim t As Object, linha As String, partes As Variant
    If Not FSO.FileExists(ThisWorkbook.Path & "\diario-operacoes.tsv") Then Exit Function
    Set t = FSO.OpenTextFile(ThisWorkbook.Path & "\diario-operacoes.tsv", 1, False, -1)
    Do Until t.AtEndOfStream
        linha = t.ReadLine: partes = Split(linha, vbTab)
        If UBound(partes) >= 4 Then
            If partes(1) = op Then EstadoDuravel = partes(4)
        End If
    Loop
    t.Close
End Function
Private Function Pendente(ByVal ep As String) As Boolean
    Dim i As Long, estado As String
    For i = 2 To Ultima(WS("_F6_Operacoes"))
        If WS("_F6_Operacoes").Cells(i, 3).Value2 = ep Then
            estado = EstadoDuravel(CStr(WS("_F6_Operacoes").Cells(i, 1).Value2))
            If estado <> "CONCLUIDA" And estado <> "COMPENSADA" Then Pendente = True: Exit Function
        End If
    Next i
End Function
Private Function ResolverOrigem(ByVal ep As String, ByRef sh As Worksheet, ByRef linha As Long) As Boolean
    Dim e As Long, p As Long, v As Long, pid As String, vid As String, i As Long, ativos As Long
    e = Unica(WS("_F6_Episodios"), 1, ep): If e = 0 Then Exit Function
    If WS("_F6_Episodios").Cells(e, 3).Value2 <> "ATIVO" Then Exit Function
    pid = CStr(WS("_F6_Episodios").Cells(e, 5).Value2)
    vid = CStr(WS("_F6_Episodios").Cells(e, 2).Value2)
    For i = 2 To Ultima(WS("_F6_Episodios"))
        If WS("_F6_Episodios").Cells(i, 2).Value2 = vid And WS("_F6_Episodios").Cells(i, 3).Value2 = "ATIVO" Then ativos = ativos + 1
    Next i
    If ativos <> 1 Then Exit Function
    p = Unica(WS("_F6_Vinculos"), 1, pid): If p = 0 Then Exit Function
    Set sh = WS(CStr(WS("_F6_Vinculos").Cells(p, 6).Value2))
    linha = Unica(sh, 24, pid): If linha = 0 Then Exit Function
    If sh.Cells(linha, 21).Value2 <> ep Or sh.Cells(linha, 22).Value2 <> vid Then Exit Function
    v = Unica(WS("_F6_Veiculos"), 1, vid): If v = 0 Then Exit Function
    If sh.Cells(linha, 6).Value2 <> WS("_F6_Veiculos").Cells(v, 2).Value2 Then Exit Function
    If Not VinculosConferem6(ep) Then Exit Function
    ResolverOrigem = True
End Function
Private Function Payload(ByVal ep As String, ByVal tipo As String, ByVal nf As Long, ByVal versao As Long, ByVal chaveNF As String) As String
    ' Canonico por comprimento; nao e identidade fiscal nem hash de todos os campos.
    Payload = Len(ep) & ":" & ep & Len(tipo) & ":" & tipo & ";NF=" & nf & ";VERSAO=" & versao & ";CHAVE=" & Len(chaveNF) & ":" & chaveNF
End Function
Private Function RodarOperacao(ByVal estrategia As String, ByVal tipo As String, ByVal ep As String, ByVal op As String, Optional ByVal falha As String = "", Optional ByVal nf As Long = 22001, Optional ByVal chaveNF As String = CHAVE, Optional ByVal versao As Long = 1) As String
    Dim src As Worksheet, dest As Worksheet, e As Long, r As Long, i As Long, j As Long, ro As Long, rl As Long
    Dim nomes As Variant, linhas(0 To 3) As Long, origem As Long, pid As String, mov As String, attempt As String, vid As String
    Dim ctx As clsOperacaoPlanilha, diag As String, canon As String, estado As String, nr As Long, descricao As String
    Dim salvando As Boolean, estrutura As Boolean, qtd As Long, parcelas As Long
    Call Guardar: gRecuperou = False
    On Error GoTo Falhou
    canon = Payload(ep, tipo, nf, versao, chaveNF)
    If Not DiarioCoerente6() Then RodarOperacao = "CONCILIAR_DIARIO": Exit Function
    ro = Unica(WS("_F6_Operacoes"), 1, op)
    If ro <> 0 Then
        If WS("_F6_Operacoes").Cells(ro, 4).Value2 <> canon Then RodarOperacao = "RECUSADA_PAYLOAD": Exit Function
        estado = EstadoDuravel(op)
        If estado = "CONCLUIDA" Then
            If Not VinculosConferem6(ep) Then
                WS("_F6_Operacoes").Cells(ro, 5).Value2 = "PENDENTE_CONCILIACAO"
                Diario op, CStr(WS("_F6_Operacoes").Cells(ro, 7).Value2), "DIVERGENCIA_POS_CONCLUSAO", "PENDENTE_CONCILIACAO"
                RodarOperacao = "CONCILIAR": Exit Function
            End If
            RodarOperacao = "JA_CONCLUIDA": Exit Function
        End If
        If estado <> "COMPENSADA" Then RodarOperacao = "CONCILIAR": Exit Function
    End If
    If Pendente(ep) Then RodarOperacao = "CONCILIAR": Exit Function
    If Not ResolverOrigem(ep, src, origem) Then RodarOperacao = "RECUSADA_IDENTIDADE": Exit Function
    e = Unica(WS("_F6_Episodios"), 1, ep)
    If WS("_F6_Episodios").Cells(e, 6).Value2 <> versao Then RodarOperacao = "RECUSADA_VERSAO": Exit Function
    If tipo = "Devolucao" Then
        If src.Name <> "Entradas Filial" Then RodarOperacao = "RECUSADA_LOCAL": Exit Function
        nomes = Array("Saidas Filial", "Todas Saidas Filial", "Entradas", "Entradas Anual")
    ElseIf tipo = "Transferencia" Then
        If src.Name <> "Entradas" Then RodarOperacao = "RECUSADA_LOCAL": Exit Function
        nomes = Array("Saidas", "Entradas Filial", "Entradas Anual Filial", "Todas Saidas")
    Else
        If src.Name <> "Entradas" Then RodarOperacao = "RECUSADA_LOCAL": Exit Function
        nomes = Array("Saidas", "Todas Saidas")
    End If
    If Len(chaveNF) <> 0 And (Len(chaveNF) <> 44 Or chaveNF Like "*[!0-9]*") Then RodarOperacao = "RECUSADA_CHAVE": Exit Function
    If src.ProtectContents Then RodarOperacao = "RECUSADA_PROTECAO": Exit Function
    For i = 0 To UBound(nomes)
        Set dest = WS(CStr(nomes(i)))
        If dest.ProtectContents Then RodarOperacao = "RECUSADA_PROTECAO": Exit Function
        linhas(i) = Application.Max(2, Ultima(dest) + 1)
        If InStr(1, dest.Name, "Entradas", vbTextCompare) > 0 Then linhas(i) = Application.Max(4, linhas(i))
    Next i
    If ro = 0 Then
        ro = Ultima(WS("_F6_Operacoes")) + 1: mov = NovaID()
    Else
        mov = CStr(WS("_F6_Operacoes").Cells(ro, 2).Value2)
    End If
    attempt = NovaID(): vid = CStr(WS("_F6_Episodios").Cells(e, 2).Value2)
    With WS("_F6_Operacoes")
        .Cells(ro, 1).Value2 = op: .Cells(ro, 2).Value2 = mov: .Cells(ro, 3).Value2 = ep
        .Cells(ro, 4).Value2 = canon: .Cells(ro, 5).Value2 = "PREPARADA": .Cells(ro, 6).Value2 = versao
        .Cells(ro, 7).Value2 = attempt: .Cells(ro, 8).Value2 = "": .Cells(ro, 9).Value2 = tipo: .Cells(ro, 10).Value2 = op
    End With
    Diario op, attempt, "PREPARADA", "PREPARADA"
    ThisWorkbook.Save ' Intencao registrada antes de dados; nao e commit de movimentacao.
    Set ctx = New clsOperacaoPlanilha: ctx.Iniciar ThisWorkbook
    ctx.RegistrarDestino WS("_F6_Episodios").Cells(e, 3).Resize(1, 4)
    ctx.RegistrarDestino src.Cells(origem, 1).Resize(1, 24)
    rl = Ultima(WS("_F6_Vinculos")) + 1
    ctx.RegistrarDestino WS("_F6_Vinculos").Cells(rl, 1).Resize(UBound(nomes) + 1, 9)
    For i = 0 To UBound(nomes): ctx.RegistrarDestino WS(CStr(nomes(i))).Cells(linhas(i), 1).Resize(1, 24): Next i
    If falha = "ANTES" Then Err.Raise vbObjectError + 601, , "Falha de teste antes da primeira escrita"
    For i = 0 To UBound(nomes)
        Set dest = WS(CStr(nomes(i))): r = linhas(i)
        For j = 1 To 10: dest.Cells(r, j).Value2 = src.Cells(origem, j).Value2: Next j
        dest.Cells(r, 1).Value2 = nf: dest.Cells(r, 2).Value2 = tipo
        If InStr(1, dest.Name, "Saidas", vbTextCompare) > 0 Then
            dest.Cells(r, 3).Value2 = src.Cells(origem, 1).Value2
            dest.Cells(r, 10).Value2 = 1000
            dest.Cells(r, 12).Value2 = CDbl(DateSerial(2026, 11, 1))
            dest.Cells(r, 12).NumberFormat = "dd/mm/yyyy"
        Else
            dest.Cells(r, 12).NumberFormat = "@": dest.Cells(r, 12).Value2 = CStr(chaveNF)
            dest.Cells(r, 12).Errors(xlNumberAsText).Ignore = True
        End If
        pid = Vincular(dest, r, ep, vid, mov, op, IIf(InStr(dest.Name, "Anual") > 0 Or InStr(dest.Name, "Todas") > 0, "HISTORICO", "PROJECAO"))
        If dest.Name = "Entradas" Or dest.Name = "Entradas Filial" Then WS("_F6_Episodios").Cells(e, 5).Value2 = pid
        Diario op, attempt, "GRAVACAO_" & i + 1, "EXECUTANDO"
        If falha = "GRAVACAO_" & i + 1 Then Err.Raise vbObjectError + 602, , "Falha de teste entre gravacoes"
        If falha = "COMPENSACAO" And i = 0 Then
            dest.Protect "F6QA": Err.Raise vbObjectError + 603, , "Falha de teste com destino protegido"
        End If
    Next i
    For i = 0 To UBound(nomes)
        Set dest = WS(CStr(nomes(i))): r = linhas(i)
        If dest.Cells(r, 21).Value2 <> ep Or dest.Cells(r, 22).Value2 <> vid Or dest.Cells(r, 1).Value2 <> nf Then Err.Raise 5, , "Verificacao de destino falhou"
    Next i
    If falha = "ANTES_BAIXA" Then Err.Raise vbObjectError + 604, , "Falha antes da baixa"
    If falha = "BAIXA" Then
        If estrategia = "A" Then src.Protect "F6QA" Else WS("_F6_Episodios").Protect "F6QA"
    End If
    If tipo <> "Cancelado" Then
        If estrategia = "A" Then
            ctx.MarcarBaixaIniciada: estrutura = True
            src.Cells(origem, 1).Resize(1, IIf(tipo = "Devolucao", 13, 14)).Delete Shift:=xlUp
        ElseIf estrategia = "B" Then
            src.Cells(origem, 1).Resize(1, 12).ClearContents
        End If
        WS("_F6_Episodios").Cells(e, 3).Value2 = IIf(tipo = "Venda", "ENCERRADO", "ATIVO")
        If tipo = "Transferencia" Then WS("_F6_Episodios").Cells(e, 4).Value2 = "FILIAL"
        If tipo = "Devolucao" Then WS("_F6_Episodios").Cells(e, 4).Value2 = "MATRIZ"
    End If
    WS("_F6_Episodios").Cells(e, 6).Value2 = versao + 1
    Diario op, attempt, "BAIXA", "EXECUTANDO"
    If falha = "ANTES_SAVE" Then Err.Raise vbObjectError + 605, , "Falha antes do Save"
    WS("_F6_Operacoes").Cells(ro, 5).Value2 = "CONCLUIDA"
    Diario op, attempt, "SAVE_BEGIN", "AGUARDANDO_CONFIRMACAO"
    ctx.MarcarSalvamentoIniciado: salvando = True
    If falha = "SAVE" Then
        ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\pai_inexistente\persistencia.xlsm"
    Else
        ThisWorkbook.Save
    End If
    If falha = "APOS_SAVE" Then Err.Raise vbObjectError + 606, , "Falha apos Save antes do recibo externo"
    Diario op, attempt, "SAVE_OK", "CONCLUIDA"
    ctx.Confirmar
    RodarOperacao = "CONCLUIDA"
    Exit Function
Falhou:
    nr = Err.Number: descricao = Err.Description
    If Not ctx Is Nothing Then
        If Not salvando And Not estrutura Then gRecuperou = ctx.Recuperar(diag)
        ctx.Finalizar
    End If
    estado = IIf(gRecuperou, "COMPENSADA", "PENDENTE_CONCILIACAO")
    If ro > 0 Then
        WS("_F6_Operacoes").Cells(ro, 5).Value2 = estado
        WS("_F6_Operacoes").Cells(ro, 8).Value2 = CStr(nr) & ":" & descricao & ":" & diag
    End If
    Diario op, attempt, "ERRO", estado, nr
    If gRecuperou Then ThisWorkbook.Save
    RodarOperacao = estado
End Function
Private Function Ativos(ByVal localEstoque As String) As Long
    Dim i As Long
    For i = 2 To Ultima(WS("_F6_Episodios"))
        If WS("_F6_Episodios").Cells(i, 3).Value2 = "ATIVO" Then
            If localEstoque = "TODOS" Or WS("_F6_Episodios").Cells(i, 4).Value2 = localEstoque Then Ativos = Ativos + 1
        End If
    Next i
End Function
Private Function Sentinelas(ByVal filial As Boolean) As Boolean
    Dim r As Long, placa As String
    r = IIf(filial, 5, 4): placa = IIf(filial, "FIC6D01", "FIC6A01")
    Application.Calculate
    With WS("Demonstração")
        If IsError(.Cells(r, 17).Value2) Or IsError(.Cells(r, 18).Value2) Or IsError(.Cells(r, 20).Value2) Then Exit Function
        Sentinelas = (.Cells(r, 18).Value2 = placa And .Cells(r, 19).Value2 = placa And .Cells(r, 20).Value2 = placa)
    End With
End Function
Private Function ContarMovimento(ByVal op As String) As Long
    ContarMovimento = Application.CountIf(WS("_F6_Vinculos").Columns(5), op)
End Function
Private Sub Reusar(ByVal filial As Boolean)
    Dim sh As Worksheet, p As String
    Set sh = WS(IIf(filial, "Entradas Filial", "Entradas"))
    Campos sh, 4, 33001, "FIC6G01", 700
    sh.Range("U4:X4").Value2 = "NOVO_REGISTRO_EXPERIMENTAL"
End Sub
Public Function F6_Casos() As Variant
    F6_Casos = Array("A_V_VALIDO", "A_T_VALIDO", "A_D_VALIDO", "A_D_FALHA_BAIXA", "A_T_DUP_DESTINO", "B_V_VALIDO", "B_T_VALIDO", "B_D_VALIDO", "B_REUSO", "B_LACUNAS", "C_V_VALIDO", "C_T_VALIDO", "C_D_VALIDO", "C_T_DUP_DESTINO", "C_PROJECAO_DUPLICADA", "C_HISTORICO_AMBIGUO", "C_REENTRADA", "C_TRANSFERENCIAS", "C_CHAVE_AUSENTE", "C_MESMA_OP", "C_PAYLOAD_DIVERGENTE", "C_FALHA_ANTES", "C_FALHA_1", "C_FALHA_2", "C_FALHA_3", "C_FALHA_4", "C_ANTES_BAIXA", "C_FALHA_BAIXA", "C_ANTES_SAVE", "C_FALHA_SAVE", "C_APOS_SAVE", "C_FALHA_COMPENSACAO", "C_REPETIR_RECUPERADA", "C_COLISAO", "C_REORDENACAO", "C_CANCELADO", "C_LEITOR_LEGADO", "C_POSHIST_MUTACAO", "C_MUTACAO_ORIGEM", "C_DIARIO_ORFAO", "C_DIARIO_TRUNCADO", "C_VERSAO_OBSOLETA", "C_REORDENACAO_TIPOS")
End Function
Public Sub F6_RodarLote(ByVal primeiro As Long, ByVal ultimo As Long)
    Dim casos As Variant, i As Long, pasta As String, arq As String, livro As Workbook
    Call Guardar: casos = F6_Casos()
    For i = primeiro To ultimo
        If i < 0 Or i > UBound(casos) Then Err.Raise 5
        pasta = ThisWorkbook.Path & "\" & casos(i)
        If FSO.FolderExists(pasta) Then Err.Raise 5, , "Cenario existente; preservar e criar outra execucao"
        FSO.CreateFolder pasta
        arq = pasta & "\Fixture_Arquitetura_Fase6_" & casos(i) & ".xlsm"
        ThisWorkbook.SaveCopyAs arq
        Set livro = Workbooks.Open(arq)
        Application.Run "'" & livro.Name & "'!F6_Caso", CStr(casos(i))
        livro.Close False
        Set livro = Workbooks.Open(arq)
        Application.Run "'" & livro.Name & "'!F6_VerificarReabertura", CStr(casos(i))
        livro.Close False
        Application.StatusBar = "Fase6: " & casos(i) & " executado"
    Next i
    Application.StatusBar = False
    Debug.Print "F6_LOTE_CONCLUIDO="; primeiro; ":"; ultimo
End Sub
Public Sub F6_Caso(ByVal caso As String)
    Dim estrategia As String, tipo As String, ep As String, resultado As String, falha As String, antes As String, chaveNF As String
    Dim r As Long, i As Long, qtd As Long, pid As String, vid As String, novo As String, a As Variant, b As Variant, sh As Worksheet, linha As Long
    Dim lista As AuditoriaListaFake, arqTeste As Object
    Call Guardar: gCaso = caso: gPasta = ThisWorkbook.Path: gSeq = 0: gFalhas = 0
    gEventos = Application.EnableEvents: gTela = Application.ScreenUpdating: gAlertas = Application.DisplayAlerts: gCalculo = Application.Calculation
    Set gLog = FSO.CreateTextFile(gPasta & "\" & caso & "_testes.tsv", False, True)
    gLog.WriteLine "Cenario" & vbTab & "Teste" & vbTab & "Resultado" & vbTab & "Esperado" & vbTab & "Observado"
    On Error GoTo Falhou
    Preparar
    estrategia = Left$(caso, 1): tipo = "Venda": ep = gEpM: chaveNF = CHAVE
    If InStr(caso, "_T_") > 0 Or InStr(caso, "FALHA_") > 0 Or InStr(caso, "ANTES_") > 0 Then tipo = "Transferencia"
    If InStr(caso, "_D_") > 0 Then tipo = "Devolucao": ep = gEpF
    If caso = "C_CANCELADO" Then tipo = "Cancelado"
    If InStr(caso, "DUP_DESTINO") > 0 Then WS("Entradas Filial").Range("A4").Value2 = 22001
    If caso = "C_CHAVE_AUSENTE" Then chaveNF = "": tipo = "Transferencia"
    If caso = "C_PROJECAO_DUPLICADA" Then WS("Entradas").Range("A8:X8").Value2 = WS("Entradas").Range("A4:X4").Value2
    If caso = "C_HISTORICO_AMBIGUO" Then WS("Entradas Anual").Range("A8:X8").Value2 = WS("Entradas Anual").Range("A4:X4").Value2
    If caso = "C_MUTACAO_ORIGEM" Then WS("Entradas").Range("I4").Value2 = 1234
    If caso = "C_DIARIO_ORFAO" Or caso = "C_DIARIO_TRUNCADO" Then
        Set arqTeste = FSO.CreateTextFile(ThisWorkbook.Path & "\diario-operacoes.tsv", False, True)
        If caso = "C_DIARIO_ORFAO" Then arqTeste.WriteLine "1" & vbTab & NovaID() & vbTab & "TENTATIVA_SINTETICA" & vbTab & "PREPARADA" & vbTab & "PREPARADA" & vbTab & "0" & vbTab & gEpM Else arqTeste.WriteLine "REGISTRO_TRUNCADO"
        arqTeste.Close
    End If
    If caso = "C_VERSAO_OBSOLETA" Then tipo = "Transferencia"
    Exportar "antes": ThisWorkbook.Save
    antes = AssinaturaDados(): gOp = NovaID()
    If caso = "C_COLISAO" Then
        novo = NovaID(gEpM)
        Teste "Colisao_recusada", novo = "", "ID vazio e sem substituicao", novo
        GoTo Encerrar
    End If
    If caso = "C_HISTORICO_AMBIGUO" Then
        pid = CStr(WS("Entradas Anual").Range("X4").Value2)
        r = Unica(WS("Entradas Anual"), 24, pid)
        Teste "Edicao_ambigua_recusada", r = 0 And antes = AssinaturaDados(), "zero linhas elegiveis e sem escrita", CStr(r)
        GoTo Encerrar
    End If
    If caso = "C_REORDENACAO" Or caso = "C_REORDENACAO_TIPOS" Then
        If caso = "C_REORDENACAO_TIPOS" Then
            Call PermutarLinhasComTipos6
        Else
        a = WS("Entradas").Range("A4:X4").Formula: b = WS("Entradas").Range("A6:X6").Formula
        WS("Entradas").Range("A4:X4").Formula = b: WS("Entradas").Range("A6:X6").Formula = a
        End If
        Application.Calculate
        Teste "Leitor_por_ID_apos_reordenar", ResolverOrigem(gEpM, sh, linha) And WS("Demonstração").Range("T4").Value2 = "FIC6A01", "A resolvido na linha 6", CStr(linha)
        Teste "Referencia_fisica_nao_migrada", Sentinelas(False), "todas as referencias ainda A", Limpo(WS("Demonstração").Range("S4").Value2)
        GoTo Encerrar
    End If
    If caso = "A_D_FALHA_BAIXA" Or caso = "C_FALHA_BAIXA" Then falha = "BAIXA"
    If caso = "C_FALHA_ANTES" Then falha = "ANTES"
    For i = 1 To 4
        If caso = "C_FALHA_" & i Then falha = "GRAVACAO_" & i
    Next i
    If caso = "C_ANTES_BAIXA" Then falha = "ANTES_BAIXA"
    If caso = "C_ANTES_SAVE" Then falha = "ANTES_SAVE"
    If caso = "C_FALHA_SAVE" Then falha = "SAVE"
    If caso = "C_APOS_SAVE" Then falha = "APOS_SAVE": tipo = "Transferencia"
    If caso = "C_FALHA_COMPENSACAO" Then falha = "COMPENSACAO"
    If caso = "C_REPETIR_RECUPERADA" Then falha = "GRAVACAO_2": tipo = "Transferencia"
    resultado = RodarOperacao(estrategia, tipo, ep, gOp, falha, 22001, chaveNF)
    gLog.WriteLine gCaso & vbTab & "Estado_operacao" & vbTab & "INFO" & vbTab & "" & vbTab & resultado
    If falha <> "" Then
        If falha = "BAIXA" Or falha = "SAVE" Or falha = "APOS_SAVE" Or falha = "COMPENSACAO" Then
            Teste "Estado_inseguro_bloqueia_repeticao", resultado = "PENDENTE_CONCILIACAO", "PENDENTE_CONCILIACAO", resultado
            resultado = RodarOperacao(estrategia, tipo, ep, gOp, "", 22001, chaveNF)
            Teste "Mesma_operacao_pendente_recusada", resultado = "CONCILIAR", "CONCILIAR", resultado
            resultado = RodarOperacao(estrategia, tipo, ep, NovaID(), "", 22001, chaveNF)
            Teste "Outro_ID_nao_contorna_pendencia", resultado = "CONCILIAR", "CONCILIAR", resultado
            If falha = "BAIXA" Or falha = "COMPENSACAO" Then Teste "Recuperacao_integral", gRecuperou, "True", CStr(gRecuperou)
        Else
            Teste "Compensacao_limitada", resultado = "COMPENSADA" And gRecuperou And antes = AssinaturaDados(), "COMPENSADA e dados identicos", resultado & ";Igual=" & CStr(antes = AssinaturaDados())
            If caso = "C_REPETIR_RECUPERADA" Then
                resultado = RodarOperacao("C", tipo, ep, gOp, "", 22001, chaveNF)
                Teste "Repeticao_apos_conferencia", resultado = "CONCLUIDA" And ContarMovimento(gOp) = 4, "CONCLUIDA e quatro projecoes", resultado & ";" & ContarMovimento(gOp)
            End If
        End If
        GoTo Encerrar
    End If
    If caso = "C_DIARIO_ORFAO" Or caso = "C_DIARIO_TRUNCADO" Then
        Teste "Diario_incoerente_bloqueado", resultado = "CONCILIAR_DIARIO" And antes = AssinaturaDados(), "CONCILIAR_DIARIO sem escrita", resultado
        GoTo Encerrar
    End If
    If caso = "C_MUTACAO_ORIGEM" Then
        Teste "Mutacao_origem_detectada", resultado = "RECUSADA_IDENTIDADE" And antes = AssinaturaDados(), "RECUSADA_IDENTIDADE sem escrita", resultado
        GoTo Encerrar
    End If
    If caso = "C_PROJECAO_DUPLICADA" Then
        Teste "Projecao_duplicada_recusada", resultado = "RECUSADA_IDENTIDADE" And antes = AssinaturaDados(), "RECUSADA_IDENTIDADE sem escrita", resultado
        GoTo Encerrar
    End If
    qtd = IIf(tipo = "Venda" Or tipo = "Cancelado", 2, 4)
    Teste "Destinos_esperados", resultado = "CONCLUIDA" And ContarMovimento(gOp) = qtd, CStr(qtd), CStr(ContarMovimento(gOp))
    Teste "Vinculo_sent inela", Sentinelas(tipo = "Devolucao"), "referencias preservam veiculo original", Limpo(WS("Demonstração").Cells(IIf(tipo = "Devolucao", 5, 4), 19).Value2)
    If estrategia = "B" Then
        If caso = "B_REUSO" Or caso = "B_LACUNAS" Then
            Exportar "apos_baixa"
            If caso = "B_REUSO" Then
                Reusar False: Application.Calculate
                Teste "Reuso_sem_troca_silenciosa", Sentinelas(False), "referencias historicas A", Limpo(WS("Demonstração").Range("S4").Value2)
            Else
                r = 4: Do While Not IsEmpty(WS("Entradas").Cells(r, 1)): r = r + 1: Loop
                Teste "Leitor_legado_atravessa_lacuna", r > 6, "leitor chega aos registros B/C", CStr(r)
            End If
        End If
    End If
    If estrategia = "C" Then
        Teste "Quantidade_ativa", Ativos("TODOS") = IIf(tipo = "Venda", 5, 6), CStr(IIf(tipo = "Venda", 5, 6)), CStr(Ativos("TODOS"))
        If tipo = "Transferencia" Then
            Teste "Chave_preservada", TypeName(WS("Entradas Filial").Range("L7").Value2) = "String" And WS("Entradas Filial").Range("L7").Value2 = chaveNF And WS("Entradas Filial").Range("L7").NumberFormat = "@", "String exata e @", Limpo(WS("Entradas Filial").Range("L7").Value2)
            Teste "Auxiliares_MNO", WS("Entradas Filial").Range("M7").Formula = "=I7*2" And WS("Entradas Filial").Range("N7").Formula = "=A7" And WS("Entradas Filial").Range("O7").Value2 = "DESTINO_SENTINELA", "formulas e O preservados", Limpo(WS("Entradas Filial").Range("N7").Formula)
        End If
        If caso = "C_MESMA_OP" Or caso = "C_PAYLOAD_DIVERGENTE" Then
            antes = AssinaturaDados()
            resultado = RodarOperacao("C", tipo, ep, gOp, "", IIf(caso = "C_MESMA_OP", 22001, 22002), chaveNF)
            Teste "Idempotencia_ou_payload_recusado", resultado = IIf(caso = "C_MESMA_OP", "JA_CONCLUIDA", "RECUSADA_PAYLOAD") And antes = AssinaturaDados(), "sem nova escrita", resultado
        End If
        If caso = "C_REENTRADA" Then
            vid = CStr(WS("_F6_Episodios").Cells(Unica(WS("_F6_Episodios"), 1, gEpM), 2).Value2)
            novo = CriarEpisodio("MATRIZ", 8, "FIC6A01", 11001, vid, gEpM)
            Teste "Nova_passagem_mesmo_veiculo", novo <> gEpM And Ativos("TODOS") = 6 And WS("_F6_Episodios").Cells(Unica(WS("_F6_Episodios"), 1, gEpM), 3).Value2 = "ENCERRADO", "novo EpisodeID mesmo VehicleID e anterior encerrado", novo
            Teste "Historia_nao_ressuscitada", Sentinelas(False), "A original preservado", Limpo(WS("Demonstração").Range("T4").Value2)
        End If
        If caso = "C_TRANSFERENCIAS" Then
            ' O primeiro movimento foi venda; criar novo episodio, depois duas transferencias e uma devolucao.
            vid = CStr(WS("_F6_Episodios").Cells(Unica(WS("_F6_Episodios"), 1, gEpM), 2).Value2)
            novo = CriarEpisodio("MATRIZ", 8, "FIC6A01", 11001, vid, gEpM)
            resultado = RodarOperacao("C", "Transferencia", novo, NovaID(), "", 22001, CHAVE, 1)
            Teste "Transferencia_1", resultado = "CONCLUIDA", "CONCLUIDA", resultado
            resultado = RodarOperacao("C", "Devolucao", novo, NovaID(), "", 22001, "", 2)
            Teste "Devolucao_apos_transferencia", resultado = "CONCLUIDA", "CONCLUIDA", resultado
            resultado = RodarOperacao("C", "Transferencia", novo, NovaID(), "", 22001, CHAVE, 3)
            Teste "Transferencia_sucessiva", resultado = "CONCLUIDA" And Ativos("TODOS") = 6, "CONCLUIDA e seis ativos globais", resultado & ";" & Ativos("TODOS")
        End If
        If caso = "C_T_DUP_DESTINO" Then Teste "Numero_repetido_sem_fusao", WS("Entradas Filial").Range("A4").Value2 = 22001 And WS("Entradas Filial").Range("A7").Value2 = 22001 And WS("Entradas Filial").Range("U4").Value2 <> WS("Entradas Filial").Range("U7").Value2, "NF igual; episodios distintos", "A4/A7=" & WS("Entradas Filial").Range("A7").Value2
        If caso = "C_CANCELADO" Then Teste "Cancelado_preserva_entrada", WS("Entradas").Range("A4").Value2 = 11001 And Ativos("MATRIZ") = 3, "entrada e quantidade permanecem", CStr(Ativos("MATRIZ"))
    End If
    Call VerificarComplementos6(caso, ep)
Encerrar:
    Exportar "depois"
    Teste "Estados_Application", Application.EnableEvents = gEventos And Application.ScreenUpdating = gTela And Application.DisplayAlerts = gAlertas And Application.Calculation = gCalculo, "estados anteriores", CStr(Application.EnableEvents) & ";" & CStr(Application.ScreenUpdating) & ";" & CStr(Application.DisplayAlerts) & ";" & CStr(Application.Calculation)
    gLog.Close: Set gLog = Nothing
    ThisWorkbook.SaveCopyAs gPasta & "\estado_memoria_nao_promover.xlsm"
    Debug.Print "F6_CASO=" & gCaso & ";FALHAS=" & gFalhas
    Exit Sub
Falhou:
    If Not gLog Is Nothing Then
        gLog.WriteLine gCaso & vbTab & "Erro_inesperado" & vbTab & "FAIL" & vbTab & "execucao completa" & vbTab & Err.Number & ":" & Err.Description
        gLog.Close: Set gLog = Nothing
    End If
    Debug.Print "F6_ERRO=" & gCaso & ";" & Err.Number & ";" & Err.Description
End Sub
Public Sub F6_VerificarReabertura(ByVal caso As String)
    Dim r As Long, op As String, estado As String, resultado As String, antes As String, ep As String, tipo As String, versao As Long
    Call Guardar: gCaso = caso: gPasta = ThisWorkbook.Path
    Set gLog = FSO.CreateTextFile(gPasta & "\" & caso & "_reabertura.tsv", False, True)
    gLog.WriteLine "Cenario" & vbTab & "Teste" & vbTab & "Resultado" & vbTab & "Esperado" & vbTab & "Observado"
    For r = 2 To Ultima(WS("_F6_Operacoes"))
        op = CStr(WS("_F6_Operacoes").Cells(r, 1).Value2): estado = EstadoDuravel(op)
        If estado = "CONCLUIDA" Then
            Teste "Persistencia_conclusao", WS("_F6_Operacoes").Cells(r, 5).Value2 = "CONCLUIDA", "workbook e diario concluidos", estado
        ElseIf estado <> "COMPENSADA" Then
            ep = CStr(WS("_F6_Operacoes").Cells(r, 3).Value2)
            tipo = CStr(WS("_F6_Operacoes").Cells(r, 9).Value2)
            versao = CLng(WS("_F6_Operacoes").Cells(r, 6).Value2)
            antes = AssinaturaDados()
            resultado = RodarOperacao("C", tipo, ep, op, "", 22001, CHAVE, versao)
            Teste "Pendencia_duravel_apos_reabrir", resultado = "CONCILIAR" And antes = AssinaturaDados(), "CONCILIAR e nenhuma nova escrita", resultado & ";" & estado
        End If
    Next r
    gLog.Close: Set gLog = Nothing
End Sub

Private Function IgnoreSeguro(ByVal c As Range) As String
    On Error GoTo Falhou
    IgnoreSeguro = CStr(c.Errors(xlNumberAsText).Ignore)
    Exit Function
Falhou:
    IgnoreSeguro = "NAO_LIDO:" & CStr(Err.Number)
End Function

Private Function SeloConteudo6(ByVal sh As Worksheet, ByVal r As Long) As String
    Dim c As Range, v As String, s As String
    For Each c In sh.Cells(r, 1).Resize(1, 12).Cells
        v = TypeName(c.Value2) & ":" & Limpo(c.Value2) & ":" & Limpo(c.Formula) & ":" & c.NumberFormat & ":" & IgnoreSeguro(c)
        s = s & Len(v) & ":" & v
    Next c
    SeloConteudo6 = s
End Function
Private Function VinculosConferem6(ByVal ep As String) As Boolean
    Dim i As Long, r As Long, sh As Worksheet, qtd As Long
    On Error GoTo Invalido
    For i = 2 To Ultima(WS("_F6_Vinculos"))
        If WS("_F6_Vinculos").Cells(i, 2).Value2 = ep Then
            Set sh = WS(CStr(WS("_F6_Vinculos").Cells(i, 6).Value2))
            r = Unica(sh, 24, CStr(WS("_F6_Vinculos").Cells(i, 1).Value2))
            If r = 0 Then Exit Function
            If sh.Cells(r, 21).Value2 <> ep Or sh.Cells(r, 22).Value2 <> WS("_F6_Vinculos").Cells(i, 3).Value2 Or sh.Cells(r, 23).Value2 <> WS("_F6_Vinculos").Cells(i, 4).Value2 Then Exit Function
            If CStr(WS("_F6_Vinculos").Cells(i, 9).Value2) <> SeloConteudo6(sh, r) Then Exit Function
            qtd = qtd + 1
        End If
    Next i
    VinculosConferem6 = (qtd >= 2)
Invalido:
End Function
Private Function DiarioCoerente6() As Boolean
    Dim t As Object, linha As String, p As Variant, n As Long
    On Error GoTo Invalido
    If Not FSO.FileExists(ThisWorkbook.Path & "\diario-operacoes.tsv") Then
        DiarioCoerente6 = (Ultima(WS("_F6_Operacoes")) = 1)
        Exit Function
    End If
    Set t = FSO.OpenTextFile(ThisWorkbook.Path & "\diario-operacoes.tsv", 1, False, -1)
    Do Until t.AtEndOfStream
        linha = t.ReadLine: p = Split(linha, vbTab)
        If UBound(p) <> 6 Then GoTo Invalido
        If Unica(WS("_F6_Operacoes"), 1, CStr(p(1))) = 0 Then GoTo Invalido
        If Not IsNumeric(p(0)) Then GoTo Invalido
        If CLng(p(0)) <= n Then GoTo Invalido
        n = CLng(p(0))
    Loop
    t.Close: Set t = Nothing
    If n > gSeq Then gSeq = n
    DiarioCoerente6 = (n > 0)
    Exit Function
Invalido:
    If Not t Is Nothing Then t.Close
    DiarioCoerente6 = False
End Function

Public Sub F6_TestarGuardas()
    Dim previo As Variant, erroTeste As Long, arq As Object
    Call Guardar
    previo = ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2
    On Error GoTo Restaurar
    ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 = "MARCADOR_INVALIDO_SINTETICO"
    On Error Resume Next
    Call Guardar
    erroTeste = Err.Number
    Err.Clear
    On Error GoTo Restaurar
    ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 = previo
    Call Guardar
    Set arq = FSO.CreateTextFile(ThisWorkbook.Path & "\guardas-ambiente.tsv", False, True)
    arq.WriteLine "Cenario" & vbTab & "Teste" & vbTab & "Resultado" & vbTab & "Esperado" & vbTab & "Observado"
    arq.WriteLine "R3_GUARDAS" & vbTab & "Marcador_invalido_recusado" & vbTab & IIf(erroTeste = 5, "PASS", "FAIL") & vbTab & "5" & vbTab & CStr(erroTeste)
    arq.WriteLine "R3_GUARDAS" & vbTab & "Marcador_restaurado_aceito" & vbTab & "PASS" & vbTab & "AUDITORIA_VBA_FICTICIO" & vbTab & CStr(ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2)
    arq.Close
    Debug.Print "F6_GUARDAS_CONCLUIDAS"
    Exit Sub
Restaurar:
    ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 = previo
    Err.Raise 5, , "Falha no teste de guardas; marcador restaurado"
End Sub

Private Sub VerificarComplementos6(ByVal caso As String, ByVal ep As String)
    Dim lista As AuditoriaListaFake, antes As String, resultado As String
    Dim vid As String, novo As String, opNova As String
    If caso = "C_LEITOR_LEGADO" Then
        Set lista = New AuditoriaListaFake
        PesquisarEntradasRapido "FIC6A01", lista
        Teste "Pesquisa_legada_exclui_baixado", lista.ListCount = 0, "0 episodios ativos de A", CStr(lista.ListCount)
        Teste "Contagem_fisica_igual_ativa", Application.WorksheetFunction.Count(WS("Entradas").Range("A4:A6")) = Ativos("MATRIZ"), "somente 2 ativos matriz", CStr(Application.WorksheetFunction.Count(WS("Entradas").Range("A4:A6")))
        vid = CStr(WS("_F6_Episodios").Cells(Unica(WS("_F6_Episodios"), 1, ep), 2).Value2)
        novo = CriarEpisodio("MATRIZ", 8, "FIC6A01", 11001, vid, ep)
        WS("Entradas Anual").Range("P8").Formula = "=IF(COUNTIF('Todas Saidas'!F2:F100,F8)>=1,""SIM"",""NAO"")"
        Application.Calculate
        Teste "Formula_placa_nao_encerra_reentrada", WS("Entradas Anual").Range("P8").Value2 = "NAO", "NAO para novo episodio sem saida", CStr(WS("Entradas Anual").Range("P8").Value2)
    ElseIf caso = "C_POSHIST_MUTACAO" Then
        WS("Todas Saidas").Range("F2").Value2 = "VEICULO_DIVERGENTE_SINTETICO"
        antes = AssinaturaDados()
        resultado = RodarOperacao("C", "Venda", ep, gOp, "", 22001, CHAVE)
        Teste "Historico_mutado_exige_conciliacao", resultado = "CONCILIAR" And antes = AssinaturaDados(), "CONCILIAR sem escrita de negocio", resultado
    ElseIf caso = "C_VERSAO_OBSOLETA" Then
        antes = AssinaturaDados(): opNova = NovaID()
        resultado = RodarOperacao("C", "Devolucao", ep, opNova, "", 22002, "", 1)
        Teste "Versao_obsoleta_recusada", resultado = "RECUSADA_VERSAO" And antes = AssinaturaDados(), "RECUSADA_VERSAO sem escrita", resultado
    End If
End Sub
Private Sub PermutarLinhasComTipos6()
    Dim c As Long, a As Range, b As Range
    Dim va As Variant, vb As Variant, fa As Variant, fb As Variant
    Dim na As Variant, nb As Variant, ha As Boolean, hb As Boolean, ia As Boolean, ib As Boolean
    For c = 1 To 24
        Set a = WS("Entradas").Cells(4, c): Set b = WS("Entradas").Cells(6, c)
        va = a.Value2: vb = b.Value2: fa = a.Formula: fb = b.Formula
        na = a.NumberFormat: nb = b.NumberFormat: ha = a.HasFormula: hb = b.HasFormula
        ia = a.Errors(xlNumberAsText).Ignore: ib = b.Errors(xlNumberAsText).Ignore
        a.NumberFormat = nb: b.NumberFormat = na
        If hb Then a.Formula = fb Else a.Value2 = vb
        If ha Then b.Formula = fa Else b.Value2 = va
        a.Errors(xlNumberAsText).Ignore = ib: b.Errors(xlNumberAsText).Ignore = ia
    Next c
End Sub
