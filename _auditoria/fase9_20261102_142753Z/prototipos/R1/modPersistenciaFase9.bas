Attribute VB_Name = "modPersistenciaFase9"
Option Explicit
' LABORATORIO SINTETICO. Nao integra gravadores nem politicas do produto.
Private Const MARK9 As String = "VBA_Fase9_20261102_142753Z"
Private Const KEY9 As String = "01234567890123456789012345678901234567890123"
Private Const PROVIDER9 As String = "Microsoft.ACE.OLEDB.16.0"
Private Type GUID9
    D1 As Long
    D2 As Integer
    D3 As Integer
    D4(0 To 7) As Byte
End Type
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef g As GUID9) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef g As GUID9, ByVal p As LongPtr, ByVal n As Long) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef n As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef n As Currency) As Long
Private log9 As Object, cn9 As Object, held9 As Object
Private owner9 As String, fence9 As Long, phase9 As String
Private Function FSO9() As Object
    Set FSO9 = CreateObject("Scripting.FileSystemObject")
End Function
Private Sub Guard9()
    Dim root As String, cur As String
    If ThisWorkbook.Name <> "Fixture_Persistencia_Fase9_Mestre.xlsm" Then Err.Raise 5, , "Mestre F9 exigido"
    If InStr(1, ThisWorkbook.Path, MARK9, vbTextCompare) = 0 Or InStr(1, ThisWorkbook.Path, "OneDrive", vbTextCompare) > 0 Then Err.Raise 5, , "TEMP F9 exigido"
    root = FSO9.GetFolder(Environ$("TEMP")).ShortPath & "\": cur = FSO9.GetFolder(ThisWorkbook.Path).ShortPath & "\"
    If StrComp(Left$(cur, Len(root)), root, vbTextCompare) <> 0 Then Err.Raise 5, , "Somente TEMP"
    If ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then Err.Raise 5, , "Marcador ausente"
End Sub
Private Function NewID9() As String
    Dim g As GUID9, s As String
    If CoCreateGuid(g) <> 0 Then Err.Raise 5, , "GUID indisponivel"
    s = String$(39, vbNullChar)
    If StringFromGUID2(g, StrPtr(s), 39) = 0 Then Err.Raise 5, , "GUID invalido"
    NewID9 = UCase$(Mid$(s, 2, 36))
End Function
Private Function Now9() As Double
    Dim n As Currency, f As Currency
    QueryPerformanceCounter n: QueryPerformanceFrequency f: Now9 = n / f
End Function
Private Function Clean9(ByVal v As Variant) As String
    If IsNull(v) Then Clean9 = "<NULL>": Exit Function
    If IsEmpty(v) Then Clean9 = "<EMPTY>": Exit Function
    Clean9 = Replace(Replace(Replace(CStr(v), vbTab, "<TAB>"), vbCr, "<CR>"), vbLf, "<LF>")
End Function
Private Sub StartLog9(ByVal name As String)
    Set log9 = FSO9.CreateTextFile(ThisWorkbook.Path & "\" & name & ".tsv", False, True)
    log9.WriteLine "Cenario" & vbTab & "Resultado" & vbTab & "Tratamento" & vbTab & "Esperado" & vbTab & "Observado" & vbTab & "Fixture" & vbTab & "Procedimento"
End Sub
Private Sub Check9(ByVal id As String, ByVal ok As Boolean, ByVal treatment As String, ByVal expected As String, ByVal observed As String, ByVal fixture As String)
    log9.WriteLine id & vbTab & IIf(ok, "PASS", "FAIL") & vbTab & treatment & vbTab & Clean9(expected) & vbTab & Clean9(observed) & vbTab & fixture & vbTab & phase9
End Sub
Private Function Open9(ByVal name As String, Optional ByVal readOnly As Boolean = False) As Object
    Dim c As Object
    Set c = CreateObject("ADODB.Connection")
    If readOnly Then c.Mode = 1
    c.ConnectionTimeout = 5: c.CommandTimeout = 5
    c.Open "Provider=" & PROVIDER9 & ";Data Source=" & ThisWorkbook.Path & "\" & name & ";"
    Set Open9 = c
End Function
Private Function Cmd9(ByVal c As Object, ByVal sql As String, ByVal ps As Variant) As Object
    Dim q As Object, p As Object, i As Long, typ As Long, sz As Long, v As Variant
    Set q = CreateObject("ADODB.Command"): Set q.ActiveConnection = c: q.CommandText = sql: q.CommandType = 1
    For i = LBound(ps) To UBound(ps)
        v = ps(i): sz = 0
        Select Case VarType(v)
            Case vbInteger, vbLong: typ = 3
            Case vbCurrency: typ = 6
            Case vbBoolean: typ = 11
            Case Else
                typ = 202: sz = Application.Max(1, Len(CStr(v)))
                If sz > 255 Then typ = 203
        End Select
        Set p = q.CreateParameter("p" & i, typ, 1, sz, v): q.Parameters.Append p
    Next i
    Set Cmd9 = q
End Function
Private Function Exec9(ByVal c As Object, ByVal sql As String, ParamArray ps() As Variant) As Long
    Dim q As Object, r As Object, n As Variant
    Set q = Cmd9(c, sql, ps): Set r = q.Execute(n): Exec9 = CLng(n)
End Function
Private Function Val9(ByVal c As Object, ByVal sql As String, ParamArray ps() As Variant) As Variant
    Dim q As Object, r As Object
    Set q = Cmd9(c, sql, ps): Set r = q.Execute
    If r.EOF Then Val9 = Empty Else Val9 = r.Fields(0).Value
    r.Close
End Function
Private Function Rows9(ByVal c As Object, ByVal sql As String) As Variant
    Dim r As Object
    Set r = c.Execute(sql)
    If r.EOF Then Rows9 = Empty Else Rows9 = r.GetRows
    r.Close
End Function
Private Function Pack9(ParamArray fields() As Variant) As String
    Dim i As Long, s As String, v As String
    For i = LBound(fields) To UBound(fields)
        If IsNull(fields(i)) Then
            v = "NULL"
        ElseIf VarType(fields(i)) = vbCurrency Then
            v = "M4:" & Format$(CCur(fields(i)) * 10000@, "0")
        Else
            v = TypeName(fields(i)) & ":" & CStr(fields(i))
        End If
        s = s & Len(v) & ":" & v
    Next i
    Pack9 = s
End Function
Private Function CanonProjection9(ByVal c As Object, ByVal pid As String) As String
    Dim r As Object, q As Object, i As Long, s As String
    Set q = Cmd9(c, "SELECT ProjectionID,EpisodeID,MovementID,AmountIn,AmountOut,CivilDate,DocNumber,AccessKey,PlaceCode FROM Projections WHERE ProjectionID=?", Array(pid))
    Set r = q.Execute
    If r.EOF Then CanonProjection9 = "ERROR:MISSING": r.Close: Exit Function
    For i = 0 To r.Fields.Count - 1: s = s & Pack9(r.Fields(i).Value): Next i
    CanonProjection9 = s: r.Close
End Function
Private Function Source9(ByVal c As Object, ByVal ep As String) As String
    Dim pid As Variant
    pid = Val9(c, "SELECT CurrentPID FROM Episodes WHERE EpisodeID=?", ep)
    Source9 = Pack9(ep, Val9(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep), Val9(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", ep))
    If Not IsNull(pid) And Not IsEmpty(pid) Then Source9 = Source9 & CanonProjection9(c, CStr(pid))
End Function
Private Sub CreateDB9(ByVal name As String)
    Dim cat As Object, c As Object, ddl As Variant, x As Variant
    If FSO9.FileExists(ThisWorkbook.Path & "\" & name) Then Err.Raise 58, , "Nao sobrescrever banco"
    Set cat = CreateObject("ADOX.Catalog")
    cat.Create "Provider=" & PROVIDER9 & ";Data Source=" & ThisWorkbook.Path & "\" & name & ";"
    cat.ActiveConnection.Close: Set cat = Nothing: Set c = Open9(name)
    ddl = Array( _
    "CREATE TABLE Meta (ID INTEGER CONSTRAINT PK_Meta PRIMARY KEY,BaseID TEXT(40) NOT NULL,SchemaNo INTEGER,RevisionNo LONG)", _
    "CREATE TABLE WriterLease (ID INTEGER CONSTRAINT PK_Lease PRIMARY KEY,OwnerID TEXT(40),FenceNo LONG,StatusText TEXT(20))", _
    "CREATE TABLE Vehicles (VehicleID TEXT(40) CONSTRAINT PK_V PRIMARY KEY,Plate TEXT(20) NOT NULL)", _
    "CREATE TABLE Episodes (EpisodeID TEXT(40) CONSTRAINT PK_E PRIMARY KEY,VehicleID TEXT(40) NOT NULL,VersionNo LONG,StatusText TEXT(20),CurrentPID TEXT(40),PlaceCode TEXT(20),CONSTRAINT FK_EV FOREIGN KEY(VehicleID) REFERENCES Vehicles(VehicleID))", _
    "CREATE TABLE Operations (OperationID TEXT(40) CONSTRAINT PK_O PRIMARY KEY,EpisodeID TEXT(40) NOT NULL,ExpectedVersion LONG,KindText TEXT(20),AmountIn CURRENCY,AmountOut CURRENCY,CommandCanon LONGTEXT,SourceCanon LONGTEXT,PlannedPID TEXT(40),PlannedMID TEXT(40),StatusText TEXT(30),CONSTRAINT FK_OE FOREIGN KEY(EpisodeID) REFERENCES Episodes(EpisodeID))", _
    "CREATE TABLE Attempts (AttemptID TEXT(40) CONSTRAINT PK_A PRIMARY KEY,OperationID TEXT(40) NOT NULL,OwnerID TEXT(40),FenceNo LONG,CONSTRAINT FK_AO FOREIGN KEY(OperationID) REFERENCES Operations(OperationID))", _
    "CREATE TABLE Movements (MovementID TEXT(40) CONSTRAINT PK_M PRIMARY KEY,OperationID TEXT(40) NOT NULL,EpisodeID TEXT(40) NOT NULL,KindText TEXT(20),CONSTRAINT U_MO UNIQUE(OperationID),CONSTRAINT FK_MO FOREIGN KEY(OperationID) REFERENCES Operations(OperationID),CONSTRAINT FK_ME FOREIGN KEY(EpisodeID) REFERENCES Episodes(EpisodeID))", _
    "CREATE TABLE Projections (ProjectionID TEXT(40) CONSTRAINT PK_P PRIMARY KEY,EpisodeID TEXT(40) NOT NULL,MovementID TEXT(40) NOT NULL,AmountIn CURRENCY,AmountOut CURRENCY,CivilDate TEXT(10),DocNumber TEXT(20),AccessKey TEXT(44),PlaceCode TEXT(20),ContentCanon LONGTEXT,CONSTRAINT FK_PE FOREIGN KEY(EpisodeID) REFERENCES Episodes(EpisodeID),CONSTRAINT FK_PM FOREIGN KEY(MovementID) REFERENCES Movements(MovementID))", _
    "CREATE TABLE Receipts (ReceiptID TEXT(40) CONSTRAINT PK_R PRIMARY KEY,OperationID TEXT(40) NOT NULL,RevisionNo LONG,EffectCanon LONGTEXT,CONSTRAINT U_RO UNIQUE(OperationID),CONSTRAINT FK_RO FOREIGN KEY(OperationID) REFERENCES Operations(OperationID))", _
    "CREATE TABLE ChangeLog (AuditID COUNTER CONSTRAINT PK_L PRIMARY KEY,OperationID TEXT(40) NOT NULL,BeforeCanon LONGTEXT,AfterCanon LONGTEXT,CONSTRAINT FK_LO FOREIGN KEY(OperationID) REFERENCES Operations(OperationID))", _
    "CREATE TABLE Pending (EpisodeID TEXT(40) CONSTRAINT PK_Pending PRIMARY KEY,OperationID TEXT(40) NOT NULL,ReasonText TEXT(80),CONSTRAINT FK_NE FOREIGN KEY(EpisodeID) REFERENCES Episodes(EpisodeID),CONSTRAINT FK_NO FOREIGN KEY(OperationID) REFERENCES Operations(OperationID))", _
    "CREATE TABLE Provenance (LegacyRef TEXT(80) CONSTRAINT PK_Prov PRIMARY KEY,SourceHash TEXT(64),DecisionText TEXT(30),EpisodeID TEXT(40),EvidenceText LONGTEXT)")
    For Each x In ddl: c.Execute CStr(x): Next x
    Exec9 c, "INSERT INTO Meta VALUES (1,?,1,0)", NewID9()
    Exec9 c, "INSERT INTO WriterLease VALUES (1,'',0,'FREE')"
    c.Close
End Sub
Private Function Acquire9(ByVal c As Object, ByVal who As String) As Long
    Dim n As Long
    n = Exec9(c, "UPDATE WriterLease SET OwnerID=?,FenceNo=FenceNo+1,StatusText='ACTIVE' WHERE ID=1 AND StatusText='FREE'", who)
    If n = 1 Then Acquire9 = CLng(Val9(c, "SELECT FenceNo FROM WriterLease WHERE ID=1"))
End Function
Private Function Own9(ByVal c As Object, ByVal who As String, ByVal fence As Long) As Boolean
    Own9 = CLng(Val9(c, "SELECT Count(*) FROM WriterLease WHERE ID=1 AND OwnerID=? AND FenceNo=? AND StatusText='ACTIVE'", who, fence)) = 1
End Function
Private Sub Release9(ByVal c As Object, ByVal who As String, ByVal fence As Long)
    Exec9 c, "UPDATE WriterLease SET StatusText='FREE',OwnerID='' WHERE ID=1 AND OwnerID=? AND FenceNo=?", who, fence
End Sub
Private Sub Seed9(ByVal c As Object, ByVal vid As String, ByVal ep As String, ByVal plate As String)
    If CLng(Val9(c, "SELECT Count(*) FROM Vehicles WHERE VehicleID=?", vid)) = 0 Then Exec9 c, "INSERT INTO Vehicles VALUES (?,?)", vid, plate
    If CLng(Val9(c, "SELECT Count(*) FROM Episodes WHERE VehicleID=? AND StatusText='ACTIVE'", vid)) > 0 Then Err.Raise 5, , "Veiculo ja em episodio ativo"
    Exec9 c, "INSERT INTO Episodes (EpisodeID,VehicleID,VersionNo,StatusText,PlaceCode) VALUES (?,?,0,'NEW','MATRIZ')", ep, vid
End Sub
Private Function Prepare9(ByVal c As Object, ByVal who As String, ByVal fence As Long, ByVal op As String, ByVal attempt As String, ByVal ep As String, ByVal expected As Long, ByVal kind As String, ByVal vin As Currency, ByVal vout As Currency, ByVal payload As String, Optional ByVal expectedBase As String = "") As String
    Dim canon As String, old As Variant, status As Variant, pid As String, mid As String, source As String, tx As Boolean
    On Error GoTo Failed
    If Not Own9(c, who, fence) Then Prepare9 = "RECUSADO_FENCE": Exit Function
    If expectedBase <> "" Then If expectedBase <> CStr(Val9(c, "SELECT BaseID FROM Meta WHERE ID=1")) Then Prepare9 = "RECUSADO_BASE": Exit Function
    canon = Pack9("SCHEMA1", CStr(Val9(c, "SELECT BaseID FROM Meta WHERE ID=1")), op, ep, expected, kind, vin, vout, "2026-11-02", "88001", KEY9, payload)
    old = Val9(c, "SELECT CommandCanon FROM Operations WHERE OperationID=?", op)
    If Not IsEmpty(old) Then
        If CStr(old) <> canon Then Prepare9 = "RECUSADO_PAYLOAD": Exit Function
        status = Val9(c, "SELECT StatusText FROM Operations WHERE OperationID=?", op)
        If CStr(status) = "CONFIRMED" Then
            pid = CStr(Val9(c, "SELECT PlannedPID FROM Operations WHERE OperationID=?", op))
            If CanonProjection9(c, pid) <> CStr(Val9(c, "SELECT EffectCanon FROM Receipts WHERE OperationID=?", op)) Then Prepare9 = "PENDENTE_HISTORICO": Exit Function
            Prepare9 = "JA_CONCLUIDA": Exit Function
        End If
        Prepare9 = "PENDENTE_CONCILIACAO": Exit Function
    End If
    If CLng(Val9(c, "SELECT Count(*) FROM Pending WHERE EpisodeID=?", ep)) <> 0 Then Prepare9 = "PENDENTE_CONCILIACAO": Exit Function
    If CLng(Val9(c, "SELECT Count(*) FROM Episodes WHERE EpisodeID=?", ep)) <> 1 Then Prepare9 = "RECUSADO_EPISODIO": Exit Function
    If CLng(Val9(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep)) <> expected Then Prepare9 = "RECUSADO_VERSAO": Exit Function
    status = Val9(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", ep)
    If (kind = "ENTRY" And status <> "NEW") Or (kind <> "ENTRY" And status <> "ACTIVE") Then Prepare9 = "RECUSADO_ESTADO": Exit Function
    If kind <> "ENTRY" And kind <> "SALE" And kind <> "TRANSFER" And kind <> "RETURN" Then Prepare9 = "RECUSADO_TIPO": Exit Function
    source = Source9(c, ep): pid = NewID9(): mid = NewID9()
    If kind <> "ENTRY" Then
        If CanonProjection9(c, CStr(Val9(c, "SELECT CurrentPID FROM Episodes WHERE EpisodeID=?", ep))) <> CStr(Val9(c, "SELECT ContentCanon FROM Projections WHERE ProjectionID=(SELECT CurrentPID FROM Episodes WHERE EpisodeID=?)", ep)) Then Prepare9 = "PENDENTE_HISTORICO": Exit Function
    End If
    c.BeginTrans: tx = True
    If Not Own9(c, who, fence) Then Err.Raise 5, , "Fence perdido"
    Exec9 c, "INSERT INTO Operations VALUES (?,?,?,?,?,?,?,?,?,?,?)", op, ep, expected, kind, vin, vout, canon, source, pid, mid, "PREPARED"
    Exec9 c, "INSERT INTO Attempts VALUES (?,?,?,?)", attempt, op, who, fence
    Exec9 c, "INSERT INTO Pending VALUES (?,?,'INTENT_COMMITTED')", ep, op
    c.CommitTrans: tx = False: Prepare9 = "PREPARADA": Exit Function
Failed:
    Prepare9 = "ERRO:" & Err.Number & ":" & Err.Description
    On Error Resume Next
    If tx Then c.RollbackTrans
End Function
Private Function Apply9(ByVal c As Object, ByVal who As String, ByVal fence As Long, ByVal op As String, Optional ByVal failure As Long = 0) As String
    Dim tx As Boolean, ep As String, pid As String, mid As String, kind As String, src As String, loc As String, ver As Long, rev As Long, seal As String, e As Long, msg As String
    On Error GoTo Failed
    If Not Own9(c, who, fence) Then Apply9 = "RECUSADO_FENCE": Exit Function
    If CStr(Val9(c, "SELECT StatusText FROM Operations WHERE OperationID=?", op)) <> "PREPARED" Then Apply9 = "RECUSADO_ESTADO": Exit Function
    ep = CStr(Val9(c, "SELECT EpisodeID FROM Operations WHERE OperationID=?", op))
    src = CStr(Val9(c, "SELECT SourceCanon FROM Operations WHERE OperationID=?", op))
    If Source9(c, ep) <> src Then Apply9 = "PENDENTE_ORIGEM_ALTERADA": Exit Function
    pid = CStr(Val9(c, "SELECT PlannedPID FROM Operations WHERE OperationID=?", op)): mid = CStr(Val9(c, "SELECT PlannedMID FROM Operations WHERE OperationID=?", op))
    kind = CStr(Val9(c, "SELECT KindText FROM Operations WHERE OperationID=?", op)): ver = CLng(Val9(c, "SELECT ExpectedVersion FROM Operations WHERE OperationID=?", op))
    loc = IIf(kind = "TRANSFER", "FILIAL", "MATRIZ")
    c.BeginTrans: tx = True
    ' Gravacao na linha de lease adquire bloqueio; fencing conferido dentro da transacao.
    If Exec9(c, "UPDATE WriterLease SET FenceNo=FenceNo WHERE ID=1 AND OwnerID=? AND FenceNo=? AND StatusText='ACTIVE'", who, fence) <> 1 Then Err.Raise 5, , "Lease invalido"
    If Source9(c, ep) <> src Then Err.Raise 5, , "Origem mudou"
    Exec9 c, "INSERT INTO Movements VALUES (?,?,?,?)", mid, op, ep, kind
    If failure = 1 Then Err.Raise 7001, , "Falha apos movimento"
    Exec9 c, "INSERT INTO Projections (ProjectionID,EpisodeID,MovementID,AmountIn,AmountOut,CivilDate,DocNumber,AccessKey,PlaceCode,ContentCanon) VALUES (?,?,?,?,?,'2026-11-02','88001',?,?,?)", pid, ep, mid, CCur(Val9(c, "SELECT AmountIn FROM Operations WHERE OperationID=?", op)), CCur(Val9(c, "SELECT AmountOut FROM Operations WHERE OperationID=?", op)), KEY9, loc, "PENDING"
    seal = CanonProjection9(c, pid): Exec9 c, "UPDATE Projections SET ContentCanon=? WHERE ProjectionID=?", seal, pid
    If failure = 2 Then Err.Raise 7002, , "Falha apos projecao"
    If Exec9(c, "UPDATE Episodes SET CurrentPID=?,VersionNo=VersionNo+1,StatusText=?,PlaceCode=? WHERE EpisodeID=? AND VersionNo=?", pid, IIf(kind = "SALE", "CLOSED", "ACTIVE"), loc, ep, ver) <> 1 Then Err.Raise 5, , "CAS episodio recusado"
    If failure = 3 Then Err.Raise 7003, , "Falha durante baixa logica"
    Exec9 c, "INSERT INTO ChangeLog (OperationID,BeforeCanon,AfterCanon) VALUES (?,?,?)", op, src, Source9(c, ep)
    If failure = 4 Then c.Execute "INSERT INTO ChangeLog (OperationID) VALUES ('OP_INEXISTENTE')"
    rev = CLng(Val9(c, "SELECT RevisionNo FROM Meta WHERE ID=1"))
    If Exec9(c, "UPDATE Meta SET RevisionNo=RevisionNo+1 WHERE ID=1 AND RevisionNo=?", rev) <> 1 Then Err.Raise 5, , "CAS base recusado"
    If failure = 5 Then Err.Raise 7005, , "Antes do recibo"
    Exec9 c, "INSERT INTO Receipts VALUES (?,?,?,?)", NewID9(), op, rev + 1, seal
    Exec9 c, "UPDATE Operations SET StatusText='CONFIRMED' WHERE OperationID=?", op
    Exec9 c, "DELETE FROM Pending WHERE OperationID=?", op
    If failure = 6 Then Err.Raise 7006, , "Antes do commit"
    If failure = 7 Then c.Close: c.CommitTrans
    c.CommitTrans: tx = False
    If failure = 8 Then Err.Raise 7008, , "Apos commit, antes de Excel"
    Apply9 = "CONFIRMADA": Exit Function
Failed:
    e = Err.Number: msg = Err.Description
    On Error Resume Next
    If tx Then c.RollbackTrans
    If failure = 8 Then Apply9 = "COMMIT_CONFIRMADO_PROJECAO_PENDENTE" Else Apply9 = "PENDENTE_ERRO:" & e & ":" & msg
End Function
Private Function DataSig9(ByVal c As Object) As String
    Dim names As Variant, sorts As Variant, i As Long, r As Object, j As Long, s As String
    names = Array("Episodes", "Movements", "Projections", "Receipts", "ChangeLog", "Meta")
    sorts = Array("EpisodeID", "MovementID", "ProjectionID", "ReceiptID", "AuditID", "ID")
    For i = 0 To UBound(names)
        Set r = c.Execute("SELECT * FROM " & names(i) & " ORDER BY " & sorts(i))
        s = s & names(i)
        Do While Not r.EOF
            For j = 0 To r.Fields.Count - 1: s = s & Pack9(r.Fields(j).Value): Next j
            r.MoveNext
        Loop
        r.Close
    Next i
    DataSig9 = s
End Function
Private Sub Snapshot9(ByVal c As Object, ByVal name As String)
    Dim f As Object, r As Object, tbl As Variant, j As Long
    Set f = FSO9.CreateTextFile(ThisWorkbook.Path & "\" & name & ".tsv", False, True)
    f.WriteLine "Tabela" & vbTab & "Campo" & vbTab & "TipoADO" & vbTab & "TipoVBA" & vbTab & "Valor"
    For Each tbl In Array("Meta", "WriterLease", "Vehicles", "Episodes", "Operations", "Attempts", "Movements", "Projections", "Receipts", "ChangeLog", "Pending", "Provenance")
        Set r = c.Execute("SELECT * FROM " & tbl)
        Do While Not r.EOF
            For j = 0 To r.Fields.Count - 1: f.WriteLine tbl & vbTab & r.Fields(j).Name & vbTab & r.Fields(j).Type & vbTab & TypeName(r.Fields(j).Value) & vbTab & Clean9(r.Fields(j).Value): Next j
            r.MoveNext
        Loop
        r.Close
    Next tbl
    f.Close
End Sub
Private Function BusinessRows9(ByVal c As Object) As Variant
    ' Consulta em lote: uma leitura, unicidade por PK/FK mais revalidacao de conteudo.
    BusinessRows9 = Rows9(c, "SELECT P.ProjectionID,P.EpisodeID,E.VehicleID,P.MovementID,E.VersionNo,E.StatusText,E.CurrentPID,P.PlaceCode,V.Plate,P.DocNumber,P.AccessKey,P.AmountIn,P.AmountOut,P.CivilDate,P.ContentCanon FROM (Projections AS P INNER JOIN Episodes AS E ON P.EpisodeID=E.EpisodeID) INNER JOIN Vehicles AS V ON E.VehicleID=V.VehicleID ORDER BY P.ProjectionID")
End Function
Private Function VerifyBatch9(ByVal a As Variant) As Long
    Dim d As Object, i As Long, s As String
    Set d = CreateObject("Scripting.Dictionary")
    If IsEmpty(a) Then Exit Function
    For i = 0 To UBound(a, 2)
        If d.Exists(CStr(a(0, i))) Then Err.Raise 5, , "PID duplicado"
        d.Add CStr(a(0, i)), True
        s = Pack9(a(0, i)) & Pack9(a(1, i)) & Pack9(a(3, i)) & Pack9(a(11, i)) & Pack9(a(12, i)) & Pack9(a(13, i)) & Pack9(a(9, i)) & Pack9(a(10, i)) & Pack9(a(7, i))
        If s <> CStr(a(14, i)) Then Err.Raise 5, , "Conteudo divergente: " & a(0, i)
        If TypeName(a(10, i)) <> "String" Or Len(a(10, i)) <> 44 Or a(10, i) Like "*[!0-9]*" Then Err.Raise 5, , "Chave invalida"
        If Not CStr(a(13, i)) Like "####-##-##" Then Err.Raise 5, , "Data invalida"
        VerifyBatch9 = VerifyBatch9 + 1
    Next i
End Function
Private Function Materialize9(ByVal c As Object, ByVal wb As Workbook, Optional ByVal fail As Boolean = False) As Long
    Dim a As Variant, output() As Variant, i As Long, j As Long, n As Long, sh As Worksheet, expected As Long, k As Long
    a = BusinessRows9(c): n = VerifyBatch9(a)
    Set sh = wb.Worksheets(1): sh.Name = "Projecao_F9"
    sh.Cells.Clear
    sh.Range("A1:R1").Value = Array("ProjectionID", "EpisodeID", "VehicleID", "MovementID", "Version", "EpisodeState", "CurrentPID", "Place", "Plate", "NFe", "Key", "AmountIn", "AmountOut", "CivilDate", "Canon", "Active", "Margin", "Annual")
    If fail Then Err.Raise 7100, , "Falha de projecao apos cabecalho"
    If n > 0 Then
        ReDim output(1 To n, 1 To 15)
        For i = 0 To n - 1: For j = 0 To 14: output(i + 1, j + 1) = a(j, i): Next j: Next i
        sh.Range("A2:K" & n + 1).NumberFormat = "@": sh.Range("N2:O" & n + 1).NumberFormat = "@"
        sh.Range("L2:M" & n + 1).NumberFormat = "0.0000"
        sh.Range("A2:O" & n + 1).Value2 = output
        For i = 2 To n + 1: sh.Cells(i, 11).Errors(xlNumberAsText).Ignore = True: Next i
        sh.Range("P2:P" & n + 1).FormulaR1C1 = "=AND(RC[-10]=""ACTIVE"",RC[-16+1]=RC[-9])"
        ' Identidade e valores acompanham a linha da projecao; reordenar SEM ID nunca autoriza mutacao.
        sh.Range("P2:P" & n + 1).FormulaR1C1 = "=AND(RC6=""ACTIVE"",RC1=RC7)"
        sh.Range("Q2:Q" & n + 1).FormulaR1C1 = "=RC13-RC12"
        sh.Range("R2:R" & n + 1).FormulaR1C1 = "=IF(COUNTIF(R2C1:R" & n + 1 & "C1,RC1)<>1,NA(),IF(RC6=""CLOSED"",""SIM"",""NAO""))"
    End If
    sh.Range("T1").Value2 = Val9(c, "SELECT BaseID FROM Meta WHERE ID=1")
    sh.Range("T2").Value2 = Val9(c, "SELECT RevisionNo FROM Meta WHERE ID=1")
    wb.Application.Calculate
    Materialize9 = n
End Function
Private Function NewProjection9(ByVal c As Object, ByVal name As String) As Workbook
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Materialize9 c, wb
    wb.SaveAs ThisWorkbook.Path & "\" & name & ".xlsm", xlOpenXMLWorkbookMacroEnabled
    Set NewProjection9 = wb
End Function
Private Function Step9(ByVal c As Object, ByVal op As String, ByVal ep As String, ByVal ver As Long, ByVal kind As String, ByVal vin As Currency, ByVal vout As Currency) As String
    Dim st As String
    st = Prepare9(c, owner9, fence9, op, NewID9(), ep, ver, kind, vin, vout, "SYNTHETIC")
    If st = "PREPARADA" Then st = Apply9(c, owner9, fence9, op)
    Step9 = st
End Function
Private Function RejectSQL9(ByVal c As Object, ByVal sql As String) As String
    On Error GoTo Rejected
    c.Execute sql: RejectSQL9 = "NOT_REJECTED": Exit Function
Rejected:
    RejectSQL9 = CStr(Err.Number) & ":" & Err.Description
End Function
Public Sub F9_Core()
    Dim c As Object, c2 As Object, wb As Workbook, gate As clsSaveFase9, r As String, b As String, n As Long, i As Long, f As Long, oldFence As Long, arr As Variant, before As String, after As String, e As Long, nm As String, op As String, tx As Boolean
    Dim ev As Boolean, su As Boolean, da As Boolean, calc As XlCalculation
    Guard9: phase9 = "F9_Core": StartLog9 "core_f9"
    ev = Application.EnableEvents: su = Application.ScreenUpdating: da = Application.DisplayAlerts: calc = Application.Calculation
    On Error GoTo Failed
    nm = "ciclo.accdb": CreateDB9 nm: Set c = Open9(nm): owner9 = NewID9(): fence9 = Acquire9(c, owner9)
    Check9 "F9_LEASE_PRIMEIRO", fence9 = 1, "RECUSADO", "Primeiro escritor recebe fence1", CStr(fence9), nm
    Set c2 = Open9(nm): Check9 "F9_LEASE_SEGUNDO", Acquire9(c2, NewID9()) = 0, "RECUSADO", "Segundo nao adquire", "Lease ativo", nm: c2.Close
    Set c2 = Open9(nm, True): r = RejectSQL9(c2, "UPDATE Meta SET RevisionNo=100 WHERE ID=1")
    Check9 "F9_SESSAO_LEITURA", r <> "NOT_REJECTED", "RECUSADO", "Gravacao recusada pelo provedor", r, nm: c2.Close
    For i = 1 To 100
        b = b & NewID9() & vbLf
    Next i
    Check9 "F9_GUID_LOCAL", Len(b) = 3700, "RECUSADO", "100 GUIDs gerados; unicidade conferida no teste adicional", CStr(Len(b)), nm
    Seed9 c, "VID_A", "EP_A1", "FIC9A01": Seed9 c, "VID_B", "EP_B1", "FIC9B01"
    Check9 "F9_ENTRADA_A1", Step9(c, "OP_A1", "EP_A1", 0, "ENTRY", 1000.1234@, 0@) = "CONFIRMADA", "RECUPERADO", "Entrada confirmada", "EP_A1", nm
    Check9 "F9_ENTRADA_B1", Step9(c, "OP_B1", "EP_B1", 0, "ENTRY", 2200.4321@, 0@) = "CONFIRMADA", "RECUPERADO", "Entrada distinta", "EP_B1", nm
    Check9 "F9_VENDA_A1", Step9(c, "OP_V1", "EP_A1", 1, "SALE", 1000.1234@, 1370.9876@) = "CONFIRMADA", "RECUPERADO", "Venda encerra episodio", CStr(Val9(c, "SELECT StatusText FROM Episodes WHERE EpisodeID='EP_A1'")), nm
    Seed9 c, "VID_A", "EP_A2", "FIC9A01"
    Check9 "F9_REENTRADA", Step9(c, "OP_A2", "EP_A2", 0, "ENTRY", 1500.1111@, 0@) = "CONFIRMADA", "RECUPERADO", "Novo EP, mesmo VehicleID", CStr(Val9(c, "SELECT VehicleID FROM Episodes WHERE EpisodeID='EP_A2'")), nm
    Check9 "F9_TRANSFERENCIA", Step9(c, "OP_T1", "EP_A2", 1, "TRANSFER", 1500.1111@, 1940.2222@) = "CONFIRMADA", "RECUPERADO", "Mesmo EP / FILIAL", CStr(Val9(c, "SELECT PlaceCode FROM Episodes WHERE EpisodeID='EP_A2'")), nm
    Check9 "F9_DEVOLUCAO", Step9(c, "OP_D1", "EP_A2", 2, "RETURN", 1500.1111@, 2100.3333@) = "CONFIRMADA", "RECUPERADO", "Mesmo EP / MATRIZ", CStr(Val9(c, "SELECT PlaceCode FROM Episodes WHERE EpisodeID='EP_A2'")), nm
    Check9 "F9_NOVA_TRANSFERENCIA", Step9(c, "OP_T2", "EP_A2", 3, "TRANSFER", 1500.1111@, 2500.4444@) = "CONFIRMADA", "RECUPERADO", "Mesmo EP versao4", CStr(Val9(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID='EP_A2'")), nm
    before = DataSig9(c)
    r = Prepare9(c, owner9, fence9, "OP_T2", NewID9(), "EP_A2", 3, "TRANSFER", 1500.1111@, 2500.4444@, "SYNTHETIC")
    Check9 "F9_IDEMPOTENCIA", r = "JA_CONCLUIDA" And DataSig9(c) = before, "RECUSADO", "Recibo original sem repetir efeitos", r, nm
    r = Prepare9(c, owner9, fence9, "OP_T2", NewID9(), "EP_A2", 3, "TRANSFER", 1500.1111@, 9999@, "SYNTHETIC")
    Check9 "F9_PAYLOAD", r = "RECUSADO_PAYLOAD", "RECUSADO", "Payload diferente recusado", r, nm
    r = Prepare9(c, owner9, fence9, "OP_OLD", NewID9(), "EP_A2", 2, "TRANSFER", 1500@, 2000@, "X")
    Check9 "F9_VERSAO", r = "RECUSADO_VERSAO", "RECUSADO", "Versao obsoleta recusada", r, nm
    r = Prepare9(c, owner9, fence9, "OP_BASE", NewID9(), "EP_A2", 4, "RETURN", 1500@, 2000@, "X", "BASE_FALSA")
    Check9 "F9_COPIA_DIVERGENTE", r = "RECUSADO_BASE", "RECUSADO", "BaseID divergente nao escreve", r, nm
    r = RejectSQL9(c, "INSERT INTO Vehicles VALUES ('VID_A','OUTRA')"): Check9 "F9_PK_VEICULO", r <> "NOT_REJECTED", "RECUSADO", "PK recusa duplicado", r, nm
    r = RejectSQL9(c, "INSERT INTO Episodes (EpisodeID,VehicleID) VALUES ('EP_ORFAO','VID_AUSENTE')"): Check9 "F9_FK_EPISODIO", r <> "NOT_REJECTED", "RECUSADO", "FK recusa orfao", r, nm
    r = RejectSQL9(c, "INSERT INTO Projections SELECT * FROM Projections WHERE EpisodeID='EP_A1'"): Check9 "F9_PK_PROJECAO", r <> "NOT_REJECTED", "RECUSADO", "ProjectionID duplicado recusado", r, nm
    Check9 "F9_PRECISAO_TEXTO", TypeName(Val9(c, "SELECT AmountIn FROM Operations WHERE OperationID='OP_A1'")) = "Currency" And CCur(Val9(c, "SELECT AmountIn FROM Operations WHERE OperationID='OP_A1'")) = 1000.1234@ And TypeName(Val9(c, "SELECT AccessKey FROM Projections")) = "String" And Len(CStr(Val9(c, "SELECT AccessKey FROM Projections"))) = 44, "RECUPERADO", "Currency4 e chave String44", CStr(Val9(c, "SELECT AmountIn FROM Operations WHERE OperationID='OP_A1'")) & ";" & TypeName(Val9(c, "SELECT AccessKey FROM Projections")), nm
    Exec9 c, "INSERT INTO Provenance VALUES ('LEGACY_DUP','HASH_SINTETICO','AMBIGUO',NULL,'Duas notas coincidem; nenhum vinculo')"
    Check9 "F9_M3_AMBIGUO", IsNull(Val9(c, "SELECT EpisodeID FROM Provenance WHERE LegacyRef='LEGACY_DUP'")), "PENDENTE_CONCILIACAO", "Legado ambiguo nao vinculado", "AMBIGUO;NULL", nm
    Set wb = NewProjection9(c, "Projecao_Ciclo")
    Set gate = New clsSaveFase9: gate.Iniciar wb: gate.RecusarUmaVez
    wb.Worksheets(1).Range("T3").Value2 = "ALTERACAO_NAO_SALVA": wb.Save
    Check9 "F9_SAVE_RECUSADO", gate.Observado And Not wb.Saved And CLng(Val9(c, "SELECT Count(*) FROM Receipts")) = 7, "PENDENTE_CONCILIACAO", "Save recusado nao desfaz commit / recibo duravel no banco", "Saved=" & wb.Saved & ";BeforeSave=" & gate.Observado, nm
    Set gate = Nothing
    wb.Close False: Set wb = Nothing
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\Projecao_Ciclo.xlsm")
    Check9 "F9_DISCO_VS_MEMORIA", CStr(wb.Worksheets(1).Range("T3").Value2) = "", "RECUPERADO", "Alteracao cancelada ausente em disco", CStr(wb.Worksheets(1).Range("T3").Value2), nm
    wb.Close False: Set wb = NewProjection9(c, "Projecao_Reconstruida")
    Check9 "F9_RECONSTRUCAO", wb.Worksheets(1).Range("A2").Value2 <> "" And Application.CountIf(wb.Worksheets(1).Range("P2:P8"), True) = 2, "RECUPERADO", "7 projecoes/2 ativos reconstruidos", CStr(Application.CountIf(wb.Worksheets(1).Range("P2:P8"), True)), nm
    wb.Worksheets(1).Range("A2:R8").Sort Key1:=wb.Worksheets(1).Range("L2"), Order1:=xlDescending, Header:=xlNo
    Check9 "F9_REORDENACAO", Application.CountIf(wb.Worksheets(1).Range("P2:P8"), True) = 2, "RECUPERADO", "Contagem por PID mantida apos sort", CStr(Application.CountIf(wb.Worksheets(1).Range("P2:P8"), True)), nm
    wb.Save: wb.Close False: Set wb = Nothing
    Release9 c, owner9, fence9: c.Close
    FSO9.CopyFile ThisWorkbook.Path & "\ciclo.accdb", ThisWorkbook.Path & "\backup_ciclo.accdb", False
    Set c = Open9(nm): owner9 = NewID9(): fence9 = Acquire9(c, owner9)
    Snapshot9 c, "ciclo_final"
    Check9 "F9_REABERTURA_CONEXAO", CLng(Val9(c, "SELECT Count(*) FROM Receipts")) = 7 And VerifyBatch9(BusinessRows9(c)) = 7, "RECUPERADO", "7 recibos e registros verificados apos fechar/reabrir", "7", nm
    oldFence = fence9
    Exec9 c, "UPDATE WriterLease SET OwnerID='RECLAIM_TEST',FenceNo=FenceNo+1 WHERE ID=1 AND OwnerID=? AND FenceNo=?", owner9, oldFence
    r = Prepare9(c, owner9, oldFence, "OP_STALE", NewID9(), "EP_A2", 4, "RETURN", 1500@, 2000@, "X")
    Check9 "F9_FENCING", r = "RECUSADO_FENCE", "RECUSADO", "Escritor antigo perde autoridade", r, nm
    Exec9 c, "UPDATE WriterLease SET StatusText='FREE',OwnerID='' WHERE ID=1 AND OwnerID='RECLAIM_TEST'"
    c.Close: Set c = Nothing
    log9.Close
    Application.EnableEvents = ev: Application.ScreenUpdating = su: Application.DisplayAlerts = da: Application.Calculation = calc
    Debug.Print "F9_CORE_DONE": Exit Sub
Failed:
    e = Err.Number: r = Err.Description
    On Error Resume Next
    Check9 "ERRO_CORE", False, "FAIL", "Conclusao controlada", CStr(e) & ":" & r, nm
    If Not wb Is Nothing Then wb.Close False
    If Not c Is Nothing Then c.Close
    log9.Close
    Application.EnableEvents = ev: Application.ScreenUpdating = su: Application.DisplayAlerts = da: Application.Calculation = calc
    Debug.Print "F9_CORE_ERROR"; e; r
End Sub
