Attribute VB_Name = "modPersistenciaFase10"
Option Explicit
' LABORATORIO SINTETICO. Nao integra gravadores nem politicas do produto.
Private Const MARK10 As String = "VBA_Fase10_20261102_155920Z"
Private Const KEY10 As String = "01234567890123456789012345678901234567890123"
Private Const PROVIDER10 As String = "Microsoft.ACE.OLEDB.16.0"
Private Type GUID10
    D1 As Long
    D2 As Integer
    D3 As Integer
    D4(0 To 7) As Byte
End Type
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef g As GUID10) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef g As GUID10, ByVal p As LongPtr, ByVal n As Long) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef n As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef n As Currency) As Long
Public log10 As Object, cn10 As Object, held10 As Object
Public owner10 As String, fence10 As Long, phase10 As String

Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32.dll" (ByRef phProv As LongPtr, ByVal container As LongPtr, ByVal provider As LongPtr, ByVal typ As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" (ByVal hp As LongPtr, ByVal alg As Long, ByVal key As LongPtr, ByVal flags As Long, ByRef hh As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" (ByVal hh As LongPtr, ByRef data As Any, ByVal n As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" (ByVal hh As LongPtr, ByVal param As Long, ByRef data As Any, ByRef n As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" (ByVal hh As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hp As LongPtr, ByVal flags As Long) As Long
Private heldTx10 As Boolean
Public hook10 As Boolean
Public Function FSO10() As Object
    Set FSO10 = CreateObject("Scripting.FileSystemObject")
End Function
Public Sub Guard10()
    Dim p As String
    If ThisWorkbook.Name <> "Estoque_Fera_Veiculos_Candidato_Integrado.xlsm" Then Err.Raise 5, , "Candidata sintetica obrigatoria"
    p = FSO10.GetFolder(ThisWorkbook.Path).ShortPath
    If InStr(1, ThisWorkbook.Path, "Entrega1_", vbTextCompare) = 0 Or InStr(1, ThisWorkbook.Path, "OneDrive", vbTextCompare) > 0 Then Err.Raise 5, , "Executar somente laboratorio TEMP autorizado"
    If StrComp(Left$(p, Len(FSO10.GetFolder(Environ$("TEMP")).ShortPath)), FSO10.GetFolder(Environ$("TEMP")).ShortPath, vbTextCompare) <> 0 Then Err.Raise 5, , "Somente TEMP"
    If ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then Err.Raise 5, , "Marcador sintetico ausente"
End Sub
Public Function NewID10() As String
    Dim g As GUID10, s As String
    If CoCreateGuid(g) <> 0 Then Err.Raise 5, , "GUID indisponivel"
    s = String$(39, vbNullChar)
    If StringFromGUID2(g, StrPtr(s), 39) = 0 Then Err.Raise 5, , "GUID invalido"
    NewID10 = UCase$(mid$(s, 2, 36))
End Function
Public Function Now10() As Double
    Dim n As Currency, f As Currency
    QueryPerformanceCounter n: QueryPerformanceFrequency f: Now10 = n / f
End Function
Public Function Clean10(ByVal v As Variant) As String
    If IsNull(v) Then Clean10 = "<NULL>": Exit Function
    If IsEmpty(v) Then Clean10 = "<EMPTY>": Exit Function
    Clean10 = Replace(Replace(Replace(CStr(v), vbTab, "<TAB>"), vbCr, "<CR>"), vbLf, "<LF>")
End Function
Public Sub StartLog10(ByVal name As String)
    Set log10 = FSO10.CreateTextFile(ThisWorkbook.Path & "\" & name & ".tsv", False, True)
    log10.WriteLine "Cenario" & vbTab & "Resultado" & vbTab & "Tratamento" & vbTab & "Esperado" & vbTab & "Observado" & vbTab & "Fixture" & vbTab & "Procedimento"
End Sub
Public Sub Check10(ByVal id As String, ByVal ok As Boolean, ByVal treatment As String, ByVal expected As String, ByVal observed As String, ByVal fixture As String)
    log10.WriteLine id & vbTab & IIf(ok, "PASS", "FAIL") & vbTab & treatment & vbTab & Clean10(expected) & vbTab & Clean10(observed) & vbTab & fixture & vbTab & phase10
End Sub
Public Function Open10(ByVal name As String, Optional ByVal readOnly As Boolean = False) As Object
    Dim c As Object
    Call Guard10
    If InStr(name, "\") > 0 Or InStr(name, "/") > 0 Or InStr(name, "..") > 0 Or LCase$(Right$(name, 6)) <> ".accdb" Then Err.Raise 5, , "Nome de banco local exigido"
    Set c = CreateObject("ADODB.Connection")
    If readOnly Then c.Mode = 1
    c.ConnectionTimeout = 5: c.CommandTimeout = 5
    c.Open "Provider=" & PROVIDER10 & ";Data Source=" & ThisWorkbook.Path & "\" & name & ";"
    Set Open10 = c
End Function
Public Function Cmd10(ByVal c As Object, ByVal sql As String, ByVal ps As Variant) As Object
    Dim q As Object, p As Object, i As Long, typ As Long, sz As Long, v As Variant
    Set q = CreateObject("ADODB.Command"): Set q.ActiveConnection = c: q.commandText = sql: q.CommandType = 1
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
    Set Cmd10 = q
End Function
Public Function Exec10(ByVal c As Object, ByVal sql As String, ParamArray ps() As Variant) As Long
    Dim q As Object, r As Object, n As Variant
    Set q = Cmd10(c, sql, ps): Set r = q.Execute(n): Exec10 = CLng(n)
End Function
Public Function Val10(ByVal c As Object, ByVal sql As String, ParamArray ps() As Variant) As Variant
    Dim q As Object, r As Object
    Set q = Cmd10(c, sql, ps): Set r = q.Execute
    If r.EOF Then Val10 = Empty Else Val10 = r.fields(0).value
    r.Close
End Function
Public Function Rows10(ByVal c As Object, ByVal sql As String) As Variant
    Dim r As Object
    Set r = c.Execute(sql)
    If r.EOF Then Rows10 = Empty Else Rows10 = r.GetRows
    r.Close
End Function
Public Function Pack10(ParamArray fields() As Variant) As String
    Dim i As Long, s As String, v As String
    For i = LBound(fields) To UBound(fields)
        If IsNull(fields(i)) Then
            v = "NULL"
        ElseIf VarType(fields(i)) = vbCurrency Then
            v = "M4:" & Format$(CDec(fields(i)) * CDec(10000), "0")
        Else
            v = TypeName(fields(i)) & ":" & CStr(fields(i))
        End If
        s = s & Len(v) & ":" & v
    Next i
    Pack10 = s
End Function
Public Function CanonProjection10(ByVal c As Object, ByVal pid As String) As String
    Dim r As Object, q As Object, i As Long, s As String
    Set q = Cmd10(c, "SELECT ProjectionID,EpisodeID,MovementID,AmountIn,AmountOut,CivilDate,DocNumber,AccessKey,PlaceCode FROM Projections WHERE ProjectionID=?", Array(pid))
    Set r = q.Execute
    If r.EOF Then CanonProjection10 = "ERROR:MISSING": r.Close: Exit Function
    For i = 0 To r.fields.count - 1: s = s & Pack10(r.fields(i).value): Next i
    CanonProjection10 = s: r.Close
End Function
Public Function Source10(ByVal c As Object, ByVal ep As String) As String
    Dim pid As Variant
    pid = Val10(c, "SELECT CurrentPID FROM Episodes WHERE EpisodeID=?", ep)
    Source10 = Pack10(ep, Val10(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep), Val10(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", ep))
    If Not IsNull(pid) And Not IsEmpty(pid) Then Source10 = Source10 & CanonProjection10(c, CStr(pid))
End Function
Public Sub LegacyDB10(ByVal name As String)
    Dim cat As Object, c As Object, ddl As Variant, x As Variant
    If FSO10.FileExists(ThisWorkbook.Path & "\" & name) Then Err.Raise 58, , "Nao sobrescrever banco"
    Set cat = CreateObject("ADOX.Catalog")
    cat.Create "Provider=" & PROVIDER10 & ";Data Source=" & ThisWorkbook.Path & "\" & name & ";"
    cat.ActiveConnection.Close: Set cat = Nothing: Set c = Open10(name)
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
    Exec10 c, "INSERT INTO Meta VALUES (1,?,1,0)", NewID10()
    Exec10 c, "INSERT INTO WriterLease VALUES (1,'',0,'FREE')"
    c.Close
End Sub
Public Function Acquire10(ByVal c As Object, ByVal who As String) As Long
    Dim n As Long
    c.BeginTrans
    n = Exec10(c, "UPDATE WriterLease SET OwnerID=?,FenceNo=FenceNo+1,StatusText='ACTIVE' WHERE ID=1 AND StatusText='FREE'", who)
    If n = 1 Then Acquire10 = CLng(Val10(c, "SELECT FenceNo FROM WriterLease WHERE ID=1"))
    c.CommitTrans
End Function
Public Function Own10(ByVal c As Object, ByVal who As String, ByVal fence As Long) As Boolean
    Own10 = CLng(Val10(c, "SELECT Count(*) FROM WriterLease WHERE ID=1 AND OwnerID=? AND FenceNo=? AND StatusText='ACTIVE'", who, fence)) = 1
End Function
Public Sub Release10(ByVal c As Object, ByVal who As String, ByVal fence As Long)
    Exec10 c, "UPDATE WriterLease SET StatusText='FREE',OwnerID='' WHERE ID=1 AND OwnerID=? AND FenceNo=?", who, fence
End Sub
Public Sub Seed10(ByVal c As Object, ByVal vid As String, ByVal ep As String, ByVal plate As String)
    If CLng(Val10(c, "SELECT Count(*) FROM Vehicles WHERE VehicleID=?", vid)) = 0 Then Exec10 c, "INSERT INTO Vehicles VALUES (?,?)", vid, plate
    If CLng(Val10(c, "SELECT Count(*) FROM Episodes WHERE VehicleID=? AND StatusText='ACTIVE'", vid)) > 0 Then Err.Raise 5, , "Veiculo ja em episodio ativo"
    Exec10 c, "INSERT INTO Episodes (EpisodeID,VehicleID,VersionNo,StatusText,PlaceCode) VALUES (?,?,0,'NEW','MATRIZ')", ep, vid
End Sub
Public Function Prepare10(ByVal c As Object, ByVal who As String, ByVal fence As Long, ByVal op As String, ByVal attempt As String, ByVal ep As String, ByVal expected As Long, ByVal kind As String, ByVal vin As Currency, ByVal vout As Currency, ByVal payload As String, Optional ByVal expectedBase As String = "", Optional ByVal civil As String = "2026-11-02", Optional ByVal doc As String = "88001", Optional ByVal access As String = KEY10) As String
    Dim canon As String, old As Variant, status As String, pid As String, mid As String, source As String, tx As Boolean, result As String, st As String, ver As Long
    On Error GoTo Failed
    If Not IDOK10(attempt) Or Len(payload) = 0 Or expected < 0 Then Prepare10 = "RECUSADO_ID_PAYLOAD": Exit Function
    If Not DateOK10(civil) Or Len(doc) = 0 Or Len(doc) > 20 Or doc Like "*[!0-9]*" Or (Len(access) <> 0 And Len(access) <> 44) Or access Like "*[!0-9]*" Then Prepare10 = "RECUSADO_DOCUMENTO": Exit Function
    If Abs(CDec(vin)) > CDec(1000000000) Or Abs(CDec(vout)) > CDec(1000000000) Then Prepare10 = "RECUSADO_PRECISAO_EXCEL": Exit Function
    c.BeginTrans: tx = True
    If Not LockWriter10(c, who, fence) Then result = "RECUSADO_FENCE": GoTo Done
    If expectedBase <> "" Then If expectedBase <> CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")) Then result = "RECUSADO_BASE": GoTo Done
    old = Val10(c, "SELECT CommandCanon FROM Operations WHERE OperationID=?", op)
    If CStr(Val10(c, "SELECT WriteState FROM Meta WHERE ID=1")) <> "READY" Then result = "PENDENTE_CONCILIACAO": GoTo Done
    st = Audit10(c)
    If st <> "OK" Then
        Exec10 c, "UPDATE Meta SET WriteState='RECONCILE',IntegrityReason=? WHERE ID=1", st
        result = "PENDENTE_INTEGRIDADE:" & st: GoTo Done
    End If
    If Not IsEmpty(old) Then
        ver = CLng(Val10(c, "SELECT CommandVersion FROM Operations WHERE OperationID=?", op))
        pid = CStr(Val10(c, "SELECT PlannedPID FROM Operations WHERE OperationID=?", op)): mid = CStr(Val10(c, "SELECT PlannedMID FROM Operations WHERE OperationID=?", op))
        canon = Command10(c, op, ep, expected, kind, vin, vout, payload, civil, doc, access, pid, mid, ver)
        If canon <> CStr(old) Then
            result = "RECUSADO_PAYLOAD"
        ElseIf CStr(Val10(c, "SELECT StatusText FROM Operations WHERE OperationID=?", op)) = "CONFIRMED" Then
            result = "JA_CONCLUIDA"
        Else
            result = "PENDENTE_CONCILIACAO"
        End If
        LogAttempt10 c, attempt, op, who, fence, result, canon, expected, CStr(Val10(c, "SELECT StatusText FROM Operations WHERE OperationID=?", op)), CStr(Val10(c, "SELECT StatusText FROM Operations WHERE OperationID=?", op))
        GoTo Done
    End If
    If Not IDOK10(op) Or Not IDOK10(ep) Then result = "RECUSADO_ID": GoTo Done
    If CLng(Val10(c, "SELECT Count(*) FROM Pending WHERE EpisodeID=?", ep)) <> 0 Then result = "PENDENTE_CONCILIACAO": GoTo Done
    If CLng(Val10(c, "SELECT Count(*) FROM Episodes WHERE EpisodeID=?", ep)) <> 1 Then result = "RECUSADO_EPISODIO": GoTo Done
    If CLng(Val10(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep)) <> expected Then result = "RECUSADO_VERSAO": GoTo Done
    status = CStr(Val10(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", ep))
    If (kind = "ENTRY" And status <> "NEW") Or (kind <> "ENTRY" And status <> "ACTIVE") Then result = "RECUSADO_ESTADO": GoTo Done
    If kind <> "ENTRY" And kind <> "SALE" And kind <> "TRANSFER" And kind <> "RETURN" Then result = "RECUSADO_TIPO": GoTo Done
    source = Source10(c, ep): pid = NewID10(): mid = NewID10()
    canon = Command10(c, op, ep, expected, kind, vin, vout, payload, civil, doc, access, pid, mid, 2)
    Exec10 c, "INSERT INTO Operations (OperationID,EpisodeID,ExpectedVersion,KindText,AmountIn,AmountOut,CommandCanon,SourceCanon,PlannedPID,PlannedMID,StatusText,CommandHash,SourceHash,CivilDate,DocNumber,AccessKey,PayloadText,CommandVersion) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", op, ep, expected, kind, vin, vout, canon, source, pid, mid, "PREPARED", SHA10(canon), SHA10(source), civil, doc, access, payload, 2
    LogAttempt10 c, attempt, op, who, fence, "PREPARED", canon, expected, "ABSENT", "PREPARED"
    Exec10 c, "INSERT INTO Pending VALUES (?,?,'INTENT_COMMITTED')", ep, op
    result = "PREPARADA"
Done:
    c.CommitTrans: tx = False: Prepare10 = result: Exit Function
Failed:
    Prepare10 = "PENDENTE_ERRO:" & Err.Number & ":" & Err.Description
    On Error Resume Next
    If tx Then c.RollbackTrans
End Function
Public Function Command10(ByVal c As Object, ByVal op As String, ByVal ep As String, ByVal expected As Long, ByVal kind As String, ByVal vin As Currency, ByVal vout As Currency, ByVal payload As String, ByVal civil As String, ByVal doc As String, ByVal access As String, ByVal pid As String, ByVal mid As String, ByVal ver As Long) As String
    Command10 = Pack10("SCHEMA" & ver, CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")), op, ep, expected, kind, vin, vout, civil, doc, access, payload)
    If ver = 2 Then Command10 = Command10 & Pack10(pid, mid)
End Function
Public Function LastPayload10(ByVal canon As String) As String
    Dim p As Long, colon As Long, n As Long, token As String
    p = 1
    Do While p <= Len(canon)
        colon = InStr(p, canon, ":"): If colon = 0 Then Err.Raise 5, , "Canon invalido"
        n = CLng(mid$(canon, p, colon - p)): token = mid$(canon, colon + 1, n): p = colon + 1 + n
    Loop
    If Left$(token, 7) <> "String:" Then Err.Raise 5, , "Payload legado sem tipo"
    LastPayload10 = mid$(token, 8)
End Function
Public Function Apply10(ByVal c As Object, ByVal who As String, ByVal fence As Long, ByVal op As String, Optional ByVal failure As Long = 0) As String
    Dim audit As String, prev As String, aft As String, tx As Boolean, ep As String, pid As String, mid As String, kind As String, src As String, loc As String, ver As Long, rev As Long, seal As String, e As Long, msg As String
    On Error GoTo Failed
    If failure = -1 Then Apply10 = "RECUSADO_ANTES_BEGIN": Exit Function
    If Not Own10(c, who, fence) Then Apply10 = "RECUSADO_FENCE": Exit Function
    If CStr(Val10(c, "SELECT StatusText FROM Operations WHERE OperationID=?", op)) <> "PREPARED" Then Apply10 = "RECUSADO_ESTADO": Exit Function
    If CLng(Val10(c, "SELECT Count(*) FROM Attempts WHERE OperationID=? AND OwnerID=? AND FenceNo=? AND OutcomeText='PREPARED'", op, who, fence)) <> 1 Then Apply10 = "PENDENTE_OWNER_CHANGED": Exit Function
    ep = CStr(Val10(c, "SELECT EpisodeID FROM Operations WHERE OperationID=?", op))
    src = CStr(Val10(c, "SELECT SourceCanon FROM Operations WHERE OperationID=?", op))
    If Source10(c, ep) <> src Then Apply10 = "PENDENTE_ORIGEM_ALTERADA": Exit Function
    pid = CStr(Val10(c, "SELECT PlannedPID FROM Operations WHERE OperationID=?", op)): mid = CStr(Val10(c, "SELECT PlannedMID FROM Operations WHERE OperationID=?", op))
    kind = CStr(Val10(c, "SELECT KindText FROM Operations WHERE OperationID=?", op)): ver = CLng(Val10(c, "SELECT ExpectedVersion FROM Operations WHERE OperationID=?", op))
    loc = C1_LocalOperacao(c, op)
    c.BeginTrans: tx = True
    ' Gravacao na linha de lease adquire bloqueio; fencing conferido dentro da transacao.
    If Not LockWriter10(c, who, fence) Then Err.Raise 5, , "Lease invalido"
    If CLng(Val10(c, "SELECT Count(*) FROM Attempts WHERE OperationID=? AND OwnerID=? AND FenceNo=? AND OutcomeText='PREPARED'", op, who, fence)) <> 1 Then Err.Raise 5, , "Dono da intencao divergente"
    If failure = -2 Then Err.Raise 7012, , "Depois de BeginTrans"
    audit = Audit10(c): If audit <> "OK" Then Err.Raise 5, , "Integridade: " & audit
    If Source10(c, ep) <> src Then Err.Raise 5, , "Origem mudou"
    Exec10 c, "INSERT INTO Movements VALUES (?,?,?,?)", mid, op, ep, kind
    If failure = 1 Then Err.Raise 7001, , "Falha apos movimento"
    Exec10 c, "INSERT INTO Projections (ProjectionID,EpisodeID,MovementID,AmountIn,AmountOut,CivilDate,DocNumber,AccessKey,PlaceCode,ContentCanon) VALUES (?,?,?,?,?,?,?,?,?,?)", pid, ep, mid, CCur(Val10(c, "SELECT AmountIn FROM Operations WHERE OperationID=?", op)), CCur(Val10(c, "SELECT AmountOut FROM Operations WHERE OperationID=?", op)), CStr(Val10(c, "SELECT CivilDate FROM Operations WHERE OperationID=?", op)), CStr(Val10(c, "SELECT DocNumber FROM Operations WHERE OperationID=?", op)), CStr(Val10(c, "SELECT AccessKey FROM Operations WHERE OperationID=?", op)), loc, "PENDING"
    seal = CanonProjection10(c, pid): Exec10 c, "UPDATE Projections SET ContentCanon=? WHERE ProjectionID=?", seal, pid
    If failure = 2 Then Err.Raise 7002, , "Falha apos projecao"
    If Exec10(c, "UPDATE Episodes SET CurrentEpisodeID=EpisodeID,CurrentPID=?,VersionNo=VersionNo+1,StatusText=?,PlaceCode=? WHERE EpisodeID=? AND VersionNo=?", pid, IIf(kind = "SALE", "CLOSED", "ACTIVE"), loc, ep, ver) <> 1 Then Err.Raise 5, , "CAS episodio recusado"
    If failure = 3 Then Err.Raise 7003, , "Falha durante baixa logica"
    C1_RegistrarViews c, op, ep, mid, pid, kind, loc
    rev = CLng(Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1"))
    aft = Source10(c, ep)
    Exec10 c, "INSERT INTO ChangeLog (OperationID,BeforeCanon,AfterCanon,RevisionNo,AuditHash) VALUES (?,?,?,?,?)", op, src, aft, rev + 1, SHA10(Pack10(op, rev + 1, src, aft))
    If failure = 4 Then c.Execute "INSERT INTO ChangeLog (OperationID) VALUES ('OP_INEXISTENTE')"
    rev = CLng(Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1"))
    If Exec10(c, "UPDATE Meta SET RevisionNo=RevisionNo+1 WHERE ID=1 AND RevisionNo=?", rev) <> 1 Then Err.Raise 5, , "CAS base recusado"
    If failure = 5 Then Err.Raise 7005, , "Antes do recibo"
    prev = "GENESIS"
    If rev > 0 Then prev = CStr(Val10(c, "SELECT ReceiptHash FROM Receipts WHERE RevisionNo=?", rev))
    Exec10 c, "INSERT INTO Receipts (ReceiptID,OperationID,RevisionNo,EffectCanon,BaseID,EpisodeID,VehicleID,MovementID,ProjectionID,PreviousHash,ReceiptHash) VALUES (?,?,?,?,?,?,?,?,?,?,?)", NewID10(), op, rev + 1, seal, CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")), ep, CStr(Val10(c, "SELECT VehicleID FROM Episodes WHERE EpisodeID=?", ep)), mid, pid, prev, String$(64, "0")
    Exec10 c, "UPDATE Receipts SET ReceiptHash=? WHERE OperationID=?", ReceiptHash10(c, op), op
    Exec10 c, "UPDATE Operations SET StatusText='CONFIRMED' WHERE OperationID=?", op
    Exec10 c, "DELETE FROM Pending WHERE OperationID=?", op
    If failure = 6 Then Err.Raise 7006, , "Antes do commit"
    If failure = 7 Then c.Close: c.CommitTrans
    audit = Audit10(c): If audit <> "OK" Then Err.Raise 5, , "Pos-validacao: " & audit
    c.CommitTrans: tx = False
    If failure = 8 Then Err.Raise 7008, , "Apos commit, antes de Excel"
    Apply10 = "CONFIRMADA": Exit Function
Failed:
    e = Err.Number: msg = Err.Description
    On Error Resume Next
    If tx Then c.RollbackTrans
    If failure = 8 Then Apply10 = "COMMIT_CONFIRMADO_PROJECAO_PENDENTE" Else Apply10 = "PENDENTE_ERRO:" & e & ":" & msg
End Function
Public Function DataSig10(ByVal c As Object) As String
    Dim names As Variant, sorts As Variant, i As Long, r As Object, j As Long, s As String
    names = Array("Episodes", "Movements", "Projections", "Receipts", "ChangeLog", "Meta")
    sorts = Array("EpisodeID", "MovementID", "ProjectionID", "ReceiptID", "AuditID", "ID")
    For i = 0 To UBound(names)
        Set r = c.Execute("SELECT * FROM " & names(i) & " ORDER BY " & sorts(i))
        s = s & names(i)
        Do While Not r.EOF
            For j = 0 To r.fields.count - 1: s = s & Pack10(r.fields(j).value): Next j
            r.MoveNext
        Loop
        r.Close
    Next i
    DataSig10 = s
End Function
Public Sub Snapshot10(ByVal c As Object, ByVal name As String)
    Dim f As Object, r As Object, tbl As Variant, j As Long
    Set f = FSO10.CreateTextFile(ThisWorkbook.Path & "\" & name & ".tsv", False, True)
    f.WriteLine "Tabela" & vbTab & "Campo" & vbTab & "TipoADO" & vbTab & "TipoVBA" & vbTab & "Valor"
    For Each tbl In Array("Meta", "WriterLease", "Vehicles", "Episodes", "Operations", "Attempts", "Movements", "Projections", "Receipts", "ChangeLog", "Pending", "Provenance")
        Set r = c.Execute("SELECT * FROM " & tbl)
        Do While Not r.EOF
            For j = 0 To r.fields.count - 1: f.WriteLine tbl & vbTab & r.fields(j).name & vbTab & r.fields(j).Type & vbTab & TypeName(r.fields(j).value) & vbTab & Clean10(r.fields(j).value): Next j
            r.MoveNext
        Loop
        r.Close
    Next tbl
    f.Close
End Sub
Public Function BusinessRows10(ByVal c As Object) As Variant
    ' Consulta em lote: uma leitura, unicidade por PK/FK mais revalidacao de conteudo.
    BusinessRows10 = Rows10(c, "SELECT P.ProjectionID,P.EpisodeID,E.VehicleID,P.MovementID,E.VersionNo,E.StatusText,E.CurrentPID,P.PlaceCode,V.Plate,P.DocNumber,P.AccessKey,P.AmountIn,P.AmountOut,P.CivilDate,P.ContentCanon FROM (Projections AS P INNER JOIN Episodes AS E ON P.EpisodeID=E.EpisodeID) INNER JOIN Vehicles AS V ON E.VehicleID=V.VehicleID ORDER BY P.ProjectionID")
End Function
Public Function VerifyBatch10(ByVal a As Variant) As Long
    Dim d As Object, i As Long, s As String
    Set d = CreateObject("Scripting.Dictionary")
    If IsEmpty(a) Then Exit Function
    For i = 0 To UBound(a, 2)
        If d.Exists(CStr(a(0, i))) Then Err.Raise 5, , "PID duplicado"
        d.Add CStr(a(0, i)), True
        s = Pack10(a(0, i)) & Pack10(a(1, i)) & Pack10(a(3, i)) & Pack10(a(11, i)) & Pack10(a(12, i)) & Pack10(a(13, i)) & Pack10(a(9, i)) & Pack10(a(10, i)) & Pack10(a(7, i))
        If s <> CStr(a(14, i)) Then Err.Raise 5, , "Conteudo divergente: " & a(0, i)
        If TypeName(a(10, i)) <> "String" Or (Len(a(10, i)) <> 0 And Len(a(10, i)) <> 44) Or a(10, i) Like "*[!0-9]*" Then Err.Raise 5, , "Chave invalida"
        If Not DateOK10(CStr(a(13, i))) Then Err.Raise 5, , "Data invalida"
        VerifyBatch10 = VerifyBatch10 + 1
    Next i
End Function
Public Function Materialize10(ByVal c As Object, ByVal wb As Workbook, Optional ByVal fail As Boolean = False) As Long
    Dim a As Variant, output() As Variant, i As Long, j As Long, n As Long, sh As Worksheet, expected As Long, k As Long
    a = BusinessRows10(c): n = VerifyBatch10(a)
    Set sh = wb.Worksheets(1): sh.name = "Projecao_F10"
    sh.Cells.Clear
    sh.Range("A1:R1").value = Array("ProjectionID", "EpisodeID", "VehicleID", "MovementID", "Version", "EpisodeState", "CurrentPID", "Place", "Plate", "NFe", "Key", "AmountIn", "AmountOut", "CivilDate", "Canon", "Active", "Margin", "Annual")
    If fail Then Err.Raise 7100, , "Falha de projecao apos cabecalho"
    If n > 0 Then
        ReDim output(1 To n, 1 To 15)
        For i = 0 To n - 1: For j = 0 To 14: output(i + 1, j + 1) = a(j, i): Next j: Next i
        sh.Range("A2:K" & n + 1).NumberFormat = "@": sh.Range("N2:O" & n + 1).NumberFormat = "@"
        sh.Range("L2:M" & n + 1).NumberFormat = "0.0000"
        sh.Range("A2:O" & n + 1).Value2 = output
        For i = 2 To n + 1: sh.Cells(i, 11).Errors(xlNumberAsText).Ignore = True: Next i
        ' Identidade e valores acompanham a linha da projecao; reordenar SEM ID nunca autoriza mutacao.
        sh.Range("P2:P" & n + 1).FormulaR1C1 = "=AND(RC6=""ACTIVE"",RC1=RC7)"
        sh.Range("Q2:Q" & n + 1).FormulaR1C1 = "=RC13-RC12"
        sh.Range("R2:R" & n + 1).FormulaR1C1 = "=IF(COUNTIF(R2C1:R" & n + 1 & "C1,RC1)<>1,NA(),IF(RC6=""CLOSED"",""SIM"",""NAO""))"
    End If
    sh.Range("T1").Value2 = Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")
    sh.Range("T2").Value2 = Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1")
    wb.Application.Calculate
    Materialize10 = n
End Function
Public Function NewProjection10(ByVal c As Object, ByVal name As String) As Workbook
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Materialize10 c, wb
    wb.SaveAs ThisWorkbook.Path & "\" & name & ".xlsm", xlOpenXMLWorkbookMacroEnabled
    Set NewProjection10 = wb
End Function
Public Function Step10(ByVal c As Object, ByVal op As String, ByVal ep As String, ByVal ver As Long, ByVal kind As String, ByVal vin As Currency, ByVal vout As Currency) As String
    Dim st As String
    st = Prepare10(c, owner10, fence10, op, NewID10(), ep, ver, kind, vin, vout, "SYNTHETIC")
    If st = "PREPARADA" Then st = Apply10(c, owner10, fence10, op)
    Step10 = st
End Function
Public Function RejectSQL10(ByVal c As Object, ByVal sql As String) As String
    On Error GoTo Rejected
    c.Execute sql: RejectSQL10 = "NOT_REJECTED": Exit Function
Rejected:
    RejectSQL10 = CStr(Err.Number) & ":" & Err.Description
End Function
Public Function Resolve10(ByVal c As Object, ByVal pid As String, ByVal ep As String, ByVal vid As String, ByVal expected As Long, Optional ByVal historical As Boolean = False) As String
    Dim q As Object, r As Object, a As Variant
    Set q = Cmd10(c, "SELECT P.ProjectionID,P.EpisodeID,E.VehicleID,P.MovementID,E.VersionNo,E.StatusText,E.CurrentPID,P.PlaceCode,V.Plate,P.DocNumber,P.AccessKey,P.AmountIn,P.AmountOut,P.CivilDate,P.ContentCanon FROM (Projections AS P INNER JOIN Episodes AS E ON P.EpisodeID=E.EpisodeID) INNER JOIN Vehicles AS V ON E.VehicleID=V.VehicleID WHERE P.ProjectionID=?", Array(pid))
    Set r = q.Execute
    If r.EOF Then Resolve10 = "INEXISTENTE": r.Close: Exit Function
    a = r.GetRows: r.Close
    If a(1, 0) <> ep Or a(2, 0) <> vid Then Resolve10 = "IDENTIDADE_AMBIGUA": Exit Function
    If a(4, 0) <> expected Then Resolve10 = "VERSAO_OBSOLETA": Exit Function
    If VerifyBatch10(a) <> 1 Then Resolve10 = "CONTEUDO_DIVERGENTE": Exit Function
    If CLng(Val10(c, "SELECT Count(*) FROM Pending WHERE EpisodeID=?", ep)) <> 0 Then Resolve10 = "PENDENTE_CONCILIACAO": Exit Function
    If Not historical Then
        If a(5, 0) <> "ACTIVE" Or a(0, 0) <> a(6, 0) Then Resolve10 = "BAIXADO": Exit Function
    End If
    Resolve10 = "ENCONTRADO"
End Function

Public Function HashBytes10(ByVal bytes As Variant) As String
    Dim hp As LongPtr, hh As LongPtr, b(0 To 31) As Byte, raw() As Byte, n As Long, i As Long, s As String, e As Long
    On Error GoTo Failed
    If CryptAcquireContextW(hp, 0, 0, 24, &HF0000000) = 0 Then Err.Raise 5, , "Crypto context"
    If CryptCreateHash(hp, &H800C&, 0, 0, hh) = 0 Then Err.Raise 5, , "Crypto SHA256"
    raw = bytes: n = UBound(raw) - LBound(raw) + 1
    If n > 0 Then If CryptHashData(hh, raw(LBound(raw)), n, 0) = 0 Then Err.Raise 5, , "Crypto data"
    n = 32
    If CryptGetHashParam(hh, 2, b(0), n, 0) = 0 Then Err.Raise 5, , "Crypto result"
    For i = 0 To 31: s = s & Right$("0" & Hex$(b(i)), 2): Next i
    HashBytes10 = LCase$(s)
    CryptDestroyHash hh: CryptReleaseContext hp, 0: Exit Function
Failed:
    e = Err.Number
    If hh <> 0 Then CryptDestroyHash hh
    If hp <> 0 Then CryptReleaseContext hp, 0
    Err.Raise e, , "SHA256 falhou"
End Function
Public Function SHA10(ByVal value As String) As String
    Dim s As Object, b As Variant
    Set s = CreateObject("ADODB.Stream"): s.Type = 2: s.Charset = "utf-8": s.Open: s.WriteText value
    s.Position = 0: s.Type = 1: s.Position = 3: b = s.Read: s.Close
    SHA10 = HashBytes10(b)
End Function
Public Function FileSHA10(ByVal file As String) As String
    Dim s As Object, b As Variant
    Set s = CreateObject("ADODB.Stream"): s.Type = 1: s.Open: s.LoadFromFile file: b = s.Read: s.Close
    FileSHA10 = HashBytes10(b)
End Function
Public Function DateOK10(ByVal s As String) As Boolean
    Dim d As Date
    On Error GoTo Invalid
    If Not s Like "####-##-##" Then Exit Function
    If CLng(Left$(s, 4)) < 1900 Or CLng(Left$(s, 4)) > 9999 Then Exit Function
    d = DateSerial(CInt(Left$(s, 4)), CInt(mid$(s, 6, 2)), CInt(Right$(s, 2)))
    DateOK10 = Year(d) = CLng(Left$(s, 4)) And Month(d) = CLng(mid$(s, 6, 2)) And Day(d) = CLng(Right$(s, 2))
Invalid:
End Function
Public Function IDOK10(ByVal s As String) As Boolean
    ' IDs F9 legados sao preservados; novos IDs devem ser GUIDs. Sem identidade por placa/NFe.
    IDOK10 = Len(s) = 36 And s Like "????????-????-????-????-????????????" And Not Replace(s, "-", "") Like "*[!0-9A-F]*"
End Function
Public Function MoneyOK10(ByVal s As String, ByRef result As Currency) As Boolean
    Dim re As Object, d As Variant
    On Error GoTo Invalid
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^-?[0-9]+(\.[0-9]{1,4})?$"
    If Not re.Test(s) Then Exit Function
    d = CDec(Replace(s, ".", Application.International(xlDecimalSeparator)))
    result = CCur(d)
    MoneyOK10 = CDec(result) = d
Invalid:
End Function
Public Function SchemaFingerprint10(ByVal c As Object) As String
    Dim cat As Object, t As Object, col As Object, k As Object, ix As Object, d As Object, a As Variant, i As Long, j As Long, tmp As String, s As String
    Set d = CreateObject("Scripting.Dictionary"): Set cat = CreateObject("ADOX.Catalog"): Set cat.ActiveConnection = c
    For Each t In cat.Tables
        If t.Type = "TABLE" Then
            For Each col In t.Columns
                tmp = t.name & ".COL." & col.name: d.Add tmp, CStr(col.Type) & ":" & col.DefinedSize & ":" & col.Attributes
            Next col
            For Each k In t.keys
                tmp = t.name & ".KEY." & k.name: s = CStr(k.Type) & ":" & k.RelatedTable & ":" & k.DeleteRule & ":" & k.UpdateRule
                For Each col In k.Columns: s = s & ":" & col.name & ">" & col.RelatedColumn: Next col
                d.Add tmp, s
            Next k
            For Each ix In t.Indexes
                tmp = t.name & ".IDX." & ix.name: s = CStr(ix.Unique) & ":" & ix.PrimaryKey
                For Each col In ix.Columns: s = s & ":" & col.name: Next col
                If Not d.Exists(tmp) Then d.Add tmp, s
            Next ix
        End If
    Next t
    a = d.keys
    For i = 0 To UBound(a) - 1: For j = i + 1 To UBound(a): If a(j) < a(i) Then tmp = a(i): a(i) = a(j): a(j) = tmp
    Next j: Next i
    s = "SCHEMA4"
    For i = 0 To UBound(a): s = s & Pack10(CStr(a(i)), CStr(d(a(i)))): Next i
    SchemaFingerprint10 = SHA10(s & RulesFingerprint10(c))
End Function
Public Function ReceiptHash10(ByVal c As Object, ByVal op As String) As String
    Dim q As Object, r As Object, s As String, i As Long
    Set q = Cmd10(c, "SELECT R.ReceiptID,R.OperationID,R.RevisionNo,R.EffectCanon,R.BaseID,R.MovementID,R.ProjectionID,R.PreviousHash,R.EpisodeID,R.VehicleID,O.EpisodeID,O.CommandHash,O.SourceHash,L.AuditHash FROM (Receipts AS R INNER JOIN Operations AS O ON R.OperationID=O.OperationID) INNER JOIN ChangeLog AS L ON O.OperationID=L.OperationID WHERE R.OperationID=?", Array(op))
    Set r = q.Execute
    If r.EOF Then ReceiptHash10 = "ERROR:MISSING": r.Close: Exit Function
    For i = 0 To r.fields.count - 1: s = s & Pack10(r.fields(i).value): Next i
    r.Close: ReceiptHash10 = SHA10(s)
End Function
Public Function AuditSlow10(ByVal c As Object) As String
    Dim r As Object, q As Object, op As String, pid As String, previous As String, n As Long, want As String, ep As String, rev As Long
    On Error GoTo Invalid
    If CLng(Val10(c, "SELECT SchemaNo FROM Meta WHERE ID=1")) <> 4 Then AuditSlow10 = "SCHEMA": Exit Function
    If CStr(Val10(c, "SELECT SchemaHash FROM Meta WHERE ID=1")) <> SchemaFingerprint10(c) Then AuditSlow10 = "SCHEMA_HASH": Exit Function
    If CLng(Val10(c, "SELECT Count(*) FROM Operations WHERE StatusText='CONFIRMED'")) <> CLng(Val10(c, "SELECT Count(*) FROM Receipts")) Then AuditSlow10 = "RECEIPT_COUNT": Exit Function
    Set r = c.Execute("SELECT O.OperationID,O.EpisodeID,O.ExpectedVersion,O.KindText,O.CommandCanon,O.CommandHash,O.SourceCanon,O.SourceHash,O.PlannedPID,O.PlannedMID,R.RevisionNo,R.EffectCanon,R.ReceiptHash,R.PreviousHash,R.BaseID,L.BeforeCanon,L.AfterCanon,L.AuditHash FROM (Operations AS O INNER JOIN Receipts AS R ON O.OperationID=R.OperationID) INNER JOIN ChangeLog AS L ON O.OperationID=L.OperationID ORDER BY R.RevisionNo")
    previous = "GENESIS"
    Do While Not r.EOF
        op = CStr(r(0)): ep = CStr(r(1)): pid = CStr(r(8)): n = n + 1: rev = CLng(r(10))
        If CommandStored10(c, op) <> CStr(r(4)) Then AuditSlow10 = "COMMAND_FIELDS": GoTo Done
        If SHA10(CStr(r(4))) <> CStr(r(5)) Or SHA10(CStr(r(6))) <> CStr(r(7)) Then AuditSlow10 = "PAYLOAD_HASH": GoTo Done
        If CanonProjection10(c, pid) <> CStr(r(11)) Or CanonProjection10(c, pid) <> CStr(Val10(c, "SELECT ContentCanon FROM Projections WHERE ProjectionID=?", pid)) Then AuditSlow10 = "HISTORY": GoTo Done
        want = Pack10(ep, CLng(r(2)) + 1, CStr(IIf(CStr(r(3)) = "SALE", "CLOSED", "ACTIVE"))) & CanonProjection10(c, pid)
        If CStr(r(15)) <> CStr(r(6)) Or CStr(r(16)) <> want Or CStr(r(17)) <> SHA10(Pack10(op, rev, CStr(r(15)), CStr(r(16)))) Then AuditSlow10 = "CHANGELOG": GoTo Done
        If CStr(r(14)) <> CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")) Or CStr(r(13)) <> previous Or rev <> n Then AuditSlow10 = "CHAIN": GoTo Done
        If CStr(r(12)) <> ReceiptHash10(c, op) Then AuditSlow10 = "RECEIPT_HASH": GoTo Done
        previous = CStr(r(12)): r.MoveNext
    Loop
    If n <> CLng(Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1")) Or n <> CLng(Val10(c, "SELECT Count(*) FROM Movements")) Or n <> CLng(Val10(c, "SELECT Count(*) FROM ChangeLog")) Then AuditSlow10 = "COUNTS" Else AuditSlow10 = "OK"
Done:
    r.Close
    If AuditSlow10 = "OK" Then AuditSlow10 = CurrentLinks10(c)
    If AuditSlow10 = "OK" Then AuditSlow10 = AttemptsCheck10(c)
    Exit Function
Invalid:
    AuditSlow10 = "ERROR:" & Err.Number & ":" & Err.Description
End Function
Public Sub SealExisting10(ByVal c As Object)
    Dim r As Object, op As String, rev As Long, prev As String, h As String
    Set r = c.Execute("SELECT OperationID,CommandCanon,SourceCanon,PlannedPID FROM Operations")
    Do While Not r.EOF
        Exec10 c, "UPDATE Operations SET CivilDate='2026-11-02',DocNumber='88001',AccessKey=?,PayloadText=?,CommandVersion=1,CommandHash=?,SourceHash=? WHERE OperationID=?", KEY10, LastPayload10(CStr(r(1))), SHA10(CStr(r(1))), SHA10(CStr(r(2))), CStr(r(0))
        r.MoveNext
    Loop: r.Close
    Set r = c.Execute("SELECT R.OperationID,R.RevisionNo,L.BeforeCanon,L.AfterCanon FROM Receipts AS R INNER JOIN ChangeLog AS L ON R.OperationID=L.OperationID ORDER BY R.RevisionNo")
    prev = "GENESIS"
    Do While Not r.EOF
        op = CStr(r(0)): rev = CLng(r(1))
        Exec10 c, "UPDATE ChangeLog SET RevisionNo=?,AuditHash=? WHERE OperationID=?", rev, SHA10(Pack10(op, rev, CStr(r(2)), CStr(r(3)))), op
        Exec10 c, "UPDATE Receipts SET BaseID=?,MovementID=?,ProjectionID=?,PreviousHash=?,EpisodeID=?,VehicleID=? WHERE OperationID=?", CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")), CStr(Val10(c, "SELECT PlannedMID FROM Operations WHERE OperationID=?", op)), CStr(Val10(c, "SELECT PlannedPID FROM Operations WHERE OperationID=?", op)), prev, CStr(Val10(c, "SELECT EpisodeID FROM Operations WHERE OperationID=?", op)), CStr(Val10(c, "SELECT VehicleID FROM Episodes WHERE EpisodeID=(SELECT EpisodeID FROM Operations WHERE OperationID=?)", op)), op
        h = ReceiptHash10(c, op): Exec10 c, "UPDATE Receipts SET ReceiptHash=? WHERE OperationID=?", h, op: prev = h
        r.MoveNext
    Loop: r.Close
End Sub
Public Sub Migrate10(ByVal name As String)
    Dim c As Object, x As Variant, r As Object, n As Long, a As Variant, h As String
    Set c = Open10(name)
    If CLng(Val10(c, "SELECT SchemaNo FROM Meta WHERE ID=1")) = 4 Then
        h = Audit10(c): c.Close: If h <> "OK" Then Err.Raise 5, , h Else Exit Sub
    End If
    ' DDL ACE nao e tratado como migracao atomicamente reversivel. Backup externo obrigatorio.
    a = BusinessRows10(c): n = VerifyBatch10(a)
    If CLng(Val10(c, "SELECT Count(*) FROM Movements")) <> CLng(Val10(c, "SELECT Count(*) FROM Receipts")) Then Err.Raise 5, , "Legado sem recibos nao migravel automaticamente"
    For Each x In Array( _
    "ALTER TABLE Episodes ADD COLUMN CurrentEpisodeID TEXT(40)", _
    "ALTER TABLE Meta ADD COLUMN WriteState TEXT(20)", "ALTER TABLE Meta ADD COLUMN IntegrityReason LONGTEXT", _
    "ALTER TABLE Meta ADD COLUMN LineageID TEXT(40)", "ALTER TABLE Meta ADD COLUMN SchemaHash TEXT(64)", _
    "ALTER TABLE Operations ADD COLUMN CivilDate TEXT(10)", "ALTER TABLE Operations ADD COLUMN DocNumber TEXT(20)", "ALTER TABLE Operations ADD COLUMN AccessKey TEXT(44)", _
    "ALTER TABLE Operations ADD COLUMN PayloadText LONGTEXT", "ALTER TABLE Operations ADD COLUMN CommandVersion INTEGER", _
    "ALTER TABLE Operations ADD COLUMN CommandHash TEXT(64)", "ALTER TABLE Operations ADD COLUMN SourceHash TEXT(64)", _
    "ALTER TABLE Attempts ADD COLUMN RequestCanon LONGTEXT", "ALTER TABLE Attempts ADD COLUMN ExpectedVersion LONG", _
    "ALTER TABLE Attempts ADD COLUMN StateBefore TEXT(30)", "ALTER TABLE Attempts ADD COLUMN StateAfter TEXT(30)", "ALTER TABLE Attempts ADD COLUMN AttemptHash TEXT(64)", _
    "ALTER TABLE Attempts ADD COLUMN OutcomeText TEXT(40)", "ALTER TABLE Attempts ADD COLUMN PayloadHash TEXT(64)", _
    "ALTER TABLE ChangeLog ADD COLUMN RevisionNo LONG", "ALTER TABLE ChangeLog ADD COLUMN AuditHash TEXT(64)", _
    "ALTER TABLE Receipts ADD COLUMN EpisodeID TEXT(40)", "ALTER TABLE Receipts ADD COLUMN VehicleID TEXT(40)", "ALTER TABLE Receipts ADD COLUMN BaseID TEXT(40)", "ALTER TABLE Receipts ADD COLUMN MovementID TEXT(40)", "ALTER TABLE Receipts ADD COLUMN ProjectionID TEXT(40)", _
    "ALTER TABLE Receipts ADD COLUMN PreviousHash TEXT(64)", "ALTER TABLE Receipts ADD COLUMN ReceiptHash TEXT(64)", _
    "CREATE TABLE ProjectionAcks (BuildID TEXT(40) CONSTRAINT PK_ACK PRIMARY KEY,BaseID TEXT(40) NOT NULL,RevisionNo LONG NOT NULL,FilePath LONGTEXT,FileHash TEXT(64),StatusText TEXT(30),LineageID TEXT(40),AuthoritySeal TEXT(64))")
        c.Execute CStr(x)
    Next x
    Exec10 c, "UPDATE Meta SET LineageID=BaseID,WriteState='READY',IntegrityReason='NONE' WHERE ID=1"
    c.Execute "UPDATE Episodes SET CurrentEpisodeID=EpisodeID WHERE CurrentPID Is Not Null"
    SealExisting10 c
    Exec10 c, "UPDATE Attempts AS A INNER JOIN Operations AS O ON A.OperationID=O.OperationID SET A.OutcomeText='MIGRATED',A.PayloadHash=O.CommandHash"
    BackfillAttempts10 c
    For Each x In Array( _
    "CREATE UNIQUE INDEX U_O_EP_ID ON Operations(EpisodeID,OperationID)", _
    "CREATE UNIQUE INDEX U_E_VID ON Episodes(EpisodeID,VehicleID)", _
    "CREATE UNIQUE INDEX U_OP_PLANMID ON Operations(OperationID,PlannedMID)", _
    "CREATE UNIQUE INDEX U_OP_PLANPID ON Operations(PlannedMID,PlannedPID)", _
    "CREATE UNIQUE INDEX U_M_EP_ID ON Movements(EpisodeID,MovementID)", _
    "CREATE UNIQUE INDEX U_M_OP_ID ON Movements(OperationID,MovementID)", _
    "CREATE UNIQUE INDEX U_P_EP_ID ON Projections(EpisodeID,ProjectionID)", _
    "CREATE UNIQUE INDEX U_P_M_ID ON Projections(MovementID,ProjectionID)", _
    "CREATE UNIQUE INDEX U_R_REV ON Receipts(RevisionNo)", _
    "CREATE UNIQUE INDEX U_LOG_OP ON ChangeLog(OperationID)", _
    "ALTER TABLE Movements ADD CONSTRAINT FK_MO_EP FOREIGN KEY(EpisodeID,OperationID) REFERENCES Operations(EpisodeID,OperationID)", _
    "ALTER TABLE Projections ADD CONSTRAINT FK_PM_EP FOREIGN KEY(EpisodeID,MovementID) REFERENCES Movements(EpisodeID,MovementID)", _
    "ALTER TABLE Episodes ADD CONSTRAINT FK_CURRENT_EP FOREIGN KEY(CurrentEpisodeID,CurrentPID) REFERENCES Projections(EpisodeID,ProjectionID)", _
    "ALTER TABLE Movements ADD CONSTRAINT FK_M_PLAN FOREIGN KEY(OperationID,MovementID) REFERENCES Operations(OperationID,PlannedMID)", _
    "ALTER TABLE Projections ADD CONSTRAINT FK_P_PLAN FOREIGN KEY(MovementID,ProjectionID) REFERENCES Operations(PlannedMID,PlannedPID)", _
    "ALTER TABLE Receipts ADD CONSTRAINT FK_R_VEHICLE FOREIGN KEY(EpisodeID,VehicleID) REFERENCES Episodes(EpisodeID,VehicleID)", _
    "ALTER TABLE Receipts ADD CONSTRAINT FK_R_MOV FOREIGN KEY(OperationID,MovementID) REFERENCES Movements(OperationID,MovementID)", _
    "ALTER TABLE Receipts ADD CONSTRAINT FK_R_PROJ FOREIGN KEY(MovementID,ProjectionID) REFERENCES Projections(MovementID,ProjectionID)", _
    "ALTER TABLE Pending ADD CONSTRAINT FK_PENDING_EP FOREIGN KEY(EpisodeID,OperationID) REFERENCES Operations(EpisodeID,OperationID)", _
    "ALTER TABLE Provenance ADD CONSTRAINT FK_PROV_EP FOREIGN KEY(EpisodeID) REFERENCES Episodes(EpisodeID)")
        c.Execute CStr(x)
    Next x
    Exec10 c, "UPDATE Meta SET SchemaNo=4 WHERE ID=1"
    h = SchemaFingerprint10(c): Exec10 c, "UPDATE Meta SET SchemaHash=? WHERE ID=1", h
    c.Close
    Rules10 name
    Set c = Open10(name): h = SchemaFingerprint10(c): Exec10 c, "UPDATE Meta SET SchemaHash=? WHERE ID=1", h: c.Close
End Sub
Public Sub CreateDB10(ByVal name As String)
    LegacyDB10 name: Migrate10 name
End Sub

Public Function RulePlan10() As Variant
    RulePlan10 = Array( _
    Array("Meta", "[ID]=1 And [SchemaNo]=4 And [RevisionNo]>=0 And [WriteState] In ('READY','RECONCILE')", "ID,BaseID,SchemaNo,RevisionNo,LineageID,SchemaHash,WriteState,IntegrityReason"), _
    Array("WriterLease", "[ID]=1 And [FenceNo]>=0 And [StatusText] In ('FREE','ACTIVE') And ([StatusText]='FREE' Or Len([OwnerID])>0)", "ID,OwnerID,FenceNo,StatusText"), _
    Array("Vehicles", "Len([VehicleID])>0 And Len([Plate])>0", "VehicleID,Plate"), _
    Array("Episodes", "[VersionNo]>=0 And [StatusText] In ('NEW','ACTIVE','CLOSED') And (([StatusText]='NEW' And [CurrentPID] Is Null And [CurrentEpisodeID] Is Null) Or ([StatusText]<>'NEW' And [CurrentPID] Is Not Null And [CurrentEpisodeID]=[EpisodeID]))", "EpisodeID,VehicleID,VersionNo,StatusText,PlaceCode"), _
    Array("Operations", "[ExpectedVersion]>=0 And [KindText] In ('ENTRY','SALE','TRANSFER','RETURN') And [StatusText] In ('PREPARED','CONFIRMED') And Len([CommandHash])=64 And Len([SourceHash])=64", "OperationID,EpisodeID,ExpectedVersion,KindText,AmountIn,AmountOut,CommandCanon,SourceCanon,PlannedPID,PlannedMID,StatusText,CommandHash,SourceHash,CivilDate,DocNumber,AccessKey,PayloadText,CommandVersion"), _
    Array("Attempts", "[FenceNo]>=0 And Len([PayloadHash])=64 And Len([AttemptHash])=64 And [ExpectedVersion]>=0", "AttemptID,OperationID,OwnerID,FenceNo,OutcomeText,PayloadHash,RequestCanon,ExpectedVersion,StateBefore,StateAfter,AttemptHash"), _
    Array("Movements", "[KindText] In ('ENTRY','SALE','TRANSFER','RETURN')", "MovementID,OperationID,EpisodeID,KindText"), _
    Array("Projections", "Len([AccessKey])=44 And Len(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace([AccessKey],'0',''),'1',''),'2',''),'3',''),'4',''),'5',''),'6',''),'7',''),'8',''),'9',''))=0 And Len([CivilDate])=10 And Val(Left([CivilDate],4))>=1900 And [CivilDate]=Format(DateSerial(Val(Left([CivilDate],4)),Val(Mid([CivilDate],6,2)),Val(Right([CivilDate],2))), 'yyyy-mm-dd') And Len([DocNumber])>0", "ProjectionID,EpisodeID,MovementID,AmountIn,AmountOut,CivilDate,DocNumber,AccessKey,PlaceCode,ContentCanon"), _
    Array("Receipts", "[RevisionNo]>0 And Len([ReceiptHash])=64 And (Len([PreviousHash])=64 Or [PreviousHash]='GENESIS')", "ReceiptID,OperationID,RevisionNo,EffectCanon,BaseID,EpisodeID,VehicleID,MovementID,ProjectionID,PreviousHash,ReceiptHash"), _
    Array("ChangeLog", "[RevisionNo]>0 And Len([AuditHash])=64", "AuditID,OperationID,BeforeCanon,AfterCanon,RevisionNo,AuditHash"), _
    Array("Pending", "Len([ReasonText])>0", "EpisodeID,OperationID,ReasonText"), _
    Array("Provenance", "[DecisionText] In ('CONFIRMADO','AMBIGUO','SEM_EVIDENCIA') And ([DecisionText]<>'CONFIRMADO' Or [EpisodeID] Is Not Null)", "LegacyRef,SourceHash,DecisionText,EvidenceText"), _
    Array("ProjectionAcks", "[RevisionNo]>=0 And [StatusText] In ('INTENT','PENDING','CONFIRMED') And ([StatusText]<>'CONFIRMED' Or Len([FileHash])=64)", "BuildID,BaseID,RevisionNo,FilePath,StatusText,LineageID,AuthoritySeal"))
End Function
Public Sub Rules10(ByVal name As String)
    Dim engine As Object, db As Object, plan As Variant, row As Variant, f As Variant
    Set engine = CreateObject("DAO.DBEngine.120"): Set db = engine.OpenDatabase(ThisWorkbook.Path & "\" & name)
    plan = RulePlan10()
    For Each row In plan
        db.TableDefs(CStr(row(0))).ValidationRule = CStr(row(1))
        db.TableDefs(CStr(row(0))).ValidationText = "Contrato sintetico F10: " & row(0)
        For Each f In Split(CStr(row(2)), ","): db.TableDefs(CStr(row(0))).fields(CStr(f)).Required = True: Next f
    Next row
    db.Close
End Sub
Public Function RulesFingerprint10(ByVal c As Object) As String
    Dim engine As Object, db As Object, row As Variant, f As Variant, plan As Variant, s As String
    Set engine = CreateObject("DAO.DBEngine.120"): Set db = engine.OpenDatabase(c.Properties("Data Source").value, False, True)
    plan = RulePlan10()
    For Each row In plan
        s = s & Pack10(CStr(row(0)), CStr(db.TableDefs(CStr(row(0))).ValidationRule))
        For Each f In Split(CStr(row(2)), ","): s = s & Pack10(CStr(f), CBool(db.TableDefs(CStr(row(0))).fields(CStr(f)).Required)): Next f
    Next row
    db.Close: RulesFingerprint10 = s
End Function

Public Function CommandStored10(ByVal c As Object, ByVal op As String) As String
    Dim q As Object, r As Object
    Set q = Cmd10(c, "SELECT EpisodeID,ExpectedVersion,KindText,AmountIn,AmountOut,PayloadText,CivilDate,DocNumber,AccessKey,PlannedPID,PlannedMID,CommandVersion FROM Operations WHERE OperationID=?", Array(op))
    Set r = q.Execute
    CommandStored10 = Command10(c, op, CStr(r(0)), CLng(r(1)), CStr(r(2)), CCur(r(3)), CCur(r(4)), CStr(r(5)), CStr(r(6)), CStr(r(7)), CStr(r(8)), CStr(r(9)), CStr(r(10)), CLng(r(11)))
    r.Close
End Function

Public Sub Assert10(ByVal id As String, ByVal ok As Boolean, ByVal effect As String, ByVal expected As String, ByVal observed As String, ByVal fixture As String)
    Check10 id, ok, effect, expected, observed, fixture
End Sub
Public Sub F10_Schema()
    Dim c As Object, c2 As Object, a As String, b As String, r As String, vin As Currency, n As Long, ep As String, op As String, mid As String, pid As String
    Call Guard10: phase10 = "F10_Schema": StartLog10 "schema_f10"
    On Error GoTo Failed
    Assert10 "F10_SHA256", SHA10("abc") = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "VERIFICADO", "SHA256 UTF8 conhecido", SHA10("abc"), "CryptoAPI"
    CreateDB10 "novo.accdb": Migrate10 "legado_f9.accdb"
    Set c = Open10("novo.accdb"): Set c2 = Open10("legado_f9.accdb")
    a = SchemaFingerprint10(c): b = SchemaFingerprint10(c2)
    Assert10 "F10_SCHEMA_EQUIVALENTE", a = b, "VERIFICADO", "Novo e migrado possuem esquema/regras identicos", a & ";" & b, "novo.accdb;legado_f9.accdb"
    Assert10 "F10_MIGRADO_AUDIT", Audit10(c2) = "OK", "VERIFICADO", "7 movimentos/recibos auditados", Audit10(c2), "legado_f9.accdb"
    b = DataSig10(c2): c2.Close: Migrate10 "legado_f9.accdb": Set c2 = Open10("legado_f9.accdb")
    Assert10 "F10_MIGRACAO_REPLAY", b = DataSig10(c2) And Audit10(c2) = "OK", "RECUSADO", "Reaplicar nao duplica ou altera fatos", Audit10(c2), "legado_f9.accdb"
    Assert10 "F10_DATAS", DateOK10("2024-02-29") And Not DateOK10("2026-02-29") And Not DateOK10("02/11/2026") And Not DateOK10("2026-13-01"), "RECUSADO", "Calendario civil ISO estrito", "2024-02-29;2026-02-29;02/11/2026;2026-13-01", "Contrato"
    Assert10 "F10_MONEY", MoneyOK10("1000.1234", vin) And vin = 1000.1234@ And Not MoneyOK10("1.12345", vin) And Not MoneyOK10("9999999999999999", vin), "RECUSADO", "Quatro casas exatas; sem arredondar quinta casa ou overflow", "1000.1234;1.12345;9999999999999999", "Contrato"
    Assert10 "F10_ID", IDOK10(NewID10()) And Not IDOK10("") And Not IDOK10("PLACA123"), "RECUSADO", "Novos IDs GUID; legado so migracao", "GUID;vazio;PLACA123", "Contrato"
    ep = "EP_B1": mid = CStr(Val10(c2, "SELECT PlannedMID FROM Operations WHERE OperationID='OP_A1'"))
    r = RejectSQL10(c2, "UPDATE Projections SET MovementID='" & mid & "' WHERE EpisodeID='EP_B1'")
    Assert10 "F10_FK_EP_MOV", r <> "NOT_REJECTED", "RECUSADO", "Nao associar B1 ao movimento A1", r, "legado_f9.accdb"
    r = RejectSQL10(c2, "UPDATE Movements SET EpisodeID='EP_B1' WHERE OperationID='OP_A1'")
    Assert10 "F10_FK_MOV_OP", r <> "NOT_REJECTED", "RECUSADO", "Episodio do movimento coincide com operacao", r, "legado_f9.accdb"
    pid = CStr(Val10(c2, "SELECT CurrentPID FROM Episodes WHERE EpisodeID='EP_A2'"))
    r = RejectSQL10(c2, "UPDATE Episodes SET CurrentPID='" & pid & "' WHERE EpisodeID='EP_B1'")
    Assert10 "F10_FK_CURRENT", r <> "NOT_REJECTED", "RECUSADO", "CurrentPID de outro episodio recusado", r, "legado_f9.accdb"
    r = RejectSQL10(c2, "UPDATE Episodes SET VehicleID='VID_B' WHERE EpisodeID='EP_A1'")
    Assert10 "F10_FK_VEHICLE_HISTORY", r <> "NOT_REJECTED", "RECUSADO", "Historico impede trocar veiculo do episodio", r, "legado_f9.accdb"
    r = RejectSQL10(c2, "DELETE FROM Episodes WHERE EpisodeID='EP_A1'")
    Assert10 "F10_DELETE_PARENT", r <> "NOT_REJECTED", "RECUSADO", "Nao excluir episodio com fatos", r, "legado_f9.accdb"
    r = RejectSQL10(c2, "UPDATE Episodes SET StatusText='UNKNOWN' WHERE EpisodeID='EP_A1'")
    Assert10 "F10_STATE_RULE", r <> "NOT_REJECTED", "RECUSADO", "Estado fora do dominio recusado no motor", r, "legado_f9.accdb"
    r = RejectSQL10(c2, "UPDATE Projections SET AccessKey='0123' WHERE EpisodeID='EP_A1'")
    Assert10 "F10_KEY_SHORT", r <> "NOT_REJECTED", "RECUSADO", "Chave curta recusada no motor", r, "legado_f9.accdb"
    r = RejectSQL10(c2, "INSERT INTO Vehicles (VehicleID,Plate) VALUES ('','FIC')")
    Assert10 "F10_EMPTY_ID_ENGINE", r <> "NOT_REJECTED", "RECUSADO", "ID vazio recusado", r, "legado_f9.accdb"
    r = RejectSQL10(c2, "INSERT INTO Episodes (EpisodeID,VehicleID,VersionNo,StatusText,PlaceCode) VALUES ('ORPHAN','ABSENT',0,'NEW','MATRIZ')")
    Assert10 "F10_ORPHAN", r <> "NOT_REJECTED", "RECUSADO", "FK recusa veiculo inexistente", r, "legado_f9.accdb"
    Snapshot10 c2, "migrado_final": c.Close: c2.Close: log10.Close: Debug.Print "F10_SCHEMA_DONE": Exit Sub
Failed:
    n = Err.Number: r = Err.Description
    On Error Resume Next
    Assert10 "F10_SCHEMA_ERROR", False, "FAIL", "Concluir esquema", n & ":" & r, "schema"
    If Not c Is Nothing Then c.Close
    If Not c2 Is Nothing Then c2.Close
    log10.Close: Debug.Print "F10_SCHEMA_ERROR"; n; r
End Sub

Public Sub F10_Core()
    Dim c As Object, ep1 As String, ep2 As String, epb As String, va As String, vb As String, op As String, st As String, n As Long, before As String, pid As String
    Call Guard10: phase10 = "F10_Core": StartLog10 "core_f10"
    On Error GoTo Failed
    CreateDB10 "ciclo_f10.accdb": Set c = Open10("ciclo_f10.accdb"): owner10 = NewID10(): fence10 = Acquire10(c, owner10)
    va = NewID10(): vb = NewID10(): ep1 = NewID10(): epb = NewID10(): ep2 = NewID10()
    Seed10 c, va, ep1, "F10A001": Seed10 c, vb, epb, "F10B001"
    st = Step10(c, NewID10(), ep1, 0, "ENTRY", 1000.1234@, 0@): Assert10 "F10_ENTRADA", st = "CONFIRMADA", "CONFIRMADO", "Entrada A1", st, "ciclo_f10.accdb"
    st = Step10(c, NewID10(), epb, 0, "ENTRY", 2200.4321@, 0@): Assert10 "F10_ENTRADA_B", st = "CONFIRMADA", "CONFIRMADO", "Entrada B", st, "ciclo_f10.accdb"
    st = Step10(c, NewID10(), ep1, 1, "SALE", 1000.1234@, 1370.9876@): Assert10 "F10_VENDA", st = "CONFIRMADA", "CONFIRMADO", "Encerra A1", st, "ciclo_f10.accdb"
    Seed10 c, va, ep2, "F10A001"
    st = Step10(c, NewID10(), ep2, 0, "ENTRY", 1500.1111@, 0@): Assert10 "F10_REENTRADA", st = "CONFIRMADA", "CONFIRMADO", "Mesmo VID novo episodio", st, "ciclo_f10.accdb"
    st = Step10(c, NewID10(), ep2, 1, "TRANSFER", 1500.1111@, 1940.2222@): Assert10 "F10_TRANSFERENCIA", st = "CONFIRMADA", "CONFIRMADO", "Preserva episodio; filial", st, "ciclo_f10.accdb"
    st = Step10(c, NewID10(), ep2, 2, "RETURN", 1500.1111@, 2100.3333@): Assert10 "F10_DEVOLUCAO", st = "CONFIRMADA", "CONFIRMADO", "Mesmo episodio; matriz", st, "ciclo_f10.accdb"
    op = NewID10(): st = Step10(c, op, ep2, 3, "TRANSFER", 1500.1111@, 2500.4444@): Assert10 "F10_TRANSFERENCIA_2", st = "CONFIRMADA", "CONFIRMADO", "Custodia filial; versao4", st, "ciclo_f10.accdb"
    before = DataSig10(c): n = CLng(Val10(c, "SELECT Count(*) FROM Attempts"))
    st = Prepare10(c, owner10, fence10, op, NewID10(), ep2, 3, "TRANSFER", 1500.1111@, 2500.4444@, "SYNTHETIC")
    Assert10 "F10_REPLAY_AUDITADO", st = "JA_CONCLUIDA" And before = DataSig10(c) And CLng(Val10(c, "SELECT Count(*) FROM Attempts")) = n + 1, "RECUSADO", "Mesmo recibo; nova tentativa auditada sem efeitos", st, "ciclo_f10.accdb"
    st = Prepare10(c, owner10, fence10, op, NewID10(), ep2, 3, "TRANSFER", 1500.1111@, 2500.4444@, "DIVERGENTE")
    Assert10 "F10_REPLAY_PAYLOAD", st = "RECUSADO_PAYLOAD" And before = DataSig10(c), "RECUSADO", "Payload divergente auditado", st, "ciclo_f10.accdb"
    st = Prepare10(c, owner10, fence10, NewID10(), NewID10(), ep2, 3, "TRANSFER", 1500@, 2600@, "SYNTHETIC")
    Assert10 "F10_VERSAO_OBSOLETA", st = "RECUSADO_VERSAO" And before = DataSig10(c), "RECUSADO", "ExpectedVersion3 nao autoriza4", st, "ciclo_f10.accdb"
    st = Prepare10(c, owner10, fence10, NewID10(), NewID10(), ep2, 4, "TRANSFER", 1500@, 2600@, "SYNTHETIC", NewID10())
    Assert10 "F10_BASE_DIVERGENTE", st = "RECUSADO_BASE" And before = DataSig10(c), "RECUSADO", "Outra BaseID nao escreve", st, "ciclo_f10.accdb"
    Assert10 "F10_CICLO_AUDIT", Audit10(c) = "OK" And CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = 7 And CLng(Val10(c, "SELECT Count(*) FROM Episodes WHERE StatusText='ACTIVE'")) = 2, "VERIFICADO", "7 movimentos/recibos;2 ativos;trilha integra", Audit10(c), "ciclo_f10.accdb"
    pid = CStr(Val10(c, "SELECT CurrentPID FROM Episodes WHERE EpisodeID=?", ep2))
    st = Resolve10(c, pid, ep2, va, 3): Assert10 "F10_LISTBOX_OBSOLETA", st = "VERSAO_OBSOLETA", "RECUSADO", "Token antigo re-resolvido", st, "ciclo_f10.accdb"
    Assert10 "F10_CHAVE_CURRENCY", CStr(Val10(c, "SELECT AccessKey FROM Projections WHERE ProjectionID=?", pid)) = KEY10 And CCur(Val10(c, "SELECT AmountOut FROM Projections WHERE ProjectionID=?", pid)) = 2500.4444@, "VERIFICADO", "Zero inicial/44 e escala4", KEY10, "ciclo_f10.accdb"
    Snapshot10 c, "ciclo_f10_final": Release10 c, owner10, fence10: c.Close
    FSO10.CopyFile ThisWorkbook.Path & "\ciclo_f10.accdb", ThisWorkbook.Path & "\backup_ciclo_f10.accdb", False
    Assert10 "F10_BACKUP_HASH", FileSHA10(ThisWorkbook.Path & "\ciclo_f10.accdb") = FileSHA10(ThisWorkbook.Path & "\backup_ciclo_f10.accdb"), "VERIFICADO", "Banco fechado copiado SHA256 igual", FileSHA10(ThisWorkbook.Path & "\backup_ciclo_f10.accdb"), "backup_ciclo_f10.accdb"
    log10.Close: Debug.Print "F10_CORE_DONE": Exit Sub
Failed:
    n = Err.Number: st = Err.Description
    On Error Resume Next
    Assert10 "F10_CORE_ERROR", False, "FAIL", "Concluir ciclo", n & ":" & st, "ciclo_f10.accdb"
    If Not c Is Nothing Then c.Close
    log10.Close: Debug.Print "F10_CORE_ERROR"; n; st
End Sub
Public Sub F10_RevalidarEsquema()
    Dim c As Object, c2 As Object, h As String, st As String
    Call Guard10: phase10 = "F10_RevalidarEsquema": StartLog10 "schema_regras_f10"
    Rules10 "novo.accdb": Rules10 "legado_f9.accdb"
    Set c = Open10("novo.accdb"): h = SchemaFingerprint10(c): Exec10 c, "UPDATE Meta SET SchemaHash=? WHERE ID=1", h
    Set c2 = Open10("legado_f9.accdb"): h = SchemaFingerprint10(c2): Exec10 c2, "UPDATE Meta SET SchemaHash=? WHERE ID=1", h
    Assert10 "F10_SCHEMA_FINAL_EQUIVALENTE", SchemaFingerprint10(c) = SchemaFingerprint10(c2) And Audit10(c2) = "OK", "VERIFICADO", "Novo/migrado equivalentes apos corrigir regra ANSI", h & ";" & Audit10(c2), "novo.accdb;legado_f9.accdb"
    st = RejectSQL10(c2, "UPDATE Projections SET AccessKey=AccessKey WHERE EpisodeID='EP_B1'")
    Assert10 "F10_REGRA_CONTROLE_POSITIVO", st = "NOT_REJECTED", "VERIFICADO", "Dados validos nao sao recusados", st, "legado_f9.accdb"
    st = RejectSQL10(c2, "UPDATE Projections SET CivilDate='2026-02-29' WHERE EpisodeID='EP_B1'")
    Assert10 "F10_DATA_INVALIDA_ENGINE", st <> "NOT_REJECTED", "RECUSADO", "Calendario invalido recusado no ACE", st, "legado_f9.accdb"
    st = RejectSQL10(c2, "UPDATE Projections SET AccessKey='0123456789012345678901234567890123456789012X' WHERE EpisodeID='EP_B1'")
    Assert10 "F10_KEY_NONDIGIT_ENGINE", st <> "NOT_REJECTED", "RECUSADO", "44chars com letra recusado", st, "legado_f9.accdb"
    st = RejectSQL10(c2, "UPDATE Projections SET MovementID=(SELECT PlannedMID FROM Operations WHERE OperationID='OP_A1') WHERE EpisodeID='EP_B1'")
    Assert10 "F10_FK_EP_MOV_FINAL", st <> "NOT_REJECTED" And InStr(st, "Projections") = 0, "RECUSADO", "Dados validos; mismatch relacional recusado pela FK", st, "legado_f9.accdb"
    c.Close: c2.Close: Migrate10 "legado_f9.accdb": log10.Close: Debug.Print "F10_SCHEMA_RULES_DONE"
End Sub

Public Function LockWriter10(ByVal c As Object, ByVal who As String, ByVal fence As Long) As Boolean
    Dim f As Object, deadline As Double
    LockWriter10 = Exec10(c, "UPDATE WriterLease SET FenceNo=FenceNo WHERE ID=1 AND OwnerID=? AND FenceNo=? AND StatusText='ACTIVE'", who, fence) = 1
    If LockWriter10 And hook10 Then
        hook10 = False: Set f = FSO10.CreateTextFile(ThisWorkbook.Path & "\prepare_locked.marker", False): f.WriteLine who & ";" & fence: f.Close
        deadline = Now10() + 35
        Do While Not FSO10.FileExists(ThisWorkbook.Path & "\worker_prepare_done.json")
            DoEvents
            If Now10() > deadline Then Err.Raise 5, , "Timeout handshake concorrencia"
        Loop
    End If
End Function

Public Function CurrentLinks10(ByVal c As Object) As String
    Dim r As Object, q As Object, x As Object
    Set r = c.Execute("SELECT EpisodeID,VersionNo,StatusText,CurrentPID FROM Episodes")
    Do While Not r.EOF
        If CLng(r(1)) > 0 Then
            Set q = Cmd10(c, "SELECT PlannedPID,KindText FROM Operations WHERE EpisodeID=? AND ExpectedVersion=? AND StatusText='CONFIRMED'", Array(CStr(r(0)), CLng(r(1)) - 1)): Set x = q.Execute
            If x.EOF Then CurrentLinks10 = "CURRENT_MISSING": GoTo Done
            If CStr(x(0)) <> CStr(r(3)) Or CStr(r(2)) <> CStr(IIf(CStr(x(1)) = "SALE", "CLOSED", "ACTIVE")) Then CurrentLinks10 = "CURRENT_IDENTITY": GoTo Done
            x.MoveNext: If Not x.EOF Then CurrentLinks10 = "CURRENT_DUPLICATE": GoTo Done
            x.Close
        End If
        r.MoveNext
    Loop
    CurrentLinks10 = "OK"
Done:
    If Not x Is Nothing Then If x.state <> 0 Then x.Close
    r.Close
End Function

Public Sub LogAttempt10(ByVal c As Object, ByVal attempt As String, ByVal op As String, ByVal who As String, ByVal fence As Long, ByVal outcome As String, ByVal canon As String, ByVal expected As Long, ByVal before As String, ByVal after As String)
    Dim h As String, ph As String
    ph = SHA10(canon): h = SHA10(Pack10(attempt, op, who, fence, outcome, ph, canon, expected, before, after))
    Exec10 c, "INSERT INTO Attempts (AttemptID,OperationID,OwnerID,FenceNo,OutcomeText,PayloadHash,RequestCanon,ExpectedVersion,StateBefore,StateAfter,AttemptHash) VALUES (?,?,?,?,?,?,?,?,?,?,?)", attempt, op, who, fence, outcome, ph, canon, expected, before, after, h
End Sub
Public Sub BackfillAttempts10(ByVal c As Object)
    Dim r As Object, s As String, ph As String, h As String, v As Long
    Set r = c.Execute("SELECT A.AttemptID,A.OperationID,A.OwnerID,A.FenceNo,O.CommandCanon,O.ExpectedVersion FROM Attempts AS A INNER JOIN Operations AS O ON A.OperationID=O.OperationID")
    Do While Not r.EOF
        s = CStr(r(4)): ph = SHA10(s): v = CLng(r(5))
        h = SHA10(Pack10(CStr(r(0)), CStr(r(1)), CStr(r(2)), CLng(r(3)), "MIGRATED", ph, s, v, "LEGACY", "LEGACY"))
        Exec10 c, "UPDATE Attempts SET RequestCanon=?,ExpectedVersion=?,StateBefore='LEGACY',StateAfter='LEGACY',AttemptHash=? WHERE AttemptID=?", s, v, h, CStr(r(0))
        r.MoveNext
    Loop: r.Close
End Sub
Public Function AttemptsCheck10(ByVal c As Object) As String
    Dim a As Variant, i As Long, h As String
    a = Rows10(c, "SELECT AttemptID,OperationID,OwnerID,FenceNo,OutcomeText,PayloadHash,RequestCanon,ExpectedVersion,StateBefore,StateAfter,AttemptHash FROM Attempts")
    AttemptsCheck10 = "OK": If IsEmpty(a) Then Exit Function
    For i = 0 To UBound(a, 2)
        h = SHA10(Pack10(CStr(a(0, i)), CStr(a(1, i)), CStr(a(2, i)), CLng(a(3, i)), CStr(a(4, i)), CStr(a(5, i)), CStr(a(6, i)), CLng(a(7, i)), CStr(a(8, i)), CStr(a(9, i))))
        If SHA10(CStr(a(6, i))) <> CStr(a(5, i)) Or h <> CStr(a(10, i)) Then AttemptsCheck10 = "ATTEMPT_HASH": Exit Function
    Next i
End Function

Public Function Audit10(ByVal c As Object) As String
    Audit10 = AuditBatch10(c): If Audit10 = "OK" Then Audit10 = PreparedAudit10(c)
    If Audit10 = "OK" Then Audit10 = C1_AuditarViews(c)
End Function
