Attribute VB_Name = "modFechamentoEntrega2"
Option Explicit
Public Const DB_FINAL As String = "Banco_Candidato_Entrega2_Final.accdb"
Private mGate As Object, mGateOwner As String
Private mGateGeneration As Long

Private Function RawC3() As Object
    Set RawC3 = Open10(DB_FINAL)
End Function
Private Sub TempC3(ByVal folder As String)
    Dim p As String, t As String
    p = FSO10.GetFolder(folder).ShortPath: t = FSO10.GetFolder(Environ$("LOCALAPPDATA") & "\Temp").ShortPath & "\"
    If InStr(1, folder, "OneDrive", vbTextCompare) > 0 Or StrComp(Left$(p & "\", Len(t)), t, vbTextCompare) <> 0 Then Err.Raise 5, , "Somente TEMP sintetico"
End Sub
Private Function RegistryC3(ByVal c As Object) As Object
    Dim R As Object, file As String
    file = CStr(Val10(c, "SELECT RegistryPath FROM AuthorityBinding WHERE ID=1"))
    TempC3 FSO10.GetParentFolderName(file)
    If Not FSO10.FileExists(file) Then Err.Raise 5, , "Registro de autoridade ausente; nao eleger copia"
    Set R = CreateObject("ADODB.Connection"): R.ConnectionTimeout = 2: R.CommandTimeout = 2
    R.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & file & ";"
    Set RegistryC3 = R
End Function
Private Sub MatchC3(ByVal c As Object, ByVal R As Object)
    If CStr(Val10(c, "SELECT AuthorityID FROM AuthorityBinding WHERE ID=1")) <> CStr(Val10(R, "SELECT AuthorityID FROM AuthorityControl WHERE ID=1")) Then Err.Raise 5, , "AuthorityID divergente"
    If CLng(Val10(c, "SELECT Generation FROM AuthorityBinding WHERE ID=1")) <> CLng(Val10(R, "SELECT Generation FROM AuthorityControl WHERE ID=1")) Then Err.Raise 5, , "Autoridade revogada por geracao"
    If CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")) <> CStr(Val10(R, "SELECT BaseID FROM AuthorityControl WHERE ID=1")) Or CStr(Val10(c, "SELECT LineageID FROM Meta WHERE ID=1")) <> CStr(Val10(R, "SELECT LineageID FROM AuthorityControl WHERE ID=1")) Then Err.Raise 5, , "Base ou linhagem divergente"
    If Not C1_MesmoArquivoAutoridade(CStr(Val10(R, "SELECT DataPath FROM AuthorityControl WHERE ID=1")), CStr(c.Properties("Data Source").value)) Then Err.Raise 5, , "Copia de ACCDB nao e autoridade"
    If Not C1_MesmoArquivoAutoridade(CStr(Val10(R, "SELECT WorkbookPath FROM AuthorityControl WHERE ID=1")), ThisWorkbook.FullName) Then Err.Raise 5, , "Copia de XLSM sem autorizacao"
End Sub
Public Sub C3_ValidarVinculo(ByVal c As Object)
    Dim R As Object, n As Long, msg As String
    On Error GoTo Bad
    If CLng(Val10(c, "SELECT SchemaNo FROM Meta WHERE ID=1")) <> 7 Then Err.Raise 5, , "Instalacao explicita do contrato7 exigida"
    If Not mGate Is Nothing Then
        MatchC3 c, mGate
    Else
        Set R = RegistryC3(c): MatchC3 c, R: R.Close
    End If
    Exit Sub
Bad:
    n = Err.Number: msg = Err.Description
    On Error Resume Next
    If Not R Is Nothing Then R.Close
    On Error GoTo 0: Err.Raise n, "C3_ValidarVinculo", msg
End Sub
Public Sub C3_AdquirirPortao(ByVal c As Object, ByVal owner As String)
    Dim R As Object, n As Long, msg As String
    On Error GoTo Bad
    If Not mGate Is Nothing Then Err.Raise 5, , "Portao ja ocupado nesta sessao"
    Set R = RegistryC3(c): R.BeginTrans
    If Exec10(R, "UPDATE AuthorityControl SET OwnerID=? WHERE ID=1 AND OwnerID Is Null", owner) <> 1 Then Err.Raise 5, , "Outro escritor ou realocacao em curso"
    MatchC3 c, R
    mGateGeneration = CLng(Val10(R, "SELECT Generation FROM AuthorityControl WHERE ID=1"))
    Set mGate = R: mGateOwner = owner: Exit Sub
Bad:
    n = Err.Number: msg = Err.Description
    On Error Resume Next
    If Not R Is Nothing Then R.RollbackTrans: R.Close
    On Error GoTo 0: Err.Raise n, "C3_AdquirirPortao", msg
End Sub
Public Sub C3_ConferirPortao(ByVal c As Object, ByVal owner As String)
    If mGate Is Nothing Then Err.Raise 5, , "Portao transacional ausente"
    If mGateOwner <> owner Then Err.Raise 5, , "Sessao do portao divergente"
    MatchC3 c, mGate
    If CLng(Val10(mGate, "SELECT Generation FROM AuthorityControl WHERE ID=1")) <> mGateGeneration Then Err.Raise 5, , "Geracao obsoleta"
End Sub
Public Sub C3_LiberarPortao(ByVal owner As String)
    Dim e As String
    If mGate Is Nothing Then Exit Sub
    If owner <> mGateOwner Then Err.Raise 5, , "Proprietario do portao divergente"
    On Error GoTo Failed
    If Exec10(mGate, "UPDATE AuthorityControl SET OwnerID=Null WHERE ID=1 AND OwnerID=?", owner) <> 1 Then Err.Raise 5, , "CAS de liberacao recusado"
    mGate.CommitTrans: mGate.Close: Set mGate = Nothing: mGateOwner = "": Exit Sub
Failed:
    e = Err.Description
    On Error Resume Next
    mGate.RollbackTrans: mGate.Close: Set mGate = Nothing: mGateOwner = ""
    On Error GoTo 0: Err.Raise 5, "C3_LiberarPortao", "Confirmacao do portao incerta: " & e
End Sub

' Instala apenas a copia sintetica de referencia com hash e revisao conhecidos.
' Nao altera o arquivo congelado nem escolhe autoridade pelo nome/data.
Public Sub C3_InstalarReferencia(ByVal responsible As String, ByVal reason As String, ByVal expectedHash As String)
    Dim c As Object, cat As Object, R As Object, engine As Object, db As Object, registry As String, authority As String, file As String
Guard10:     TempC3 ThisWorkbook.Path
    If Len(responsible) = 0 Or Len(reason) = 0 Then Err.Raise 5, , "Responsavel e justificativa exigidos"
    file = ThisWorkbook.Path & "\" & DB_FINAL
    If FileSHA10(file) <> expectedHash Or expectedHash <> "1390a79c2b36fcea13843bb211dcb56ffbfe3fe8523e2200f70f32e1dbc35581" Then Err.Raise 5, , "Hash da referencia16 divergente"
    Set c = RawC3()
    If Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1") <> 16 Or Audit10(c) <> "OK" Then Err.Raise 5, , "Referencia nao auditada"
    If Val10(c, "SELECT Count(*) FROM Pending") <> 0 Or Val10(c, "SELECT StatusText FROM WriterLease WHERE ID=1") <> "FREE" Then Err.Raise 5, , "Pendencia ou lease impede instalacao"
    registry = FSO10.GetFolder(Environ$("TEMP")).Path & "\Controle_E2_143C55EC132B48398413268B265117AF.accdb"
    If FSO10.FileExists(registry) Then Err.Raise 5, , "Registro existente; use realocacao controlada"
    c.Close
    FSO10.CopyFile file, ThisWorkbook.Path & "\Checkpoint_Contrato6_" & NewID10() & ".accdb", False
    Set cat = CreateObject("ADOX.Catalog"): cat.Create "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & registry & ";Jet OLEDB:Engine Type=5;"
    Set R = cat.ActiveConnection
    R.Execute "CREATE TABLE AuthorityControl (ID INTEGER CONSTRAINT PK_AC PRIMARY KEY,AuthorityID TEXT(40) NOT NULL,BaseID TEXT(40) NOT NULL,LineageID TEXT(40) NOT NULL,Generation LONG NOT NULL,DataPath LONGTEXT,WorkbookPath LONGTEXT,OwnerID TEXT(40))"
    R.Execute "CREATE TABLE AuthorityHistory (EventID TEXT(40) CONSTRAINT PK_AH PRIMARY KEY,EventCanon LONGTEXT,EventHash TEXT(64))"
    Set c = RawC3()
    c.Execute "CREATE TABLE AuthorityBinding (ID INTEGER CONSTRAINT PK_AB PRIMARY KEY,AuthorityID TEXT(40) NOT NULL,RegistryPath LONGTEXT,Generation LONG NOT NULL,WorkbookPath LONGTEXT)"
    c.Execute "CREATE TABLE ReconciliationPlans (PlanID TEXT(40) CONSTRAINT PK_RP PRIMARY KEY,OperationID TEXT(40) NOT NULL,AttemptID TEXT(40) NOT NULL,SnapshotCanon LONGTEXT,PlanHash TEXT(64) NOT NULL,ActionText TEXT(40) NOT NULL,StateText TEXT(20) NOT NULL,CONSTRAINT FK_RP_O FOREIGN KEY(OperationID) REFERENCES Operations(OperationID),CONSTRAINT FK_RP_A FOREIGN KEY(AttemptID) REFERENCES Attempts(AttemptID))"
    c.Execute "CREATE TABLE ReconciliationReceipts (ResolutionID TEXT(40) CONSTRAINT PK_RR PRIMARY KEY,PlanID TEXT(40) NOT NULL,OperationID TEXT(40) NOT NULL,AttemptID TEXT(40) NOT NULL,Responsible TEXT(100) NOT NULL,ReasonText LONGTEXT,ResolutionCanon LONGTEXT,ResolutionHash TEXT(64) NOT NULL,CONSTRAINT U_RR_PLAN UNIQUE(PlanID),CONSTRAINT U_RR_OP UNIQUE(OperationID),CONSTRAINT U_RR_ATT UNIQUE(AttemptID),CONSTRAINT FK_RR_PLAN FOREIGN KEY(PlanID) REFERENCES ReconciliationPlans(PlanID),CONSTRAINT FK_RR_OP FOREIGN KEY(OperationID) REFERENCES Operations(OperationID),CONSTRAINT FK_RR_ATT FOREIGN KEY(AttemptID) REFERENCES Attempts(AttemptID))"
    authority = NewID10()
    Exec10 c, "INSERT INTO AuthorityBinding VALUES (1,?,?,1,?)", authority, registry, ThisWorkbook.FullName
    Exec10 R, "INSERT INTO AuthorityControl (ID,AuthorityID,BaseID,LineageID,Generation,DataPath,WorkbookPath) VALUES (1,?,?,?,1,?,?)", authority, CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")), CStr(Val10(c, "SELECT LineageID FROM Meta WHERE ID=1")), file, ThisWorkbook.FullName
    C3_Evento R, Pack10("INSTALACAO_LAB", authority, expectedHash, responsible, reason, file, ThisWorkbook.FullName)
    R.Close: Set cat = Nothing: c.Close
    Set engine = CreateObject("DAO.DBEngine.120"): Set db = engine.OpenDatabase(file)
    db.TableDefs("Meta").ValidationRule = Replace(db.TableDefs("Meta").ValidationRule, "[SchemaNo]=6", "[SchemaNo]=7")
    db.Close: Set db = Nothing: Set engine = Nothing
    Set c = RawC3()
    Exec10 c, "UPDATE Meta SET SchemaNo=7 WHERE ID=1"
    Exec10 c, "UPDATE CandidateSettings SET AuthorityPath=?,ContractVersion=3,ProjectionState='PENDING' WHERE ID=1", file
    c.Close: Set c = RawC3()
    Exec10 c, "UPDATE Meta SET SchemaHash=? WHERE ID=1", SchemaFingerprint10(c)
    If Audit10(c) <> "OK" Then Err.Raise 5, , "Auditoria apos instalacao: " & Audit10(c)
    c.Close: CandidatoReconstruir
    Debug.Print "INSTALACAO_EXPLICITA_CONCLUIDA"
End Sub
Private Sub C3_Evento(ByVal R As Object, ByVal canon As String)
    Exec10 R, "INSERT INTO AuthorityHistory VALUES (?,?,?)", NewID10(), canon, SHA10(canon)
End Sub
Public Function C3_Conectar() As Object
    Dim c As Object
    Set c = RawC3(): C3_ValidarVinculo c: Set C3_Conectar = c
End Function

Private Function SnapshotC3(ByVal c As Object, ByVal op As String) As String
    Dim ep As String, R As Object, i As Long, s As String
    ep = CStr(Val10(c, "SELECT EpisodeID FROM Operations WHERE OperationID=?", op))
    If Not IDOK10(ep) Then Err.Raise 5, , "OperationID inexistente"
    Set R = c.Execute("SELECT * FROM Operations WHERE OperationID='" & Replace(op, "'", "''") & "'")
    For i = 0 To R.fields.count - 1: s = s & Pack10(R.fields(i).name, Clean10(R.fields(i).value)): Next i
    R.Close
    SnapshotC3 = Pack10(CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")), CLng(Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1")), op, Source10(c, ep), s, Clean10(Val10(c, "SELECT ReasonText FROM Pending WHERE OperationID=?", op)), CLng(Val10(c, "SELECT Count(*) FROM Receipts WHERE OperationID=?", op)), CLng(Val10(c, "SELECT Count(*) FROM Movements WHERE OperationID=?", op)), CLng(Val10(c, "SELECT Count(*) FROM ChangeLog WHERE OperationID=?", op)))
End Function
Private Function ActionC3(ByVal c As Object, ByVal op As String) As String
    Dim ep As String, pid As String
    ActionC3 = "PENDENTE_CONCILIACAO"
    If Audit10(c) <> "OK" Or Val10(c, "SELECT WriteState FROM Meta WHERE ID=1") <> "READY" Then Exit Function
    If Val10(c, "SELECT StatusText FROM Operations WHERE OperationID=?", op) = "CONFIRMED" Then ActionC3 = "JA_CONCLUIDA": Exit Function
    If Val10(c, "SELECT StatusText FROM Operations WHERE OperationID=?", op) <> "PREPARED" Then Exit Function
    ep = CStr(Val10(c, "SELECT EpisodeID FROM Operations WHERE OperationID=?", op))
    ' ENTRY/rascunho tem regra de abandono ainda nao aprovada; nao liberar automaticamente.
    If Val10(c, "SELECT KindText FROM Operations WHERE OperationID=?", op) = "ENTRY" Then Exit Function
    pid = CStr(Val10(c, "SELECT PlannedPID FROM Operations WHERE OperationID=?", op))
    If CLng(Val10(c, "SELECT Count(*) FROM Projections WHERE ProjectionID=?", pid)) <> 0 Then Exit Function
    If CLng(Val10(c, "SELECT Count(*) FROM Pending WHERE OperationID=? AND EpisodeID=?", op, ep)) <> 1 Then Exit Function
    If CLng(Val10(c, "SELECT Count(*) FROM Receipts WHERE OperationID=?", op)) <> 0 Or CLng(Val10(c, "SELECT Count(*) FROM Movements WHERE OperationID=?", op)) <> 0 Or CLng(Val10(c, "SELECT Count(*) FROM ChangeLog WHERE OperationID=?", op)) <> 0 Then Exit Function
    If Source10(c, ep) <> CStr(Val10(c, "SELECT SourceCanon FROM Operations WHERE OperationID=?", op)) Then Exit Function
    ActionC3 = "ABORTAR_INTENCAO_SEM_EFEITOS"
End Function
Public Function C3_Analisar(ByVal op As String) As String
    Dim c As Object, plan As String, snap As String, attempt As String, action As String, who As String, fence As Long
    On Error GoTo Failed
    Set c = C3_Conectar(): action = ActionC3(c, op)
    If action <> "ABORTAR_INTENCAO_SEM_EFEITOS" Then c.Close: C3_Analisar = action: Exit Function
    who = NewID10(): fence = Acquire10(c, who)
    If fence = 0 Then Err.Raise 5, , "Lease ocupado; nenhuma liberacao por idade"
    c.BeginTrans
    If Not LockWriter10(c, who, fence) Then Err.Raise 5, , "Autoridade perdida"
    If ActionC3(c, op) <> action Then Err.Raise 5, , "Estado mudou durante analise"
    attempt = CStr(Val10(c, "SELECT AttemptID FROM Attempts WHERE OperationID=? AND OutcomeText='PREPARED'", op))
    plan = NewID10(): snap = SnapshotC3(c, op)
    Exec10 c, "INSERT INTO ReconciliationPlans VALUES (?,?,?,?,?,?,?)", plan, op, attempt, snap, SHA10(Pack10(plan, op, attempt, snap, action)), action, "ANALISADO"
    c.CommitTrans: Release10 c, who, fence: c.Close
    C3_Analisar = plan: Exit Function
Failed:
    C3_Analisar = "PENDENTE_CONCILIACAO: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.RollbackTrans: Release10 c, who, fence: c.Close
End Function
Public Function C3_TokenConfirmacao(ByVal plan As String) As String
    Dim c As Object
    Set c = C3_Conectar()
    C3_TokenConfirmacao = "CONFIRMAR|" & plan & "|" & CStr(Val10(c, "SELECT PlanHash FROM ReconciliationPlans WHERE PlanID=?", plan))
    c.Close
End Function
Public Function C3_Resolver(ByVal plan As String, ByVal responsible As String, ByVal reason As String, ByVal confirmation As String) As String
    Dim c As Object, who As String, fence As Long, op As String, snap As String, attempt As String, canon As String, action As String, originalAttempt As String, planHash As String, tx As Boolean, e As String
    On Error GoTo Failed
    If Len(Trim$(responsible)) = 0 Or Len(responsible) > 100 Or Len(Trim$(reason)) = 0 Then Err.Raise 5, , "Responsavel e motivo exigidos"
    Set c = C3_Conectar()
    planHash = CStr(Val10(c, "SELECT PlanHash FROM ReconciliationPlans WHERE PlanID=?", plan))
    If confirmation <> "CONFIRMAR|" & plan & "|" & planHash Then Err.Raise 5, , "Confirmacao explicita do plano exigida"
    who = NewID10(): fence = Acquire10(c, who): If fence = 0 Then Err.Raise 5, , "Lease ocupado ou orfao"
    c.BeginTrans: tx = True
    If Not LockWriter10(c, who, fence) Then Err.Raise 5, , "Autoridade obsoleta"
    op = CStr(Val10(c, "SELECT OperationID FROM ReconciliationPlans WHERE PlanID=?", plan))
    If CLng(Val10(c, "SELECT Count(*) FROM ReconciliationReceipts WHERE PlanID=?", plan)) = 1 Then
        If C3_Auditar(c) <> "OK" Then Err.Raise 5, , "Recibo de conciliacao divergente"
        c.RollbackTrans: tx = False: Release10 c, who, fence: c.Close: C3_Resolver = "JA_CONCILIADO": Exit Function
    End If
    snap = CStr(Val10(c, "SELECT SnapshotCanon FROM ReconciliationPlans WHERE PlanID=?", plan))
    originalAttempt = CStr(Val10(c, "SELECT AttemptID FROM ReconciliationPlans WHERE PlanID=?", plan))
    action = CStr(Val10(c, "SELECT ActionText FROM ReconciliationPlans WHERE PlanID=?", plan))
    If planHash <> SHA10(Pack10(plan, op, originalAttempt, snap, action)) Or snap <> SnapshotC3(c, op) Then Err.Raise 5, , "Plano obsoleto ou divergente; reanalisar"
    If ActionC3(c, op) <> "ABORTAR_INTENCAO_SEM_EFEITOS" Then Err.Raise 5, , "Efeitos incertos; nenhuma anulacao automatica"
    attempt = NewID10()
    LogAttempt10 c, attempt, op, who, fence, "ABORTED_BY_RECONCILIATION", CStr(Val10(c, "SELECT CommandCanon FROM Operations WHERE OperationID=?", op)), CLng(Val10(c, "SELECT ExpectedVersion FROM Operations WHERE OperationID=?", op)), "PREPARED", "ABORTED"
    canon = Pack10(plan, op, originalAttempt, attempt, responsible, reason, snap, planHash, Format$(Now, "yyyy-mm-dd hh:nn:ss"), "CONCILIADO_SEM_EFEITOS")
    Exec10 c, "INSERT INTO ReconciliationReceipts VALUES (?,?,?,?,?,?,?,?)", NewID10(), plan, op, attempt, responsible, reason, canon, SHA10(canon)
    If Exec10(c, "UPDATE Operations SET StatusText='ABORTED' WHERE OperationID=? AND StatusText='PREPARED'", op) <> 1 Then Err.Raise 5, , "CAS da intencao recusado"
    ' Pending e uma guarda ativa, arquivada integralmente no snapshot do recibo.
    ' Operations e Attempts originais permanecem, inclusive comando e motivo anteriores.
    If Exec10(c, "DELETE FROM Pending WHERE OperationID=?", op) <> 1 Then Err.Raise 5, , "Guarda pendente divergente"
    Exec10 c, "UPDATE ReconciliationPlans SET StateText='APLICADO' WHERE PlanID=? AND StateText='ANALISADO'", plan
    If Audit10(c) <> "OK" Then Err.Raise 5, , "Pos-validacao: " & Audit10(c)
    c.CommitTrans: tx = False: Release10 c, who, fence: c.Close
    C3_Resolver = "CONCILIADO_SEM_EFEITOS": Exit Function
Failed:
    e = Err.Description
    On Error Resume Next
    If Not c Is Nothing Then
        If tx Then c.RollbackTrans
        Release10 c, who, fence: c.Close
    End If
    C3_Resolver = "RECUSADO_OU_PENDENTE: " & e
End Function
Public Function C3_Auditar(ByVal c As Object) As String
    Dim R As Object, plan As String, op As String, attempt As String, snap As String, action As String
    On Error GoTo Failed
    C3_Auditar = "OK": If CLng(Val10(c, "SELECT SchemaNo FROM Meta WHERE ID=1")) < 7 Then Exit Function
    If CLng(Val10(c, "SELECT Count(*) FROM Operations WHERE StatusText='ABORTED'")) <> CLng(Val10(c, "SELECT Count(*) FROM ReconciliationReceipts")) Then C3_Auditar = "RECONCILIATION_COUNT": Exit Function
    Set R = c.Execute("SELECT PlanID,OperationID,AttemptID,SnapshotCanon,PlanHash,ActionText,StateText FROM ReconciliationPlans")
    Do While Not R.EOF
        If SHA10(Pack10(CStr(R(0)), CStr(R(1)), CStr(R(2)), CStr(R(3)), CStr(R(5)))) <> CStr(R(4)) Then C3_Auditar = "PLAN_HASH": GoTo Done
        R.MoveNext
    Loop
    R.Close
    Set R = c.Execute("SELECT R.PlanID,R.OperationID,R.AttemptID,R.Responsible,R.ReasonText,R.ResolutionCanon,R.ResolutionHash,P.SnapshotCanon,P.PlanHash,P.AttemptID FROM ReconciliationReceipts R INNER JOIN ReconciliationPlans P ON R.PlanID=P.PlanID")
    Do While Not R.EOF
        plan = CStr(R(0)): op = CStr(R(1)): attempt = CStr(R(2)): snap = CStr(R(7))
        If SHA10(CStr(R(5))) <> CStr(R(6)) Or Left$(CStr(R(5)), Len(Pack10(plan, op, CStr(R(9)), attempt, CStr(R(3)), CStr(R(4)), snap, CStr(R(8))))) <> Pack10(plan, op, CStr(R(9)), attempt, CStr(R(3)), CStr(R(4)), snap, CStr(R(8))) Then C3_Auditar = "RECONCILIATION_HASH": GoTo Done
        If Val10(c, "SELECT StatusText FROM Operations WHERE OperationID=?", op) <> "ABORTED" Or CLng(Val10(c, "SELECT Count(*) FROM Pending WHERE OperationID=?", op)) <> 0 Or CLng(Val10(c, "SELECT Count(*) FROM Receipts WHERE OperationID=?", op)) <> 0 Or CLng(Val10(c, "SELECT Count(*) FROM Movements WHERE OperationID=?", op)) <> 0 Then C3_Auditar = "RECONCILIATION_EFFECT": GoTo Done
        If CLng(Val10(c, "SELECT Count(*) FROM Attempts WHERE AttemptID=? AND OperationID=? AND OutcomeText='ABORTED_BY_RECONCILIATION'", attempt, op)) <> 1 Then C3_Auditar = "RECONCILIATION_ATTEMPT": GoTo Done
        R.MoveNext
    Loop
Done:
    R.Close: Exit Function
Failed:
    C3_Auditar = "RECONCILIATION_ERROR:" & Err.Description
End Function
Public Sub CandidatoConciliarIntencao()
    Dim op As String, plan As String, responsible As String, reason As String, answer As VbMsgBoxResult
    op = Trim$(InputBox("OperationID da intencao; somente laboratorio sintetico.", "Conciliacao experimental"))
    If op = "" Then Exit Sub
    plan = C3_Analisar(op)
    If Not IDOK10(plan) Then MsgBox plan, vbExclamation, "Nao disponivel para resolucao automatica": Exit Sub
    responsible = Trim$(InputBox("Responsavel experimental pela decisao:", "Responsavel"))
    reason = Trim$(InputBox("Justificativa da decisao:", "Motivo"))
    answer = MsgBox("Plano " & plan & vbCrLf & "OperationID " & op & vbCrLf & "Abortar SOMENTE a intencao comprovadamente sem efeitos?" & vbCrLf & "Comando e tentativas serao preservados. A origem sera revalidada.", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmacao explicita")
    If answer <> vbYes Then Exit Sub
    MsgBox C3_Resolver(plan, responsible, reason, C3_TokenConfirmacao(plan)), vbInformation, "Resultado da conciliacao"
End Sub

' Realocacao explicita: registry compartilhado e atualizado antes de habilitar o destino.
' Uma falha intermediaria pode exigir conciliacao; nao ha transacao entre arquivos.
Public Function C3_Realocar(ByVal destination As String, ByVal responsible As String, ByVal reason As String, ByVal expectedHash As String) As String
    Dim c As Object, d As Object, who As String, fence As Long, targetDB As String, targetBook As String, generation As Long, copyHash As String, revision As Long, sourceFile As String, e As String, published As Boolean
    On Error GoTo Failed
Guard10:     TempC3 destination
    If Len(Trim$(responsible)) = 0 Or Len(Trim$(reason)) = 0 Then Err.Raise 5, , "Responsavel e motivo exigidos"
    sourceFile = ThisWorkbook.Path & "\" & DB_FINAL
    If FileSHA10(sourceFile) <> expectedHash Then Err.Raise 5, , "Hash anterior divergente"
    targetDB = destination & "\" & DB_FINAL: targetBook = destination & "\" & ThisWorkbook.name
    If FSO10.FileExists(targetDB) Or FSO10.FileExists(targetBook) Then Err.Raise 5, , "Destino deve ser independente e vazio"
    Set c = C3_Conectar()
    If Audit10(c) <> "OK" Or Val10(c, "SELECT WriteState FROM Meta WHERE ID=1") <> "READY" Or Val10(c, "SELECT Count(*) FROM Pending") <> 0 Or Val10(c, "SELECT StatusText FROM WriterLease WHERE ID=1") <> "FREE" Then Err.Raise 5, , "Autoridade nao esta pronta para realocacao"
    revision = CLng(Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1"))
    who = NewID10(): fence = Acquire10(c, who): If fence = 0 Then Err.Raise 5, , "Escritor ativo ou orfao"
    c.Close
    ThisWorkbook.SaveCopyAs targetBook
    FSO10.CopyFile sourceFile, targetDB, False
    copyHash = FileSHA10(sourceFile)
    If FileSHA10(targetDB) <> copyHash Then Err.Raise 5, , "Copia do banco nao confere"
    Set d = CreateObject("ADODB.Connection"): d.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & targetDB & ";"
    generation = mGateGeneration + 1
    d.BeginTrans
    Exec10 d, "UPDATE AuthorityBinding SET Generation=?,WorkbookPath=? WHERE ID=1", generation, targetBook
    Exec10 d, "UPDATE CandidateSettings SET AuthorityPath=?,ProjectionState='PENDING' WHERE ID=1", targetDB
    Exec10 d, "UPDATE WriterLease SET StatusText='FREE',OwnerID='',FenceNo=FenceNo+1 WHERE ID=1"
    If Audit10(d) <> "OK" Or CLng(Val10(d, "SELECT RevisionNo FROM Meta WHERE ID=1")) <> revision Then Err.Raise 5, , "Destino nao confere"
    d.CommitTrans: d.Close: Set d = Nothing
    Set c = RawC3(): c.BeginTrans
    If Not LockWriter10(c, who, fence) Then Err.Raise 5, , "Autorizacao perdida durante realocacao"
    Exec10 c, "UPDATE Meta SET WriteState='RECONCILE',IntegrityReason='AUTHORITY_RELOCATED' WHERE ID=1"
    c.CommitTrans
    If Exec10(mGate, "UPDATE AuthorityControl SET Generation=?,DataPath=?,WorkbookPath=? WHERE ID=1 AND Generation=? AND OwnerID=?", generation, targetDB, targetBook, mGateGeneration, who) <> 1 Then Err.Raise 5, , "CAS do registro recusado"
    C3_Evento mGate, Pack10("REALOCACAO_EXPLICITA", responsible, reason, revision, copyHash, FileSHA10(targetDB), sourceFile, targetDB, targetBook, generation)
    Release10 c, who, fence: published = True: c.Close
    C3_Realocar = "REALOCADO_REQUER_REABERTURA_E_VERIFICACAO|" & targetBook: Exit Function
Failed:
    e = Err.Description
    On Error Resume Next
    If Not d Is Nothing Then d.RollbackTrans: d.Close
    If Not c Is Nothing Then
        If c.state <> 0 Then c.RollbackTrans: Release10 c, who, fence: c.Close
    End If
    C3_LiberarPortao who
    C3_Realocar = "PENDENTE_OU_RECUSADO_REALOCACAO: " & e & ";nao repetir nem apagar copias sem diagnostico"
End Function

