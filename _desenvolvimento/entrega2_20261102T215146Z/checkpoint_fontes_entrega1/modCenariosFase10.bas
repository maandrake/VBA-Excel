Attribute VB_Name = "modCenariosFase10"
Option Explicit
Private Const KEYF10 As String = "01234567890123456789012345678901234567890123"
Private sessionC10 As Object, sessionEP10 As String, sessionOP10 As String, sessionOwner10 As String, sessionFence10 As Long
Private Function Clone10(ByVal name As String) As Object
    Guard10
    FSO10.CopyFile ThisWorkbook.Path & "\backup_ciclo_f10.accdb", ThisWorkbook.Path & "\" & name, False
    Set Clone10 = Open10(name)
End Function
Public Function EpisodeB10(ByVal c As Object) As String
    If CLng(Val10(c, "SELECT Count(*) FROM Episodes AS E INNER JOIN Vehicles AS V ON E.VehicleID=V.VehicleID WHERE V.Plate='F10B001'")) <> 1 Then Err.Raise 5, , "Identidade B ambigua"
    EpisodeB10 = CStr(Val10(c, "SELECT E.EpisodeID FROM Episodes AS E INNER JOIN Vehicles AS V ON E.VehicleID=V.VehicleID WHERE V.Plate='F10B001'"))
End Function
Public Sub F10_Transacoes()
    Dim c As Object, ep As String, who As String, fence As Long, op As String, st As String, before As String, n As Long, point As Variant, nm As String, i As Long
    Call Guard10: phase10 = "F10_Transacoes": StartLog10 "transacoes_f10"
    On Error GoTo Failed
    For Each point In Array(-1, -2, 1, 2, 3, 4, 5, 6, 7, 8)
        nm = "transacao_" & Replace(CStr(point), "-", "antes") & ".accdb": Set c = Clone10(nm): ep = EpisodeB10(c): who = NewID10(): fence = Acquire10(c, who): op = NewID10()
        st = Prepare10(c, who, fence, op, NewID10(), ep, 1, "TRANSFER", 2200.4321@, 3900.9876@, "FAULT")
        If st <> "PREPARADA" Then Err.Raise 5, , st
        before = DataSig10(c): Snapshot10 c, "tx_" & Replace(CStr(point), "-", "antes") & "_before"
        st = Apply10(c, who, fence, op, CLng(point))
        If point < 8 Then
            Assert10 "F10_TX_" & point, before = DataSig10(c) And Audit10(c) = "OK" And CLng(Val10(c, "SELECT Count(*) FROM Receipts WHERE OperationID=?", op)) = 0, IIf(point = -1, "RECUSADO", "RECUPERADO"), "Efeitos iguais; intencao continua pendente; nao e falha fisica", st & ";AUDIT=" & Audit10(c), nm
            st = Prepare10(c, who, fence, NewID10(), NewID10(), ep, 1, "TRANSFER", 2200.4321@, 3900.9876@, "OTHER")
            Assert10 "F10_TX_NOVO_ID_" & point, st = "PENDENTE_CONCILIACAO", "PENDENTE_CONCILIACAO", "Novo ID nao contorna pendencia", st, nm
        Else
            Assert10 "F10_TX_APOS_COMMIT", before <> DataSig10(c) And Audit10(c) = "OK" And CLng(Val10(c, "SELECT Count(*) FROM Receipts WHERE OperationID=?", op)) = 1, "PENDENTE_CONCILIACAO", "Oitavo recibo confirmado; projecao ainda nao salva", st & ";AUDIT=" & Audit10(c), nm
        End If
        Snapshot10 c, "tx_" & Replace(CStr(point), "-", "antes") & "_after": Release10 c, who, fence: c.Close
    Next point
    log10.Close: Debug.Print "F10_TRANSACTIONS_DONE": Exit Sub
Failed:
    n = Err.Number: st = Err.Description
    On Error Resume Next
    Assert10 "F10_TX_ERROR", False, "FAIL", "Concluir cenarios independentes", n & ":" & st, nm
    If Not c Is Nothing Then c.Close
    log10.Close: Debug.Print "F10_TX_ERROR"; n; st
End Sub
Public Sub F10_AuditoriaAdversarial()
    Dim c As Object, ep As String, who As String, fence As Long, op As String, st As String, before As String, nm As String, i As Long, a As Variant, pid As String, n As Long
    Call Guard10: phase10 = "F10_AuditoriaAdversarial": StartLog10 "auditoria_f10"
    On Error GoTo Failed
    For i = 1 To 7
        nm = "audit_" & i & ".accdb": Set c = Clone10(nm): ep = EpisodeB10(c): who = NewID10(): fence = Acquire10(c, who)
        op = CStr(Val10(c, "SELECT OperationID FROM Operations WHERE EpisodeID=? AND ExpectedVersion=0", ep))
        pid = CStr(Val10(c, "SELECT PlannedPID FROM Operations WHERE OperationID=?", op))
        Select Case i
            Case 1: Exec10 c, "DELETE FROM Receipts WHERE OperationID=?", op
            Case 2: Exec10 c, "UPDATE Receipts SET ReceiptHash=? WHERE OperationID=?", String$(64, "1"), op
            Case 3: Exec10 c, "UPDATE Projections SET AmountOut=123.9876 WHERE ProjectionID=?", pid
            Case 4: Exec10 c, "UPDATE ChangeLog SET BeforeCanon='ADULTERADO' WHERE OperationID=?", op
            Case 5: Exec10 c, "UPDATE Operations SET AmountIn=123 WHERE OperationID=?", op
            Case 6: Exec10 c, "DELETE FROM ChangeLog WHERE OperationID=?", op
            Case 7: Exec10 c, "UPDATE Attempts SET OutcomeText='ADULTERADO' WHERE OperationID=?", op
        End Select
        before = DataSig10(c): st = Audit10(c)
        Assert10 "F10_AUDIT_DETECTA_" & i, st <> "OK", "PENDENTE_CONCILIACAO", "Alteracao controlada detectada; nao autentica contra adulteracao coordenada", st, nm
        st = Prepare10(c, who, fence, op, NewID10(), ep, 0, "ENTRY", 2200.4321@, 0@, "SYNTHETIC")
        Assert10 "F10_AUDIT_BLOQUEIO_DURAVEL_" & i, InStr(st, "PENDENTE_INTEGRIDADE:") = 1, "PENDENTE_CONCILIACAO", "Replay nao produz efeitos e congela autoridade", st, nm
        c.Close: Set c = Open10(nm)
        Assert10 "F10_AUDIT_FRONTEIRA_FRIA_" & i, CStr(Val10(c, "SELECT WriteState FROM Meta WHERE ID=1")) = "RECONCILE", "PENDENTE_CONCILIACAO", "Reabertura preserva exigencia de conciliacao", CStr(Val10(c, "SELECT WriteState FROM Meta WHERE ID=1")), nm
        c.Close
    Next i
    Set c = Clone10("constraints_f10.accdb")
    ep = EpisodeB10(c): who = NewID10(): fence = Acquire10(c, who): op = NewID10()
    st = Prepare10(c, who, fence, op, NewID10(), ep, 1, "TRANSFER", 2200.4321@, 3900.9876@, "CONSTRAINT")
    If st <> "PREPARADA" Then Err.Raise 5, , st
    pid = CStr(Val10(c, "SELECT EpisodeID FROM Episodes WHERE StatusText='CLOSED'"))
    a = Val10(c, "SELECT PlannedMID FROM Operations WHERE OperationID=?", op)
    st = RejectSQL10(c, "INSERT INTO Movements (MovementID,OperationID,EpisodeID,KindText) VALUES ('" & CStr(a) & "','" & op & "','" & pid & "','TRANSFER')")
    Assert10 "F10_FK_ISOLADA_OP_EP", st <> "NOT_REJECTED" And InStr(st, "Operations") > 0, "RECUSADO", "Novo movimento sem filhos; FK composta recusa episodio errado", st, "constraints_f10.accdb"
    st = RejectSQL10(c, "INSERT INTO Movements (MovementID,OperationID,EpisodeID,KindText) VALUES ('" & NewID10() & "','" & op & "','" & ep & "','TRANSFER')")
    Assert10 "F10_PLANO_MID_IMUTAVEL", st <> "NOT_REJECTED", "RECUSADO", "Movimento diferente do reservado recusado", st, "constraints_f10.accdb"
    st = RejectSQL10(c, "UPDATE Projections SET AccessKey=AccessKey")
    Assert10 "F10_CONTROLE_POSITIVO_FINAL", st = "NOT_REJECTED", "VERIFICADO", "Regra aceita dados validos; rejeicao anterior nao mascara FKs", st, "constraints_f10.accdb"
    st = RejectSQL10(c, "UPDATE Projections SET CivilDate='2026-02-29' WHERE EpisodeID='" & ep & "'")
    Assert10 "F10_CALENDARIO_ENGINE_FINAL", st <> "NOT_REJECTED", "RECUSADO", "Data real invalida recusada", st, "constraints_f10.accdb"
    st = RejectSQL10(c, "INSERT INTO Receipts (ReceiptID,OperationID,RevisionNo,EffectCanon,BaseID,EpisodeID,VehicleID,MovementID,ProjectionID,PreviousHash,ReceiptHash) SELECT ReceiptID,OperationID,RevisionNo,EffectCanon,BaseID,EpisodeID,VehicleID,MovementID,ProjectionID,PreviousHash,ReceiptHash FROM Receipts")
    Assert10 "F10_RECIBO_DUPLICADO", st <> "NOT_REJECTED", "RECUSADO", "PK/unique impedem recibo duplicado", st, "constraints_f10.accdb"
    Snapshot10 c, "constraints_final": c.Close: log10.Close: Debug.Print "F10_AUDIT_DONE": Exit Sub
Failed:
    n = Err.Number: st = Err.Description
    On Error Resume Next
    Assert10 "F10_AUDIT_ERROR", False, "FAIL", "Concluir adversariais", n & ":" & st, nm
    If Not c Is Nothing Then c.Close
    log10.Close: Debug.Print "F10_AUDIT_ERROR"; n; st
End Sub

Public Function AuditBatch10(ByVal c As Object) As String
    Dim a As Variant, e As Variant, i As Long, j As Long, n As Long, rev As Long, bid As String, prev As String, canon As String, seal As String, aft As String, rh As String, src As String, ep As String
    Dim states As Object, versions As Object, pointers As Object, beforeState As String, afterState As String, op As String, st As String
    On Error GoTo Invalid
    AuditBatch10 = "SCHEMA"
    If CLng(Val10(c, "SELECT SchemaNo FROM Meta WHERE ID=1")) <> 5 Then Exit Function
    If CStr(Val10(c, "SELECT SchemaHash FROM Meta WHERE ID=1")) <> SchemaFingerprint10(c) Then AuditBatch10 = "SCHEMA_HASH": Exit Function
    rev = CLng(Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1")): bid = CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1"))
    For Each e In Array("Receipts", "Movements", "Projections", "ChangeLog")
        If CLng(Val10(c, "SELECT Count(*) FROM " & CStr(e))) <> rev Then AuditBatch10 = "COUNT_" & e: Exit Function
    Next e
    If CLng(Val10(c, "SELECT Count(*) FROM Operations WHERE StatusText='CONFIRMED'")) <> rev Then AuditBatch10 = "COUNT_OPERATION": Exit Function
    a = Rows10(c, "SELECT O.OperationID,O.EpisodeID,O.ExpectedVersion,O.KindText,O.AmountIn,O.AmountOut,O.CommandCanon,O.CommandHash,O.SourceCanon,O.SourceHash,O.PlannedPID,O.PlannedMID,O.CivilDate,O.DocNumber,O.AccessKey,O.PayloadText,O.CommandVersion,R.ReceiptID,R.RevisionNo,R.EffectCanon,R.PreviousHash,R.ReceiptHash,R.BaseID,R.EpisodeID,R.VehicleID,L.BeforeCanon,L.AfterCanon,L.AuditHash,P.ProjectionID,P.EpisodeID,P.MovementID,P.AmountIn,P.AmountOut,P.CivilDate,P.DocNumber,P.AccessKey,P.PlaceCode,P.ContentCanon,M.MovementID,M.EpisodeID,M.OperationID,M.KindText,E.VehicleID FROM (((((Operations AS O INNER JOIN Receipts AS R ON O.OperationID=R.OperationID) INNER JOIN ChangeLog AS L ON O.OperationID=L.OperationID) INNER JOIN Projections AS P ON O.PlannedPID=P.ProjectionID) INNER JOIN Movements AS M ON O.PlannedMID=M.MovementID) INNER JOIN Episodes AS E ON O.EpisodeID=E.EpisodeID) WHERE O.StatusText='CONFIRMED' ORDER BY R.RevisionNo")
    Set states = CreateObject("Scripting.Dictionary"): Set versions = CreateObject("Scripting.Dictionary"): Set pointers = CreateObject("Scripting.Dictionary"): prev = "GENESIS"
    If Not IsEmpty(a) Then
        For i = 0 To UBound(a, 2)
            n = n + 1: op = CStr(a(0, i)): ep = CStr(a(1, i))
            canon = Pack10("SCHEMA" & CLng(a(16, i)), bid, op, ep, CLng(a(2, i)), CStr(a(3, i)), CCur(a(4, i)), CCur(a(5, i)), CStr(a(12, i)), CStr(a(13, i)), CStr(a(14, i)), CStr(a(15, i)))
            If CLng(a(16, i)) = 2 Then canon = canon & Pack10(CStr(a(10, i)), CStr(a(11, i)))
            If canon <> CStr(a(6, i)) Or SHA10(canon) <> CStr(a(7, i)) Or SHA10(CStr(a(8, i))) <> CStr(a(9, i)) Then AuditBatch10 = "COMMAND_HASH": Exit Function
            seal = Pack10(a(28, i)) & Pack10(a(29, i)) & Pack10(a(30, i)) & Pack10(a(31, i)) & Pack10(a(32, i)) & Pack10(a(33, i)) & Pack10(a(34, i)) & Pack10(a(35, i)) & Pack10(a(36, i))
            If seal <> CStr(a(37, i)) Or seal <> CStr(a(19, i)) Then AuditBatch10 = "HISTORY": Exit Function
            If a(28, i) <> a(10, i) Or a(29, i) <> ep Or a(30, i) <> a(11, i) Or a(38, i) <> a(11, i) Or a(39, i) <> ep Or a(40, i) <> op Or a(41, i) <> a(3, i) Or a(23, i) <> ep Or a(24, i) <> a(42, i) Then AuditBatch10 = "RELATION": Exit Function
            If CCur(a(4, i)) <> CCur(a(31, i)) Or CCur(a(5, i)) <> CCur(a(32, i)) Or a(12, i) <> a(33, i) Or a(13, i) <> a(34, i) Or a(14, i) <> a(35, i) Then AuditBatch10 = "EFFECT_PAYLOAD": Exit Function
            If Not DateOK10(CStr(a(33, i))) Or (Len(CStr(a(35, i))) <> 0 And Len(CStr(a(35, i))) <> 44) Or CStr(a(35, i)) Like "*[!0-9]*" Then AuditBatch10 = "DATA_CONTRACT": Exit Function
            If states.Exists(ep) Then
                src = states(ep): If CLng(a(2, i)) <> CLng(versions(ep)) Then AuditBatch10 = "EPISODE_VERSION": Exit Function
            Else
                src = Pack10(ep, CLng(0), "NEW"): If CLng(a(2, i)) <> 0 Or CStr(a(3, i)) <> "ENTRY" Then AuditBatch10 = "EPISODE_GENESIS": Exit Function
            End If
            aft = Pack10(ep, CLng(a(2, i)) + 1, CStr(IIf(CStr(a(3, i)) = "SALE", "CLOSED", "ACTIVE"))) & seal
            If CStr(a(8, i)) <> src Or CStr(a(25, i)) <> src Or CStr(a(26, i)) <> aft Or CStr(a(27, i)) <> SHA10(Pack10(op, CLng(a(18, i)), src, aft)) Then AuditBatch10 = "CHANGELOG": Exit Function
            If a(22, i) <> bid Or CLng(a(18, i)) <> n Or CStr(a(20, i)) <> prev Then AuditBatch10 = "CHAIN": Exit Function
            rh = SHA10(Pack10(a(17, i), a(0, i), a(18, i), a(19, i), a(22, i), a(11, i), a(10, i), a(20, i), a(23, i), a(24, i), a(1, i), a(7, i), a(9, i), a(27, i)))
            If rh <> CStr(a(21, i)) Then AuditBatch10 = "RECEIPT_HASH": Exit Function
            prev = rh: states(ep) = aft: versions(ep) = CLng(a(2, i)) + 1: pointers(ep) = CStr(a(10, i))
        Next i
    End If
    If n <> rev Then AuditBatch10 = "JOIN_COUNT": Exit Function
    st = CurrentLinksBatch10(c): If st <> "OK" Then AuditBatch10 = st: Exit Function
    st = AttemptsCheck10(c): If st <> "OK" Then AuditBatch10 = st: Exit Function
    AuditBatch10 = "OK": Exit Function
Invalid:
    AuditBatch10 = "ERROR:" & Err.Number & ":" & Err.Description
End Function

Public Function CurrentLinksBatch10(ByVal c As Object) As String
    Dim a As Variant, e As Variant, d As Object, i As Long, ep As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    a = Rows10(c, "SELECT EpisodeID,ExpectedVersion,KindText,PlannedPID FROM Operations WHERE StatusText='CONFIRMED'")
    If Not IsEmpty(a) Then
        For i = 0 To UBound(a, 2)
            ep = CStr(a(0, i))
            If Not d.Exists(ep) Then
                d.Add ep, Array(CLng(a(1, i)) + 1, CStr(a(3, i)), CStr(IIf(a(2, i) = "SALE", "CLOSED", "ACTIVE")))
            Else
                v = d(ep): If CLng(a(1, i)) + 1 > CLng(v(0)) Then d(ep) = Array(CLng(a(1, i)) + 1, CStr(a(3, i)), CStr(IIf(a(2, i) = "SALE", "CLOSED", "ACTIVE")))
            End If
        Next i
    End If
    a = Rows10(c, "SELECT EpisodeID,VersionNo,CurrentPID,StatusText FROM Episodes")
    If Not IsEmpty(a) Then
        For i = 0 To UBound(a, 2)
            ep = CStr(a(0, i))
            If CLng(a(1, i)) = 0 Then
                If Not IsNull(a(2, i)) Or a(3, i) <> "NEW" Or d.Exists(ep) Then CurrentLinksBatch10 = "CURRENT_GENESIS": Exit Function
            Else
                If Not d.Exists(ep) Then CurrentLinksBatch10 = "CURRENT_MISSING": Exit Function
                v = d(ep)
                If CLng(a(1, i)) <> CLng(v(0)) Or CStr(a(2, i)) <> CStr(v(1)) Or CStr(a(3, i)) <> CStr(v(2)) Then CurrentLinksBatch10 = "CURRENT_IDENTITY": Exit Function
            End If
        Next i
    End If
    CurrentLinksBatch10 = "OK"
End Function

Public Function SnapshotFull10(ByVal c As Object, ByVal who As String, ByVal fence As Long) As Object
    Dim d As Object, st As String, a As Variant, r As Variant, rev As Long, tx As Boolean, e As Long, msg As String
    On Error GoTo Failed
    c.BeginTrans: tx = True
    If Not LockWriter10(c, who, fence) Then Err.Raise 5, , "FENCE"
    If Exec10(c, "UPDATE Meta SET RevisionNo=RevisionNo WHERE ID=1 AND WriteState='READY'") <> 1 Then Err.Raise 5, , "RECONCILE"
    st = Audit10(c): If st <> "OK" Then Err.Raise 5, , st
    Set d = CreateObject("Scripting.Dictionary")
    rev = CLng(Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1"))
    d.Add "Base", CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")): d.Add "Revision", rev
    d.Add "Lineage", CStr(Val10(c, "SELECT LineageID FROM Meta WHERE ID=1")): d.Add "Schema", CStr(Val10(c, "SELECT SchemaHash FROM Meta WHERE ID=1"))
    d.Add "Seal", "GENESIS": If rev > 0 Then d("Seal") = CStr(Val10(c, "SELECT ReceiptHash FROM Receipts WHERE RevisionNo=?", rev))
    a = BusinessRows10(c): d.Add "Rows", a: r = Rows10(c, "SELECT ProjectionID,ReceiptHash FROM Receipts ORDER BY ProjectionID"): d.Add "Receipts", r
    d.Add "Count", VerifyBatch10(a): d.Add "Active", CLng(Val10(c, "SELECT Count(*) FROM Episodes WHERE StatusText='ACTIVE'"))
    c.CommitTrans: tx = False: Set SnapshotFull10 = d: Exit Function
Failed:
    e = Err.Number: msg = Err.Description
    On Error Resume Next
    If tx Then c.RollbackTrans
    On Error GoTo 0
    Err.Raise e, , msg
End Function
Public Function AnnualFormula10(ByVal n As Long) As String
    AnnualFormula10 = "=IF(COUNTIF(R2C1:R" & n + 1 & "C1,RC1)<>1,NA(),IF(RC6=""CLOSED"",""SIM"",""NAO""))"
End Function
Public Function BuildFull10(ByVal c As Object, ByVal wb As Workbook, ByVal who As String, ByVal fence As Long, ByVal build As String, Optional ByVal failure As Long = 0) As Long
    Dim d As Object, a As Variant, receipts As Variant, rd As Object, sh As Worksheet, output() As Variant, n As Long, i As Long, j As Long
    Set d = SnapshotFull10(c, who, fence): a = d("Rows"): receipts = d("Receipts"): n = CLng(d("Count"))
    Set rd = CreateObject("Scripting.Dictionary"): If n > 0 Then For i = 0 To UBound(receipts, 2): rd.Add CStr(receipts(0, i)), CStr(receipts(1, i)): Next i
    Set sh = wb.Worksheets(1): sh.name = "Projecao_F10": sh.Cells.Clear
    sh.Range("A1:S1").Value2 = Array("ProjectionID", "EpisodeID", "VehicleID", "MovementID", "Version", "EpisodeState", "CurrentPID", "Place", "Plate", "NFe", "Key", "AmountIn", "AmountOut", "CivilDate", "Canon", "Active", "Margin", "Annual", "ReceiptHash")
    If failure = 1 Then Err.Raise 7100, , "Reconstrucao interrompida apos cabecalho"
    If n > 0 Then
        ReDim output(1 To n, 1 To 15)
        For i = 0 To n - 1: For j = 0 To 14: output(i + 1, j + 1) = a(j, i): Next j: Next i
        sh.Range("A2:K" & n + 1).NumberFormat = "@": sh.Range("N2:O" & n + 1).NumberFormat = "@": sh.Range("S2:S" & n + 1).NumberFormat = "@"
        sh.Range("L2:M" & n + 1).NumberFormat = "0.0000": sh.Range("Q2:Q" & n + 1).NumberFormat = "0.0000": sh.Range("R2:R" & n + 1).NumberFormat = "General"
        sh.Range("A2:O" & n + 1).Value2 = output
        For i = 2 To n + 1: sh.Cells(i, 11).Errors(xlNumberAsText).Ignore = True: Next i
        sh.Range("P2:P" & n + 1).FormulaR1C1 = "=AND(RC6=""ACTIVE"",RC1=RC7)"
        sh.Range("Q2:Q" & n + 1).FormulaR1C1 = "=RC13-RC12": sh.Range("R2:R" & n + 1).FormulaR1C1 = AnnualFormula10(n)
        ReDim output(1 To n, 1 To 1)
        For i = 0 To n - 1: output(i + 1, 1) = rd(CStr(a(0, i))): Next i
        sh.Range("S2:S" & n + 1).Value2 = output
    End If
    sh.Range("T1:T10").NumberFormat = "@"
    sh.Range("T1").Value2 = d("Base"): sh.Range("T2").Value2 = CStr(d("Revision")): sh.Range("T3").Value2 = d("Lineage"): sh.Range("T4").Value2 = d("Schema")
    sh.Range("T5").Value2 = build: sh.Range("T6").Value2 = "BUILT": sh.Range("T7").Value2 = d("Seal"): sh.Range("T8").Value2 = CStr(n): sh.Range("T9").Value2 = CStr(n): sh.Range("T10").Value2 = CStr(n)
    sh.Range("U1").Formula = "=COUNTA(A2:A" & n + 1 & ")": sh.Range("U2").Formula = "=COUNTIF(P2:P" & n + 1 & ",TRUE)"
    sh.Range("U3").Formula = "=SUMIF(P2:P" & n + 1 & ",TRUE,L2:L" & n + 1 & ")": sh.Range("U4").Formula = "=SUMIF(P2:P" & n + 1 & ",TRUE,M2:M" & n + 1 & ")"
    sh.Calculate: BuildFull10 = n
End Function
Public Function Uniform10(ByVal rng As Range, ByVal fmt As String) As Boolean
    Dim v As Variant
    v = rng.NumberFormat: If IsNull(v) Then Exit Function
    Uniform10 = CStr(v) = fmt
End Function
Public Function VerifyFull10(ByVal c As Object, ByVal wb As Workbook, ByVal who As String, ByVal fence As Long) As String
    Dim d As Object, a As Variant, v As Variant, f As Variant, rr As Variant, rd As Object, idx As Object, seen As Object, sh As Worksheet, i As Long, j As Long, k As Long, n As Long, key As String, expected As Variant, headers As Variant, active As Long, sumIn As Currency, sumOut As Currency, ig As Variant
    On Error GoTo Invalid
    Set d = SnapshotFull10(c, who, fence): a = d("Rows"): rr = d("Receipts"): n = CLng(d("Count")): Set sh = wb.Worksheets(1)
    If sh.name <> "Projecao_F10" Then VerifyFull10 = "SHEET": Exit Function
    If CStr(sh.Range("T1").Value2) <> d("Base") Or CStr(sh.Range("T2").Value2) <> CStr(d("Revision")) Or CStr(sh.Range("T3").Value2) <> d("Lineage") Or CStr(sh.Range("T4").Value2) <> d("Schema") Or CStr(sh.Range("T7").Value2) <> d("Seal") Then VerifyFull10 = "BASE_REVISION_LINEAGE_SEAL": Exit Function
    If Not IDOK10(CStr(sh.Range("T5").Value2)) Or sh.Range("T6").Value2 <> "BUILT" Then VerifyFull10 = "BUILD_STATE": Exit Function
    If sh.Cells(sh.rows.count, 1).End(xlUp).row <> n + 1 Then VerifyFull10 = "ROW_COUNT": Exit Function
    headers = Array("ProjectionID", "EpisodeID", "VehicleID", "MovementID", "Version", "EpisodeState", "CurrentPID", "Place", "Plate", "NFe", "Key", "AmountIn", "AmountOut", "CivilDate", "Canon", "Active", "Margin", "Annual", "ReceiptHash")
    For j = 0 To UBound(headers): If sh.Cells(1, j + 1).Value2 <> headers(j) Then VerifyFull10 = "HEADER": Exit Function
    Next j
    If Not Uniform10(sh.Range("T1:T10"), "@") Then VerifyFull10 = "FOOTER_FORMAT": Exit Function
    If n = 0 Then VerifyFull10 = "VERIFIED": Exit Function
    If sh.Range("U1").Formula <> "=COUNTA(A2:A" & n + 1 & ")" Or sh.Range("U2").Formula <> "=COUNTIF(P2:P" & n + 1 & ",TRUE)" Or sh.Range("U3").Formula <> "=SUMIF(P2:P" & n + 1 & ",TRUE,L2:L" & n + 1 & ")" Or sh.Range("U4").Formula <> "=SUMIF(P2:P" & n + 1 & ",TRUE,M2:M" & n + 1 & ")" Then VerifyFull10 = "TOTAL_FORMULA": Exit Function
    For Each expected In Array("A2:K", "N2:O", "S2:S")
        If Not Uniform10(sh.Range(CStr(expected) & n + 1), "@") Then VerifyFull10 = "TEXT_FORMAT": Exit Function
    Next expected
    If Not Uniform10(sh.Range("L2:M" & n + 1), "0.0000") Or Not Uniform10(sh.Range("Q2:Q" & n + 1), "0.0000") Or Not Uniform10(sh.Range("R2:R" & n + 1), "General") Or Not Uniform10(sh.Range("P2:P" & n + 1), "General") Then VerifyFull10 = "NUMBER_FORMAT": Exit Function
    For i = 2 To n + 1
        If Not sh.Cells(i, 11).Errors(xlNumberAsText).Ignore Then VerifyFull10 = "IGNORE": Exit Function
    Next i
    Set idx = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary"): Set rd = CreateObject("Scripting.Dictionary")
    For i = 0 To n - 1: idx.Add CStr(a(0, i)), i: rd.Add CStr(rr(0, i)), CStr(rr(1, i)): Next i
    v = sh.Range("A2:S" & n + 1).Value2: f = sh.Range("P2:R" & n + 1).FormulaR1C1
    For i = 1 To n
        key = CStr(v(i, 1))
        If Not idx.Exists(key) Or seen.Exists(key) Then VerifyFull10 = "IDENTITY_DUP_MISSING": Exit Function
        seen.Add key, True: k = CLng(idx(key))
        For j = 1 To 15
            Select Case j
                Case 5: If CDbl(v(i, j)) <> CDbl(a(j - 1, k)) Then VerifyFull10 = "VERSION": Exit Function
                Case 12, 13: If CCur(v(i, j)) <> CCur(a(j - 1, k)) Then VerifyFull10 = "MONEY": Exit Function
                Case Else: If TypeName(v(i, j)) <> "String" Or CStr(v(i, j)) <> CStr(a(j - 1, k)) Then VerifyFull10 = "IDENTITY_VALUE_TYPE_" & j: Exit Function
            End Select
        Next j
        If f(i, 1) <> "=AND(RC6=""ACTIVE"",RC1=RC7)" Or f(i, 2) <> "=RC13-RC12" Or f(i, 3) <> AnnualFormula10(n) Then VerifyFull10 = "FORMULA": Exit Function
        expected = a(5, k) = "ACTIVE" And a(0, k) = a(6, k)
        If CBool(v(i, 16)) <> CBool(expected) Or CCur(v(i, 17)) <> CCur(a(12, k)) - CCur(a(11, k)) Or CStr(v(i, 18)) <> CStr(IIf(a(5, k) = "CLOSED", "SIM", "NAO")) Or CStr(v(i, 19)) <> rd(key) Then VerifyFull10 = "DERIVED_RECEIPT": Exit Function
        If expected Then active = active + 1: sumIn = sumIn + CCur(a(11, k)): sumOut = sumOut + CCur(a(12, k))
    Next i
    If active <> CLng(d("Active")) Or CLng(sh.Range("U1").Value2) <> n Or CLng(sh.Range("U2").Value2) <> active Or CCur(sh.Range("U3").Value2) <> sumIn Or CCur(sh.Range("U4").Value2) <> sumOut Then VerifyFull10 = "TOTALS": Exit Function
    For i = 8 To 10: If CStr(sh.Cells(i, 20).Value2) <> CStr(n) Then VerifyFull10 = "HISTORY_COUNTS": Exit Function
    Next i
    VerifyFull10 = "VERIFIED": Exit Function
Invalid:
    VerifyFull10 = "ERROR:" & Err.Number & ":" & Err.Description
End Function
Public Sub IntentAck10(ByVal c As Object, ByVal build As String, ByVal file As String, ByVal who As String, ByVal fence As Long)
    Dim d As Object
    Set d = SnapshotFull10(c, who, fence)
    c.BeginTrans
    If Not LockWriter10(c, who, fence) Then Err.Raise 5, , "FENCE"
    Exec10 c, "INSERT INTO ProjectionAcks (BuildID,BaseID,RevisionNo,FilePath,StatusText,LineageID,AuthoritySeal) VALUES (?,?,?,?,?,?,?)", build, CStr(d("Base")), CLng(d("Revision")), file, "INTENT", CStr(d("Lineage")), CStr(d("Seal"))
    c.CommitTrans
End Sub
Public Function ConfirmAck10(ByVal c As Object, ByVal build As String, ByVal who As String, ByVal fence As Long) As String
    Dim file As String, h As String, wb As Workbook, st As String, tx As Boolean, n As Long, e As Long, msg As String
    On Error GoTo Failed
    file = CStr(Val10(c, "SELECT FilePath FROM ProjectionAcks WHERE BuildID=?", build))
    If Left$(file, Len(ThisWorkbook.Path) + 1) <> ThisWorkbook.Path & "\" Or Not FSO10.FileExists(file) Then ConfirmAck10 = "PENDING_FILE": Exit Function
    h = FileSHA10(file): Set wb = Workbooks.Open(file, UpdateLinks:=0, readOnly:=True)
    st = VerifyFull10(c, wb, who, fence)
    If wb.Worksheets(1).Range("T5").Value2 <> build Then st = "BUILD_ID"
    wb.Close False: Set wb = Nothing
    If st <> "VERIFIED" Or h <> FileSHA10(file) Then ConfirmAck10 = "PENDING_VERIFY:" & st: Exit Function
    c.BeginTrans: tx = True
    If Not LockWriter10(c, who, fence) Then Err.Raise 5, , "FENCE"
    If Exec10(c, "UPDATE Meta SET RevisionNo=RevisionNo WHERE ID=1 AND WriteState='READY'") <> 1 Then Err.Raise 5, , "RECONCILE"
    If Audit10(c) <> "OK" Then Err.Raise 5, , "AUDIT"
    n = Exec10(c, "UPDATE ProjectionAcks SET FileHash=?,StatusText='CONFIRMED' WHERE BuildID=? AND BaseID=? AND LineageID=? AND RevisionNo=?", h, build, CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")), CStr(Val10(c, "SELECT LineageID FROM Meta WHERE ID=1")), CLng(Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1")))
    If n <> 1 Then Err.Raise 5, , "ACK_REVISION"
    c.CommitTrans: tx = False: ConfirmAck10 = "CONFIRMED:" & h: Exit Function
Failed:
    e = Err.Number: msg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If tx Then c.RollbackTrans
    ConfirmAck10 = "PENDING_ERROR:" & e & ":" & msg
End Function

Public Sub F10_Projecoes()
    Dim c As Object, wb As Workbook, sh As Worksheet, who As String, fence As Long, build As String, file As String, st As String, before As String, i As Long, n As Long, gate As clsSaveFase9, e As Long, msg As String
    Dim ev As Boolean, su As Boolean, da As Boolean, cc As XlCalculation
    Call Guard10: phase10 = "F10_Projecoes": StartLog10 "projecoes_f10"
    ev = Application.EnableEvents: su = Application.ScreenUpdating: da = Application.DisplayAlerts: cc = Application.Calculation
    On Error GoTo Failed
    Set c = Clone10("projecao_f10.accdb"): who = NewID10(): fence = Acquire10(c, who): before = DataSig10(c)
    Set wb = Workbooks.Add(xlWBATWorksheet): build = NewID10(): n = BuildFull10(c, wb, who, fence, build)
    st = VerifyFull10(c, wb, who, fence): Assert10 "F10_PROJECAO_INTEGRAL", st = "VERIFIED" And n = 7, "RECUPERADO", "7 registros/2 ativos e metadados/fatos/recibos/formulas", st, "projecao_f10.accdb"
    If st <> "VERIFIED" Then Err.Raise 7101, , "Baseline invalida: " & st
    Set sh = wb.Worksheets(1)
    For i = 1 To 13
        Select Case i
            Case 1: sh.Range("L2").Value2 = 999999
            Case 2: sh.Range("Q2").FormulaR1C1 = "=RC13-RC12+0"
            Case 3: sh.Range("L2").NumberFormat = "0.00"
            Case 4: sh.Range("K2").Value2 = String$(44, "9")
            Case 5: sh.Range("A3").Value2 = sh.Range("A2").Value2
            Case 6: sh.Range("A2").ClearContents
            Case 7: sh.Range("T2").Value2 = "6"
            Case 8: sh.Range("T3").Value2 = NewID10()
            Case 9: sh.Range("S2").Value2 = String$(64, "0")
            Case 10: sh.Range("K2").Errors(xlNumberAsText).Ignore = False
            Case 11: sh.Range("B2").Value2 = NewID10()
            Case 12: sh.Range("T6").Value2 = "CONFIRMED_EARLY"
            Case 13: sh.Range("U1").Formula = "=7"
        End Select
        sh.Calculate: st = VerifyFull10(c, wb, who, fence)
        Assert10 "F10_PROJECAO_ADULTERACAO_" & i, st <> "VERIFIED" And DataSig10(c) = before, "RECUSADO", "Verificador recusa divergencia sem escrever na autoridade", st, "projecao_f10.accdb"
        n = BuildFull10(c, wb, who, fence, build)
    Next i
    sh.Range("A1:S8").Sort Key1:=sh.Range("M2"), Order1:=xlDescending, Header:=xlYes
    st = VerifyFull10(c, wb, who, fence)
    Assert10 "F10_SORT_METADATA_DETECTADO", InStr(st, "IGNORE") > 0, "RECUSADO", "Reproduzir perda de Ignore do Sort; FAIL F9 permanece", st, "projecao_f10.accdb"
    n = BuildFull10(c, wb, who, fence, build): st = VerifyFull10(c, wb, who, fence)
    Assert10 "F10_SORT_RECONSTRUIDO", st = "VERIFIED", "RECUPERADO", "Reconstruir e conferir vinculos discriminantes/metadados", st, "projecao_f10.accdb"
    file = ThisWorkbook.Path & "\Projecao_F10_Confirmada.xlsm": IntentAck10 c, build, file, who, fence
    st = ConfirmAck10(c, build, who, fence)
    Assert10 "F10_ACK_ANTES_SAVE", st = "PENDING_FILE" And CStr(Val10(c, "SELECT StatusText FROM ProjectionAcks WHERE BuildID=?", build)) = "INTENT", "PENDENTE_CONCILIACAO", "Sem arquivo nao confirmar", st, file
    wb.SaveAs file, xlOpenXMLWorkbookMacroEnabled: wb.Close False: Set wb = Nothing
    Assert10 "F10_SAVE_SEM_ACK", CStr(Val10(c, "SELECT StatusText FROM ProjectionAcks WHERE BuildID=?", build)) = "INTENT", "PENDENTE_CONCILIACAO", "Save sozinho nao confirma projecao", "INTENT", file
    st = ConfirmAck10(c, build, who, fence)
    Assert10 "F10_ACK_REABERTO", Left$(st, 10) = "CONFIRMED:" And CStr(Val10(c, "SELECT FileHash FROM ProjectionAcks WHERE BuildID=?", build)) = FileSHA10(file), "RECUPERADO", "Reabrir somente leitura; verificar e persistir hash exato", st, file
    Set wb = Workbooks.Add(xlWBATWorksheet): build = NewID10(): n = BuildFull10(c, wb, who, fence, build)
    file = ThisWorkbook.Path & "\Projecao_F10_SaveCancelado.xlsm": wb.SaveAs file, xlOpenXMLWorkbookMacroEnabled
    build = NewID10(): IntentAck10 c, build, file, who, fence: wb.Worksheets(1).Range("T5").Value2 = build
    Set gate = New clsSaveFase9: gate.Iniciar wb: gate.RecusarUmaVez: Application.EnableEvents = True: wb.save
    Assert10 "F10_SAVE_CANCELADO_REAL", gate.Observado And Not wb.Saved, "PENDENTE_CONCILIACAO", "BeforeSave cancela chamada real", "Observado=" & gate.Observado & ";Saved=" & wb.Saved, file
    Set gate = Nothing: wb.Close False: Set wb = Nothing: st = ConfirmAck10(c, build, who, fence)
    Assert10 "F10_ACK_SAVE_CANCELADO", st = "PENDING_VERIFY:BUILD_ID" And CStr(Val10(c, "SELECT StatusText FROM ProjectionAcks WHERE BuildID=?", build)) = "INTENT", "PENDENTE_CONCILIACAO", "Nao confirmar geracao apenas em memoria", st, file
    Set wb = Workbooks.Add(xlWBATWorksheet): n = BuildFull10(c, wb, who, fence, NewID10())
    On Error Resume Next
    wb.SaveAs ThisWorkbook.Path & "\PARENT_MISSING\invalid.xlsm", xlOpenXMLWorkbookMacroEnabled
    e = Err.Number: msg = Err.Description: Err.Clear
    On Error GoTo Failed
    Assert10 "F10_SAVEAS_INVALIDO", e = 1004 And before = DataSig10(c), "PENDENTE_CONCILIACAO", "Erro real1004 sem perder fatos/recibos", e & ":" & msg, "projecao_f10.accdb"
    wb.Close False: Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error Resume Next
    n = BuildFull10(c, wb, who, fence, NewID10(), 1): e = Err.Number: msg = Err.Description: Err.Clear
    On Error GoTo Failed
    st = VerifyFull10(c, wb, who, fence)
    Assert10 "F10_REBUILD_INTERROMPIDO", e = 7100 And st <> "VERIFIED" And before = DataSig10(c), "PENDENTE_CONCILIACAO", "Cabecalho parcial nao autoriza; autoridade igual", e & ":" & st, "projecao_f10.accdb"
    wb.Close False: Set wb = Nothing: Release10 c, who, fence: Snapshot10 c, "projecao_autoridade_final": c.Close: log10.Close
    Application.EnableEvents = ev: Application.ScreenUpdating = su: Application.DisplayAlerts = da: Application.Calculation = cc
    Debug.Print "F10_PROJECTION_DONE": Exit Sub
Failed:
    e = Err.Number: msg = Err.Description
    On Error Resume Next
    Assert10 "F10_PROJECTION_ERROR", False, "FAIL", "Concluir cenarios", e & ":" & msg, "projecao_f10.accdb"
    If Not wb Is Nothing Then wb.Close False
    If Not c Is Nothing Then c.Close
    log10.Close
    Application.EnableEvents = ev: Application.ScreenUpdating = su: Application.DisplayAlerts = da: Application.Calculation = cc
    Debug.Print "F10_PROJECTION_ERROR"; e; msg
End Sub

Public Sub F10_ConcorrenciaInicio()
    Call Guard10
    Set sessionC10 = Clone10("concorrencia_f10.accdb"): sessionEP10 = EpisodeB10(sessionC10): sessionOwner10 = NewID10(): sessionFence10 = Acquire10(sessionC10, sessionOwner10): sessionOP10 = NewID10()
    Debug.Print "F10_CONC_START"; sessionOwner10; sessionFence10; sessionOP10
End Sub
Public Sub F10_ConcorrenciaPreparar()
    Dim st As String
    Call Guard10: phase10 = "F10_ConcorrenciaPreparar": StartLog10 "concorrencia_vba_f10"
    hook10 = True
    st = Prepare10(sessionC10, sessionOwner10, sessionFence10, sessionOP10, NewID10(), sessionEP10, 1, "TRANSFER", 2200.4321@, 3900.9876@, "INTERLEAVING")
    Assert10 "F10_PREPARE_COM_DISPUTA", st = "PREPARADA" And Audit10(sessionC10) = "OK", "RECUSADO", "Revogacao concorrente nao atravessa transacao de preparacao", st, "concorrencia_f10.accdb"
    Release10 sessionC10, sessionOwner10, sessionFence10
    log10.Close: Debug.Print "F10_PREPARE_RELEASED"; sessionOwner10; sessionFence10
End Sub
Public Sub F10_ConcorrenciaObsoleto(ByVal newOwner As String, ByVal newFence As Long)
    Dim st As String, before As String
    Call Guard10: phase10 = "F10_ConcorrenciaObsoleto": StartLog10 "fencing_vba_f10"
    before = DataSig10(sessionC10): st = Apply10(sessionC10, sessionOwner10, sessionFence10, sessionOP10)
    Assert10 "F10_OLD_OWNER_COMMIT", st = "RECUSADO_FENCE" And before = DataSig10(sessionC10), "RECUSADO", "A antigo nao confirma apos B obter nova geracao", st, "concorrencia_f10.accdb"
    st = Apply10(sessionC10, newOwner, newFence, sessionOP10)
    Assert10 "F10_NOVO_DONO_NAO_HERDA_INTENCAO", st = "PENDENTE_OWNER_CHANGED" And before = DataSig10(sessionC10), "PENDENTE_CONCILIACAO", "B nao conclui automaticamente intencao de A", st, "concorrencia_f10.accdb"
    st = Prepare10(sessionC10, newOwner, newFence, NewID10(), NewID10(), sessionEP10, 1, "TRANSFER", 2200.4321@, 3900.9876@, "NOVO_ID")
    Assert10 "F10_NOVO_DONO_NOVO_ID", st = "PENDENTE_CONCILIACAO", "PENDENTE_CONCILIACAO", "Novo ID nao contorna intencao orfa", st, "concorrencia_f10.accdb"
    sessionC10.Close: Set sessionC10 = Nothing: log10.Close: Debug.Print "F10_STALE_DONE"
End Sub
Public Sub F10_BackupRestauracao()
    Dim c As Object, newer As Object, wb As Workbook, who As String, fence As Long, st As String, n As Long, before As String
    Call Guard10: phase10 = "F10_BackupRestauracao": StartLog10 "backup_f10"
    Set c = Clone10("restauracao_independente_f10.accdb"): who = NewID10(): fence = Acquire10(c, who)
    Assert10 "F10_BACKUP_REABERTO", Audit10(c) = "OK" And CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = 7, "RECUPERADO", "Backup fechado restaura esquema/fatos/recibos/trilha", Audit10(c), "restauracao_independente_f10.accdb"
    Set wb = Workbooks.Add(xlWBATWorksheet): n = BuildFull10(c, wb, who, fence, NewID10()): st = VerifyFull10(c, wb, who, fence)
    Assert10 "F10_BACKUP_RECONSTRUIDO", n = 7 And st = "VERIFIED", "RECUPERADO", "Reconstrucao integral do backup restaurado", st, "restauracao_independente_f10.accdb"
    Set newer = Open10("transacao_8.accdb")
    Assert10 "F10_RESTAURACAO_FATOS_POSTERIORES", CLng(Val10(newer, "SELECT RevisionNo FROM Meta WHERE ID=1")) = 8 And CLng(Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1")) = 7, "PENDENTE_CONCILIACAO", "Copia antiga nao apaga oitavo fato; preservar ambas", "OLD=7;NEW=8;NAO_PROMOVIDO", "restauracao_independente_f10.accdb;transacao_8.accdb"
    wb.Close False: Release10 c, who, fence: c.Close: newer.Close: log10.Close: Debug.Print "F10_BACKUP_DONE"
End Sub
Public Sub F10_Frio()
    Dim c As Object, who As String, fence As Long, wb As Workbook, st As String, n As Long, r As Object, file As String, bid As String, h As String
    Call Guard10: phase10 = "F10_Frio": StartLog10 "frio_f10"
    Set c = Open10("ciclo_f10.accdb", True)
    Assert10 "F10_FRIO_DADOS_RECIBOS", Audit10(c) = "OK" And CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = 7, "RECUPERADO", "Nova instancia recupera esquema4/dados/recibos/tentativas", Audit10(c), "ciclo_f10.accdb": c.Close
    Set c = Open10("projecao_f10.accdb"): who = NewID10(): fence = Acquire10(c, who)
    Set r = c.Execute("SELECT BuildID,FilePath,FileHash FROM ProjectionAcks WHERE StatusText='CONFIRMED'")
    If r.EOF Then Err.Raise 5, , "ACK confirmado ausente"
    bid = CStr(r(0)): file = CStr(r(1)): h = CStr(r(2)): r.Close
    Set wb = Workbooks.Open(file, UpdateLinks:=0, readOnly:=True): st = VerifyFull10(c, wb, who, fence): wb.Close False
    Assert10 "F10_FRIO_ACK_PERSISTIDO", st = "VERIFIED" And FileSHA10(file) = h, "RECUPERADO", "Arquivo reaberto e hash coincide com confirmacao persistida", st & ";" & h, "projecao_f10.accdb"
    Assert10 "F10_FRIO_ACK_PENDENTE", CLng(Val10(c, "SELECT Count(*) FROM ProjectionAcks WHERE StatusText='INTENT'")) > 0, "PENDENTE_CONCILIACAO", "Save cancelado nao vira confirmacao em sessao fria", "INTENT", "projecao_f10.accdb"
    Release10 c, who, fence: c.Close
    Set c = Open10("audit_1.accdb", True): Assert10 "F10_FRIO_CONCILIACAO_DURAVEL", CStr(Val10(c, "SELECT WriteState FROM Meta WHERE ID=1")) = "RECONCILE" And Audit10(c) <> "OK", "PENDENTE_CONCILIACAO", "Autoridade divergente permanece bloqueada", Audit10(c), "audit_1.accdb": c.Close
    Set c = Open10("transacao_1.accdb", True): Assert10 "F10_FRIO_INTENCAO", CLng(Val10(c, "SELECT Count(*) FROM Pending")) = 1 And Audit10(c) = "OK", "PENDENTE_CONCILIACAO", "Intencao persiste; fatos anteriores integros", Audit10(c), "transacao_1.accdb": c.Close
    Set c = Open10("concorrencia_f10.accdb", True): Assert10 "F10_FRIO_LEASE_ORFAO", CStr(Val10(c, "SELECT StatusText FROM WriterLease WHERE ID=1")) = "ACTIVE", "PENDENTE_CONCILIACAO", "Nao liberar por horario", "ACTIVE", "concorrencia_f10.accdb": c.Close
    log10.Close: Debug.Print "F10_COLD_DONE"
End Sub

Private Sub SeedTable10(ByVal c As Object, ByVal name As String, ByVal fields As Variant, ByVal values As Variant, ByVal n As Long)
    Dim r As Object, i As Long, j As Long, row() As Variant
    Set r = CreateObject("ADODB.Recordset"): r.Open name, c, 3, 3, 2
    ReDim row(0 To UBound(fields))
    For i = 0 To n - 1
        For j = 0 To UBound(fields): row(j) = values(i, j): Next j
        r.AddNew fields, row: r.Update
    Next i: r.Close
End Sub
Private Function SeedComplete10(ByVal c As Object, ByVal n As Long, ByVal who As String, ByVal fence As Long) As Double
    Dim v As Variant, ep As Variant, ops As Variant, ats As Variant, ms As Variant, ps As Variant, rs As Variant, logs As Variant
    Dim i As Long, bid As String, vid As String, eid As String, op As String, mid As String, pid As String, aid As String, rid As String, canon As String, src As String, seal As String, aft As String, ah As String, rh As String, prev As String, ph As String, vin As Currency, vout As Currency, r As Object, t As Double
    ReDim v(0 To n - 1, 0 To 1): ReDim ep(0 To n - 1, 0 To 4): ReDim ops(0 To n - 1, 0 To 17): ReDim ats(0 To n - 1, 0 To 10): ReDim ms(0 To n - 1, 0 To 3): ReDim ps(0 To n - 1, 0 To 9): ReDim rs(0 To n - 1, 0 To 10): ReDim logs(0 To n - 1, 0 To 4)
    bid = CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")): prev = "GENESIS"
    For i = 0 To n - 1
        vid = NewID10(): eid = NewID10(): op = NewID10(): mid = NewID10(): pid = NewID10(): aid = NewID10(): rid = NewID10(): vin = CCur(i + 1) + 0.1234@: vout = CCur(i + 1) + 110.9876@
        src = Pack10(eid, CLng(0), "NEW"): canon = Pack10("SCHEMA2", bid, op, eid, CLng(0), "ENTRY", vin, vout, "2026-11-02", "88001", KEYF10, "BENCH") & Pack10(pid, mid): ph = SHA10(canon)
        seal = Pack10(pid) & Pack10(eid) & Pack10(mid) & Pack10(vin) & Pack10(vout) & Pack10("2026-11-02") & Pack10("88001") & Pack10(KEYF10) & Pack10("MATRIZ")
        aft = Pack10(eid, CLng(1), "ACTIVE") & seal: ah = SHA10(Pack10(op, CLng(i + 1), src, aft))
        rh = SHA10(Pack10(rid, op, CLng(i + 1), seal, bid, mid, pid, prev, eid, vid, eid, ph, SHA10(src), ah))
        v(i, 0) = vid: v(i, 1) = "FIC" & Format$(i + 1, "000000")
        ep(i, 0) = eid: ep(i, 1) = vid: ep(i, 2) = CLng(0): ep(i, 3) = "NEW": ep(i, 4) = "MATRIZ"
        FillSeedRow10 ops, i, Array(op, eid, CLng(0), "ENTRY", vin, vout, canon, src, pid, mid, "CONFIRMED", ph, SHA10(src), "2026-11-02", "88001", KEYF10, "BENCH", CInt(2))
        FillSeedRow10 ats, i, Array(aid, op, who, fence, "PREPARED", ph, canon, CLng(0), "ABSENT", "PREPARED", SHA10(Pack10(aid, op, who, fence, "PREPARED", ph, canon, CLng(0), "ABSENT", "PREPARED")))
        FillSeedRow10 ms, i, Array(mid, op, eid, "ENTRY")
        FillSeedRow10 ps, i, Array(pid, eid, mid, vin, vout, "2026-11-02", "88001", KEYF10, "MATRIZ", seal)
        FillSeedRow10 logs, i, Array(op, src, aft, CLng(i + 1), ah)
        FillSeedRow10 rs, i, Array(rid, op, CLng(i + 1), seal, bid, eid, vid, mid, pid, prev, rh): prev = rh
    Next i
    c.BeginTrans
    SeedTable10 c, "Vehicles", Array("VehicleID", "Plate"), v, n
    SeedTable10 c, "Episodes", Array("EpisodeID", "VehicleID", "VersionNo", "StatusText", "PlaceCode"), ep, n
    SeedTable10 c, "Operations", Array("OperationID", "EpisodeID", "ExpectedVersion", "KindText", "AmountIn", "AmountOut", "CommandCanon", "SourceCanon", "PlannedPID", "PlannedMID", "StatusText", "CommandHash", "SourceHash", "CivilDate", "DocNumber", "AccessKey", "PayloadText", "CommandVersion"), ops, n
    SeedTable10 c, "Attempts", Array("AttemptID", "OperationID", "OwnerID", "FenceNo", "OutcomeText", "PayloadHash", "RequestCanon", "ExpectedVersion", "StateBefore", "StateAfter", "AttemptHash"), ats, n
    SeedTable10 c, "Movements", Array("MovementID", "OperationID", "EpisodeID", "KindText"), ms, n
    SeedTable10 c, "Projections", Array("ProjectionID", "EpisodeID", "MovementID", "AmountIn", "AmountOut", "CivilDate", "DocNumber", "AccessKey", "PlaceCode", "ContentCanon"), ps, n
    Set r = CreateObject("ADODB.Recordset"): r.Open "Episodes", c, 3, 3, 2
    Do While Not r.EOF
        ' Uma busca no array de projecoes e evitada por indice em memoria.
        r.MoveNext
    Loop: r.Close
    c.Execute "UPDATE Episodes AS E INNER JOIN Projections AS P ON E.EpisodeID=P.EpisodeID SET E.CurrentEpisodeID=E.EpisodeID,E.CurrentPID=P.ProjectionID,E.VersionNo=1,E.StatusText='ACTIVE'"
    SeedTable10 c, "ChangeLog", Array("OperationID", "BeforeCanon", "AfterCanon", "RevisionNo", "AuditHash"), logs, n
    SeedTable10 c, "Receipts", Array("ReceiptID", "OperationID", "RevisionNo", "EffectCanon", "BaseID", "EpisodeID", "VehicleID", "MovementID", "ProjectionID", "PreviousHash", "ReceiptHash"), rs, n
    Exec10 c, "UPDATE Meta SET RevisionNo=? WHERE ID=1", n
    If Audit10(c) <> "OK" Then Err.Raise 5, , "Seed nao auditado: " & Audit10(c)
    t = Now10(): c.CommitTrans: SeedComplete10 = Now10() - t
End Function
Private Sub FillSeedRow10(ByRef target As Variant, ByVal i As Long, ByVal values As Variant)
    Dim j As Long
    For j = 0 To UBound(values): target(i, j) = values(j): Next j
End Sub
Public Sub F10_Benchmark(ByVal n As Long)
    Dim c As Object, wb As Workbook, frm As Cad_NFE, who As String, fence As Long, rep As Long, st As String, nm As String, file As String, a As Variant, list() As Variant, d As Object, i As Long, j As Long, k As Long, cnt As Long, e As Long
    Dim t As Double, seed As Double, commit As Double, query As Double, audit As Double, snap As Double, verified As Double, candidate As Double, listing As Double, bind As Double, rebuild As Double, check As Double, calc As Double, save As Double, reopen As Double, commitDiag As Double, snapObj As Object
    Dim ev As Boolean, su As Boolean, da As Boolean, cc As XlCalculation
    Call Guard10
    If n <> 100 And n <> 1000 And n <> 5000 And n <> 10000 Then Err.Raise 5, , "Volume"
    phase10 = "F10_Benchmark": StartLog10 "bench_f10_" & n
    ev = Application.EnableEvents: su = Application.ScreenUpdating: da = Application.DisplayAlerts: cc = Application.Calculation
    On Error GoTo Failed
    Application.ScreenUpdating = False
    nm = "bench_f10_" & n & ".accdb": CreateDB10 nm: Set c = Open10(nm): who = NewID10(): fence = Acquire10(c, who)
    t = Now10(): commit = SeedComplete10(c, n, who, fence): seed = Now10() - t
    For rep = 1 To 3
        t = Now10(): a = BusinessRows10(c): query = Now10() - t
        t = Now10(): st = Audit10(c): audit = Now10() - t: If st <> "OK" Then Err.Raise 5, , st
        t = Now10(): Set snapObj = SnapshotFull10(c, who, fence): snap = Now10() - t
        t = Now10(): Set d = CreateObject("Scripting.Dictionary")
        For i = 0 To n - 1: d.Add CStr(a(0, i)), i: Next i
        For i = 1 To 200: k = d(CStr(a(0, (i * 47) Mod n))): Next i: candidate = Now10() - t
        t = Now10()
        For i = 1 To 20
            k = (i * 47) Mod n: st = Resolve10(c, CStr(a(0, k)), CStr(a(1, k)), CStr(a(2, k)), CLng(a(4, k)))
            If st <> "ENCONTRADO" Then Err.Raise 5, , st
        Next i: verified = Now10() - t
        t = Now10(): a = snapObj("Rows"): ReDim list(0 To n - 1, 0 To 13)
        For i = 0 To n - 1: For j = 0 To 13: list(i, j) = a(j, i): Next j: Next i: listing = Now10() - t
        Set frm = New Cad_NFE: t = Now10(): frm.ListBoxLista.ColumnCount = 14: frm.ListBoxLista.list = list: bind = Now10() - t
        If frm.ListBoxLista.ListCount <> n Then Err.Raise 5, , "Lista"
        Unload frm: Set frm = Nothing
        Set wb = Workbooks.Add(xlWBATWorksheet): t = Now10(): cnt = BuildFull10(c, wb, who, fence, NewID10()): rebuild = Now10() - t
        t = Now10(): st = VerifyFull10(c, wb, who, fence): check = Now10() - t: If st <> "VERIFIED" Then Err.Raise 5, , st
        t = Now10(): wb.Worksheets(1).Calculate: calc = Now10() - t
        file = ThisWorkbook.Path & "\Bench_F10_" & n & "_" & rep & ".xlsm": t = Now10(): wb.SaveAs file, xlOpenXMLWorkbookMacroEnabled: save = Now10() - t: wb.Close False
        t = Now10(): Set wb = Workbooks.Open(file, UpdateLinks:=0, readOnly:=True): reopen = Now10() - t
        st = VerifyFull10(c, wb, who, fence): wb.Close False: Set wb = Nothing
        t = Now10(): c.BeginTrans: If Not LockWriter10(c, who, fence) Then Err.Raise 5, , "Fence"
        Exec10 c, "UPDATE Meta SET IntegrityReason=? WHERE ID=1", "BENCH_" & rep: c.CommitTrans: commitDiag = Now10() - t
        Assert10 "F10_BENCH_" & n & "_R" & rep, cnt = n And st = "VERIFIED", "VERIFICADO", "Mesmo esquema completo; 3 amostras por volume; tempos segundos", "SEED=" & seed & ";COMMIT_SEED=" & commit & ";QUERY=" & query & ";AUDIT=" & audit & ";SNAPSHOT=" & snap & ";INDEX_BUILD_CANDIDATE200=" & candidate & ";RESOLVE20=" & verified & ";LIST_FROM_SNAPSHOT=" & listing & ";MSFORMS=" & bind & ";REBUILD=" & rebuild & ";VERIFY_FULL=" & check & ";RECALC=" & calc & ";SAVE=" & save & ";REOPEN=" & reopen & ";COMMIT_DIAG=" & commitDiag, nm
    Next rep
    Release10 c, who, fence: c.Close: log10.Close
    Application.EnableEvents = ev: Application.ScreenUpdating = su: Application.DisplayAlerts = da: Application.Calculation = cc
    Debug.Print "F10_BENCH_DONE"; n: Exit Sub
Failed:
    e = Err.Number: st = Err.Description
    On Error Resume Next
    Assert10 "F10_BENCH_ERROR_" & n, False, "FAIL", "Concluir volume", e & ":" & st, nm
    If Not frm Is Nothing Then Unload frm
    If Not wb Is Nothing Then wb.Close False
    If Not c Is Nothing Then c.RollbackTrans: c.Close
    log10.Close
    Application.EnableEvents = ev: Application.ScreenUpdating = su: Application.DisplayAlerts = da: Application.Calculation = cc
    Debug.Print "F10_BENCH_ERROR"; n; e; st
End Sub
