Attribute VB_Name = "modFronteirasFase10"
Option Explicit
Private fc10 As Object, fw10 As String, ff10 As Long, fo10 As String
Public Function PreparedAudit10(ByVal c As Object) As String
    Dim a As Variant, i As Long, op As String, ep As String, canon As String, n As Long
    On Error GoTo Bad
    PreparedAudit10 = "PREPARED_LINK"
    If CLng(Val10(c, "SELECT Count(*) FROM Pending AS P INNER JOIN Operations AS O ON P.OperationID=O.OperationID WHERE O.StatusText<>'PREPARED'")) <> 0 Then Exit Function
    a = Rows10(c, "SELECT OperationID,EpisodeID,ExpectedVersion,KindText,AmountIn,AmountOut,CommandCanon,CommandHash,SourceCanon,SourceHash,PlannedPID,PlannedMID,CivilDate,DocNumber,AccessKey,PayloadText,CommandVersion FROM Operations WHERE StatusText='PREPARED'")
    If Not IsEmpty(a) Then
        For i = 0 To UBound(a, 2)
            op = CStr(a(0, i)): ep = CStr(a(1, i))
            If CLng(Val10(c, "SELECT Count(*) FROM Pending WHERE OperationID=? AND EpisodeID=?", op, ep)) <> 1 Then Exit Function
            If CLng(Val10(c, "SELECT Count(*) FROM Attempts WHERE OperationID=? AND OutcomeText='PREPARED'", op)) <> 1 Then Exit Function
            canon = Command10(c, op, ep, CLng(a(2, i)), CStr(a(3, i)), CCur(a(4, i)), CCur(a(5, i)), CStr(a(15, i)), CStr(a(12, i)), CStr(a(13, i)), CStr(a(14, i)), CStr(a(10, i)), CStr(a(11, i)), CLng(a(16, i)))
            If canon <> CStr(a(6, i)) Or SHA10(canon) <> CStr(a(7, i)) Or SHA10(CStr(a(8, i))) <> CStr(a(9, i)) Then PreparedAudit10 = "PREPARED_HASH": Exit Function
            If Source10(c, ep) <> CStr(a(8, i)) Then PreparedAudit10 = "PREPARED_SOURCE": Exit Function
            If CLng(Val10(c, "SELECT Count(*) FROM Receipts WHERE OperationID=?", op)) <> 0 Or CLng(Val10(c, "SELECT Count(*) FROM Movements WHERE OperationID=?", op)) <> 0 Then Exit Function
        Next i
    End If
    PreparedAudit10 = "OK": Exit Function
Bad:
    PreparedAudit10 = "PREPARED_ERROR:" & Err.Number & ":" & Err.Description
End Function
Public Sub F10_FronteiraInicio(ByVal stage As String)
    Dim ep As String, st As String
    Call Guard10
    If stage <> "apply" And stage <> "snapshot" Then Err.Raise 5, , "Etapa invalida"
    FSO10.CopyFile ThisWorkbook.Path & "\backup_ciclo_f10.accdb", ThisWorkbook.Path & "\fronteira_" & stage & "_f10.accdb", False
    Set fc10 = Open10("fronteira_" & stage & "_f10.accdb"): fw10 = NewID10(): ff10 = Acquire10(fc10, fw10): fo10 = NewID10()
    If stage = "apply" Then
        ep = CStr(Val10(fc10, "SELECT EpisodeID FROM Episodes WHERE VehicleID IN (SELECT VehicleID FROM Vehicles WHERE Plate='F10B001')"))
        st = Prepare10(fc10, fw10, ff10, fo10, NewID10(), ep, 1, "TRANSFER", 2200.4321@, 3900.9876@, "APPLY_DISPUTA")
        If st <> "PREPARADA" Then Err.Raise 5, , st
    End If
    Debug.Print "F10_BOUNDARY_READY:" & stage
End Sub
Public Sub F10_FronteiraExecutar(ByVal stage As String)
    Dim st As String, d As Object
    Call Guard10: phase10 = "F10_FronteiraExecutar": StartLog10 "fronteira_" & stage & "_vba_f10"
    hook10 = True
    If stage = "apply" Then
        st = Apply10(fc10, fw10, ff10, fo10)
        Assert10 "F10_APPLY_COM_DISPUTA", st = "CONFIRMADA" And Audit10(fc10) = "OK" And CLng(Val10(fc10, "SELECT RevisionNo FROM Meta WHERE ID=1")) = 8, "RECUPERADO", "Confirmacao serializada contra revogacao concorrente", st, "fronteira_apply_f10.accdb"
    Else
        Set d = SnapshotFull10(fc10, fw10, ff10)
        Assert10 "F10_SNAPSHOT_COM_DISPUTA", CLng(d("Revision")) = 7 And CLng(d("Count")) = 7 And Audit10(fc10) = "OK", "VERIFICADO", "Snapshot coerente sob lease e fronteira transacional", "Revision=" & d("Revision") & ";Count=" & d("Count"), "fronteira_snapshot_f10.accdb"
    End If
    Release10 fc10, fw10, ff10: fc10.Close: Set fc10 = Nothing: log10.Close
    Debug.Print "F10_BOUNDARY_DONE:" & stage
End Sub
Public Sub F10_IntencaoIncompleta()
    Dim c As Object, i As Long, who As String, fence As Long, op As String, ep As String, st As String, baseline As String, nm As String
    Call Guard10: phase10 = "F10_IntencaoIncompleta": StartLog10 "intencoes_f10"
    For i = 1 To 4
        nm = "intencao_" & i & "_f10.accdb"
        FSO10.CopyFile ThisWorkbook.Path & "\backup_ciclo_f10.accdb", ThisWorkbook.Path & "\" & nm, False
        Set c = Open10(nm): who = NewID10(): fence = Acquire10(c, who): op = NewID10()
        ep = CStr(Val10(c, "SELECT EpisodeID FROM Episodes WHERE VehicleID IN (SELECT VehicleID FROM Vehicles WHERE Plate='F10B001')"))
        st = Prepare10(c, who, fence, op, NewID10(), ep, 1, "TRANSFER", 2200.4321@, 3900.9876@, "INCOMPLETE")
        Assert10 "F10_INTENCAO_BASE_" & i, st = "PREPARADA" And Audit10(c) = "OK", "VERIFICADO", "Intencao integra antes da adulteracao", st, nm
        Select Case i
            Case 1: Exec10 c, "DELETE FROM Pending WHERE OperationID=?", op
            Case 2: Exec10 c, "DELETE FROM Attempts WHERE OperationID=?", op
            Case 3: Exec10 c, "UPDATE Operations SET PayloadText='TAMPER' WHERE OperationID=?", op
            Case 4: Exec10 c, "UPDATE Episodes SET PlaceCode='FILIAL' WHERE EpisodeID=?", ep
        End Select
        st = Audit10(c)
        Assert10 "F10_INTENCAO_DIVERGENTE_" & i, st <> "OK", "RECUSADO", "Diagnosticar vinculo/hash/origem incompleta", st, nm
        st = Prepare10(c, who, fence, NewID10(), NewID10(), ep, 1, "TRANSFER", 2200.4321@, 3900.9876@, "NEW_ID")
        Assert10 "F10_INTENCAO_CONCILIAR_" & i, Left$(st, Len("PENDENTE_INTEGRIDADE:")) = "PENDENTE_INTEGRIDADE:" And Val10(c, "SELECT WriteState FROM Meta WHERE ID=1") = "RECONCILE", "PENDENTE_CONCILIACAO", "Persistir bloqueio por divergencia sem restauracao ficticia", st, nm
        Snapshot10 c, "intencao_" & i & "_final": Release10 c, who, fence: c.Close
    Next i
    log10.Close: Debug.Print "F10_INTENT_DONE"
End Sub

Public Sub F10_FKTransacao()
    Dim c As Object, who As String, fence As Long, prior As String, n As Long, msg As String, st As String
    Call Guard10: phase10 = "F10_FKTransacao": StartLog10 "fk_transacao_f10"
    FSO10.CopyFile ThisWorkbook.Path & "\backup_ciclo_f10.accdb", ThisWorkbook.Path & "\fk_transacao_f10.accdb", False
    Set c = Open10("fk_transacao_f10.accdb"): who = NewID10(): fence = Acquire10(c, who): prior = DataSig10(c)
    c.BeginTrans
    If Not LockWriter10(c, who, fence) Then Err.Raise 5, , "FENCE"
    On Error Resume Next
    Exec10 c, "INSERT INTO Episodes (EpisodeID,VehicleID,VersionNo,StatusText,PlaceCode) VALUES (?,?,0,'NEW','MATRIZ')", NewID10(), NewID10()
    n = Err.Number: msg = Err.Description: Err.Clear
    On Error GoTo 0
    c.RollbackTrans
    st = Audit10(c)
    Assert10 "F10_FK_DENTRO_TRANSACAO", n <> 0 And InStr(msg, "Vehicles") > 0 And prior = DataSig10(c) And st = "OK", "RECUPERADO", "Rejeicao da FK Vehicles; rollback e estado conferidos", n & ":" & msg & ";AUDIT=" & st, "fk_transacao_f10.accdb"
    Snapshot10 c, "fk_transacao_final": Release10 c, who, fence: c.Close: log10.Close
    Debug.Print "F10_FK_TRANSACTION_DONE"
End Sub
