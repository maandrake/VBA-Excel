Attribute VB_Name = "modVerificacaoFase9"
Option Explicit
' Verificacao adicional independente, somente no laboratorio sintetico.
Private Function Root9X() As String
    Dim f As Object, a As String, t As String
    Set f = CreateObject("Scripting.FileSystemObject")
    a = f.GetFolder(ThisWorkbook.Path).ShortPath & "\": t = f.GetFolder(Environ$("TEMP")).ShortPath & "\"
    If ThisWorkbook.Name <> "Fixture_Persistencia_Fase9_Mestre.xlsm" Or InStr(1, ThisWorkbook.Path, "VBA_Fase9_20261102_142753Z", vbTextCompare) = 0 Or InStr(1, a, "OneDrive", vbTextCompare) > 0 Or StrComp(Left$(a, Len(t)), t, vbTextCompare) <> 0 Then Err.Raise 5, , "Fixture TEMP F9 exigida"
    If ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then Err.Raise 5, , "Marcador"
    Root9X = ThisWorkbook.Path & "\"
End Function
Private Function Open9X(ByVal name As String) As Object
    Dim c As Object
    Set c = CreateObject("ADODB.Connection"): c.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & Root9X() & name & ";": Set Open9X = c
End Function
Private Function Scalar9X(ByVal c As Object, ByVal sql As String) As Variant
    Dim r As Object
    Set r = c.Execute(sql): If Not r.EOF Then Scalar9X = r.Fields(0).Value
    r.Close
End Function
Private Sub C9X(ByVal f As Object, ByVal id As String, ByVal ok As Boolean, ByVal treatment As String, ByVal expected As String, ByVal actual As String, ByVal fixture As String)
    f.WriteLine id & vbTab & IIf(ok, "PASS", "FAIL") & vbTab & treatment & vbTab & expected & vbTab & Replace(Replace(Replace(actual, vbTab, " "), vbCr, " "), vbLf, " ") & vbTab & fixture & vbTab & "F9_Complemento"
End Sub
Private Function VerifyProjection9X(ByVal c As Object, ByVal wb As Workbook) As String
    Dim sh As Worksheet, r As Object, d As Object, i As Long, a As Variant, pid As String, expected As Long, seen As Object
    Set sh = wb.Worksheets(1)
    If CStr(sh.Range("T1").Value2) <> CStr(Scalar9X(c, "SELECT BaseID FROM Meta WHERE ID=1")) Then VerifyProjection9X = "RECUSADO_BASE": Exit Function
    If sh.Range("T2").Value2 <> Scalar9X(c, "SELECT RevisionNo FROM Meta WHERE ID=1") Then VerifyProjection9X = "RECUSADO_REVISAO": Exit Function
    Set d = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    Set r = c.Execute("SELECT P.ProjectionID,P.EpisodeID,E.VehicleID,P.AmountIn,P.AmountOut,P.AccessKey,E.VersionNo,E.StatusText,E.CurrentPID FROM Projections AS P INNER JOIN Episodes AS E ON P.EpisodeID=E.EpisodeID")
    Do While Not r.EOF
        d.Add CStr(r.Fields(0).Value), Array(r.Fields(1).Value, r.Fields(2).Value, r.Fields(3).Value, r.Fields(4).Value, r.Fields(5).Value, r.Fields(6).Value, r.Fields(7).Value, r.Fields(8).Value)
        r.MoveNext
    Loop
    r.Close
    For i = 2 To sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
        pid = CStr(sh.Cells(i, 1).Value2)
        If Not d.Exists(pid) Or seen.Exists(pid) Then VerifyProjection9X = "RECUSADO_IDENTIDADE": Exit Function
        seen.Add pid, True: a = d(pid)
        If sh.Cells(i, 2).Value2 <> a(0) Or sh.Cells(i, 3).Value2 <> a(1) Or CCur(sh.Cells(i, 12).Value2) <> a(2) Or CCur(sh.Cells(i, 13).Value2) <> a(3) Or sh.Cells(i, 11).Value2 <> a(4) Or sh.Cells(i, 5).Value2 <> a(5) Or sh.Cells(i, 6).Value2 <> a(6) Or sh.Cells(i, 7).Value2 <> a(7) Then VerifyProjection9X = "RECUSADO_CONTEUDO": Exit Function
        If TypeName(sh.Cells(i, 11).Value2) <> "String" Or sh.Cells(i, 11).NumberFormat <> "@" Then VerifyProjection9X = "RECUSADO_FORMATO": Exit Function
        If Not sh.Cells(i, 11).Errors(xlNumberAsText).Ignore Then VerifyProjection9X = "RECUSADO_IGNORE": Exit Function
        If Not sh.Cells(i, 17).HasFormula Or Not sh.Cells(i, 18).HasFormula Then VerifyProjection9X = "RECUSADO_FORMULA": Exit Function
        If Abs(sh.Cells(i, 17).Value2 - CDbl(a(3) - a(2))) > 0.0000001 Then VerifyProjection9X = "RECUSADO_RESULTADO": Exit Function
    Next i
    If seen.Count <> d.Count Then VerifyProjection9X = "RECUSADO_QUANTIDADE" Else VerifyProjection9X = "VERIFICADA"
End Function
Public Sub F9_Complemento(ByVal sourceHash As String)
    Dim fso As Object, f As Object, c As Object, wb As Workbook, root As String, r As String, e As Long, errText As String, cnt As Long, i As Long, pid As String, rs As Object, a As Variant, sh As Worksheet, rowB As Long
    root = Root9X(): Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(root & "complemento_f9.tsv", False, True)
    f.WriteLine "Cenario" & vbTab & "Resultado" & vbTab & "Tratamento" & vbTab & "Esperado" & vbTab & "Observado" & vbTab & "Fixture" & vbTab & "Procedimento"
    On Error GoTo Failed
    If Len(sourceHash) <> 64 Or sourceHash Like "*[!0-9a-f]*" Then Err.Raise 5, , "Hash proveniencia invalido"
    fso.CopyFile root & "backup_ciclo.accdb", root & "migracao.accdb", False
    Set c = Open9X("migracao.accdb")
    c.BeginTrans
    c.Execute "INSERT INTO Provenance VALUES ('SYN_A1','" & sourceHash & "','CONFIRMADO','EP_A1','Origem sintetica explicita, movimento OP_A1')"
    c.Execute "INSERT INTO Provenance VALUES ('SYN_A2','" & sourceHash & "','CONFIRMADO','EP_A2','Reentrada explicita, movimento OP_A2')"
    c.Execute "INSERT INTO Provenance VALUES ('SYN_AMBIGUO','" & sourceHash & "','AMBIGUO',NULL,'Mesma placa e NF; episodio nao demonstrado')"
    c.CommitTrans
    C9X f, "F9_M0_M4_PROVENIENCIA", Scalar9X(c, "SELECT Count(*) FROM Provenance WHERE SourceHash='" & sourceHash & "' AND DecisionText='CONFIRMADO'") = 2 And IsNull(Scalar9X(c, "SELECT EpisodeID FROM Provenance WHERE LegacyRef='SYN_AMBIGUO'")), "PENDENTE_CONCILIACAO", "Duas correspondencias explicitas euma ambigua sem vinculo", "2 CONFIRMADO;1 AMBIGUO;SHA256=" & sourceHash, "migracao.accdb"
    c.Execute "CREATE UNIQUE INDEX U_M_EP_ID ON Movements (EpisodeID,MovementID)"
    c.Execute "ALTER TABLE Projections ADD CONSTRAINT FK_PM_EP FOREIGN KEY(EpisodeID,MovementID) REFERENCES Movements(EpisodeID,MovementID)"
    Set rs = c.Execute("SELECT TOP 1 MovementID FROM Movements WHERE EpisodeID='EP_A1'"): pid = CStr(rs.Fields(0).Value): rs.Close
    On Error Resume Next
    c.Execute "INSERT INTO Projections (ProjectionID,EpisodeID,MovementID) VALUES ('PID_MISMATCH','EP_B1','" & pid & "')"
    e = Err.Number: errText = Err.Description: Err.Clear
    On Error GoTo Failed
    C9X f, "F9_FK_COMPOSTA_EP_MOV", e <> 0 And Scalar9X(c, "SELECT Count(*) FROM Projections WHERE ProjectionID='PID_MISMATCH'") = 0, "RECUSADO", "Movimento de A1 nao pode projetar episodio B1", CStr(e) & ":" & errText, "migracao.accdb"
    c.Close
    Set c = Open9X("ciclo.accdb")
    fso.CopyFile root & "Projecao_Reparada.xlsm", root & "Projecao_Inspecao_Metadata.xlsm", False
    Set wb = Workbooks.Open(root & "Projecao_Inspecao_Metadata.xlsm"): Set sh = wb.Worksheets(1)
    ' A projecao reparada pertence ao banco falha_8, que tem uma revisao posterior.
    r = VerifyProjection9X(c, wb)
    C9X f, "F9_PROJECAO_REVISAO_DIVERGENTE", r = "RECUSADO_REVISAO", "RECUSADO", "Revisao de outra autoridade nao autoriza", r, "Projecao_Inspecao_Metadata.xlsm"
    c.Close: Set c = Open9X("falha_8.accdb")
    r = VerifyProjection9X(c, wb)
    C9X f, "F9_METADATA_APOS_SORT_ANTES", r = "VERIFICADA", IIf(r = "VERIFICADA", "VERIFICADO", "FAIL"), "Ignore/formato preservados alem da associacao", r, "Projecao_Inspecao_Metadata.xlsm"
    ' Reconstituir metadado derivado da chave validada por identidade, sem mudar fatos/selos.
    For i = 2 To sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
        If TypeName(sh.Cells(i, 11).Value2) = "String" And Len(sh.Cells(i, 11).Value2) = 44 Then sh.Cells(i, 11).Errors(xlNumberAsText).Ignore = True
    Next i
    r = VerifyProjection9X(c, wb)
    C9X f, "F9_METADATA_RECONSTRUIDO", r = "VERIFICADA", "RECUPERADO", "Metadado reparado; identidade e valores conferidos", r, "Projecao_Inspecao_Metadata.xlsm"
    sh.Range("L2").Value2 = 999999@
    r = VerifyProjection9X(c, wb)
    C9X f, "F9_EDICAO_PROJECAO_RECUSADA", r = "RECUSADO_CONTEUDO", "RECUSADO", "Edicao local divergente nao autoriza", r, "Projecao_Inspecao_Metadata.xlsm"
    wb.Save: wb.Close False: Set wb = Nothing
    c.Close
    ' Tentativa real da API sobre conexao fechada: nao simula falha fisica de commit.
    On Error Resume Next
    c.CommitTrans: e = Err.Number: errText = Err.Description: Err.Clear
    On Error GoTo Failed
    C9X f, "F9_COMMIT_CONEXAO_FECHADA", e <> 0, "RECUSADO", "API recusa commit sem conexao; I/O commit nao exercitado", CStr(e) & ":" & errText, "falha_8.accdb"
    f.Close: Debug.Print "F9_COMPLEMENT_DONE": Exit Sub
Failed:
    e = Err.Number: errText = Err.Description
    On Error Resume Next
    C9X f, "ERRO_COMPLEMENTO", False, "FAIL", "Conclusao", CStr(e) & ":" & errText, "migracao.accdb"
    If Not wb Is Nothing Then wb.Close False
    c.Close: f.Close: Debug.Print "F9_COMPLEMENT_ERROR"; e; errText
End Sub
