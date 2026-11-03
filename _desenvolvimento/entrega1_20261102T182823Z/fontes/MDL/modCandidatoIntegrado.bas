Attribute VB_Name = "modCandidatoIntegrado"
Option Explicit
Private Const DB_C1 As String = "Banco_Candidato_Sintetico.accdb"
Private Const PROTECT_C1 As String = "CANDIDATO_SINTETICO"
Private mCtx As Object, mCache As Object
Private mBusy As Boolean, mInstalled As Boolean, mUI As Boolean
Public C1_UltimoResultado As String
Public C1_FornecedorXML As Object
Public C1_PoliticaFalha As Object
Public C1_ChaveEntrada As String

Private Function ConnC1() As Object
    Dim c As Object, stored As String
    Call Guard10
    Set c = Open10(DB_C1)
    If CLng(Val10(c, "SELECT SchemaNo FROM Meta WHERE ID=1")) <> 5 Then c.Close: Err.Raise 5, , "Execute CandidatoInstalarLaboratorio"
    stored = CStr(Val10(c, "SELECT AuthorityPath FROM CandidateSettings WHERE ID=1"))
    If Not C1_MesmoArquivoAutoridade(stored, ThisWorkbook.Path & "\" & DB_C1) Then c.Close: Err.Raise 5, , "Copia divergente: caminho da autoridade nao confere"
    If CStr(Val10(c, "SELECT BaseID FROM CandidateSettings WHERE ID=1")) <> CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")) Then c.Close: Err.Raise 5, , "BaseID divergente"
    Set ConnC1 = c
End Function
Public Function C1_Context(ByVal f As Object) As Object
    Dim key As String
    If mCtx Is Nothing Then Set mCtx = CreateObject("Scripting.Dictionary")
    key = TypeName(f)
    If Not mCtx.Exists(key) Then mCtx.Add key, CreateObject("Scripting.Dictionary")
    Set C1_Context = mCtx(key)
End Function
Public Sub C1_LimparContexto(ByVal f As Object)
    Dim d As Object
    Set d = C1_Context(f): d.RemoveAll
    C1_ChaveEntrada = ""
End Sub
Private Function CV(ByVal f As Object, ByVal name As String) As String
    CV = Trim$(CStr(f.Controls(name).value))
End Function
Private Sub SV(ByVal f As Object, ByVal name As String, ByVal value As Variant)
    f.Controls(name).value = value
End Sub
Private Sub StatusC1(ByVal f As Object, ByVal text As String)
    C1_UltimoResultado = text
    f.Controls("LabelStatus").CAPTION = text
End Sub
Private Function XmlC1(ByVal fields As Object) As String
    Dim d As Object, r As Object, e As Object, k As Variant
    Set d = CreateObject("MSXML2.DOMDocument.6.0")
    Set r = d.createElement("CandidatePayload"): r.setAttribute "version", "1": d.appendChild r
    For Each k In fields.keys
        Set e = d.createElement("f"): e.setAttribute "n", CStr(k): e.text = CStr(fields(k)): r.appendChild e
    Next k
    XmlC1 = d.XML
End Function
Private Function ReadXmlC1(ByVal text As String) As Object
    Dim d As Object, e As Object, result As Object, k As String
    Set d = CreateObject("MSXML2.DOMDocument.6.0")
    d.async = False: d.resolveExternals = False: d.setProperty "ProhibitDTD", True
    If Not d.LoadXML(text) Then Err.Raise 5, , "Payload de interface invalido"
    If d.documentElement.nodeName <> "CandidatePayload" Or d.documentElement.getAttribute("version") <> "1" Then Err.Raise 5, , "Contrato do payload divergente"
    Set result = CreateObject("Scripting.Dictionary")
    For Each e In d.documentElement.childNodes
        k = CStr(e.getAttribute("n")): If result.Exists(k) Then Err.Raise 5, , "Campo duplicado no payload"
        result.Add k, CStr(e.text)
    Next e
    Set ReadXmlC1 = result
End Function
Private Function FieldC1(ByVal d As Object, ByVal key As String, Optional ByVal fallback As String = "") As String
    If d.Exists(key) Then FieldC1 = CStr(d(key)) Else FieldC1 = fallback
End Function
Private Function AmountC1(ByVal s As String) As Currency
    Dim n As Currency
    s = Replace(Replace(Trim$(s), "R$", ""), ChrW(160), "")
    If InStr(s, ",") > 0 Then s = Replace(Replace(s, ".", ""), ",", ".")
    If Not MoneyOK10(s, n) Then Err.Raise 5, , "Valor monetario invalido: " & s
    AmountC1 = n
End Function
Private Function MoneyC1(ByVal n As Currency) As String
    MoneyC1 = Replace(Format$(n, "0.0000"), Application.International(xlDecimalSeparator), ".")
End Function
Private Function DateC1(ByVal text As String) As String
    Dim a As Variant, iso As String
    text = Trim$(text)
    If DateOK10(text) Then DateC1 = text: Exit Function
    a = Split(text, "/")
    If UBound(a) <> 2 Then Err.Raise 5, , "Data invalida; use dd/mm/aaaa"
    If Len(a(0)) <> 2 Or Len(a(1)) <> 2 Or Len(a(2)) <> 4 Then Err.Raise 5, , "Data incompleta"
    iso = a(2) & "-" & a(1) & "-" & a(0)
    If Not DateOK10(iso) Then Err.Raise 5, , "Data civil invalida"
    DateC1 = iso
End Function
Private Function AsDateC1(ByVal iso As String) As Date
    If Not DateOK10(iso) Then Err.Raise 5, , "Data persistente invalida"
    AsDateC1 = DateSerial(CInt(Left$(iso, 4)), CInt(mid$(iso, 6, 2)), CInt(Right$(iso, 2)))
End Function
Private Function FailC1(ByVal point As String) As Long
    If Not C1_PoliticaFalha Is Nothing Then FailC1 = C1_PoliticaFalha.Ponto(point)
End Function
Private Sub ReadyC1(ByVal c As Object)
    Dim st As String
    st = Audit10(c)
    If st <> "OK" Then Err.Raise 5, , "Conciliacao obrigatoria: " & st
    If Val10(c, "SELECT WriteState FROM Meta WHERE ID=1") <> "READY" Then Err.Raise 5, , "Autoridade em conciliacao"
End Sub
Public Sub CandidatoInstalarLaboratorio()
    Dim c As Object, dbe As Object, db As Object
    Call Guard10
    If Not FSO10.FileExists(ThisWorkbook.Path & "\" & DB_C1) Then Err.Raise 5, , "Banco sintetico de instalacao ausente"
    Set c = Open10(DB_C1)
    If CLng(Val10(c, "SELECT SchemaNo FROM Meta WHERE ID=1")) = 4 Then
        If CLng(Val10(c, "SELECT Count(*) FROM Movements")) <> 0 Then Err.Raise 5, , "Instalacao exige banco sintetico vazio"
        If CStr(Val10(c, "SELECT SchemaHash FROM Meta WHERE ID=1")) <> "8698a2b7f44693510141e9fe75fd123fe41b4d99f7c60a03c1f9d2a99625d26b" Then Err.Raise 5, , "Schema4 de origem desconhecido"
        c.Execute "CREATE TABLE CandidateSettings (ID INTEGER CONSTRAINT PK_CS PRIMARY KEY,BaseID TEXT(40) NOT NULL,AuthorityPath LONGTEXT,WorkbookID TEXT(40),ContractVersion LONG,ProjectionState TEXT(30))"
        c.Execute "CREATE TABLE SheetProjections (ViewID TEXT(40) CONSTRAINT PK_SV PRIMARY KEY,ProjectionID TEXT(40) NOT NULL,EpisodeID TEXT(40) NOT NULL,MovementID TEXT(40) NOT NULL,SheetKey TEXT(40) NOT NULL,RoleText TEXT(20) NOT NULL,ViewNo COUNTER,ViewHash TEXT(64),CONSTRAINT U_SV UNIQUE(ProjectionID,SheetKey),CONSTRAINT FK_SV_P FOREIGN KEY(EpisodeID,ProjectionID) REFERENCES Projections(EpisodeID,ProjectionID),CONSTRAINT FK_SV_M FOREIGN KEY(EpisodeID,MovementID) REFERENCES Movements(EpisodeID,MovementID))"
        c.Close
        Set dbe = CreateObject("DAO.DBEngine.120"): Set db = dbe.OpenDatabase(ThisWorkbook.Path & "\" & DB_C1)
        db.TableDefs("Meta").ValidationRule = "[ID]=1 And [SchemaNo]=5 And [RevisionNo]>=0 And [WriteState] In ('READY','RECONCILE')"
        db.TableDefs("Operations").ValidationRule = Replace(db.TableDefs("Operations").ValidationRule, "'PREPARED','CONFIRMED'", "'PREPARED','CONFIRMED','ABORTED'")
        db.TableDefs("Operations").fields("AccessKey").AllowZeroLength = True
        db.TableDefs("Projections").fields("AccessKey").AllowZeroLength = True
        db.TableDefs("Operations").ValidationRule = Replace(db.TableDefs("Operations").ValidationRule, "Len([AccessKey])=44", "(Len([AccessKey])=0 Or Len([AccessKey])=44)")
        db.TableDefs("Projections").ValidationRule = Replace(db.TableDefs("Projections").ValidationRule, "Len([AccessKey])=44", "(Len([AccessKey])=0 Or Len([AccessKey])=44)")
        db.Close: Set db = Nothing: Set dbe = Nothing
        Set c = Open10(DB_C1)
        Exec10 c, "UPDATE Meta SET SchemaNo=5 WHERE ID=1"
        Exec10 c, "INSERT INTO CandidateSettings VALUES (1,?,?,?,?,?)", CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")), ThisWorkbook.Path & "\" & DB_C1, NewID10(), 1, "PENDING"
        Exec10 c, "UPDATE Meta SET SchemaHash=? WHERE ID=1", SchemaFingerprint10(c)
    End If
    c.Close: CandidatoReconstruir
    Debug.Print "ENTREGA1_INSTALADA"
End Sub
Public Function C1_LocalOperacao(ByVal c As Object, ByVal op As String) As String
    Dim d As Object
    Set d = ReadXmlC1(CStr(Val10(c, "SELECT PayloadText FROM Operations WHERE OperationID=?", op)))
    C1_LocalOperacao = FieldC1(d, "DESTINO")
    If C1_LocalOperacao <> "MATRIZ" And C1_LocalOperacao <> "FILIAL" Then Err.Raise 5, , "Custodia experimental ausente"
End Function
Private Function PlanC1(ByVal kind As String, ByVal loc As String) As Variant
    Select Case kind
        Case "ENTRY"
            If loc = "MATRIZ" Then PlanC1 = Array("Entradas|CUSTODY", "Entradas Anual|HISTORY_IN") Else PlanC1 = Array("Entradas Filial|CUSTODY", "Entradas Anual Filial|HISTORY_IN")
        Case "SALE"
            If loc = "MATRIZ" Then PlanC1 = Array("Saidas|EXIT", "Todas Saidas|HISTORY_OUT") Else PlanC1 = Array("Saidas Filial|EXIT", "Todas Saidas Filial|HISTORY_OUT")
        Case "TRANSFER": PlanC1 = Array("Saidas|EXIT", "Todas Saidas|HISTORY_OUT", "Entradas Filial|CUSTODY", "Entradas Anual Filial|HISTORY_IN")
        Case "RETURN": PlanC1 = Array("Saidas Filial|EXIT", "Todas Saidas Filial|HISTORY_OUT", "Entradas|CUSTODY", "Entradas Anual|HISTORY_IN")
        Case Else: Err.Raise 5, , "Operacao ainda nao integrada"
    End Select
End Function
Public Sub C1_RegistrarViews(ByVal c As Object, ByVal op As String, ByVal ep As String, ByVal mid As String, ByVal pid As String, ByVal kind As String, ByVal loc As String)
    Dim item As Variant, p As Variant, id As String, canon As String
    For Each item In PlanC1(kind, loc)
        p = Split(CStr(item), "|"): id = NewID10(): canon = Pack10(id, pid, ep, mid, CStr(p(0)), CStr(p(1)))
        Exec10 c, "INSERT INTO SheetProjections (ViewID,ProjectionID,EpisodeID,MovementID,SheetKey,RoleText,ViewHash) VALUES (?,?,?,?,?,?,?)", id, pid, ep, mid, CStr(p(0)), CStr(p(1)), SHA10(canon)
    Next item
End Sub
Public Function C1_AuditarViews(ByVal c As Object) As String
    Dim r As Object, rows As Variant, expected As Long, actual As Long, op As String, pid As String, ep As String, mid As String, kind As String, loc As String, item As Variant, p As Variant
    On Error GoTo Bad
    Set r = c.Execute("SELECT OperationID,PlannedPID,PlannedMID,EpisodeID,KindText FROM Operations WHERE StatusText='CONFIRMED'")
    Do While Not r.EOF
        op = CStr(r(0)): pid = CStr(r(1)): mid = CStr(r(2)): ep = CStr(r(3)): kind = CStr(r(4)): loc = C1_LocalOperacao(c, op)
        For Each item In PlanC1(kind, loc)
            p = Split(CStr(item), "|"): expected = expected + 1
            If CLng(Val10(c, "SELECT Count(*) FROM SheetProjections WHERE ProjectionID=? AND SheetKey=? AND RoleText=? AND EpisodeID=? AND MovementID=?", pid, CStr(p(0)), CStr(p(1)), ep, mid)) <> 1 Then C1_AuditarViews = "PLANO_PROJECAO_INCOMPLETO": r.Close: Exit Function
        Next item
        r.MoveNext
    Loop
    r.Close
    actual = CLng(Val10(c, "SELECT Count(*) FROM SheetProjections"))
    If actual <> expected Then C1_AuditarViews = "VIEW_COUNT": Exit Function
    Set r = c.Execute("SELECT ViewID,ProjectionID,EpisodeID,MovementID,SheetKey,RoleText,ViewHash FROM SheetProjections")
    Do While Not r.EOF
        If CStr(r(6)) <> SHA10(Pack10(CStr(r(0)), CStr(r(1)), CStr(r(2)), CStr(r(3)), CStr(r(4)), CStr(r(5)))) Then C1_AuditarViews = "VIEW_HASH": r.Close: Exit Function
        r.MoveNext
    Loop
    r.Close: C1_AuditarViews = "OK": Exit Function
Bad:
    C1_AuditarViews = "VIEW_ERROR:" & Err.Number & ":" & Err.Description
End Function

Private Function EntradaFieldsC1(ByVal f As Object, ByVal filial As Boolean) As Object
    Dim d As Object, s As String, v As Currency, key As String
    s = IIf(filial, "Filial", ""): Set d = CreateObject("Scripting.Dictionary")
    d.Add "MODEL", CV(f, "ComboBox_VEÍCULO" & s): d.Add "COD", CV(f, "COD" & s)
    d.Add "PLATE", UCase$(CV(f, "PLACA" & s)): d.Add "ANO", CV(f, "Combo_ANO_MOD" & s)
    d.Add "COR", CV(f, "ComboBox_COR" & s): d.Add "UN", CV(f, "UN" & s)
    d.Add "OUTRO", CV(f, "ComboBox_OutroEnvio" & s): d.Add "NFE", CV(f, "N_Nfe" & s)
    d.Add "DATA", DateC1(CV(f, IIf(filial, "DataFilial", "data")))
    v = AmountC1(CV(f, "txt_Valor" & s)): If v <= 0 Then Err.Raise 5, , "Valor de entrada deve ser positivo"
    d.Add "VALOR", MoneyC1(v): d.Add "VIN", MoneyC1(v): d.Add "VOUT", "0.0000"
    key = C1_ChaveEntrada: If filial Then key = CV(f, "TextBoxChaveDeAcessoFilial")
    If Len(key) > 0 Then If Len(key) <> 44 Or key Like "*[!0-9]*" Then Err.Raise 5, , "Chave deve conter exatamente44 digitos"
    If filial And Len(key) <> 44 Then Err.Raise 5, , "Informe a chave da entrada filial"
    d.Add "KEY", key: d.Add "SOURCE", IIf(filial, "FILIAL", "MATRIZ"): d.Add "DESTINO", d("SOURCE")
    d.Add "OPERACAO", "Compra": d.Add "KIND", "ENTRY": d.Add "ENTRYDOC", d("NFE")
    If d("NFE") = "" Or Len(d("NFE")) > 20 Or d("NFE") Like "*[!0-9]*" Then Err.Raise 5, , "Numero da NF-e invalido"
    If d("PLATE") = "" Or d("MODEL") = "" Or d("COD") = "" Or d("UN") = "" Then Err.Raise 5, , "Preencha veiculo, placa, codigo e unidade"
    Set EntradaFieldsC1 = d
End Function
Private Function DupEntradaC1(ByVal c As Object, ByVal doc As String, ByVal loc As String) As Boolean
    DupEntradaC1 = CLng(Val10(c, "SELECT Count(*) FROM Episodes E INNER JOIN Projections P ON E.CurrentPID=P.ProjectionID WHERE E.StatusText='ACTIVE' AND P.PlaceCode=? AND P.DocNumber=?", loc, doc)) > 0
End Function
Private Sub NewAdmissionC1(ByVal c As Object, ByVal ctx As Object, ByVal d As Object, ByVal who As String, ByVal fence As Long)
    Dim vid As String, ep As String
    If ctx.Exists("ReentryVehicleID") Then
        vid = CStr(ctx("ReentryVehicleID"))
        If CLng(Val10(c, "SELECT Count(*) FROM Vehicles WHERE VehicleID=?", vid)) <> 1 Then Err.Raise 5, , "Veiculo da reentrada ausente"
    Else
        vid = NewID10()
    End If
    ep = NewID10(): c.BeginTrans
    If Not LockWriter10(c, who, fence) Then c.RollbackTrans: Err.Raise 5, , "Fencing recusado"
    If ctx.Exists("ReentryVehicleID") Then
        If CLng(Val10(c, "SELECT Count(*) FROM Episodes WHERE VehicleID=? AND StatusText='ACTIVE'", vid)) <> 0 Then c.RollbackTrans: Err.Raise 5, , "Reentrada recusada: veiculo ainda ativo"
        If CStr(Val10(c, "SELECT Plate FROM Vehicles WHERE VehicleID=?", vid)) <> CStr(d("PLATE")) Then c.RollbackTrans: Err.Raise 5, , "Placa diverge do VehicleID explicitamente selecionado"
    Else
        Exec10 c, "INSERT INTO Vehicles VALUES (?,?)", vid, CStr(d("PLATE"))
    End If
    Exec10 c, "INSERT INTO Episodes (EpisodeID,VehicleID,VersionNo,StatusText,PlaceCode) VALUES (?,?,0,'NEW',?)", ep, vid, CStr(d("SOURCE"))
    c.CommitTrans
    ctx("VehicleID") = vid: ctx("EpisodeID") = ep: ctx("Version") = 0: ctx("OperationID") = NewID10()
End Sub
Public Sub C1_GravarEntrada(ByVal f As Object, ByVal filial As Boolean)
    Dim c As Object, d As Object, ctx As Object, who As String, fence As Long, st As String, own As Boolean, e As String, op As String
    If mBusy Then Exit Sub
    mBusy = True: On Error GoTo Failed
    Set d = EntradaFieldsC1(f, filial): Set ctx = C1_Context(f)
    If ctx.Exists("ViewID") And Not ctx.Exists("OperationID") Then Err.Raise 5, , "Cadastro selecionado e somente consulta; use Limpar para nova entrada"
    Set c = ConnC1(): who = NewID10(): fence = Acquire10(c, who): own = True: ReadyC1 c
    If Not ctx.Exists("OperationID") Then
        If Val10(c, "SELECT ProjectionState FROM CandidateSettings WHERE ID=1") <> "VERIFIED" Then Err.Raise 5, , "Reconstrua a projecao antes de nova operacao"
        If DupEntradaC1(c, CStr(d("NFE")), CStr(d("SOURCE"))) Then Err.Raise 5, , "Numero da NFe ja cadastrado no estoque deste local"
        If CLng(Val10(c, "SELECT Count(*) FROM Operations WHERE StatusText='PREPARED' AND DocNumber=?", CStr(d("NFE")))) > 0 Then Err.Raise 5, , "Documento com intencao pendente; conciliar antes de repetir"
        NewAdmissionC1 c, ctx, d, who, fence
    End If
    op = CStr(ctx("OperationID"))
    st = Prepare10(c, who, fence, op, NewID10(), CStr(ctx("EpisodeID")), 0, "ENTRY", AmountC1(CStr(d("VIN"))), 0@, XmlC1(d), CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")), CStr(d("DATA")), CStr(d("NFE")), CStr(d("KEY")))
    If st = "PREPARADA" Then st = Apply10(c, who, fence, op, FailC1("APPLY"))
    If st <> "CONFIRMADA" And st <> "JA_CONCLUIDA" And st <> "COMMIT_CONFIRMADO_PROJECAO_PENDENTE" Then Err.Raise 5, , st
    Exec10 c, "UPDATE CandidateSettings SET ProjectionState='PENDING' WHERE ID=1"
    Release10 c, who, fence: own = False: c.Close: Set c = Nothing
    CandidatoReconstruir
    StatusC1 f, "ENTRADA_CONFIRMADA; " & op
    C1_LimparContexto f
    mBusy = False: Exit Sub
Failed:
    e = Err.Description
    On Error Resume Next
    If own Then Release10 c, who, fence
    If Not c Is Nothing Then c.Close
    StatusC1 f, "ENTRADA_RECUSADA_OU_PENDENTE: " & e
    mBusy = False
End Sub
Private Function ViewC1(ByVal c As Object, ByVal id As String) As Object
    Dim r As Object, d As Object, p As Object, k As Variant
    Set r = c.Execute("SELECT S.ViewID,S.ProjectionID,S.EpisodeID,S.MovementID,S.SheetKey,S.RoleText,E.VehicleID,E.VersionNo,E.StatusText,E.CurrentPID,P.PlaceCode,P.DocNumber,P.AccessKey,P.AmountIn,P.AmountOut,P.CivilDate,O.PayloadText FROM (((SheetProjections S INNER JOIN Projections P ON S.ProjectionID=P.ProjectionID) INNER JOIN Episodes E ON S.EpisodeID=E.EpisodeID) INNER JOIN Movements M ON S.MovementID=M.MovementID) INNER JOIN Operations O ON M.OperationID=O.OperationID WHERE S.ViewID='" & Replace(id, "'", "''") & "'")
    If r.EOF Then Err.Raise 5, , "Projecao ausente"
    Set d = ReadXmlC1(CStr(r(16)))
    d("ViewID") = CStr(r(0)): d("PID") = CStr(r(1)): d("EpisodeID") = CStr(r(2)): d("MovementID") = CStr(r(3))
    d("SHEET") = CStr(r(4)): d("ROLE") = CStr(r(5)): d("VehicleID") = CStr(r(6)): d("Version") = CLng(r(7))
    d("STATE") = CStr(r(8)): d("CURRENT") = CStr(r(9)): d("PLACE") = CStr(r(10))
    d("NFE") = CStr(r(11)): d("KEY") = CStr(r(12)): d("VIN") = MoneyC1(CCur(r(13))): d("VOUT") = MoneyC1(CCur(r(14))): d("DATA") = CStr(r(15))
    d("ACTIVE") = (d("ROLE") = "CUSTODY" And d("STATE") = "ACTIVE" And d("PID") = d("CURRENT"))
    d("EXITED") = IIf(d("PID") = d("CURRENT") And d("STATE") = "ACTIVE", "NÃO", "SIM")
    If d("KIND") = "ENTRY" Or (d("KIND") = "RETURN" And (d("ROLE") = "CUSTODY" Or d("ROLE") = "HISTORY_IN")) Then d("VALOR") = d("VIN") Else d("VALOR") = d("VOUT")
    r.MoveNext: If Not r.EOF Then Err.Raise 5, , "Projecao duplicada"
    r.Close: Set ViewC1 = d
End Function
Private Sub RecheckC1(ByVal c As Object, ByVal ctx As Object, ByVal d As Object)
    If Not ctx.Exists("ViewID") Then Err.Raise 5, , "Pesquise e selecione o veiculo pela identidade"
    If CStr(ctx("EpisodeID")) <> CStr(d("EpisodeID")) Or CStr(ctx("VehicleID")) <> CStr(d("VehicleID")) Then Err.Raise 5, , "Identidade divergente"
    If CLng(ctx("Version")) <> CLng(d("Version")) Then Err.Raise 5, , "Selecao obsoleta; pesquise novamente"
    If Not CBool(d("ACTIVE")) Then Err.Raise 5, , "Registro baixado ou custodia anterior"
    If CLng(Val10(c, "SELECT Count(*) FROM Pending WHERE EpisodeID=?", CStr(d("EpisodeID")))) <> 0 Then Err.Raise 5, , "Episodio pendente de conciliacao"
End Sub
Public Sub C1_GravarSaida(ByVal f As Object, ByVal filial As Boolean)
    Dim c As Object, d As Object, ctx As Object, Data As Object, who As String, fence As Long, st As String, op As String, s As String, kind As String, doc As Object, file As String, owner As Boolean, e As String
    If mBusy Then Exit Sub
    mBusy = True: On Error GoTo Failed
    s = IIf(filial, "Filial", ""): Set ctx = C1_Context(f)
    If Not ctx.Exists("ViewID") Then Err.Raise 5, , "Selecione uma entrada pelo cadastro antes de gravar"
    If CV(f, "ComboBox_Operação" & s) = "Cancelado" Then Err.Raise 5, , "Cancelado ainda nao integrado; nenhuma regra fiscal foi alterada"
    If filial Then
        If CV(f, "ComboBox_OperaçãoFilial") <> "Devolução" Then Err.Raise 5, , "Saida filial integrada somente para Devolucao"
        kind = "RETURN"
    Else
        Select Case CV(f, "ComboBox_Operação")
            Case "Venda": kind = "SALE"
            Case "Transferência": kind = "TRANSFER"
            Case Else: Err.Raise 5, , "Operacao ainda nao integrada"
        End Select
        If Not CarregarXMLNFe(f, doc, file) Then Err.Raise 5, , "XML nao carregado"
        If Not ValidarXMLNFe(f, doc, file) Then Err.Raise 5, , "XML divergente; gravacao recusada"
    End If
    Set c = ConnC1(): who = NewID10(): fence = Acquire10(c, who): owner = True: ReadyC1 c
    Set Data = ViewC1(c, CStr(ctx("ViewID"))): RecheckC1 c, ctx, Data
    If Data("PLACE") <> IIf(filial, "FILIAL", "MATRIZ") Then Err.Raise 5, , "Custodia de origem divergente"
    If CV(f, "TextBox_N_Entrada" & s) <> Data("NFE") Or UCase$(CV(f, "TextBox_PLACA" & s)) <> Data("PLATE") Or CV(f, "COD" & s) <> Data("COD") Then Err.Raise 5, , "Campos do formulario divergem da identidade selecionada"
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "MODEL", Data("MODEL"): d.Add "COD", Data("COD"): d.Add "PLATE", Data("PLATE"): d.Add "ANO", Data("ANO"): d.Add "COR", Data("COR"): d.Add "UN", Data("UN"): d.Add "OUTRO", CV(f, "ComboBox_OutroEnvio" & s)
    d.Add "NFE", CV(f, "TextBox_N_Saida" & s): d.Add "ENTRYDOC", Data("NFE"): d.Add "DATA", DateC1(CV(f, "TextBox_Data" & s))
    d.Add "VIN", MoneyC1(AmountC1(CV(f, "TextBox_Valor_Entrada" & s))): d.Add "VOUT", MoneyC1(AmountC1(CV(f, "TextBox_Valor_Saida" & s)))
    If AmountC1(CStr(d("VIN"))) <> AmountC1(CStr(Data("VALOR"))) Then Err.Raise 5, , "Valor de origem diverge do banco"
    If AmountC1(CStr(d("VOUT"))) <= 0 Then Err.Raise 5, , "Valor de saida invalido"
    d.Add "KEY", CV(f, "TextBoxChaveDeAcesso" & s)
    If Len(d("KEY")) <> 44 Or CStr(d("KEY")) Like "*[!0-9]*" Then Err.Raise 5, , "Chave de saida invalida"
    d.Add "OPERACAO", CV(f, "ComboBox_Operação" & s): d.Add "KIND", kind: d.Add "SOURCE", Data("PLACE")
    d.Add "DESTINO", IIf(kind = "TRANSFER", "FILIAL", "MATRIZ")
    d.Add "ICMS", MoneyC1(AmountC1(CV(f, "TextBox_ICMS" & s))): d.Add "PIS", MoneyC1(AmountC1(CV(f, "TextBox_PIS" & s))): d.Add "COFINS", MoneyC1(AmountC1(CV(f, "TextBox_COFINS" & s))): d.Add "POR", MoneyC1(AmountC1(CV(f, "ComboBox_Por" & s)))
    If filial Then d.Add "XMLHASH", "NAO_EXIGIDO_PELO_FLUXO_LEGADO" Else d.Add "XMLHASH", FileSHA10(file)
    If Len(d("NFE")) = 0 Or CStr(d("NFE")) Like "*[!0-9]*" Then Err.Raise 5, , "Numero de saida invalido"
    If Not ctx.Exists("OperationID") Then
        If Val10(c, "SELECT ProjectionState FROM CandidateSettings WHERE ID=1") <> "VERIFIED" Then Err.Raise 5, , "Projecao pendente; reconstruir antes de gravar"
        If CLng(Val10(c, "SELECT Count(*) FROM SheetProjections S INNER JOIN Projections P ON S.ProjectionID=P.ProjectionID WHERE S.SheetKey=? AND P.DocNumber=?", IIf(filial, "Saidas Filial", "Saidas"), CStr(d("NFE")))) <> 0 Then Err.Raise 5, , "Numero da NFe ja cadastrado na saida"
        If kind = "TRANSFER" Then If DupEntradaC1(c, CStr(d("NFE")), "FILIAL") Then Err.Raise 5, , "Destino filial ja possui documento ativo; identidade ambigua recusada"
        ctx("OperationID") = NewID10()
    End If
    op = CStr(ctx("OperationID"))
    st = Prepare10(c, who, fence, op, NewID10(), CStr(Data("EpisodeID")), CLng(ctx("Version")), kind, AmountC1(CStr(d("VIN"))), AmountC1(CStr(d("VOUT"))), XmlC1(d), CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")), CStr(d("DATA")), CStr(d("NFE")), CStr(d("KEY")))
    If st = "PREPARADA" Then st = Apply10(c, who, fence, op, FailC1("APPLY"))
    If st <> "CONFIRMADA" And st <> "JA_CONCLUIDA" And st <> "COMMIT_CONFIRMADO_PROJECAO_PENDENTE" Then Err.Raise 5, , st
    Exec10 c, "UPDATE CandidateSettings SET ProjectionState='PENDING' WHERE ID=1"
    Release10 c, who, fence: owner = False: c.Close: Set c = Nothing
    Call CandidatoReconstruir: StatusC1 f, "SAIDA_CONFIRMADA; " & op
    C1_LimparContexto f: mBusy = False: Exit Sub
Failed:
    e = Err.Description
    On Error Resume Next
    If owner Then Release10 c, who, fence
    If Not c Is Nothing Then c.Close
    StatusC1 f, "SAIDA_RECUSADA_OU_PENDENTE: " & e
    mBusy = False
End Sub

Private Function ViewListC1(ByVal c As Object) As Collection
    Dim r As Object, result As New Collection
    Set r = c.Execute("SELECT ViewID FROM SheetProjections ORDER BY ViewNo")
    Do While Not r.EOF: result.Add CStr(r(0)): r.MoveNext: Loop
    r.Close: Set ViewListC1 = result
End Function
Private Sub CacheC1(ByVal c As Object)
    Dim id As Variant, d As Object
    Set mCache = CreateObject("Scripting.Dictionary")
    For Each id In ViewListC1(c)
        Set d = ViewC1(c, CStr(id)): mCache.Add CStr(id), d
    Next id
    mCache.Add "@Revision", CLng(Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1"))
End Sub
Public Function CampoCandidato(ByVal id As String, ByVal field As String, ByVal revision As Variant) As Variant
    Dim d As Object, days As Long
    On Error GoTo Missing
    If mCache Is Nothing Then GoTo Missing
    If Not mCache.Exists(id) Then GoTo Missing
    If CLng(revision) <> CLng(mCache("@Revision")) Then GoTo Missing
    Set d = mCache(id)
    days = DateDiff("d", AsDateC1(CStr(d("DATA"))), Date)
    Select Case field
        Case "MES": CampoCandidato = Month(AsDateC1(CStr(d("DATA"))))
        Case "DIAS": CampoCandidato = days
        Case "PRAZO": CampoCandidato = IIf(days < 90, "Dentro do prazo", "Fora do prazo")
        Case "SAIU": CampoCandidato = d("EXITED")
        Case "LUCRO": If AmountC1(CStr(d("VOUT"))) > AmountC1(CStr(d("VIN"))) Then CampoCandidato = CDbl(AmountC1(CStr(d("VOUT"))) - AmountC1(CStr(d("VIN")))) Else CampoCandidato = 0
        Case Else: CampoCandidato = CVErr(xlErrValue)
    End Select
    Exit Function
Missing:
    CampoCandidato = CVErr(xlErrNA)
End Function
Private Function StartC1(ByVal sheet As String) As Long
    If Left$(sheet, 8) = "Entradas" Then StartC1 = 4 Else StartC1 = 2
End Function
Private Function RowC1(ByVal d As Object, ByVal row As Long) As Variant
    Dim a(1 To 16) As Variant, j As Long
    For j = 1 To 16: a(j) = "": Next j
    If d("ROLE") = "CUSTODY" Or d("ROLE") = "HISTORY_IN" Then
        a(1) = CStr(d("NFE")): a(2) = "Compra": a(3) = CStr(d("UN")): a(4) = CStr(d("MODEL")): a(5) = CStr(d("COD")): a(6) = CStr(d("PLATE")): a(7) = CStr(d("ANO")): a(8) = CStr(d("COR")): a(9) = CDbl(AmountC1(CStr(d("VALOR")))): a(10) = CDbl(AsDateC1(CStr(d("DATA"))))
        If d("ROLE") = "CUSTODY" Then a(11) = CStr(d("OUTRO"))
        If d("SHEET") = "Entradas Filial" Then a(12) = CStr(d("KEY"))
        If d("ROLE") = "HISTORY_IN" Then
            a(11) = FormulaC1(row, "MES"): a(12) = FormulaC1(row, "DIAS"): a(13) = FormulaC1(row, "PRAZO")
            a(16) = FormulaC1(row, "SAIU")
        ElseIf d("SHEET") = "Entradas Filial" Then
            a(13) = FormulaC1(row, "MES"): a(14) = FormulaC1(row, "DIAS"): a(15) = FormulaC1(row, "PRAZO")
        Else
            a(12) = FormulaC1(row, "MES"): a(13) = FormulaC1(row, "DIAS"): a(14) = FormulaC1(row, "PRAZO")
        End If
    Else
        a(1) = CStr(d("NFE")): a(2) = CStr(d("ENTRYDOC")): a(3) = CStr(d("OPERACAO")): a(4) = CStr(d("UN")): a(5) = CStr(d("MODEL")): a(6) = CStr(d("PLATE")): a(7) = CStr(d("ANO")): a(8) = CStr(d("COR")): a(9) = CDbl(AmountC1(CStr(d("VIN")))): a(10) = CDbl(AmountC1(CStr(d("VOUT")))): a(11) = FormulaC1(row, "LUCRO"): a(12) = CDbl(AsDateC1(CStr(d("DATA"))))
        a(13) = CDbl(AmountC1(FieldC1(d, "POR", "0"))) / 100#: a(14) = CDbl(AmountC1(FieldC1(d, "ICMS", "0"))): a(15) = CDbl(AmountC1(FieldC1(d, "PIS", "0"))): a(16) = CDbl(AmountC1(FieldC1(d, "COFINS", "0")))
    End If
    RowC1 = a
End Function
Private Function FormulaC1(ByVal row As Long, ByVal field As String) As String
    FormulaC1 = "=CampoCandidato($AZ" & row & ",""" & field & """,$BI$2)"
End Function
Private Function MetaC1(ByVal d As Object) As Variant
    MetaC1 = Array(CStr(d("ViewID")), CStr(d("EpisodeID")), CStr(d("MovementID")), CStr(d("VehicleID")), CLng(d("Version")), CBool(d("ACTIVE")), CStr(d("PID")), CStr(d("ROLE")), CStr(d("KEY")))
End Function
Private Sub PutC1(ByVal sh As Worksheet, ByVal row As Long, ByVal d As Object)
    Dim a As Variant, m As Variant, j As Long, keycol As Long
    On Error GoTo Failed
    a = RowC1(d, row): m = MetaC1(d)
    sh.Range(sh.Cells(row, 1), sh.Cells(row, 16)).NumberFormat = "@"
    For j = 1 To 16
        If VarType(a(j)) = vbDouble Then
            sh.Cells(row, j).NumberFormat = "0.0000": sh.Cells(row, j).Value2 = a(j)
        ElseIf Left$(CStr(a(j)), 1) = "=" Then
            sh.Cells(row, j).NumberFormat = "General": sh.Cells(row, j).Formula = a(j)
        Else
            sh.Cells(row, j).Value2 = CStr(a(j))
        End If
    Next j
    If d("ROLE") = "CUSTODY" Or d("ROLE") = "HISTORY_IN" Then sh.Cells(row, 10).NumberFormat = "dd""/""mm""/""yyyy" Else sh.Cells(row, 12).NumberFormat = "dd""/""mm""/""yyyy"
    sh.Range(sh.Cells(row, 52), sh.Cells(row, 60)).NumberFormat = "@"
    For j = 0 To 8
        If j = 4 Or j = 5 Then sh.Cells(row, 52 + j).NumberFormat = "General"
        sh.Cells(row, 52 + j).Value2 = m(j)
    Next j
    sh.Cells(row, 60).Errors(xlNumberAsText).Ignore = True
    If d("SHEET") = "Entradas Filial" Then sh.Cells(row, 12).Errors(xlNumberAsText).Ignore = True
    Exit Sub
Failed:
    Err.Raise Err.Number, "PutC1", sh.name & ";linha=" & row & ";coluna=" & j & ";" & Err.Description
End Sub
Private Function SheetsC1() As Variant
    SheetsC1 = Array("Entradas", "Entradas Anual", "Entradas Filial", "Entradas Anual Filial", "Saidas", "Todas Saidas", "Saidas Filial", "Todas Saidas Filial")
End Function
Private Function VerifyC1(ByVal c As Object, ByVal wb As Workbook, ByVal build As String) As String
    Dim seen As Object, sh As Worksheet, sn As Variant, row As Long, last As Long, id As String, d As Object, a As Variant, m As Variant, j As Long, actual As Variant, field As String
    On Error GoTo Invalid
    Set seen = CreateObject("Scripting.Dictionary")
    If wb.Worksheets.count <> 17 Then VerifyC1 = "SHEET_COUNT": Exit Function
    For Each sn In SheetsC1()
        Set sh = wb.Worksheets(CStr(sn))
        If CStr(sh.Range("BI1").Value2) <> CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")) Or CLng(sh.Range("BI2").Value2) <> CLng(mCache("@Revision")) Or CStr(sh.Range("BI3").Value2) <> build Then VerifyC1 = "BASE_REVISION_BUILD": Exit Function
        last = sh.Cells(sh.rows.count, 52).End(xlUp).row
        For row = StartC1(CStr(sn)) To last
            id = CStr(sh.Cells(row, 52).Value2)
            If Not mCache.Exists(id) Or seen.Exists(id) Then VerifyC1 = "IDENTITY_DUP_MISSING": Exit Function
            seen.Add id, True: Set d = mCache(id)
            If d("SHEET") <> CStr(sn) Then VerifyC1 = "SHEET_IDENTITY": Exit Function
            a = RowC1(d, row): m = MetaC1(d)
            For j = 1 To 16
                actual = sh.Cells(row, j).Value2
                If Left$(CStr(a(j)), 1) = "=" Then
                    If sh.Cells(row, j).Formula <> a(j) Then VerifyC1 = "FORMULA:" & sn & ":" & row & ":" & j: Exit Function
                    field = Split(CStr(a(j)), """")(1)
                    If IsError(actual) Then VerifyC1 = "FORMULA_ERROR": Exit Function
                    If CStr(actual) <> CStr(CampoCandidato(id, field, mCache("@Revision"))) Then VerifyC1 = "FORMULA_VALUE": Exit Function
                    If sh.Cells(row, j).NumberFormat <> "General" Then VerifyC1 = "FORMULA_FORMAT": Exit Function
                ElseIf VarType(a(j)) = vbDouble Then
                    If Not IsNumeric(actual) Then VerifyC1 = "NUMBER_TYPE": Exit Function
                    If Abs(CDbl(actual) - CDbl(a(j))) > 0.0000001 Then VerifyC1 = "NUMBER_VALUE": Exit Function
                    If (j = 10 And (d("ROLE") = "CUSTODY" Or d("ROLE") = "HISTORY_IN")) Or (j = 12 And (d("ROLE") = "EXIT" Or d("ROLE") = "HISTORY_OUT")) Then
                        If Left$(F8_CanonData(sh.Cells(row, j), "BR_LITERAL"), 5) = "ERRO:" Then VerifyC1 = "DATE_FORMAT:" & sh.Cells(row, j).NumberFormat: Exit Function
                    ElseIf sh.Cells(row, j).NumberFormat <> "0.0000" Then
                        VerifyC1 = "MONEY_FORMAT": Exit Function
                    End If
                Else
                    If CStr(actual) <> CStr(a(j)) Then VerifyC1 = "VALUE:" & sn & ":" & row & ":" & j: Exit Function
                    If Len(CStr(a(j))) > 0 And VarType(actual) <> vbString Then VerifyC1 = "TEXT_TYPE": Exit Function
                    If sh.Cells(row, j).NumberFormat <> "@" Then VerifyC1 = "TEXT_FORMAT": Exit Function
                End If
            Next j
            For j = 0 To 8
                If CStr(sh.Cells(row, 52 + j).Value2) <> CStr(m(j)) Then VerifyC1 = "META_VALUE": Exit Function
                If j <> 4 And j <> 5 Then If sh.Cells(row, 52 + j).NumberFormat <> "@" Then VerifyC1 = "META_FORMAT": Exit Function
            Next j
            If Not IgnoredC1(sh, row, 60) Then VerifyC1 = "IGNORE": Exit Function
            If sn = "Entradas Filial" Then If Not IgnoredC1(sh, row, 12) Then VerifyC1 = "KEY_IGNORE": Exit Function
        Next row
        If sh.Cells(sh.rows.count, 1).End(xlUp).row > IIf(last < StartC1(CStr(sn)), StartC1(CStr(sn)) - 1, last) Then VerifyC1 = "UNIDENTIFIED_ROW": Exit Function
    Next sn
    If seen.count <> mCache.count - 1 Then VerifyC1 = "VIEW_COUNT": Exit Function
    VerifyC1 = "VERIFIED": Exit Function
Invalid:
    VerifyC1 = "ERROR:" & Err.Number & ":" & CStr(sn) & ":r=" & row & ":j=" & j & ":" & Err.Description
End Function
Public Sub CandidatoReconstruir()
    Dim c As Object, who As String, fence As Long, owner As Boolean, ev As Boolean, su As Boolean, alerts As Boolean, calc As XlCalculation, sh As Worksheet, sn As Variant, id As Variant, d As Object, positions As Object, row As Long, build As String, file As String, wb As Workbook, st As String, h As String, e As Long, msg As String, tx As Boolean
    ev = Application.EnableEvents: su = Application.ScreenUpdating: alerts = Application.DisplayAlerts: calc = Application.Calculation
    On Error GoTo Failed
    Set c = ConnC1(): who = NewID10(): fence = Acquire10(c, who): owner = True
    If fence = 0 Then Err.Raise 5, , "Outro escritor possui a autoridade"
    ReadyC1 c: CacheC1 c: build = NewID10()
    Exec10 c, "UPDATE CandidateSettings SET ProjectionState='PENDING' WHERE ID=1"
    Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
    Set positions = CreateObject("Scripting.Dictionary")
    For Each sn In SheetsC1()
        Set sh = ThisWorkbook.Worksheets(CStr(sn)): sh.Unprotect PROTECT_C1
        sh.Range("A" & StartC1(CStr(sn)) & ":BI" & Application.Max(StartC1(CStr(sn)), sh.UsedRange.row + sh.UsedRange.rows.count)).Clear
        sh.Range("BI1:BI4").NumberFormat = "@"
        sh.Range("BI1").Value2 = CStr(Val10(c, "SELECT BaseID FROM Meta WHERE ID=1")): sh.Range("BI2").Value2 = CStr(mCache("@Revision")): sh.Range("BI3").Value2 = build: sh.Range("BI4").Value2 = "BUILT"
        positions.Add CStr(sn), StartC1(CStr(sn))
    Next sn
    For Each id In ViewListC1(c)
        Set d = mCache(CStr(id)): Set sh = ThisWorkbook.Worksheets(CStr(d("SHEET")))
        row = CLng(positions(d("SHEET"))): PutC1 sh, row, d: positions(d("SHEET")) = row + 1
        If FailC1("PROJECTION") = 1 Then Err.Raise 7101, , "Falha controlada na projecao"
    Next id
    For Each sh In ThisWorkbook.Worksheets
        sh.Unprotect PROTECT_C1: sh.Columns("AZ:BI").Hidden = True: sh.Protect PROTECT_C1
    Next sh
    For Each sh In ThisWorkbook.Worksheets: sh.Calculate: Next sh
    st = VerifyC1(c, ThisWorkbook, build): If st <> "VERIFIED" Then Err.Raise 5, , "Verificacao em memoria: " & st
    file = ThisWorkbook.Path & "\Projecao_C1_" & build & ".xlsm"
    IntentAck10 c, build, file, who, fence
    If FailC1("SAVE") = 1 Then Err.Raise 7102, , "Falha controlada antes do Save"
    Application.EnableEvents = ev
    ThisWorkbook.save
    Application.EnableEvents = False
    If Not ThisWorkbook.Saved Then Err.Raise 5, , "Save recusado; autoridade preservada, projecao pendente"
    ThisWorkbook.SaveCopyAs file: h = FileSHA10(file)
    Set wb = Workbooks.Open(file, UpdateLinks:=0, readOnly:=True)
    st = VerifyC1(c, wb, build): wb.Close False: Set wb = Nothing
    If st <> "VERIFIED" Or h <> FileSHA10(file) Then Err.Raise 5, , "Verificacao persistida: " & st
    c.BeginTrans: tx = True
    If Not LockWriter10(c, who, fence) Then Err.Raise 5, , "Fencing recusado na confirmacao"
    If Audit10(c) <> "OK" Then Err.Raise 5, , "Autoridade divergente"
    If Exec10(c, "UPDATE ProjectionAcks SET StatusText='CONFIRMED',FileHash=? WHERE BuildID=? AND RevisionNo=?", h, build, CLng(mCache("@Revision"))) <> 1 Then Err.Raise 5, , "Recibo da projecao divergente"
    Exec10 c, "UPDATE CandidateSettings SET ProjectionState='VERIFIED' WHERE ID=1"
    c.CommitTrans: tx = False
    Release10 c, who, fence: owner = False: c.Close
    Application.EnableEvents = ev: Application.ScreenUpdating = su: Application.DisplayAlerts = alerts: Application.Calculation = calc
    C1_UltimoResultado = "PROJECAO_VERIFICADA; " & build: Exit Sub
Failed:
    e = Err.Number: msg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not c Is Nothing Then
        If tx Then c.RollbackTrans
        Exec10 c, "UPDATE CandidateSettings SET ProjectionState='PENDING' WHERE ID=1"
        If owner Then Release10 c, who, fence
        c.Close
    End If
    For Each sh In ThisWorkbook.Worksheets: sh.Protect PROTECT_C1: Next sh
    Application.EnableEvents = ev: Application.ScreenUpdating = su: Application.DisplayAlerts = alerts: Application.Calculation = calc
    On Error GoTo 0
    C1_UltimoResultado = "PROJECAO_PENDENTE: " & msg
    Err.Raise e, "CandidatoReconstruir", msg
End Sub

Public Sub C1_ConfigLista(ByVal lst As Object)
    lst.Clear: lst.ColumnCount = 17
    lst.ColumnWidths = "80;60;70;40;100;70;70;60;60;80;70;70;0;0;0;0;0"
End Sub
Public Sub C1_Pesquisar(ByVal text As String, ByVal lst As Object, Optional ByVal place As String = "", Optional ByVal history As Boolean = False)
    Dim c As Object, id As Variant, d As Object, matches As New Collection, a() As Variant, i As Long, hay As String
    On Error GoTo Failed
    mUI = True: C1_ConfigLista lst
    Set c = ConnC1(): ReadyC1 c
    For Each id In ViewListC1(c)
        Set d = ViewC1(c, CStr(id))
        If (CBool(d("ACTIVE")) And Not history) Or (history And (d("ROLE") = "EXIT" Or d("ROLE") = "HISTORY_OUT")) Then
            If place = "" Or (place = "FILIAL" And InStr(d("SHEET"), "Filial") > 0) Or (place = "MATRIZ" And InStr(d("SHEET"), "Filial") = 0) Then
                hay = d("NFE") & " " & d("PLATE") & " " & d("MODEL") & " " & d("COD") & " " & d("EpisodeID")
                If InStr(1, hay, text, vbTextCompare) > 0 Then matches.Add d
            End If
        End If
    Next id
    If matches.count > 0 Then
        ReDim a(0 To matches.count - 1, 0 To 16)
        For i = 1 To matches.count
            Set d = matches(i)
            a(i - 1, 0) = d("SHEET"): a(i - 1, 1) = d("NFE"): a(i - 1, 2) = d("OPERACAO"): a(i - 1, 3) = d("UN"): a(i - 1, 4) = d("MODEL"): a(i - 1, 5) = d("COD"): a(i - 1, 6) = d("PLATE"): a(i - 1, 7) = d("ANO"): a(i - 1, 8) = d("COR"): a(i - 1, 9) = Format$(AmountC1(CStr(d("VALOR"))), "0.00"): a(i - 1, 10) = Format$(AsDateC1(d("DATA")), "dd/mm/yyyy"): a(i - 1, 11) = d("OUTRO"): a(i - 1, 12) = d("ViewID"): a(i - 1, 13) = d("KEY"): a(i - 1, 14) = d("Version"): a(i - 1, 15) = d("EpisodeID"): a(i - 1, 16) = d("VehicleID")
        Next i
        lst.list = a
    End If
    c.Close: mUI = False: Exit Sub
Failed:
    On Error Resume Next
    If Not c Is Nothing Then c.Close
    mUI = False: C1_UltimoResultado = "PESQUISA_RECUSADA: " & Err.Description
End Sub
Public Sub C1_SelecionarEntrada(ByVal f As Object, ByVal lst As Object, ByVal filial As Boolean)
    Dim c As Object, d As Object, ctx As Object, s As String, n As Long, e As String
    If mUI Then Exit Sub
    n = lst.ListIndex: If n < 0 Or n >= lst.ListCount Then Exit Sub
    On Error GoTo Failed
    Set c = ConnC1(): ReadyC1 c: Set d = ViewC1(c, CStr(lst.list(n, 12)))
    If CLng(d("Version")) <> CLng(lst.list(n, 14)) Or Not CBool(d("ACTIVE")) Then Err.Raise 5, , "Selecao obsoleta ou registro baixado"
    s = IIf(filial, "Filial", "")
    If d("PLACE") <> IIf(filial, "FILIAL", "MATRIZ") Then Err.Raise 5, , "Selecione a aba da custodia correta"
    C1_LimparContexto f: Set ctx = C1_Context(f)
    ctx("ViewID") = d("ViewID"): ctx("EpisodeID") = d("EpisodeID"): ctx("VehicleID") = d("VehicleID"): ctx("Version") = d("Version")
    SV f, "N_Nfe" & s, d("NFE"): SV f, "ComboBox_Operação" & s, "Compra": SV f, "UN" & s, d("UN"): SV f, "COD" & s, d("COD"): SV f, "ComboBox_VEÍCULO" & s, d("MODEL"): SV f, "PLACA" & s, d("PLATE"): SV f, "Combo_ANO_MOD" & s, d("ANO"): SV f, "ComboBox_COR" & s, d("COR"): SV f, "txt_Valor" & s, Format$(AmountC1(CStr(d("VALOR"))), "0.00"): SV f, IIf(filial, "DataFilial", "data"), Format$(AsDateC1(d("DATA")), "dd/mm/yyyy"): SV f, "ComboBox_OutroEnvio" & s, d("OUTRO")
    If filial Then SV f, "TextBoxChaveDeAcessoFilial", d("KEY") Else C1_ChaveEntrada = d("KEY")
    StatusC1 f, "CONSULTA_POR_ID; " & d("EpisodeID"): c.Close: Exit Sub
Failed:
    e = Err.Description: On Error Resume Next
    If Not c Is Nothing Then c.Close
    C1_LimparContexto f: StatusC1 f, "SELECAO_RECUSADA: " & e
End Sub
Public Sub C1_EnviarSaida(ByVal f As Object, ByVal filial As Boolean, Optional ByVal show As Boolean = True)
    Dim c As Object, ctx As Object, target As Object, d As Object, s As String, k As Variant
    On Error GoTo Failed
    Set ctx = C1_Context(f)
    If Not ctx.Exists("ViewID") Then Err.Raise 5, , "Pesquise e selecione uma entrada"
    Set c = ConnC1(): ReadyC1 c: Set d = ViewC1(c, CStr(ctx("ViewID"))): RecheckC1 c, ctx, d
    Load Cad_NFE_Saida: C1_LimparContexto Cad_NFE_Saida: Set target = C1_Context(Cad_NFE_Saida)
    For Each k In ctx.keys: target(k) = ctx(k): Next k
    s = IIf(filial, "Filial", ""): Cad_NFE_Saida.MultiPage1.value = IIf(filial, 1, 0)
    SV Cad_NFE_Saida, "ComboBox_Operação" & s, IIf(filial, "Devolução", "Venda")
    SV Cad_NFE_Saida, "TextBox_ESTOQUE" & s, d("UN"): SV Cad_NFE_Saida, "TextBox_N_Entrada" & s, d("NFE"): SV Cad_NFE_Saida, "COD" & s, d("COD"): SV Cad_NFE_Saida, "ComboBox_VEÍCULO" & s, d("MODEL"): SV Cad_NFE_Saida, "TextBox_PLACA" & s, d("PLATE"): SV Cad_NFE_Saida, "ComboBox_Ano" & s, d("ANO"): SV Cad_NFE_Saida, "ComboBox_COR" & s, d("COR"): SV Cad_NFE_Saida, "ComboBox_OutroEnvio" & s, d("OUTRO"): SV Cad_NFE_Saida, "TextBox_Valor_Entrada" & s, Format$(AmountC1(CStr(d("VALOR"))), "0.00"): SV Cad_NFE_Saida, "TextBox_Data" & s, Format$(Date, "dd/mm/yyyy")
    SV Cad_NFE_Saida, "TextBoxChaveDeAcesso" & s, d("KEY")
    If filial Then SV Cad_NFE_Saida, "TextBox_Valor_SaidaFilial", Format$(AmountC1(CStr(d("VALOR"))), "0.00")
    c.Close
    If show Then
        f.Hide
        If C1_FornecedorXML Is Nothing Then Cad_NFE_Saida.show Else Cad_NFE_Saida.show vbModeless
    End If
    Exit Sub
Failed:
    C1_UltimoResultado = "ENVIO_RECUSADO: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
    StatusC1 f, C1_UltimoResultado
End Sub
Public Sub C1_Navegar(ByVal lst As Object)
    Dim c As Object, d As Object, sh As Worksheet, r As Long, found As Long, n As Long
    n = lst.ListIndex: If n < 0 Then Exit Sub
    On Error GoTo Failed
    Set c = ConnC1(): ReadyC1 c: Set d = ViewC1(c, CStr(lst.list(n, 12)))
    If CLng(d("Version")) <> CLng(lst.list(n, 14)) Then Err.Raise 5, , "Token obsoleto"
    Set sh = ThisWorkbook.Worksheets(CStr(d("SHEET")))
    For r = StartC1(sh.name) To sh.Cells(sh.rows.count, 52).End(xlUp).row
        If CStr(sh.Cells(r, 52).Value2) = d("ViewID") Then
            If found <> 0 Then Err.Raise 5, , "Identidade duplicada na projecao"
            found = r
        End If
    Next r
    If found = 0 Then Err.Raise 5, , "Projecao desatualizada; reconstruir"
    If sh.Cells(found, 53).Value2 <> d("EpisodeID") Or sh.Cells(found, 55).Value2 <> d("VehicleID") Then Err.Raise 5, , "Linhagem da celula divergente"
    c.Close: ThisWorkbook.Activate: sh.Activate: Application.Goto sh.Cells(found, 1), True: Exit Sub
Failed:
    C1_UltimoResultado = "NAVEGACAO_RECUSADA: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Sub
Public Sub C1_Bloqueado(Optional ByVal context As String = "Manutencao legada")
    C1_UltimoResultado = "NAO_INTEGRADO: " & context & "; nenhuma escrita por celulas permitida"
    MsgBox C1_UltimoResultado, vbExclamation, "Candidata sintetica"
End Sub
Public Sub CandidatoAbrirEntrada()
    Call Guard10
    Cad_NFE.show
End Sub
Public Sub CandidatoReentradaSelecionada()
    Dim ctx As Object, c As Object, id As String
    Set ctx = C1_Context(Cad_NFE)
    If Not ctx.Exists("VehicleID") Then Err.Raise 5, , "Selecione explicitamente um veiculo identificado antes da reentrada"
    id = CStr(ctx("VehicleID")): Set c = ConnC1()
    If CLng(Val10(c, "SELECT Count(*) FROM Episodes WHERE VehicleID=? AND StatusText='ACTIVE'", id)) <> 0 Then c.Close: Err.Raise 5, , "Veiculo ainda ativo"
    c.Close: C1_LimparContexto Cad_NFE: Set ctx = C1_Context(Cad_NFE): ctx("ReentryVehicleID") = id
End Sub
Public Sub CandidatoBackup()
    Dim c As Object, who As String, fence As Long, owner As Boolean, folder As String, copy As Object, st As String, dbcopy As String, sourceHash As String, e As Long, msg As String, manifest As Object
    On Error GoTo Failed
    Call Guard10: Set c = ConnC1(): who = NewID10(): fence = Acquire10(c, who): owner = True
    If fence = 0 Then Err.Raise 5, , "Outro escritor ativo"
    ReadyC1 c
    If Val10(c, "SELECT ProjectionState FROM CandidateSettings WHERE ID=1") <> "VERIFIED" Then Err.Raise 5, , "Reconstrua a projecao antes do backup"
    folder = ThisWorkbook.Path & "\Backup_C1_" & Replace(NewID10(), "-", "")
    FSO10.CreateFolder folder
    ' O lease persistido permanece ocupado enquanto as conexoes do teste estao fechadas.
    c.Close: Set c = Nothing
    sourceHash = FileSHA10(ThisWorkbook.Path & "\" & DB_C1)
    dbcopy = folder & "\" & DB_C1
    FSO10.CopyFile ThisWorkbook.Path & "\" & DB_C1, dbcopy, False
    If sourceHash <> FileSHA10(dbcopy) Or sourceHash <> FileSHA10(ThisWorkbook.Path & "\" & DB_C1) Then Err.Raise 5, , "Fonte mudou durante copia fechada"
    Set copy = CreateObject("ADODB.Connection"): copy.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & dbcopy
    If Not Own10(copy, who, fence) Then Err.Raise 5, , "Proveniencia do lease da copia divergente"
    st = Audit10(copy): If st <> "OK" Then Err.Raise 5, , "Backup recusado: " & st
    ' Libera somente o lease da copia verificada; nao altera fatos ou recibos.
    Release10 copy, who, fence: copy.Close: Set copy = Nothing
    ThisWorkbook.SaveCopyAs folder & "\" & ThisWorkbook.name
    Set manifest = FSO10.CreateTextFile(folder & "\manifesto_backup.tsv", False, True)
    manifest.WriteLine "SOURCE_CLOSED_SHA256" & vbTab & sourceHash
    manifest.WriteLine DB_C1 & vbTab & FileSHA10(dbcopy)
    manifest.WriteLine ThisWorkbook.name & vbTab & FileSHA10(folder & "\" & ThisWorkbook.name)
    manifest.WriteLine "LIMITACAO" & vbTab & "Lease coordena somente escritores conformes; restauracao exige copia isolada e rebind explicito."
    manifest.Close
    Set c = ConnC1(): If Not Own10(c, who, fence) Then Err.Raise 5, , "Autoridade do backup foi perdida"
    Release10 c, who, fence: owner = False: c.Close
    C1_UltimoResultado = "BACKUP_VERIFICADO: " & folder: Exit Sub
Failed:
    e = Err.Number: msg = Err.Description
    On Error Resume Next
    If Not copy Is Nothing Then copy.Close
    If c Is Nothing Then Set c = ConnC1()
    If owner Then Release10 c, who, fence
    If Not c Is Nothing Then c.Close
    C1_UltimoResultado = "BACKUP_PENDENTE: " & msg
    On Error GoTo 0: Err.Raise e, , msg
End Sub
Public Function C1_ArquivoXML() As String
    If Not C1_FornecedorXML Is Nothing Then C1_ArquivoXML = C1_FornecedorXML.Arquivo()
    If Len(C1_ArquivoXML) > 0 Then
        If InStr(1, C1_ArquivoXML, ThisWorkbook.Path & "\", vbTextCompare) <> 1 Then Err.Raise 5, , "XML de teste fora do laboratorio"
    End If
End Function

Public Sub CandidatoCarregarCache()
    Dim c As Object
    On Error GoTo Failed
    Set c = ConnC1(): ReadyC1 c: CacheC1 c: c.Close
    C1_UltimoResultado = "CACHE_AUDITADO": Exit Sub
Failed:
    C1_UltimoResultado = "CACHE_BLOQUEADO: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Sub
Public Function C1_Diagnostico() As String
    Dim c As Object, build As String, verified As String
    On Error GoTo Failed
    Set c = ConnC1(): ReadyC1 c: CacheC1 c
    build = CStr(ThisWorkbook.Worksheets("Entradas").Range("BI3").Value2)
    verified = VerifyC1(c, ThisWorkbook, build)
    If Val10(c, "SELECT ProjectionState FROM CandidateSettings WHERE ID=1") <> "VERIFIED" Then verified = "PENDING_MEMORY_" & verified
    C1_Diagnostico = "AUDIT=" & Audit10(c) & ";PROJECTION=" & verified & ";REV=" & Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1") & ";MOV=" & Val10(c, "SELECT Count(*) FROM Movements") & ";RECEIPTS=" & Val10(c, "SELECT Count(*) FROM Receipts") & ";ACTIVE=" & Val10(c, "SELECT Count(*) FROM Episodes WHERE StatusText='ACTIVE'") & ";PENDING=" & Val10(c, "SELECT Count(*) FROM Pending")
    c.Close: Exit Function
Failed:
    C1_Diagnostico = "ERROR:" & Err.Number & ":" & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function

Private Function IgnoredC1(ByVal sh As Worksheet, ByVal row As Long, ByVal col As Long) As Boolean
    Dim protected As Boolean, e As Long, msg As String
    protected = sh.ProtectContents
    On Error GoTo Failed
    If protected Then sh.Unprotect PROTECT_C1
    IgnoredC1 = sh.Cells(row, col).Errors(xlNumberAsText).Ignore
    If protected Then sh.Protect PROTECT_C1
    Exit Function
Failed:
    e = Err.Number: msg = Err.Description
    On Error Resume Next
    If protected Then sh.Protect PROTECT_C1
    On Error GoTo 0: Err.Raise e, "IgnoredC1", msg
End Function


Public Function C1_MesmoArquivoAutoridade(ByVal stored As String, ByVal actual As String) As Boolean
    On Error GoTo Invalid
    If Not FSO10.FileExists(stored) Or Not FSO10.FileExists(actual) Then Exit Function
    ' Normaliza aliases 8.3 de arquivos existentes; nao altera o vinculo persistido.
    C1_MesmoArquivoAutoridade = (StrComp(FSO10.GetFile(stored).ShortPath, FSO10.GetFile(actual).ShortPath, vbTextCompare) = 0)
    Exit Function
Invalid:
    C1_MesmoArquivoAutoridade = False
End Function

