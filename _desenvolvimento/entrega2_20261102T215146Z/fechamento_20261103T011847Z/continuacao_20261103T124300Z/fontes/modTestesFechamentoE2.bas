Attribute VB_Name = "modTestesFechamentoE2"
Option Explicit
Private Const KEY_E1 As String = "01234567890123456789012345678901234567890123"
Private g As clsTesteCandidato
Public F2_Operation As String, F2_Plan As String

Private Sub R(ByVal id As String, ByVal ok As Boolean, ByVal expected As String, ByVal observed As String, Optional ByVal modality As String = "VBA_FUNCIONAL_EVENTOS_REAIS")
    Dim f As Object
    Set f = FSO10.OpenTextFile(ThisWorkbook.Path & "\resultados_fechamento.tsv", 8, True, -1)
    f.WriteLine id & vbTab & IIf(ok, "PASS", "FAIL") & vbTab & modality & vbTab & Replace(expected, vbTab, " ") & vbTab & Replace(Replace(observed, vbCrLf, " "), vbTab, " ") & vbTab & ThisWorkbook.FullName & vbTab & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    f.Close: Debug.Print id; " "; IIf(ok, "PASS", "FAIL"); " "; observed
End Sub
Private Function Digest(ByVal c As Object) As String
    Dim a As Variant, i As Long, j As Long, s As String
    a = BusinessRows10(c)
    If IsEmpty(a) Then Digest = SHA10("EMPTY"): Exit Function
    For i = LBound(a, 2) To UBound(a, 2)
        For j = LBound(a, 1) To UBound(a, 1): s = s & Pack10(TypeName(a(j, i)), Clean10(a(j, i))): Next j
    Next i
    Digest = SHA10(s)
End Function
Public Sub F2_Snapshot(ByVal label As String)
    Dim c As Object, t As Variant, rs As Object, f As Object, i As Long, s As String
    Set c = C3_Conectar(): Set f = FSO10.CreateTextFile(ThisWorkbook.Path & "\f2_estado_" & label & "_" & NewID10() & ".tsv", False, True)
    For Each t In Array("Meta", "Vehicles", "Episodes", "Operations", "Movements", "Projections", "SheetProjections", "Receipts", "ChangeLog", "Pending", "Attempts", "ProjectionAcks", "CandidateSettings", "WriterLease", "AuthorityBinding", "ReconciliationPlans", "ReconciliationReceipts")
        Set rs = c.Execute("SELECT * FROM [" & t & "]")
        Do While Not rs.EOF
            s = CStr(t)
            For i = 0 To rs.Fields.Count - 1: s = s & vbTab & rs.Fields(i).Name & "=" & Clean10(rs.Fields(i).Value): Next i
            f.WriteLine s: rs.MoveNext
        Loop
        rs.Close
    Next t
    c.Close: f.Close
End Sub
Private Sub FillEntryE1(ByVal nf As String, ByVal plate As String, ByVal COD As String, ByVal amount As String, Optional ByVal filial As Boolean = False)
    Dim s As String
    Load Cad_NFE
    s = IIf(filial, "Filial", "")
    Cad_NFE.MultiPage1.value = IIf(filial, 1, 0): C1_LimparContexto Cad_NFE
    With Cad_NFE
        .Controls("N_Nfe" & s).value = nf: .Controls("ComboBox_Operação" & s).value = "Compra": .Controls("UN" & s).value = "1": .Controls("COD" & s).value = COD: .Controls("ComboBox_VEÍCULO" & s).value = "FICTICIO A": .Controls("PLACA" & s).value = plate: .Controls("Combo_ANO_MOD" & s).value = "2026": .Controls("ComboBox_COR" & s).value = "BRANCO": .Controls("txt_Valor" & s).value = amount: .Controls(IIf(filial, "DataFilial", "data")).value = "01/11/2026": .Controls("ComboBox_OutroEnvio" & s).value = "TESTE"
        If filial Then .TextBoxChaveDeAcessoFilial.value = KEY_E1 Else C1_Context(Cad_NFE)("ENTRYKEY") = KEY_E1
    End With
End Sub
Private Sub SelectE1(ByVal plate As String, ByVal filial As Boolean)
    Dim lst As Object, n As Long
    Cad_NFE.MultiPage1.value = IIf(filial, 1, 0)
    Cad_NFE.Controls("ComboBox_Operação" & IIf(filial, "Filial", "")).value = "Compra"
    If filial Then
        Cad_NFE.TextBoxPesquisaFilial.value = "": Cad_NFE.TextBoxPesquisaFilial.value = plate: Set lst = Cad_NFE.ListBoxLista2
    Else
        Cad_NFE.TextBoxPesquisa.value = "": Cad_NFE.TextBoxPesquisa.value = plate: Set lst = Cad_NFE.ListBoxLista
    End If
    If lst.ListCount <> 1 Then Err.Raise 5, , "Pesquisa deveria identificar exatamente uma custodia ativa; count=" & lst.ListCount
    lst.ListIndex = 0
    If Not C1_Context(Cad_NFE).Exists("ViewID") Then Err.Raise 5, , "Evento de selecao nao estabeleceu identidade"
End Sub
Private Sub FillOutputE1(ByVal operation As String, ByVal nf As String, ByVal amount As String, ByVal filial As Boolean)
    Dim s As String
    s = IIf(filial, "Filial", "")
    With Cad_NFE_Saida
        .Controls("ComboBox_Operação" & s).value = operation
        .Controls("TextBox_N_Saida" & s).value = nf: .Controls("TextBox_Valor_Saida" & s).value = amount: .Controls("TextBox_Data" & s).value = "01/11/2026"
        .Controls("ComboBox_Por" & s).value = "12": .Controls("TextBox_ICMS" & s).value = "0": .Controls("TextBox_PIS" & s).value = "0": .Controls("TextBox_COFINS" & s).value = "0": .Controls("TextBoxChaveDeAcesso" & s).value = KEY_E1
        If Not filial Then .TextBoxBaseICMS.value = "0": .TextBoxBasePISeCOFINS.value = "0"
    End With
End Sub
Public Sub F2_Baseline()
    Dim c As Object, d As String
    On Error GoTo Bad
    Set c = C3_Conectar(): d = C1_Diagnostico()
    R "F2_INSTALACAO", Val10(c, "SELECT SchemaNo FROM Meta WHERE ID=1") = 7 And Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1") = 16 And Audit10(c) = "OK" And InStr(d, "PROJECTION=VERIFIED") > 0, "Referencia16 preservada; contrato7 auditado e projecao persistida", d, "VBA_INTEGRACAO_AUTORIDADE"
    R "F2_APP_BASELINE", Application.EnableEvents And Application.ScreenUpdating And Application.DisplayAlerts And Application.Calculation = xlCalculationAutomatic, "Estados Application preservados", CStr(Application.EnableEvents) & ";" & Application.ScreenUpdating & ";" & Application.DisplayAlerts & ";" & Application.Calculation, "VBA_INTEGRACAO_AUTORIDADE"
    c.Close: F2_Snapshot "baseline": Exit Sub
Bad:
    R "F2_BASELINE_ERRO", False, "Instalar sem regressao", Err.Number & ":" & Err.Description
End Sub
Public Sub F2_FalhaEConciliacao()
    Dim c As Object, n As Long, ep As String, v As Long, before As String, st As String, token As String
    On Error GoTo Bad
    Set g = New clsTesteCandidato: Set C1_FornecedorXML = g
    g.Etapa = "APPLY": g.Codigo = 6: Set C1_PoliticaFalha = g
    Load Cad_NFE: SelectE1 "FIC5C03", True
    ep = C1_Context(Cad_NFE)("EpisodeID"): v = C1_Context(Cad_NFE)("Version")
    Cad_NFE.CommandButtonCopiarEntradaFilial.Value = True: FillOutputE1 "Devolução", "34001", "450", True
    Set c = C3_Conectar(): n = Val10(c, "SELECT Count(*) FROM Receipts"): before = Digest(c): c.Close
    F2_Snapshot "falha_antes"
    Cad_NFE_Saida.CommandButton_GravarFilial.Value = True
    Set C1_PoliticaFalha = Nothing
    Set c = C3_Conectar(): F2_Operation = CStr(Val10(c, "SELECT OperationID FROM Pending WHERE EpisodeID=?", ep))
    R "F2_PRECOMMIT_FATOS", IDOK10(F2_Operation) And Val10(c, "SELECT Count(*) FROM Receipts") = n And Digest(c) = before And Val10(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep) = v And Audit10(c) = "OK", "Rollback de fatos; intencao preservada para conciliacao", C1_UltimoResultado & ";" & F2_Operation
    c.Close
    Cad_NFE_Saida.CommandButton_GravarFilial.Value = True
    R "F2_REPETICAO_PENDENTE", InStr(C1_UltimoResultado, "PENDENTE") > 0, "Mesmo OperationID bloqueado", C1_UltimoResultado
    SelectE1 "FIC5C03", True: Cad_NFE.CommandButtonCopiarEntradaFilial.Value = True: FillOutputE1 "Devolução", "34002", "450", True
    Cad_NFE_Saida.CommandButton_GravarFilial.Value = True
    Set c = C3_Conectar()
    R "F2_OUTRO_ID_PENDENTE", Val10(c, "SELECT Count(*) FROM Receipts") = n And Val10(c, "SELECT Count(*) FROM Pending WHERE EpisodeID=?", ep) = 1, "Novo OperationID nao contorna pendencia", C1_UltimoResultado
    c.Close
    F2_Plan = C3_Analisar(F2_Operation)
    If Not IDOK10(F2_Plan) Then Err.Raise 5, , F2_Plan
    token = C3_TokenConfirmacao(F2_Plan)
    st = C3_Resolver(F2_Plan, "QA SINTETICO", "Teste negativo sem confirmacao", "NAO_CONFIRMADO")
    R "F2_EXIGE_CONFIRMACAO", InStr(st, "RECUSADO") > 0, "Sem token plano nao e aplicado", st, "VBA_CONCILIACAO_API"
    st = C3_Resolver(F2_Plan, "QA SINTETICO", "Confirma abandono de intencao sem efeitos; falha injetada antes do commit", token)
    Set c = C3_Conectar()
    R "F2_CONCILIACAO_SEM_EFEITOS", st = "CONCILIADO_SEM_EFEITOS" And Digest(c) = before And Val10(c, "SELECT Count(*) FROM Pending WHERE EpisodeID=?", ep) = 0 And Val10(c, "SELECT Count(*) FROM ReconciliationReceipts WHERE OperationID=?", F2_Operation) = 1 And Audit10(c) = "OK", "Intencao abortada auditavelmente; fatos e versao intactos", st & ";AUDIT=" & Audit10(c), "VBA_CONCILIACAO_API_CONFIRMACAO_SINTETICA"
    c.Close
    st = C3_Resolver(F2_Plan, "QA SINTETICO", "Repeticao deliberada", token)
    R "F2_CONCILIACAO_IDEMPOTENTE", st = "JA_CONCILIADO", "Repeticao nao duplica recibo", st, "VBA_CONCILIACAO_API"
    Cad_NFE_Saida.Hide: F2_Snapshot "conciliacao_depois": Exit Sub
Bad:
    Set C1_PoliticaFalha = Nothing
    R "F2_CONCILIACAO_ERRO", False, "Compensacao limitada e conciliacao auditada", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub

Public Sub F2_Restaurada()
    Dim c As Object, d As String
    On Error GoTo Bad
    Set c = C3_Conectar(): d = C1_Diagnostico()
    R "F2_RESTAURACAO_SESSAO_FRIA", Val10(c, "SELECT Generation FROM AuthorityBinding WHERE ID=1") = 2 And Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1") = 16 And Val10(c, "SELECT Count(*) FROM ReconciliationReceipts") = 1 And Audit10(c) = "OK" And InStr(d, "PROJECTION=VERIFIED") > 0, "Nova pasta, mesma linhagem/rev16/recibo conciliacao; projecao reconstruida e verificada", d, "VBA_RESTAURACAO_REAL"
    R "F2_CAMINHOS_83_LONGO", C1_MesmoArquivoAutoridade(FSO10.GetFile(ThisWorkbook.FullName).ShortPath, ThisWorkbook.FullName), "Alias 8.3 identifica mesmo arquivo", ThisWorkbook.FullName, "VBA_CAMINHOS"
    c.Close: F2_Snapshot "restaurada": Exit Sub
Bad: R "F2_RESTAURACAO_ERRO", False, "Reabrir autoridade explicitamente realocada", Err.Number & ":" & Err.Description
End Sub
Public Sub F2_EntradasNovas()
    Dim c As Object, n As Long, j As Long, nf As String, placa As String, valor As Currency, s As String
    On Error GoTo Bad
    For j = 0 To 1
        nf = IIf(j = 0, "12001", "12002"): placa = IIf(j = 0, "F2A0001", "F2F0002"): valor = IIf(j = 0, 123.4567@, 456.7891@)
        Set c = C3_Conectar(): n = Val10(c, "SELECT Count(*) FROM Receipts"): c.Close
        FillEntryE1 nf, placa, IIf(j = 0, "F2A", "F2F"), Replace(CStr(valor), ".", ","), CBool(j)
        If j = 0 Then Cad_NFE.CommandButton_Gravar.Value = True Else Cad_NFE.CommandButton_GravarFilial.Value = True
        Set c = C3_Conectar(): s = C1_Diagnostico()
        R "F2_ENTRADA_" & IIf(j = 0, "MATRIZ", "FILIAL"), Val10(c, "SELECT Count(*) FROM Receipts") = n + 1 And CCur(Val10(c, "SELECT AmountIn FROM Projections WHERE DocNumber=?", nf)) = valor And Val10(c, "SELECT Count(*) FROM SheetProjections S INNER JOIN Projections P ON S.ProjectionID=P.ProjectionID WHERE P.DocNumber=?", nf) = 2 And InStr(s, "PROJECTION=VERIFIED") > 0, "Evento real; novo veiculo/episodio; duas projecoes; quatro casas; chave texto", C1_UltimoResultado & ";" & s
        c.Close
    Next j
    Set c = C3_Conectar(): n = Val10(c, "SELECT Count(*) FROM Receipts"): c.Close
    FillEntryE1 "12002", "F2Z0003", "F2Z", "999,9999", True: Cad_NFE.CommandButton_GravarFilial.Value = True
    Set c = C3_Conectar()
    R "F2_DUPLICIDADE_ENTRADA", Val10(c, "SELECT Count(*) FROM Receipts") = n And Val10(c, "SELECT Count(*) FROM Vehicles WHERE Plate='F2Z0003'") = 0, "Duplicidade recusa antes de criar identidade", C1_UltimoResultado
    c.Close: F2_Snapshot "entradas_novas": Exit Sub
Bad: R "F2_ENTRADAS_ERRO", False, "Cadastrar somente dados sinteticos", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Private Sub SaidaF2(ByVal id As String, ByVal placa As String, ByVal kind As String, ByVal nf As String, ByVal valor As String, ByVal filial As Boolean, ByVal xml As String)
    Dim c As Object, ep As String, vid As String, ver As Long, n As Long, mid As String, st As String, expected As String, places As Long
    On Error GoTo Bad
    Load Cad_NFE: SelectE1 placa, filial
    ep = C1_Context(Cad_NFE)("EpisodeID"): vid = C1_Context(Cad_NFE)("VehicleID"): ver = C1_Context(Cad_NFE)("Version")
    Set g = New clsTesteCandidato: g.ArquivoXML = ThisWorkbook.Path & "\" & xml: Set C1_FornecedorXML = g
    If filial Then Cad_NFE.CommandButtonCopiarEntradaFilial.Value = True Else Cad_NFE.CommandButtonCopiarEntrada.Value = True
    FillOutputE1 kind, nf, valor, filial
    Set c = C3_Conectar(): n = Val10(c, "SELECT Count(*) FROM Receipts"): c.Close
    F2_Snapshot id & "_antes"
    If filial Then Cad_NFE_Saida.CommandButton_GravarFilial.Value = True Else Cad_NFE_Saida.CommandButton_Gravar.Value = True
    Set c = C3_Conectar(): st = C1_Diagnostico()
    mid = CStr(Val10(c, "SELECT MovementID FROM Projections WHERE DocNumber=?", nf))
    expected = IIf(kind = "Transferência", "FILIAL", "MATRIZ"): places = IIf(kind = "Transferência" Or kind = "Devolução", 4, 2)
    R id, Val10(c, "SELECT Count(*) FROM Receipts") = n + 1 And Val10(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep) = ver + 1 And Val10(c, "SELECT VehicleID FROM Episodes WHERE EpisodeID=?", ep) = vid And Val10(c, "SELECT Count(*) FROM SheetProjections WHERE MovementID=?", mid) = places And Val10(c, "SELECT PlaceCode FROM Episodes WHERE EpisodeID=?", ep) = expected And InStr(st, "PROJECTION=VERIFIED") > 0, "Evento real preserva VehicleID/EpisodeID; versiona movimento; destinos corretos; projecao integral", C1_UltimoResultado & ";" & st
    If kind = "Venda" Then R id & "_BAIXA", Val10(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", ep) = "CLOSED", "Baixa logica encerra episodio, sem deslocar linhas historicas", CStr(Val10(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", ep))
    c.Close: Cad_NFE_Saida.Hide: F2_Snapshot id & "_depois": Exit Sub
Bad: R id & "_ERRO", False, "Concluir fluxo real", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub F2_Transferencia()
    SaidaF2 "F2_T_VALIDO", "F2A0001", "Transferência", "24002", "300,1111", False, "f2_transferencia.xml"
End Sub
Public Sub F2_Devolucao()
    SaidaF2 "F2_D_VALIDO", "F2A0001", "Devolução", "34011", "300,1111", True, ""
End Sub
Public Sub F2_Venda()
    SaidaF2 "F2_V_VALIDO", "F2A0001", "Venda", "24012", "400,2222", False, "f2_venda.xml"
End Sub
Public Sub F2_Cancelado()
    SaidaF2 "F2_CANCELADO", "FIC5B02", "Cancelado", "24990", "0", False, ""
End Sub
Public Sub F2_ReentradaNova()
    Dim c As Object, vid As String, ep As String, n As Long
    On Error GoTo Bad
    Set c = C3_Conectar(): vid = Val10(c, "SELECT VehicleID FROM Vehicles WHERE Plate='F2A0001'"): n = Val10(c, "SELECT Count(*) FROM Receipts"): c.Close
    Load Cad_NFE: Cad_NFE.MultiPage1.Value = 0: C1_LimparContexto Cad_NFE
    Cad_NFE.ComboBox_Operação.Value = "Reentrada": Cad_NFE.TextBoxPesquisa.Value = "": Cad_NFE.TextBoxPesquisa.Value = "F2A0001"
    If Cad_NFE.ListBoxLista.ListCount <> 1 Then Err.Raise 5, , "Escolha de episodio nao univoca"
    Cad_NFE.ListBoxLista.ListIndex = 0
    Cad_NFE.N_Nfe.Value = "12003": Cad_NFE.txt_Valor.Value = "150,3333": Cad_NFE.Data.Value = "03/11/2026": Cad_NFE.UN.Value = "1": Cad_NFE.ComboBox_OutroEnvio.Value = "TESTE": C1_Context(Cad_NFE)("ENTRYKEY") = KEY_E1
    Cad_NFE.CommandButton_Gravar.Value = True
    Set c = C3_Conectar()
    R "F2_REENTRADA_INTERFACE", Val10(c, "SELECT Count(*) FROM Receipts") = n + 1 And Val10(c, "SELECT Count(*) FROM Episodes WHERE VehicleID=?", vid) = 2 And Val10(c, "SELECT Count(*) FROM Episodes WHERE VehicleID=? AND StatusText='ACTIVE'", vid) = 1 And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Lista real seleciona episodio fechado; novo episodio mesmo VID; valor150,3333", C1_UltimoResultado & ";" & C1_Diagnostico()
    c.Close: F2_Snapshot "reentrada": Exit Sub
Bad: R "F2_REENTRADA_ERRO", False, "Identidade comprovada pela selecao real", Err.Number & ":" & Err.Description
End Sub
Public Sub F2_FiltrosReais()
    Dim c As Object, n As Long, k As Variant, p As String
    On Error GoTo Bad
    Set c = C3_Conectar(): Load UserFormFiltro
    UserFormFiltro.CommandButtonMatriz.Value = True: UserFormFiltro.CommandButtonFiltroCompra.Value = True
    n = Val10(c, "SELECT Count(*) FROM Episodes WHERE StatusText='ACTIVE' AND PlaceCode='MATRIZ'")
    R "F2_FILTRO_COMPRA_MATRIZ", Cad_NFE.ListBoxLista.ListCount = n, "Botao real lista somente custodias ativas matriz", "LISTA=" & Cad_NFE.ListBoxLista.ListCount & ";BANCO=" & n
    Cad_NFE.Hide
    UserFormFiltro.CommandButtonFilial.Value = True: UserFormFiltro.CommandButtonFiltroCompra.Value = True
    n = Val10(c, "SELECT Count(*) FROM Episodes WHERE StatusText='ACTIVE' AND PlaceCode='FILIAL'")
    R "F2_FILTRO_COMPRA_FILIAL", Cad_NFE.ListBoxLista2.ListCount = n, "Botao real lista somente custodias ativas filial", "LISTA=" & Cad_NFE.ListBoxLista2.ListCount & ";BANCO=" & n
    Cad_NFE.Hide
    UserFormFiltro.CommandButtonMatriz.Value = True: UserFormFiltro.CommandButtonVenda.Value = True
    n = Val10(c, "SELECT Count(*) FROM SheetProjections S INNER JOIN Projections P ON S.ProjectionID=P.ProjectionID WHERE S.SheetKey='Saidas' AND P.KindText='SALE'")
    R "F2_FILTRO_VENDA_HISTORICO", Cad_NFE_Saida.ListBoxLista.ListCount = n, "Venda real sem duplicar Todas Saidas", "LISTA=" & Cad_NFE_Saida.ListBoxLista.ListCount & ";BANCO=" & n
    c.Close: Cad_NFE_Saida.Hide: Exit Sub
Bad: R "F2_FILTROS_ERRO", False, "Filtros reais por identidade", Err.Number & ":" & Err.Description
End Sub
