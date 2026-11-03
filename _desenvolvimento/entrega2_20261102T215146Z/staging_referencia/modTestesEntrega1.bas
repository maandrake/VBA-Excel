Attribute VB_Name = "modTestesEntrega1"
Option Explicit
Private gate As clsTesteCandidato
Private episodeA As String, vehicleA As String, originA As String
Private Const KEY_E1 As String = "01234567890123456789012345678901234567890123"
Private Function OpenE1() As Object
    Call Guard10
    Set OpenE1 = Open10("Banco_Candidato_Operacoes_Corrigidas.accdb")
End Function
Private Sub ResultE1(ByVal id As String, ByVal ok As Boolean, ByVal expected As String, ByVal observed As String)
    Dim f As Object, file As String
    file = ThisWorkbook.Path & "\resultados_entrega2_referencia.tsv"
    Set f = FSO10.OpenTextFile(file, 8, True, -1)
    f.WriteLine id & "_R3" & vbTab & IIf(ok, "PASS", "FAIL") & vbTab & "VBA_FUNCIONAL_EVENTOS_REAIS" & vbTab & Replace(expected, vbTab, " ") & vbTab & Replace(Replace(observed, vbCrLf, " "), vbTab, " ") & vbTab & ThisWorkbook.FullName & vbTab & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    f.Close: Debug.Print id; " "; IIf(ok, "PASS", "FAIL"); " "; observed
End Sub
Private Sub DumpE1(ByVal label As String)
    Dim c As Object, t As Variant, r As Object, f As Object, i As Long, line As String, sh As Worksheet, cell As Range
    Set c = OpenE1(): Set f = FSO10.CreateTextFile(ThisWorkbook.Path & "\estado_e2_referencia_" & label & ".tsv", False, True)
    For Each t In Array("Meta", "Vehicles", "Episodes", "Operations", "Movements", "Projections", "SheetProjections", "Receipts", "ChangeLog", "Pending", "ProjectionAcks", "CandidateSettings")
        Set r = c.Execute("SELECT * FROM [" & t & "]")
        Do While Not r.EOF
            line = CStr(t)
            For i = 0 To r.fields.count - 1
                line = line & vbTab & r.fields(i).name & "=" & Clean10(r.fields(i).value)
            Next i
            f.WriteLine line: r.MoveNext
        Loop
        r.Close
    Next t
    c.Close
    For Each sh In ThisWorkbook.Worksheets
        If sh.name = "Entradas" Or sh.name = "Entradas Filial" Or sh.name = "Entradas Anual" Or sh.name = "Entradas Anual Filial" Or InStr(sh.name, "Saidas") > 0 Then
            For Each cell In sh.UsedRange.Cells
                If Len(CStr(cell.Formula)) > 0 Then f.WriteLine "CELL" & vbTab & sh.name & vbTab & cell.Address(False, False) & vbTab & TypeName(cell.Value2) & vbTab & Clean10(cell.Value2) & vbTab & cell.Formula & vbTab & cell.NumberFormat
            Next cell
        End If
    Next sh
    f.Close
End Sub
Private Sub FillEntryE1(ByVal nf As String, ByVal plate As String, ByVal COD As String, ByVal amount As String, Optional ByVal filial As Boolean = False)
    Dim s As String
    Load Cad_NFE: C1_LimparContexto Cad_NFE
    s = IIf(filial, "Filial", "")
    Cad_NFE.MultiPage1.value = IIf(filial, 1, 0)
    With Cad_NFE
        .Controls("N_Nfe" & s).value = nf: .Controls("ComboBox_Operação" & s).value = "Compra": .Controls("UN" & s).value = "1": .Controls("COD" & s).value = COD: .Controls("ComboBox_VEÍCULO" & s).value = "FICTICIO A": .Controls("PLACA" & s).value = plate: .Controls("Combo_ANO_MOD" & s).value = "2026": .Controls("ComboBox_COR" & s).value = "BRANCO": .Controls("txt_Valor" & s).value = amount: .Controls(IIf(filial, "DataFilial", "data")).value = "01/11/2026": .Controls("ComboBox_OutroEnvio" & s).value = "TESTE"
        If filial Then .TextBoxChaveDeAcessoFilial.value = KEY_E1 Else C1_ChaveEntrada = KEY_E1
    End With
End Sub
Private Sub SelectE1(ByVal plate As String, ByVal filial As Boolean)
    Dim lst As Object, n As Long
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
Public Sub E1_Entrada()
    Dim c As Object, diag As String
    On Error GoTo Failed
    Set c = OpenE1(): If CLng(Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1")) <> 0 Then Err.Raise 5, , "Teste exige candidata inicialmente vazia"
    c.Close
    DumpE1 "A_antes"
    FillEntryE1 "11001", "FIC5A01", "F5A", "900"
    Cad_NFE.show vbModeless
    Cad_NFE.CommandButton_Gravar.value = True
    diag = C1_Diagnostico(): Set c = OpenE1()
    ResultE1 "E1_ENTRADA_EVENTO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = 1 And InStr(diag, "PROJECTION=VERIFIED") > 0, "1 recibo, entrada ativa e projecao persistida verificada", C1_UltimoResultado & ";" & diag
    ResultE1 "E1_ENTRADA_VALOR", CCur(Val10(c, "SELECT AmountIn FROM Projections")) = 900@ And ThisWorkbook.Worksheets("Entradas").Range("I4").Value2 = 900#, "900 exatos no ACE e na projecao", "ACE=" & Val10(c, "SELECT AmountIn FROM Projections") & ";EXCEL=" & ThisWorkbook.Worksheets("Entradas").Range("I4").Value2
    c.Close: DumpE1 "A_depois_entrada": Exit Sub
Failed:
    ResultE1 "E1_ENTRADA_ERRO", False, "Concluir entrada pelo evento real", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub E1_Venda()
    Dim c As Object, diag As String
    On Error GoTo Failed
    Set gate = New clsTesteCandidato: gate.ArquivoXML = ThisWorkbook.Path & "\valido.xml": Set C1_FornecedorXML = gate
    SelectE1 "FIC5A01", False
    episodeA = CStr(C1_Context(Cad_NFE)("EpisodeID")): vehicleA = CStr(C1_Context(Cad_NFE)("VehicleID")): originA = CStr(C1_Context(Cad_NFE)("ViewID"))
    Cad_NFE.CommandButtonCopiarEntrada.value = True
    FillOutputE1 "Venda", "22001", "1000", False
    DumpE1 "V_antes"
    Cad_NFE_Saida.CommandButton_Gravar.value = True
    diag = C1_Diagnostico(): Set c = OpenE1()
    ResultE1 "E1_V_VALIDO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = 2 And Val10(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", episodeA) = "CLOSED" And InStr(diag, "PROJECTION=VERIFIED") > 0, "Venda confirmada, origem preservada e inelegivel, 2 destinos vinculados", C1_UltimoResultado & ";" & diag
    ResultE1 "E1_V_IDENTIDADE", CLng(Val10(c, "SELECT Count(*) FROM SheetProjections WHERE EpisodeID=?", episodeA)) = 4 And Val10(c, "SELECT VehicleID FROM Episodes WHERE EpisodeID=?", episodeA) = vehicleA, "4 registros de projecao do mesmo episodio e veiculo", diag
    ResultE1 "E1_V_CHAVE_TEXTO", ThisWorkbook.Worksheets("Entradas").Range("BH4").Value2 = KEY_E1 And TypeName(ThisWorkbook.Worksheets("Entradas").Range("BH4").Value2) = "String", "44 digitos com zero inicial como String", ThisWorkbook.Worksheets("Entradas").Range("BH4").Value2
    c.Close: DumpE1 "V_depois": Cad_NFE_Saida.Hide: Exit Sub
Failed:
    ResultE1 "E1_VENDA_ERRO", False, "Concluir venda pelo evento real", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub E1_Diagnostico()
    Debug.Print C1_Diagnostico()
End Sub

Public Sub E1_IniciarBaseLimpa()
    Call Guard10
    Call CreateDB10("Banco_Candidato_Operacoes_Corrigidas.accdb")
    Call CandidatoInstalarLaboratorio
End Sub
Public Sub E1_Transferencia()
    Dim c As Object, ep As String, vid As String, diag As String
    On Error GoTo Failed
    FillEntryE1 "11002", "FIC5B02", "F5B", "700"
    Cad_NFE.show vbModeless: Cad_NFE.CommandButton_Gravar.value = True
    Set c = OpenE1()
    ResultE1 "E1_ENTRADA_B", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = 3 And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Nova entrada 700 exatos, 3 recibos", C1_UltimoResultado & ";" & C1_Diagnostico()
    c.Close
    Set gate = New clsTesteCandidato: gate.ArquivoXML = ThisWorkbook.Path & "\transferencia_b.xml": Set C1_FornecedorXML = gate
    SelectE1 "FIC5B02", False: ep = C1_Context(Cad_NFE)("EpisodeID"): vid = C1_Context(Cad_NFE)("VehicleID")
    Cad_NFE.CommandButtonCopiarEntrada.value = True
    FillOutputE1 "Transferência", "22002", "800", False
    DumpE1 "T_antes": Cad_NFE_Saida.CommandButton_Gravar.value = True
    diag = C1_Diagnostico(): Set c = OpenE1()
    ResultE1 "E1_T_VALIDO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = 4 And Val10(c, "SELECT PlaceCode FROM Episodes WHERE EpisodeID=?", ep) = "FILIAL" And InStr(diag, "PROJECTION=VERIFIED") > 0, "Custodia filial, 4 destinos, mesmo episodio e veiculo", C1_UltimoResultado & ";" & diag
    ResultE1 "E1_T_FORMULAS", ThisWorkbook.Worksheets("Entradas Filial").Range("M4").HasFormula And ThisWorkbook.Worksheets("Entradas Filial").Range("N4").HasFormula And ThisWorkbook.Worksheets("Entradas Filial").Range("O4").HasFormula And ThisWorkbook.Worksheets("Entradas Filial").Range("I4").Value2 = 800#, "M/N/O por identidade e valor 800 distinto da entrada 700", diag
    c.Close: DumpE1 "T_depois": Cad_NFE_Saida.Hide: Exit Sub
Failed:
    ResultE1 "E1_TRANSFERENCIA_ERRO", False, "Transferencia por evento real", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub E1_Devolucao()
    Dim c As Object, ep As String, diag As String
    On Error GoTo Failed
    Set gate = New clsTesteCandidato: Set C1_FornecedorXML = gate
    SelectE1 "FIC5B02", True: ep = C1_Context(Cad_NFE)("EpisodeID")
    Cad_NFE.CommandButtonCopiarEntradaFilial.value = True
    FillOutputE1 "Devolução", "33001", "800", True
    DumpE1 "D_antes": Cad_NFE_Saida.CommandButton_GravarFilial.value = True
    diag = C1_Diagnostico(): Set c = OpenE1()
    ResultE1 "E1_D_VALIDO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = 5 And Val10(c, "SELECT PlaceCode FROM Episodes WHERE EpisodeID=?", ep) = "MATRIZ" And InStr(diag, "PROJECTION=VERIFIED") > 0, "Retorno a matriz no mesmo episodio, quatro destinos preservados", C1_UltimoResultado & ";" & diag
    ResultE1 "E1_D_VINCULOS", CLng(Val10(c, "SELECT Count(*) FROM SheetProjections WHERE EpisodeID=?", ep)) = 10 And CLng(Val10(c, "SELECT Count(*) FROM Episodes WHERE EpisodeID=? AND StatusText='ACTIVE'", ep)) = 1, "10 projecoes historicas, uma custodia ativa", diag
    c.Close: DumpE1 "D_depois": Cad_NFE_Saida.Hide: Exit Sub
Failed:
    ResultE1 "E1_DEVOLUCAO_ERRO", False, "Devolucao por evento real", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub

Public Sub E1_Seguranca()
    Dim c As Object, n As Long, i As Long, oldDate As String, diag As String, saver As clsSaveFase9, e As Long, msg As String, folder As String, b As Object, sh As Worksheet, a As Variant, z As Variant
    On Error GoTo Failed
    Set gate = New clsTesteCandidato: Set C1_FornecedorXML = gate: Set C1_PoliticaFalha = Nothing
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    FillEntryE1 "", "FICINV1", "INVAL", "123": Cad_NFE.CommandButton_Gravar.value = True
    Set c = OpenE1(): ResultE1 "E1_ENTRADA_NFE_VAZIA", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n And InStr(C1_UltimoResultado, "Numero da NF-e invalido") > 0, "Recusa sem efeito", C1_UltimoResultado: c.Close
    FillEntryE1 "11991", "FICINV1", "INVAL", "abc": Cad_NFE.CommandButton_Gravar.value = True
    Set c = OpenE1(): ResultE1 "E1_ENTRADA_VALOR_INVALIDO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n And InStr(C1_UltimoResultado, "monetario invalido") > 0, "Recusa sem efeito", C1_UltimoResultado: c.Close
    FillEntryE1 "11991", "FICINV1", "INVAL", "123": Cad_NFE.Data.value = "31/02/2026": Cad_NFE.CommandButton_Gravar.value = True
    Set c = OpenE1(): ResultE1 "E1_ENTRADA_DATA_INVALIDA", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n And InStr(C1_UltimoResultado, "Data civil invalida") > 0, "Recusa sem efeito", C1_UltimoResultado: c.Close
    FillEntryE1 "33001", "FICINV1", "INVAL", "123": Cad_NFE.CommandButton_Gravar.value = True
    Set c = OpenE1(): ResultE1 "E1_ENTRADA_DUPLICADA", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n And InStr(C1_UltimoResultado, "ja cadastrado") > 0, "Recusa de documento ativo duplicado", C1_UltimoResultado: c.Close
    Cad_NFE.TextBoxPesquisa.value = "FIC5A01": ResultE1 "E1_PESQUISA_BAIXADO", Cad_NFE.ListBoxLista.ListCount = 0, "Venda nao aparece no estoque ativo", "COUNT=" & Cad_NFE.ListBoxLista.ListCount
    SelectE1 "FIC5B02", False
    C1_Context(Cad_NFE)("Version") = CLng(C1_Context(Cad_NFE)("Version")) - 1
    Cad_NFE.CommandButtonCopiarEntrada.value = True
    ResultE1 "E1_TOKEN_OBSOLETO", InStr(C1_UltimoResultado, "obsoleta") > 0, "Selecao antiga recusada antes de abrir gravador", C1_UltimoResultado
    SelectE1 "FIC5B02", False: Cad_NFE.ListBoxLista.ListIndex = -1
    ResultE1 "E1_LISTINDEX_NEGATIVO", Cad_NFE.ListBoxLista.ListIndex = -1, "Indice negativo nao executa gravacao", "ListIndex=-1"
    Set sh = ThisWorkbook.Worksheets("Entradas")
    sh.Unprotect "CANDIDATO_SINTETICO"
    a = sh.Range("A4:BH4").Formula: z = sh.Range("A6:BH6").Formula
    sh.Range("A4:BH4").Formula = z: sh.Range("A6:BH6").Formula = a: sh.Protect "CANDIDATO_SINTETICO"
    diag = C1_Diagnostico()
    ResultE1 "E1_REORDENACAO_DETECTADA", InStr(diag, "PROJECTION=VERIFIED") = 0, "Troca fisica sem metadados corretos nao e aprovada", diag
    Call CandidatoReconstruir
    ResultE1 "E1_REORDENACAO_RECONSTRUIDA", InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Reconstrucao integral pela autoridade", C1_Diagnostico()
    Set saver = New clsSaveFase9: saver.Iniciar ThisWorkbook: saver.RecusarUmaVez
    On Error Resume Next
    Call CandidatoReconstruir
    e = Err.Number: msg = Err.Description: Err.Clear
    On Error GoTo Failed
    Set c = OpenE1(): ResultE1 "E1_SAVE_CANCELADO_REAL", saver.Observado And e <> 0 And Val10(c, "SELECT ProjectionState FROM CandidateSettings WHERE ID=1") = "PENDING" And CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n, "Evento BeforeSave cancela; ACE preservado e projecao pendente", CStr(e) & ":" & msg: c.Close
    Set saver = Nothing
    Call CandidatoReconstruir
    ResultE1 "E1_RECONSTRUCAO_APOS_SAVE", InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Recuperacao por reconstruir e reabrir geracao persistida", C1_Diagnostico()
    Call CandidatoBackup
    folder = mid$(C1_UltimoResultado, Len("BACKUP_VERIFICADO: ") + 1)
    Set b = CreateObject("ADODB.Connection"): b.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & folder & "\Banco_Candidato_Operacoes_Corrigidas.accdb"
    ResultE1 "E1_BACKUP_REABERTO", Audit10(b) = "OK" And CLng(Val10(b, "SELECT Count(*) FROM Receipts")) = n And Val10(b, "SELECT StatusText FROM WriterLease WHERE ID=1") = "FREE", "Backup com conexoes fechadas, auditoria e recibos preservados", folder & ";AUDIT=" & Audit10(b): b.Close
    DumpE1 "seguranca_depois"
    Exit Sub
Failed:
    ResultE1 "E1_SEGURANCA_ERRO", False, "Concluir testes dirigidos", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub

Public Sub E1_Reentrada()
    Dim c As Object, vid As String, ep As String, n As Long
    On Error GoTo Failed
    Set c = OpenE1()
    If CLng(Val10(c, "SELECT Count(*) FROM Projections WHERE DocNumber='11001'")) <> 1 Then Err.Raise 5, , "Origem de teste ambigua"
    ep = CStr(Val10(c, "SELECT EpisodeID FROM Projections WHERE DocNumber='11001'"))
    vid = CStr(Val10(c, "SELECT VehicleID FROM Episodes WHERE EpisodeID=?", ep))
    n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    FillEntryE1 "11003", "FIC5A01", "F5A", "650,25"
    C1_Context(Cad_NFE)("VehicleID") = vid
    Call CandidatoReentradaSelecionada
    Cad_NFE.CommandButton_Gravar.value = True
    Set c = OpenE1()
    ResultE1 "E1_REENTRADA_EPISODIO", CLng(Val10(c, "SELECT Count(*) FROM Episodes WHERE VehicleID=?", vid)) = 2 And CLng(Val10(c, "SELECT Count(*) FROM Episodes WHERE VehicleID=? AND StatusText='ACTIVE'", vid)) = 1 And CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n + 1, "VehicleID explicito, episodio novo, episodio vendido preservado", C1_UltimoResultado & ";" & C1_Diagnostico()
    ResultE1 "E1_REENTRADA_VALOR", CCur(Val10(c, "SELECT AmountIn FROM Projections WHERE DocNumber='11003'")) = 650.25@, "Decimal brasileiro preservado com quatro casas", CStr(Val10(c, "SELECT AmountIn FROM Projections WHERE DocNumber='11003'"))
    ResultE1 "E1_HISTORICO_EPISODIO", Val10(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", ep) = "CLOSED", "Reentrada nao reabre episodio anterior", CStr(Val10(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", ep))
    c.Close: DumpE1 "reentrada_depois": Exit Sub
Failed:
    ResultE1 "E1_REENTRADA_ERRO", False, "Reentrada pelo evento de cadastro e selecao explicita de identidade", Err.Number & ":" & Err.Description
End Sub
Public Sub E1_DestinoDuplicado()
    Dim c As Object, n As Long
    On Error GoTo Failed
    FillEntryE1 "22004", "FIC5C03", "F5C", "450", True
    Cad_NFE.CommandButton_GravarFilial.value = True
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    Set gate = New clsTesteCandidato: gate.ArquivoXML = ThisWorkbook.Path & "\transferencia_duplicada.xml": Set C1_FornecedorXML = gate
    SelectE1 "FIC5B02", False: Cad_NFE.CommandButtonCopiarEntrada.value = True
    FillOutputE1 "Transferência", "22004", "800", False
    Cad_NFE_Saida.CommandButton_Gravar.value = True
    Set c = OpenE1()
    ResultE1 "E1_T_DUP_DESTINO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n And InStr(C1_UltimoResultado, "Destino filial") > 0, "Documento ativo no destino recusa transferencia inteira", C1_UltimoResultado & ";" & C1_Diagnostico()
    c.Close: Cad_NFE_Saida.Hide: DumpE1 "duplicidade_depois": Exit Sub
Failed:
    ResultE1 "E1_T_DUP_ERRO", False, "Recusar sem efeitos", Err.Number & ":" & Err.Description
End Sub
Public Sub E1_NovaTransferencia()
    Dim c As Object, n As Long
    On Error GoTo Failed
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    Set gate = New clsTesteCandidato: gate.ArquivoXML = ThisWorkbook.Path & "\transferencia_b2.xml": Set C1_FornecedorXML = gate
    SelectE1 "FIC5B02", False: Cad_NFE.CommandButtonCopiarEntrada.value = True
    FillOutputE1 "Transferência", "22003", "800", False
    Cad_NFE_Saida.CommandButton_Gravar.value = True
    Set c = OpenE1()
    ResultE1 "E1_TRANSFERENCIA_SUCESSIVA", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n + 1 And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Novo movimento, mesmo episodio, projecao reconstruida", C1_UltimoResultado & ";" & C1_Diagnostico()
    c.Close: Cad_NFE_Saida.Hide: DumpE1 "transferencia2_depois": Exit Sub
Failed:
    ResultE1 "E1_TRANSFERENCIA2_ERRO", False, "Concluir segunda transferencia", Err.Number & ":" & Err.Description
End Sub
Public Sub E1_FalhaDevolucao()
    Dim c As Object, n As Long, ep As String, ver As Long
    On Error GoTo Failed
    Set gate = New clsTesteCandidato: Set C1_FornecedorXML = gate
    gate.Etapa = "APPLY": gate.Codigo = 3: Set C1_PoliticaFalha = gate
    SelectE1 "FIC5B02", True: ep = C1_Context(Cad_NFE)("EpisodeID"): ver = CLng(C1_Context(Cad_NFE)("Version"))
    Cad_NFE.CommandButtonCopiarEntradaFilial.value = True
    FillOutputE1 "Devolução", "33002", "800", True
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    DumpE1 "falha_devolucao_antes"
    Cad_NFE_Saida.CommandButton_GravarFilial.value = True
    Set c = OpenE1()
    ResultE1 "E1_D_FALHA_BAIXA", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n And Val10(c, "SELECT PlaceCode FROM Episodes WHERE EpisodeID=?", ep) = "FILIAL" And CLng(Val10(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep)) = ver And CLng(Val10(c, "SELECT Count(*) FROM Projections WHERE DocNumber='33002'")) = 0 And CLng(Val10(c, "SELECT Count(*) FROM Pending WHERE EpisodeID=?", ep)) = 1, "Rollback dos fatos e da baixa; intencao persistente exige conciliacao", C1_UltimoResultado & ";" & C1_Diagnostico(): c.Close
    Set C1_PoliticaFalha = Nothing
    Cad_NFE_Saida.CommandButton_GravarFilial.value = True
    Set c = OpenE1()
    ResultE1 "E1_D_REPETICAO_PENDENTE", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n And InStr(C1_UltimoResultado, "pendente de conciliacao") > 0, "Nao repetir nem prometer recuperacao integral", C1_UltimoResultado: c.Close
    DumpE1 "falha_devolucao_depois": Cad_NFE_Saida.Hide: Exit Sub
Failed:
    Set C1_PoliticaFalha = Nothing
    ResultE1 "E1_FALHA_D_ERRO", False, "Conter falha de baixa transacional", Err.Number & ":" & Err.Description
End Sub
Public Sub E1_FalhasEntrada()
    Dim c As Object, n As Long, point As Variant, op As String, ep As String, label As String
    On Error GoTo Failed
    For Each point In Array(1, 2, 6, 8)
        Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
        Set gate = New clsTesteCandidato: gate.Etapa = "APPLY": gate.Codigo = CLng(point): Set C1_PoliticaFalha = gate
        FillEntryE1 CStr(11900 + CLng(point)), "FX" & CStr(point), "FALHA", "321,45"
        Cad_NFE.CommandButton_Gravar.value = True
        Set c = OpenE1()
        If CLng(point) = 8 Then
            ResultE1 "E1_POS_COMMIT_RECONSTRUIDO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n + 1 And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Commit confirmado e Excel reconstruido apesar de erro apos commit", C1_UltimoResultado & ";" & C1_Diagnostico()
        Else
            op = CStr(C1_Context(Cad_NFE)("OperationID")): ep = CStr(C1_Context(Cad_NFE)("EpisodeID"))
            ResultE1 "E1_TRANSACAO_FALHA_" & CStr(point), CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n And CLng(Val10(c, "SELECT Count(*) FROM Movements WHERE OperationID=?", op)) = 0 And CLng(Val10(c, "SELECT Count(*) FROM Pending WHERE OperationID=?", op)) = 1 And Val10(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", ep) = "NEW", "Fatos revertidos; rascunho NEW e intencao pendente preservados", C1_UltimoResultado & ";" & C1_Diagnostico()
        End If
        c.Close
    Next point
    Set C1_PoliticaFalha = Nothing: DumpE1 "falhas_entrada_depois": Exit Sub
Failed:
    Set C1_PoliticaFalha = Nothing
    ResultE1 "E1_FALHAS_ENTRADA_ERRO", False, "Concluir pontos controlados", Err.Number & ":" & Err.Description
End Sub
Public Sub E1_SessaoReaberta()
    ResultE1 "E1_SESSAO_REABERTA", InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Cache e formulas apos fechar e reabrir o XLSM", C1_Diagnostico()
End Sub

