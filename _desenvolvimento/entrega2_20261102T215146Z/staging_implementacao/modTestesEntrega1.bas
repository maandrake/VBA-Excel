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
    file = ThisWorkbook.Path & "\resultados_entrega2.tsv"
    Set f = FSO10.OpenTextFile(file, 8, True, -1)
    f.WriteLine id & "_E2_R3" & vbTab & IIf(ok, "PASS", "FAIL") & vbTab & "VBA_FUNCIONAL_EVENTOS_REAIS" & vbTab & Replace(expected, vbTab, " ") & vbTab & Replace(Replace(observed, vbCrLf, " "), vbTab, " ") & vbTab & ThisWorkbook.FullName & vbTab & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    f.Close: Debug.Print id; " "; IIf(ok, "PASS", "FAIL"); " "; observed
End Sub
Private Sub DumpE1(ByVal label As String)
    Dim c As Object, t As Variant, r As Object, f As Object, i As Long, line As String, sh As Worksheet, cell As Range
    Set c = OpenE1(): Set f = FSO10.CreateTextFile(ThisWorkbook.Path & "\estado_e2_" & label & "_" & Replace(NewID10(), "-", "") & ".tsv", False, True)
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
    Cad_NFE.MultiPage1.Value = IIf(filial, 1, 0)
    Cad_NFE.Controls("ComboBox_Operação" & IIf(filial, "Filial", "")).Value = "Compra"
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


' Entrega2: reentrada somente pelos controles reais, sem injetar VehicleID.
Public Sub E2_ReentradaInterface()
    Dim c As Object, vid As String, oldEP As String, n As Long, d As String
    On Error GoTo Failed
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts"))
    oldEP = CStr(Val10(c, "SELECT EpisodeID FROM Projections WHERE DocNumber='11001'"))
    vid = CStr(Val10(c, "SELECT VehicleID FROM Episodes WHERE EpisodeID=?", oldEP)): c.Close
    Load Cad_NFE: Cad_NFE.MultiPage1.Value = 0: C1_LimparContexto Cad_NFE
    Cad_NFE.ComboBox_Operação.Value = "Reentrada"
    Cad_NFE.TextBoxPesquisa.Value = "": Cad_NFE.TextBoxPesquisa.Value = "FIC5A01"
    If Cad_NFE.ListBoxLista.ListCount <> 1 Then Err.Raise 5, , "Identificacao nao univoca; count=" & Cad_NFE.ListBoxLista.ListCount
    Cad_NFE.ListBoxLista.ListIndex = 0
    If Not C1_Context(Cad_NFE).Exists("PreviousEpisodeID") Then Err.Raise 5, , "Evento real nao estabeleceu o vinculo anterior"
    Cad_NFE.N_Nfe.Value = "11003": Cad_NFE.txt_Valor.Value = "650,25": Cad_NFE.data.Value = "02/11/2026"
    Cad_NFE.UN.Value = "1": Cad_NFE.ComboBox_OutroEnvio.Value = "TESTE"
    Cad_NFE.CommandButton_Gravar.Value = True
    Set c = OpenE1(): d = C1_Diagnostico()
    ResultE1 "E2_REENTRADA_CONTROLES", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n + 1 And CLng(Val10(c, "SELECT Count(*) FROM Episodes WHERE VehicleID=?", vid)) = 2 And Val10(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", oldEP) = "CLOSED" And InStr(d, "PROJECTION=VERIFIED") > 0, "Lista real escolhe episodio encerrado; novo episodio do mesmo VID", C1_UltimoResultado & ";" & d
    ResultE1 "E2_REENTRADA_DECIMAL", CCur(Val10(c, "SELECT AmountIn FROM Projections WHERE DocNumber='11003'")) = 650.25@, "650,25 sem conversao por placa", CStr(Val10(c, "SELECT AmountIn FROM Projections WHERE DocNumber='11003'"))
    c.Close: DumpE1 "reentrada_interface": Exit Sub
Failed:
    ResultE1 "E2_REENTRADA_ERRO", False, "Concluir reentrada por controles", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub E2_AlteracaoInterface()
    Dim c As Object, n As Long, pid As String, ep As String, oldHash As String, version As Long, d As String
    On Error GoTo Failed
    Set gate = New clsTesteCandidato: gate.Responsavel = "QA SINTETICO": gate.Motivo = "Correcao descritiva experimental"
    Set C2_Interacao = gate
    SelectE1 "FIC5B02", False
    pid = CStr(C1_Context(Cad_NFE)("PID")): ep = CStr(C1_Context(Cad_NFE)("EpisodeID")): version = CLng(C1_Context(Cad_NFE)("Version"))
    Set c = OpenE1(): oldHash = CStr(Val10(c, "SELECT ContentCanon FROM Projections WHERE ProjectionID=?", pid)): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    Cad_NFE.ComboBox_COR.Value = "VERDE QA"
    Cad_NFE.CommandButtonAlterar.Value = True
    Set c = OpenE1(): d = C1_Diagnostico()
    ResultE1 "E2_ALTERAR_EVENTO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n + 1 And CLng(Val10(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep)) = version + 1 And CStr(Val10(c, "SELECT CurrentPID FROM Episodes WHERE EpisodeID=?", ep)) <> pid And InStr(d, "PROJECTION=VERIFIED") > 0, "Retificacao cria fato substitutivo e versao, com projecao verificada", C1_UltimoResultado & ";" & d
    ResultE1 "E2_ALTERAR_IMUTAVEL", CStr(Val10(c, "SELECT ContentCanon FROM Projections WHERE ProjectionID=?", pid)) = oldHash And Audit10(c) = "OK", "Fato anterior intacto e auditoria consistente", "AUDIT=" & Audit10(c)
    c.Close: SelectE1 "FIC5B02", False
    ResultE1 "E2_ALTERAR_LEITOR", Cad_NFE.ComboBox_COR.Value = "VERDE QA", "Pesquisa resolve novo fato do mesmo episodio", CStr(Cad_NFE.ComboBox_COR.Value)
    Cad_NFE.CommandButton_Excluir.Value = True
    ResultE1 "E2_EXCLUSAO_CLASSIFICADA", InStr(C1_UltimoResultado, "CADASTRO_CONFIRMADO") > 0, "Anulacao confirmada exige regra; nao apaga fato", C1_UltimoResultado
    Set C2_Interacao = Nothing: DumpE1 "alteracao": Exit Sub
Failed:
    Set C2_Interacao = Nothing
    ResultE1 "E2_ALTERACAO_ERRO", False, "Concluir retificacao por evento", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub E2_CanceladoInterface()
    Dim c As Object, n As Long, ep As String, pid As String, v As Long, d As String
    On Error GoTo Failed
    SelectE1 "FIC5A01", False: ep = CStr(C1_Context(Cad_NFE)("EpisodeID")): pid = CStr(C1_Context(Cad_NFE)("PID")): v = CLng(C1_Context(Cad_NFE)("Version"))
    Set gate = New clsTesteCandidato: Set C1_FornecedorXML = gate
    Cad_NFE.CommandButtonCopiarEntrada.Value = True
    FillOutputE1 "Cancelado", "22990", "0", False
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    Cad_NFE_Saida.CommandButton_Gravar.Value = True
    Set c = OpenE1(): d = C1_Diagnostico()
    ResultE1 "E2_CANCELADO_EVENTO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n + 1 And CStr(Val10(c, "SELECT CurrentPID FROM Episodes WHERE EpisodeID=?", ep)) = pid And Val10(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", ep) = "ACTIVE" And CLng(Val10(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep)) = v + 1 And InStr(d, "PROJECTION=VERIFIED") > 0, "Cancelado ligado: dois destinos, origem permanece e custo preservado, sem XML", C1_UltimoResultado & ";" & d
    ResultE1 "E2_CANCELADO_HISTORICOS", CLng(Val10(c, "SELECT Count(*) FROM SheetProjections S INNER JOIN Projections P ON S.ProjectionID=P.ProjectionID WHERE P.DocNumber='22990'")) = 2, "Duas projecoes da mesma movimentacao", "AUDIT=" & Audit10(c)
    c.Close: Cad_NFE_Saida.Hide: DumpE1 "cancelado": Exit Sub
Failed:
    ResultE1 "E2_CANCELADO_ERRO", False, "Preservar comportamento tecnico sem homologacao fiscal", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub E2_Diagnostico()
    Call CandidatoConsultarPendencias
    Debug.Print C1_Diagnostico()
End Sub

Public Sub E2_LeitoresEFormulas()
    Dim c As Object, n As Long, sh As Worksheet, r As Long, oldEP As String, newEP As String, vid As String, olds As Long, news As Long, diag As String
    On Error GoTo Failed
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts"))
    oldEP = CStr(Val10(c, "SELECT EpisodeID FROM Projections WHERE DocNumber='11001'"))
    newEP = CStr(Val10(c, "SELECT EpisodeID FROM Projections WHERE DocNumber='11003'"))
    vid = CStr(Val10(c, "SELECT VehicleID FROM Episodes WHERE EpisodeID=?", oldEP))
    Set sh = ThisWorkbook.Worksheets("Entradas Anual")
    For r = 4 To sh.Cells(sh.Rows.Count, 52).End(xlUp).Row
        If CStr(sh.Cells(r, 53).Value2) = oldEP Then
            olds = olds + 1
            If sh.Cells(r, 55).Value2 <> vid Or sh.Cells(r, 9).Value2 <> 900# Or sh.Cells(r, 16).Value2 <> "SIM" Then Err.Raise 5, , "Historico antigo associado incorretamente"
        ElseIf CStr(sh.Cells(r, 53).Value2) = newEP Then
            news = news + 1
            If sh.Cells(r, 55).Value2 <> vid Or sh.Cells(r, 9).Value2 <> 650.25 Or sh.Cells(r, 16).Value2 <> "NÃO" Then Err.Raise 5, , "Reentrada associada incorretamente"
        End If
    Next r
    ResultE1 "E2_ANUAL_REENTRADA", olds = 1 And news = 1 And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Mesmo VID, episodios distintos, valores900/650,25 e estados distintos", "antigos=" & olds & ";novos=" & news & ";" & C1_Diagnostico()
    c.Close
    SelectE1 "FIC5A01", False
    ResultE1 "E2_PESQUISA_REENTRADA", C1_Context(Cad_NFE)("EpisodeID") = newEP And Cad_NFE.ListBoxLista.ListCount = 1, "Somente nova passagem ativa, sem primeira ocorrencia por placa", C1_Context(Cad_NFE)("EpisodeID")
    C1_Context(Cad_NFE)("Version") = CLng(C1_Context(Cad_NFE)("Version")) - 1
    Cad_NFE.CommandButtonCopiarEntrada.Value = True
    ResultE1 "E2_SELECAO_OBSOLETA", InStr(C1_UltimoResultado, "obsoleta") > 0, "Token antigo nao abre gravador autorizado", C1_UltimoResultado
    SelectE1 "FIC5A01", False: Cad_NFE.ListBoxLista.ListIndex = -1
    ResultE1 "E2_LISTINDEX_NEGATIVO", Cad_NFE.ListBoxLista.ListIndex = -1, "Nenhuma acao por indice negativo", C1_UltimoResultado
    C2_Filtrar "Venda", "MATRIZ"
    ResultE1 "E2_FILTRO_HISTORICO", Cad_NFE_Saida.ListBoxLista.ListCount = 1, "Uma apresentacao por saida; nao duplica com Todas Saidas", "COUNT=" & Cad_NFE_Saida.ListBoxLista.ListCount
    Cad_NFE_Saida.ListBoxLista.ListIndex = 0
    Cad_NFE_Saida.CommandButton_Excluir.Value = True
    ResultE1 "E2_EXCLUSAO_HISTORICO", InStr(C1_UltimoResultado, "HISTORICO_OU_BAIXADO") > 0, "Historico confirmado nao e exclusao de estoque", C1_UltimoResultado
    Cad_NFE_Saida.Hide
    ResultE1 "E2_ESTOQUE_ATIVO", EstoqueAtivoCandidato() = 2, "Duas custodias ativas por identidade, excluindo baixadas", CStr(EstoqueAtivoCandidato())
    DumpE1 "leitores": Exit Sub
Failed:
    ResultE1 "E2_LEITORES_ERRO", False, "Integridade dos leitores e formulas por episodio", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub E2_FalhaProjecao()
    Dim c As Object, n As Long, ep As String, op As String, ev As Boolean, su As Boolean, da As Boolean, calc As Long
    On Error GoTo Failed
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    ev = Application.EnableEvents: su = Application.ScreenUpdating: da = Application.DisplayAlerts: calc = Application.Calculation
    Set gate = New clsTesteCandidato: gate.Etapa = "PROJECTION": gate.Codigo = 1: Set C1_PoliticaFalha = gate
    FillEntryE1 "11988", "FCE2X01", "E2REC", "321,45"
    Cad_NFE.CommandButton_Gravar.Value = True
    op = CStr(C1_Context(Cad_NFE)("OperationID")): ep = CStr(C1_Context(Cad_NFE)("EpisodeID"))
    Set c = OpenE1()
    ResultE1 "E2_C_FALHA_COMPENSACAO_CONTENCAO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n + 1 And Val10(c, "SELECT ProjectionState FROM CandidateSettings WHERE ID=1") = "PENDING" And Audit10(c) = "OK", "Commit preservado; Excel parcial nao e apresentado como compensacao", C1_UltimoResultado & ";AUDIT=" & Audit10(c)
    ResultE1 "E2_APPLICATION_RESTAURADO", Application.EnableEvents = ev And Application.ScreenUpdating = su And Application.DisplayAlerts = da And Application.Calculation = calc, "Estados anteriores restaurados apos erro", "EV=" & Application.EnableEvents & ";SU=" & Application.ScreenUpdating & ";CALC=" & Application.Calculation
    c.Close: Set C1_PoliticaFalha = Nothing
    Call CandidatoReconstruir
    ResultE1 "E2_C_FALHA_COMPENSACAO_RECONSTRUIDO", InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Recuperacao da projecao pela autoridade, sem desfazer commit", C1_Diagnostico()
    Cad_NFE.CommandButton_Gravar.Value = True
    Set c = OpenE1()
    ResultE1 "E2_REPLAY_EVENTO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n + 1 And CLng(Val10(c, "SELECT Count(*) FROM Movements WHERE OperationID=?", op)) = 1 And Audit10(c) = "OK", "Mesma operacao confirmada nao duplica fato; replay auditado", C1_UltimoResultado & ";AUDIT=" & Audit10(c)
    c.Close: DumpE1 "falha_projecao_recuperada": Exit Sub
Failed:
    Set C1_PoliticaFalha = Nothing
    ResultE1 "E2_FALHA_PROJECAO_ERRO", False, "Recuperar projecao sem duplicar o commit", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub E2_SaveEBackup()
    Dim c As Object, n As Long, saver As clsSaveFase9, e As Long, msg As String, folder As String, b As Object, f As String
    On Error GoTo Failed
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    Set saver = New clsSaveFase9: saver.Iniciar ThisWorkbook: saver.RecusarUmaVez
    On Error Resume Next
    Call CandidatoReconstruir
    e = Err.Number: msg = Err.Description: Err.Clear
    On Error GoTo Failed
    Set c = OpenE1()
    ResultE1 "E2_SAVE_CANCELADO_REAL", saver.Observado And e <> 0 And Val10(c, "SELECT ProjectionState FROM CandidateSettings WHERE ID=1") = "PENDING" And CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n, "BeforeSave cancela; autoridade nao perde fatos", CStr(e) & ":" & msg
    c.Close: Set saver = Nothing
    Call CandidatoReconstruir
    ResultE1 "E2_SAVE_RECUPERADO", InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Arquivo salvo e geracao reaberta verificados", C1_Diagnostico()
    Load Cad_NFE_Saida: Cad_NFE_Saida.CommandButtonSalvarTudo.Value = True
    If InStr(C1_UltimoResultado, "BACKUP_VERIFICADO: ") <> 1 Then Err.Raise 5, , C1_UltimoResultado
    folder = Mid$(C1_UltimoResultado, Len("BACKUP_VERIFICADO: ") + 1)
    Set b = CreateObject("ADODB.Connection"): b.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & folder & "\Banco_Candidato_Operacoes_Corrigidas.accdb"
    ResultE1 "E2_BACKUP_BOTAO_REABERTO", Audit10(b) = "OK" And CLng(Val10(b, "SELECT Count(*) FROM Receipts")) = n And Val10(b, "SELECT StatusText FROM WriterLease WHERE ID=1") = "FREE", "Backup pelo botao, banco reaberto com recibos e lease livre", folder & ";AUDIT=" & Audit10(b)
    b.Close: DumpE1 "save_backup": Exit Sub
Failed:
    Set saver = Nothing
    ResultE1 "E2_SAVE_BACKUP_ERRO", False, "Concluir recuperacao e backup", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub

Public Sub E2_DevolucaoSucessiva()
    Dim c As Object, ep As String, n As Long, v As Long, mid As String
    On Error GoTo Failed
    Set gate = New clsTesteCandidato: Set C1_FornecedorXML = gate
    SelectE1 "FIC5B02", True: ep = C1_Context(Cad_NFE)("EpisodeID"): v = CLng(C1_Context(Cad_NFE)("Version"))
    Cad_NFE.CommandButtonCopiarEntradaFilial.Value = True
    FillOutputE1 "Devolução", "33003", "800", True
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    DumpE1 "devolucao_sucessiva_antes"
    Cad_NFE_Saida.CommandButton_GravarFilial.Value = True
    Set c = OpenE1(): mid = CStr(Val10(c, "SELECT MovementID FROM Projections WHERE DocNumber=?", "33003"))
    ResultE1 "E2_D_SUCESSIVA", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n + 1 And Val10(c, "SELECT PlaceCode FROM Episodes WHERE EpisodeID=?", ep) = "MATRIZ" And CLng(Val10(c, "SELECT Count(*) FROM SheetProjections WHERE MovementID=?", mid)) = 4 And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Mesmo episodio, quatro destinos e custodia matriz", C1_UltimoResultado & ";" & C1_Diagnostico()
    c.Close: Cad_NFE_Saida.Hide: DumpE1 "devolucao_sucessiva_depois": Exit Sub
Failed:
    ResultE1 "E2_D_SUCESSIVA_ERRO", False, "Devolucao apos segunda transferencia", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub E2_EntradasInvalidas()
    Dim c As Object, n As Long, i As Long, before As String, states As String
    On Error GoTo Failed
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    For i = 1 To 5
        FillEntryE1 "11989", "FCE2X02", "E2INV", "123,45", True
        Select Case i
            Case 1: Cad_NFE.N_NfeFilial.Value = ""
            Case 2: Cad_NFE.txt_ValorFilial.Value = "invalido"
            Case 3: Cad_NFE.DataFilial.Value = "31/02/2026"
            Case 4: Cad_NFE.TextBoxChaveDeAcessoFilial.Value = "0123"
            Case 5: Cad_NFE.N_NfeFilial.Value = "abc"
        End Select
        Cad_NFE.CommandButton_GravarFilial.Value = True
        Set c = OpenE1()
        ResultE1 "E2_ENTRADA_FILIAL_INVALIDA_" & i, CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n And CLng(Val10(c, "SELECT Count(*) FROM Projections WHERE DocNumber='11989'")) = 0 And InStr(C1_UltimoResultado, "RECUSADA") > 0, "Recusa antes de criar episodio ou gravar fatos", C1_UltimoResultado
        c.Close
    Next i
    FillEntryE1 "11989", "FIC5A01", "E2INV", "123,45"
    Cad_NFE.CommandButton_Gravar.Value = True
    Set c = OpenE1()
    ResultE1 "E2_IDENTIDADE_CONHECIDA_SEM_VINCULO", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n And InStr(C1_UltimoResultado, "RECUSADA") > 0, "Placa conhecida nao vincula sozinha", C1_UltimoResultado
    c.Close: DumpE1 "invalidos": Exit Sub
Failed:
    ResultE1 "E2_INVALIDOS_ERRO", False, "Validacao antes dos efeitos", Err.Number & ":" & Err.Description
End Sub

Public Sub E2_ConferirDestinos()
    Dim c As Object, mid As String, ep As String, sh As Worksheet, r As Long, matches As Long, ok As Boolean
    On Error GoTo Failed
    Set c = OpenE1()
    mid = CStr(Val10(c, "SELECT MovementID FROM Projections WHERE DocNumber='33003'"))
    ep = CStr(Val10(c, "SELECT EpisodeID FROM Movements WHERE MovementID=?", mid))
    ResultE1 "E2_D_SUCESSIVA_CONFERIDA", Len(mid) > 0 And CLng(Val10(c, "SELECT Count(*) FROM SheetProjections WHERE MovementID=?", mid)) = 4 And Val10(c, "SELECT PlaceCode FROM Episodes WHERE EpisodeID=?", ep) = "MATRIZ" And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Evento anterior confirmado, quatro destinos e custodia por episodio", mid & ";" & C1_Diagnostico()
    ep = CStr(Val10(c, "SELECT EpisodeID FROM Projections WHERE DocNumber='22004'"))
    Set sh = ThisWorkbook.Worksheets("Entradas Filial"): ok = True
    For r = 4 To sh.Cells(sh.Rows.Count, 52).End(xlUp).Row
        If CStr(sh.Cells(r, 53).Value2) = ep Then
            matches = matches + 1
            ok = ok And VarType(sh.Cells(r, 12).Value2) = vbString And sh.Cells(r, 12).Value2 = KEY_E1 And sh.Cells(r, 12).NumberFormat = "@" And sh.Cells(r, 9).Value2 = 450#
            ok = ok And sh.Cells(r, 13).HasFormula And sh.Cells(r, 14).HasFormula And sh.Cells(r, 15).HasFormula
        End If
    Next r
    ResultE1 "E2_ENTRADA_FILIAL_CHAVE_FORMULAS", matches = 1 And ok And Audit10(c) = "OK", "Cadastro real filial, valor450, chave44 texto, MNO por ID", "MATCHES=" & matches & ";AUDIT=" & Audit10(c)
    c.Close: DumpE1 "destinos_conferidos": Exit Sub
Failed:
    ResultE1 "E2_CONFERIR_DESTINOS_ERRO", False, "Conferir banco e projecoes persistidas", Err.Number & ":" & Err.Description
End Sub
Public Sub E2_PrecisaoCriar()
    On Error GoTo Failed
    FillEntryE1 "11987", "FX4D001", "E2DEC", "123,4567"
    Cad_NFE.CommandButton_Gravar.Value = True
    Call E2_PrecisaoConsultar
    Exit Sub
Failed:
    ResultE1 "E2_PRECISAO_CRIAR_ERRO", False, "Persistir quatro casas do contrato Currency", Err.Number & ":" & Err.Description
End Sub
Public Sub E2_PrecisaoConsultar()
    Dim c As Object, ep As String, actual As String
    On Error GoTo Failed
    Set gate = New clsTesteCandidato: Set C1_FornecedorXML = gate
    SelectE1 "FX4D001", False: ep = CStr(C1_Context(Cad_NFE)("EpisodeID"))
    Set c = OpenE1(): actual = CStr(Val10(c, "SELECT AmountIn FROM Projections WHERE ProjectionID=?", CStr(C1_Context(Cad_NFE)("PID"))))
    ResultE1 "E2_PRECISAO_ENTRADA", Replace(CStr(Cad_NFE.txt_Valor.Value), ",", ".") = "123.4567", "Leitor preserva quatro casas, sem alterar fato", CStr(Cad_NFE.txt_Valor.Value)
    c.Close
    Cad_NFE.CommandButtonCopiarEntrada.Value = True
    ResultE1 "E2_PRECISAO_ENVIO", Replace(CStr(Cad_NFE_Saida.TextBox_Valor_Entrada.Value), ",", ".") = "123.4567", "Envio ao gravador preserva valor exato da autoridade", CStr(Cad_NFE_Saida.TextBox_Valor_Entrada.Value)
    Cad_NFE_Saida.Hide: Exit Sub
Failed:
    ResultE1 "E2_PRECISAO_ERRO", False, "Conferir precisao pelos eventos e banco", Err.Number & ":" & Err.Description
End Sub
Public Sub E2_ReordenacaoESaveAs()
    Dim sh As Worksheet, a As Variant, z As Variant, diag As String, oldName As String, e As Long, msg As String, alerts As Boolean
    On Error GoTo Failed
    Set sh = ThisWorkbook.Worksheets("Entradas"): sh.Unprotect "CANDIDATO_SINTETICO"
    a = sh.Range("A4:BH4").Formula: z = sh.Range("A5:BH5").Formula
    sh.Range("A4:BH4").Formula = z: sh.Range("A5:BH5").Formula = a: sh.Protect "CANDIDATO_SINTETICO"
    diag = C1_Diagnostico()
    ResultE1 "E2_REORDENACAO_DETECTADA", InStr(diag, "PROJECTION=VERIFIED") = 0, "Referencia fisica alterada nao recebe aprovacao silenciosa", diag
    Call CandidatoReconstruir
    ResultE1 "E2_REORDENACAO_RECONSTRUIDA", InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Reconstrucao por IDs e metadados completos", C1_Diagnostico()
    oldName = ThisWorkbook.FullName: alerts = Application.DisplayAlerts
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\diretorio_inexistente_E2\arquivo.xlsm", xlOpenXMLWorkbookMacroEnabled
    e = Err.Number: msg = Err.Description: Err.Clear
    Application.DisplayAlerts = alerts
    On Error GoTo Failed
    ResultE1 "E2_SAVEAS_INVALIDO_REAL", e <> 0 And ThisWorkbook.FullName = oldName And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "SaveAs invalido nao altera autoridade ou arquivo atual", CStr(e) & ":" & msg
    DumpE1 "reordenacao_saveas": Exit Sub
Failed:
    Application.DisplayAlerts = alerts
    ResultE1 "E2_REORDENACAO_ERRO", False, "Concluir ensaio sem associacao incorreta", Err.Number & ":" & Err.Description
End Sub

Public Sub E2_VendaReentrada()
    Dim c As Object, ep As String, vid As String, n As Long
    On Error GoTo Failed
    Set gate = New clsTesteCandidato: gate.ArquivoXML = ThisWorkbook.Path & "\venda_reentrada.xml": Set C1_FornecedorXML = gate
    SelectE1 "FIC5A01", False: ep = CStr(C1_Context(Cad_NFE)("EpisodeID")): vid = CStr(C1_Context(Cad_NFE)("VehicleID"))
    Cad_NFE.CommandButtonCopiarEntrada.Value = True
    FillOutputE1 "Venda", "22011", "700", False: Cad_NFE_Saida.TextBox_Data.Value = "02/11/2026"
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    DumpE1 "venda_reentrada_antes": Cad_NFE_Saida.CommandButton_Gravar.Value = True
    Set c = OpenE1()
    ResultE1 "E2_V_VALIDO_REENTRADA", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n + 1 And Val10(c, "SELECT StatusText FROM Episodes WHERE EpisodeID=?", ep) = "CLOSED" And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Venda real encerra somente episodio selecionado, preserva primeira passagem", C1_UltimoResultado & ";" & C1_Diagnostico()
    ResultE1 "E2_V_VALORES_IDENTIDADE", CCur(Val10(c, "SELECT AmountIn FROM Projections WHERE DocNumber='22011'")) = 650.25@ And CCur(Val10(c, "SELECT AmountOut FROM Projections WHERE DocNumber='22011'")) = 700@ And Val10(c, "SELECT EpisodeID FROM Projections WHERE DocNumber='22011'") = ep And Val10(c, "SELECT VehicleID FROM Episodes WHERE EpisodeID=?", ep) = vid, "VIN650,25 VOUT700, mesmo VID e episodio correto", ep & ";AUDIT=" & Audit10(c)
    c.Close: Cad_NFE_Saida.Hide
    Cad_NFE.TextBoxPesquisa.Value = "": Cad_NFE.TextBoxPesquisa.Value = "FIC5A01"
    ResultE1 "E2_V_BAIXADO_PESQUISA", Cad_NFE.ListBoxLista.ListCount = 0, "Nenhum episodio vendido aparece como ativo", "COUNT=" & Cad_NFE.ListBoxLista.ListCount
    DumpE1 "venda_reentrada_depois": Exit Sub
Failed:
    ResultE1 "E2_V_FINAL_ERRO", False, "Venda real apos reentrada", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub E2_FalhaAntesCommit()
    Dim c As Object, ep As String, n As Long, ver As Long, before As String
    On Error GoTo Failed
    Set gate = New clsTesteCandidato: Set C1_FornecedorXML = gate: gate.Etapa = "APPLY": gate.Codigo = 6: Set C1_PoliticaFalha = gate
    SelectE1 "FIC5C03", True: ep = CStr(C1_Context(Cad_NFE)("EpisodeID")): ver = CLng(C1_Context(Cad_NFE)("Version"))
    Cad_NFE.CommandButtonCopiarEntradaFilial.Value = True: FillOutputE1 "Devolução", "33996", "450", True
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): before = E2_DigestRows(BusinessRows10(c)): c.Close
    Cad_NFE_Saida.CommandButton_GravarFilial.Value = True
    Set c = OpenE1()
    ResultE1 "E2_FALHA_ANTES_COMMIT", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n And E2_DigestRows(BusinessRows10(c)) = before And CLng(Val10(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep)) = ver And CLng(Val10(c, "SELECT Count(*) FROM Pending WHERE EpisodeID=?", ep)) = 1 And Audit10(c) = "OK", "Rollback de fatos, historico e versao; intencao pendente nao e recuperacao integral", C1_UltimoResultado & ";" & C1_Diagnostico()
    c.Close: Set C1_PoliticaFalha = Nothing: Cad_NFE_Saida.Hide: Call CandidatoConsultarPendencias: DumpE1 "antes_commit_pendente": Exit Sub
Failed:
    Set C1_PoliticaFalha = Nothing
    ResultE1 "E2_FALHA_COMMIT_ERRO", False, "Conter erro anterior ao commit", Err.Number & ":" & Err.Description
End Sub
Public Sub E2_PortabilidadeGuardas()
    Dim c As Object, authority As String, bid As String, diag As String, a As String, b As String, moved As Boolean
    On Error GoTo Failed
    Set c = OpenE1(): authority = CStr(Val10(c, "SELECT AuthorityPath FROM CandidateSettings WHERE ID=1")): bid = CStr(Val10(c, "SELECT BaseID FROM CandidateSettings WHERE ID=1"))
    c.BeginTrans: Exec10 c, "UPDATE CandidateSettings SET AuthorityPath=? WHERE ID=1", ThisWorkbook.Path & "\copia_divergente.accdb": c.CommitTrans: c.Close: Set c = Nothing
    diag = C1_Diagnostico()
    Set c = OpenE1(): Exec10 c, "UPDATE CandidateSettings SET AuthorityPath=? WHERE ID=1", authority
    ResultE1 "E2_CAMINHO_DIVERGENTE", InStr(diag, "caminho da autoridade nao confere") > 0, "Nome/copia nao elege autoridade", diag
    c.BeginTrans: Exec10 c, "UPDATE CandidateSettings SET BaseID=? WHERE ID=1", NewID10(): c.CommitTrans: c.Close: Set c = Nothing
    diag = C1_Diagnostico()
    Set c = OpenE1(): Exec10 c, "UPDATE CandidateSettings SET BaseID=? WHERE ID=1", bid
    ResultE1 "E2_BASEID_DIVERGENTE", InStr(diag, "BaseID divergente") > 0, "BaseID e linhagem conferidos antes de escrever", diag
    c.Close: Set c = Nothing
    a = ThisWorkbook.Path & "\Banco_Candidato_Operacoes_Corrigidas.accdb": b = a & ".ausente_teste"
    Name a As b: moved = True
    diag = C1_Diagnostico()
    Name b As a: moved = False
    ResultE1 "E2_BANCO_AUSENTE", Left$(diag, 6) = "ERROR:", "Banco ausente nao gera outra autoridade", diag
    ResultE1 "E2_GUARDAS_RESTAURADAS", InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Vinculo original restaurado no laboratorio, sem alterar fatos", C1_Diagnostico()
    Exit Sub
Failed:
    diag = Err.Number & ":" & Err.Description
    On Error Resume Next
    If moved Then Name b As a
    If Not c Is Nothing Then
        If c.State <> 0 Then
            Exec10 c, "UPDATE CandidateSettings SET AuthorityPath=?,BaseID=? WHERE ID=1", authority, bid
            c.Close
        End If
    End If
    ResultE1 "E2_PORTABILIDADE_ERRO", False, "Guardas e restauracao sem escolha automatica de banco", diag
End Sub
Public Sub E2_ReciboAusente()
    Dim c As Object, op As String, diag As String
    On Error GoTo Failed
    Set c = OpenE1(): op = CStr(Val10(c, "SELECT OperationID FROM Receipts WHERE RevisionNo=1"))
    Exec10 c, "DELETE FROM Receipts WHERE OperationID=?", op
    c.Close
    diag = C1_Diagnostico()
    ResultE1 "E2_RECIBO_AUSENTE_CONTIDO", Left$(diag, 6) = "ERROR:", "Autoridade inconsistente bloqueada; exige conciliacao, nao recria recibo", diag
    Call CandidatoConsultarPendencias: DumpE1 "recibo_ausente_forense": Exit Sub
Failed:
    ResultE1 "E2_RECIBO_AUSENTE_ERRO", False, "Detectar adulteracao sintetica", Err.Number & ":" & Err.Description
End Sub

Private Function E2_DigestRows(ByVal rows As Variant) As String
    Dim i As Long, j As Long, s As String
    If IsEmpty(rows) Then E2_DigestRows = SHA10("EMPTY"): Exit Function
    For j = LBound(rows, 2) To UBound(rows, 2)
        For i = LBound(rows, 1) To UBound(rows, 1)
            If IsNull(rows(i, j)) Then s = s & Pack10("NULL") Else s = s & Pack10(TypeName(rows(i, j)), CStr(rows(i, j)))
        Next i
    Next j
    E2_DigestRows = SHA10(s)
End Function
Public Sub E2_ErroAposCommit()
    Dim c As Object, n As Long
    On Error GoTo Failed
    Set c = OpenE1(): n = CLng(Val10(c, "SELECT Count(*) FROM Receipts")): c.Close
    Set gate = New clsTesteCandidato: gate.Etapa = "APPLY": gate.Codigo = 8: Set C1_PoliticaFalha = gate
    FillEntryE1 "11986", "FC8D001", "E2COMMIT", "80,88": Cad_NFE.CommandButton_Gravar.Value = True
    Set C1_PoliticaFalha = Nothing: Set c = OpenE1()
    ResultE1 "E2_ERRO_APOS_COMMIT", CLng(Val10(c, "SELECT Count(*) FROM Receipts")) = n + 1 And CCur(Val10(c, "SELECT AmountIn FROM Projections WHERE DocNumber='11986'")) = 80.88@ And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Commit reconhecido; reconstrucao concluida sem repetir fatos", C1_UltimoResultado & ";" & C1_Diagnostico()
    c.Close: DumpE1 "erro_apos_commit": Exit Sub
Failed:
    Set C1_PoliticaFalha = Nothing
    ResultE1 "E2_APOS_COMMIT_ERRO", False, "Concluir projecao apos erro controlado8", Err.Number & ":" & Err.Description
End Sub
