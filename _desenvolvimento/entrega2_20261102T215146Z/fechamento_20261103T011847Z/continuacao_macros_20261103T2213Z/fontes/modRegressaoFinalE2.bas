Attribute VB_Name = "modRegressaoFinalE2"
Option Explicit
Public CF_Operation As String, CF_Plan As String
Private Const KEY_CF As String = "01234567890123456789012345678901234567890123"
Private Sub CR(ByVal id As String, ByVal ok As Boolean, ByVal expected As String, ByVal observed As String, Optional ByVal modality As String = "VBA_EVENTO_REAL_COM_VERIFICACAO_ACE")
    Dim f As Object
    Set f = FSO10.OpenTextFile(ThisWorkbook.Path & "\resultados_final_e2.tsv", 8, True, -1)
    f.WriteLine id & vbTab & IIf(ok, "PASS", "FAIL") & vbTab & modality & vbTab & expected & vbTab & Replace(Replace(observed, vbCrLf, " "), vbTab, " ") & vbTab & ThisWorkbook.FullName & vbTab & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    f.Close: Debug.Print id; " "; IIf(ok, "PASS", "FAIL"); " "; observed
End Sub
Private Sub CS(ByVal plate As String, ByVal filial As Boolean)
    Dim lst As Object
    Load Cad_NFE: Cad_NFE.MultiPage1.Value = IIf(filial, 1, 0)
    Cad_NFE.Controls("ComboBox_Operação" & IIf(filial, "Filial", "")).Value = "Compra"
    If filial Then
        Cad_NFE.TextBoxPesquisaFilial.Value = "": Cad_NFE.TextBoxPesquisaFilial.Value = plate: Set lst = Cad_NFE.ListBoxLista2
    Else
        Cad_NFE.TextBoxPesquisa.Value = "": Cad_NFE.TextBoxPesquisa.Value = plate: Set lst = Cad_NFE.ListBoxLista
    End If
    If lst.ListCount <> 1 Then Err.Raise 5, , "Selecao nao univoca: " & lst.ListCount
    lst.ListIndex = 0
    If Not C1_Context(Cad_NFE).Exists("ViewID") Then Err.Raise 5, , "Selecao sem identidade"
End Sub
Private Sub CO(ByVal nf As String)
    With Cad_NFE_Saida
        .ComboBox_OperaçãoFilial.Value = "Devolução"
        .TextBox_N_SaidaFilial.Value = nf: .TextBox_Valor_SaidaFilial.Value = "450": .TextBox_DataFilial.Value = "03/11/2026"
        .ComboBox_PorFilial.Value = "12": .TextBox_ICMSFilial.Value = "0": .TextBox_PISFilial.Value = "0": .TextBox_COFINSFilial.Value = "0": .TextBoxChaveDeAcessoFilial.Value = KEY_CF
    End With
End Sub
Public Sub CF_PlanoObsoleto()
    Dim c As Object, g As clsTesteCandidato, ep As String, before As String, n As Long, views As Long, ver As Long, rev As Long, token As String, st As String, pid As String, canon As String
    On Error GoTo Bad
    F2_Snapshot "CF_baixa_antes"
    CS "FIC5C03", True: ep = C1_Context(Cad_NFE)("EpisodeID"): ver = C1_Context(Cad_NFE)("Version")
    Cad_NFE.CommandButtonCopiarEntradaFilial.Value = True: CO "34021"
    Set c = C3_Conectar(): before = Source10(c, ep): n = Val10(c, "SELECT Count(*) FROM Receipts"): views = Val10(c, "SELECT Count(*) FROM SheetProjections"): c.Close
    Set g = New clsTesteCandidato: g.Etapa = "APPLY": g.Codigo = 3: Set C1_PoliticaFalha = g
    Cad_NFE_Saida.CommandButton_GravarFilial.Value = True: Set C1_PoliticaFalha = Nothing
    Set c = C3_Conectar(): CF_Operation = Val10(c, "SELECT OperationID FROM Pending WHERE EpisodeID=?", ep)
    CR "CF_FALHA_BAIXA_LOGICA", Val10(c, "SELECT Count(*) FROM Receipts") = n And Val10(c, "SELECT Count(*) FROM SheetProjections") = views And Source10(c, ep) = before And Val10(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep) = ver And Val10(c, "SELECT Count(*) FROM Pending WHERE OperationID=?", CF_Operation) = 1 And Audit10(c) = "OK", "Rollback dos efeitos, origem intacta; intencao permanece PENDENTE_CONCILIACAO", C1_UltimoResultado
    c.Close: CF_Plan = C3_Analisar(CF_Operation): token = C3_TokenConfirmacao(CF_Plan)
    If Not IDOK10(CF_Plan) Then Err.Raise 5, , CF_Plan
    CS "F2F0002", True: pid = C1_Context(Cad_NFE)("PID")
    Set c = C3_Conectar(): canon = Val10(c, "SELECT ContentCanon FROM Projections WHERE ProjectionID=?", pid): rev = Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1"): c.Close
    Set g = New clsTesteCandidato: g.Responsavel = "QA SINTETICO CF": g.Motivo = "Teste de obsolescencia e fato substitutivo"
    Set C2_Interacao = g: Cad_NFE.ComboBox_CORFilial.Value = "AMARELO QA": Cad_NFE.CommandButtonAlterar.Value = True: Set C2_Interacao = Nothing
    Set c = C3_Conectar()
    CR "CF_ADJUST_DISCRIMINANTE", Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1") = rev + 1 And Val10(c, "SELECT ContentCanon FROM Projections WHERE ProjectionID=?", pid) = canon And Audit10(c) = "OK", "Cor distinta em novo fato; original imutavel; operacao independente", C1_UltimoResultado
    c.Close
    st = C3_Resolver(CF_Plan, "QA SINTETICO CF", "Plano anterior a mudanca de revisao", token)
    Set c = C3_Conectar()
    CR "CF_PLANO_OBSOLETO_RECUSADO", InStr(st, "obsoleto") > 0 And Val10(c, "SELECT Count(*) FROM Pending WHERE OperationID=?", CF_Operation) = 1 And Val10(c, "SELECT Count(*) FROM ReconciliationReceipts WHERE OperationID=?", CF_Operation) = 0, "Versao alterada entre analise e confirmacao; nao liberar intencao", st
    c.Close: F2_Snapshot "CF_plano_obsoleto": Debug.Print "CF OperationID=" & CF_Operation: Exit Sub
Bad:
    Set C1_PoliticaFalha = Nothing: Set C2_Interacao = Nothing
    CR "CF_PLANO_ERRO", False, "Concluir teste com estado diagnosticado", Err.Number & ":" & Err.Description & ";" & C1_UltimoResultado
End Sub
Public Sub CF_ConferirConciliacao()
    Dim c As Object, n As Long, st As String
    On Error GoTo Bad
    Set c = C3_Conectar()
    CR "CF_CONCILIACAO_INTERFACE", Val10(c, "SELECT Count(*) FROM ReconciliationReceipts WHERE OperationID=?", CF_Operation) = 1 And Val10(c, "SELECT StatusText FROM Operations WHERE OperationID=?", CF_Operation) = "ABORTED" And Val10(c, "SELECT Count(*) FROM Pending WHERE OperationID=?", CF_Operation) = 0 And Audit10(c) = "OK", "Decisao explicita na interface; recibo duravel; fatos intactos", C1_Diagnostico()
    n = Val10(c, "SELECT Count(*) FROM Receipts"): c.Close
    CS "FIC5C03", True: Cad_NFE.CommandButtonCopiarEntradaFilial.Value = True: CO "34021"
    C1_Context(Cad_NFE_Saida)("OperationID") = CF_Operation
    Cad_NFE_Saida.CommandButton_GravarFilial.Value = True
    Set c = C3_Conectar()
    CR "CF_REPLAY_ABORTADO_RECUSADO", Val10(c, "SELECT Count(*) FROM Receipts") = n And InStr(C1_UltimoResultado, "ABORTADA") > 0, "OperationID abortado nunca reabre nem duplica fatos", C1_UltimoResultado
    c.Close: Cad_NFE_Saida.Hide: F2_Snapshot "CF_conciliacao_interface": Exit Sub
Bad: CR "CF_CONCILIACAO_ERRO", False, "Conferir resolucao sem repeticao insegura", Err.Number & ":" & Err.Description
End Sub
Public Sub CF_RepetirAposConciliacao()
    Dim c As Object, ep As String, n As Long
    On Error GoTo Bad
    Set c = C3_Conectar()
    If Val10(c, "SELECT StatusText FROM Operations WHERE OperationID=?", CF_Operation) <> "ABORTED" Or Val10(c, "SELECT Count(*) FROM Pending") <> 0 Then Err.Raise 5, , "Estado nao conciliado; proibido tentar novamente"
    n = Val10(c, "SELECT Count(*) FROM Receipts"): c.Close
    CS "FIC5C03", True: ep = C1_Context(Cad_NFE)("EpisodeID")
    Cad_NFE.CommandButtonCopiarEntradaFilial.Value = True: CO "34021"
    Cad_NFE_Saida.CommandButton_GravarFilial.Value = True
    Set c = C3_Conectar()
    CR "CF_NOVA_TENTATIVA_CONCILIADA", Val10(c, "SELECT Count(*) FROM Receipts") = n + 1 And Val10(c, "SELECT PlaceCode FROM Episodes WHERE EpisodeID=?", ep) = "MATRIZ" And Val10(c, "SELECT Count(*) FROM Projections WHERE DocNumber='34021'") = 1 And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Nova operacao somente depois de conciliacao comprovada; quatro destinos", C1_UltimoResultado & ";" & C1_Diagnostico()
    c.Close: Cad_NFE_Saida.Hide: F2_Snapshot "CF_retry_conciliado": Exit Sub
Bad: CR "CF_RETRY_ERRO", False, "Tentativa somente em estado conciliado", Err.Number & ":" & Err.Description
End Sub
Public Sub CF_LeitoresMetadados()
    Dim c As Object, sh As Worksheet, sn As Variant, row As Long, first As Long, pid As String, ep As String, vid As String, key As String, ok As Boolean, count As Long, n As Long
    On Error GoTo Bad
    Set c = C3_Conectar(): ok = True
    For Each sn In Array("Entradas", "Entradas Filial", "Entradas Anual", "Entradas Anual Filial", "Saidas", "Todas Saidas", "Saidas Filial", "Todas Saidas Filial")
        Set sh = ThisWorkbook.Worksheets(CStr(sn)): first = IIf(Left$(CStr(sn), 8) = "Entradas", 4, 2)
        For row = first To sh.Cells(sh.Rows.Count, 52).End(xlUp).Row
            If Len(sh.Cells(row, 52).Value2) > 0 Then
                count = count + 1: pid = CStr(sh.Cells(row, 58).Value2): ep = CStr(sh.Cells(row, 53).Value2): vid = CStr(sh.Cells(row, 55).Value2): key = CStr(sh.Cells(row, 60).Value2)
                ok = ok And Val10(c, "SELECT EpisodeID FROM Projections WHERE ProjectionID=?", pid) = ep And Val10(c, "SELECT VehicleID FROM Episodes WHERE EpisodeID=?", ep) = vid
                If Len(key) > 0 Then ok = ok And Len(key) = 44 And VarType(sh.Cells(row, 60).Value2) = vbString And sh.Cells(row, 60).NumberFormat = "@" And CI(sh, row)
            End If
        Next row
    Next sn
    CR "CF_PROJECOES_IDENTIDADE_TIPOS", ok And InStr(C1_Diagnostico(), "PROJECTION=VERIFIED") > 0, "Todos os registros projetados conferidos por PID/EP/VID; chaves String44/@/Ignore", "REGISTROS=" & count & ";" & C1_Diagnostico(), "VBA_VERIFICACAO_PROJECAO_COMPLETA"
    c.Close
    CS "F2A0001", False: C1_Context(Cad_NFE)("Version") = CLng(C1_Context(Cad_NFE)("Version")) - 1
    Cad_NFE.CommandButtonCopiarEntrada.Value = True
    CR "CF_TOKEN_OBSOLETO", InStr(C1_UltimoResultado, "obsoleta") > 0, "Selecao antiga nao autoriza gravador", C1_UltimoResultado
    CS "F2A0001", False: Cad_NFE.ListBoxLista.ListIndex = -1
    CR "CF_LISTINDEX_NEGATIVO", Cad_NFE.ListBoxLista.ListIndex = -1, "Sem selecao nao navegar para outro registro", C1_UltimoResultado
    Cad_NFE.TextBoxPesquisa.Value = "INEXISTENTE_CF"
    CR "CF_PESQUISA_SEM_RESULTADO", Cad_NFE.ListBoxLista.ListCount = 0, "Pesquisa nao retorna primeira linha como alternativa", CStr(Cad_NFE.ListBoxLista.ListCount)
    CR "CF_APP_ESTADOS", Application.EnableEvents And Application.ScreenUpdating And Application.DisplayAlerts And Application.Calculation = xlCalculationAutomatic, "Estados globais anteriores preservados", Application.EnableEvents & ";" & Application.ScreenUpdating & ";" & Application.DisplayAlerts & ";" & Application.Calculation
    Cad_NFE.Hide: F2_Snapshot "CF_metadados": Exit Sub
Bad: CR "CF_LEITORES_ERRO", False, "Validar todos os vinculos e tipos", Err.Number & ":" & Err.Description
End Sub

Private Function CI(ByVal sh As Worksheet, ByVal row As Long) As Boolean
    Dim wasProtected As Boolean, n As Long, msg As String
    wasProtected = sh.ProtectContents: On Error GoTo Bad
    If wasProtected Then sh.Unprotect "CANDIDATO_SINTETICO"
    CI = sh.Cells(row, 60).Errors(xlNumberAsText).Ignore
    If wasProtected Then sh.Protect "CANDIDATO_SINTETICO"
    Exit Function
Bad:
    n = Err.Number: msg = Err.Description: On Error Resume Next
    If wasProtected Then sh.Protect "CANDIDATO_SINTETICO"
    On Error GoTo 0: Err.Raise n, "CI", msg
End Function
Public Sub CX_XMLNegativos()
    Dim c As Object, g As clsTesteCandidato, item As Variant, f As String, d As Object, n As Long, revision As Long, ep As String, version As Long, valid As Boolean, modality As String, mark As Object
    On Error GoTo Bad
    Set g = New clsTesteCandidato: Set C1_FornecedorXML = g
    For Each item In Array("MALFORMADO", "NUMERO_INVALIDO", "CHAVE_INVALIDA", "OPERACAO_INCOMPATIVEL", "VALOR_INCORRETO", "DADOS_FISCAIS_AUSENTES", "SOMATORIO_DIVERGENTE", "NFE_DIVERGENTE", "CHAVE_DIVERGENTE", "NUMERO_DUPLICADO")
        CS "F2A0001", False: ep = C1_Context(Cad_NFE)("EpisodeID"): version = C1_Context(Cad_NFE)("Version")
        Cad_NFE.CommandButtonCopiarEntrada.Value = True
        With Cad_NFE_Saida
            .ComboBox_Operação.Value = "Venda": .TextBox_N_Saida.Value = "25051": .TextBox_Valor_Saida.Value = "400,2222": .TextBox_Data.Value = "01/11/2026"
            .ComboBox_Por.Value = "12": .TextBox_ICMS.Value = "0": .TextBox_PIS.Value = "0": .TextBox_COFINS.Value = "0": .TextBoxBaseICMS.Value = "0": .TextBoxBasePISeCOFINS.Value = "0": .TextBoxChaveDeAcesso.Value = KEY_CF
        End With
        f = ThisWorkbook.Path & "\cx_" & CStr(item) & ".xml"
        Set mark = FSO10.CreateTextFile(ThisWorkbook.Path & "\caso_xml_atual.txt", True, True): mark.WriteLine CStr(item) & "|" & f: mark.Close
        Set c = C3_Conectar(): n = Val10(c, "SELECT Count(*) FROM Receipts"): revision = Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1"): c.Close
        modality = "VBA_EVENTO_REAL_COM_RELATORIO_XML"
        If item = "NFE_DIVERGENTE" Or item = "CHAVE_DIVERGENTE" Then
            Set d = CreateObject("MSXML2.DOMDocument.6.0"): d.async = False: d.resolveExternals = False: d.setProperty "ProhibitDTD", True
            If Not d.Load(f) Then Err.Raise 5, , "Documento sintetico direto nao carregou"
            If item = "NFE_DIVERGENTE" Then Cad_NFE_Saida.TextBox_N_Saida.Value = "25052" Else Cad_NFE_Saida.TextBoxChaveDeAcesso.Value = "11234567890123456789012345678901234567890123"
            valid = ValidarXMLNFe(Cad_NFE_Saida, d, f)
            modality = "VBA_VALIDADOR_REAL_DOCUMENTO_DIRETO_E_RELATORIO"
        Else
            Set g = New clsTesteCandidato: g.ArquivoXML = f: Set C1_FornecedorXML = g
            Cad_NFE_Saida.CommandButton_Gravar.Value = True
            valid = (InStr(C1_UltimoResultado, "SAIDA_CONFIRMADA") = 1)
        End If
        Set c = C3_Conectar()
        CR "CX_XML_" & CStr(item), Not valid And (item = "NFE_DIVERGENTE" Or item = "CHAVE_DIVERGENTE" Or InStr(C1_UltimoResultado, "XML") > 0 Or (item = "NUMERO_DUPLICADO" And InStr(C1_UltimoResultado, "Numero") > 0)) And Val10(c, "SELECT Count(*) FROM Receipts") = n And Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1") = revision And Val10(c, "SELECT VersionNo FROM Episodes WHERE EpisodeID=?", ep) = version And Audit10(c) = "OK", "Recusar documento divergente sem fatos ou versoes novas", "VALIDO=" & valid & ";" & C1_UltimoResultado & ";REV=" & revision, modality
        c.Close: Cad_NFE_Saida.Hide
    Next item
    Set C1_FornecedorXML = Nothing: F2_Snapshot "CX_xml_negativos": Exit Sub
Bad:
    Set C1_FornecedorXML = Nothing
    CR "CX_XML_ERRO", False, "Concluir cenarios XML negativos", CStr(item) & ";" & Err.Number & ":" & Err.Description
End Sub
Private Function CG(ByVal c As Object) As String
    On Error GoTo Bad
    C3_ValidarVinculo c: CG = "ACEITO": Exit Function
Bad: CG = CStr(Err.Number) & ":" & Err.Description
End Function
Public Sub CX_GuardasCopias()
    Dim c As Object, copy As Object, file As String, base As String, lineage As String, st As String, n As Long, rev As Long, registry As String, marker As String
    On Error GoTo Bad
    Set c = C3_Conectar(): n = Val10(c, "SELECT Count(*) FROM Receipts"): rev = Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1"): c.Close
    file = ThisWorkbook.Path & "\cx_copia_mesmo_baseid_" & NewID10() & ".accdb"
    FSO10.CopyFile ThisWorkbook.Path & "\" & DB_FINAL, file, False
    Set copy = CreateObject("ADODB.Connection"): copy.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & file
    base = Val10(copy, "SELECT BaseID FROM Meta WHERE ID=1"): lineage = Val10(copy, "SELECT LineageID FROM Meta WHERE ID=1")
    st = CG(copy)
    CR "CX_COPIA_MESMO_BASEID_RECUSADA", InStr(st, "Copia de ACCDB nao e autoridade") > 0, "Mesmo BaseID nao elege segunda autoridade", st, "VBA_GUARDA_REAL_EM_COPIA_ACCDB"
    Exec10 copy, "UPDATE Meta SET BaseID=? WHERE ID=1", NewID10(): st = CG(copy)
    CR "CX_BASEID_DIVERGENTE_RECUSADO", InStr(st, "Base ou linhagem divergente") > 0, "Base divergente nao escreve", st, "VBA_GUARDA_REAL_EM_COPIA_ACCDB"
    Exec10 copy, "UPDATE Meta SET BaseID=?,LineageID=? WHERE ID=1", base, NewID10(): st = CG(copy)
    CR "CX_LINHAGEM_DIVERGENTE_RECUSADA", InStr(st, "Base ou linhagem divergente") > 0, "Linhagem divergente nao escreve", st, "VBA_GUARDA_REAL_EM_COPIA_ACCDB"
    Exec10 copy, "UPDATE Meta SET LineageID=? WHERE ID=1", lineage
    registry = Val10(copy, "SELECT RegistryPath FROM AuthorityBinding WHERE ID=1")
    marker = ThisWorkbook.Path & "\registro_inexistente_" & NewID10() & ".accdb"
    Exec10 copy, "UPDATE AuthorityBinding SET RegistryPath=? WHERE ID=1", marker: st = CG(copy)
    CR "CX_REGISTRO_AUSENTE_NAO_ELEGE", InStr(st, "Registro de autoridade ausente") > 0 And Not FSO10.FileExists(marker), "Registro ausente nao cria nem elege copia", st, "VBA_GUARDA_REAL_EM_COPIA_ACCDB"
    copy.Close
    Set c = C3_Conectar()
    CR "CX_AUTORIDADE_UNICA_PRESERVADA", Val10(c, "SELECT Count(*) FROM Receipts") = n And Val10(c, "SELECT RevisionNo FROM Meta WHERE ID=1") = rev And Audit10(c) = "OK", "Somente a copia foi adulterada; autoridade atual intacta", C1_Diagnostico(), "VBA_GUARDA_REAL_EM_COPIA_ACCDB"
    c.Close: F2_Snapshot "CX_guardas": Exit Sub
Bad:
    CR "CX_GUARDAS_ERRO", False, "Concluir guardas em copia sintetica independente", Err.Number & ":" & Err.Description
    On Error Resume Next: If Not copy Is Nothing Then copy.Close
End Sub
Public Sub CX_ConcorrenciaExterna()
    Dim c As Object, n As Long, rev As Long, f As Object, marker As String
    On Error GoTo Bad
    ' Contagens lidas antes de disparar o processo concorrente ficam no arquivo do ensaio.
    Load Cad_NFE: Cad_NFE.MultiPage1.Value = 1: C1_LimparContexto Cad_NFE
    With Cad_NFE
        .ComboBox_OperaçãoFilial.Value = "Compra": .N_NfeFilial.Value = "12061": .UNFilial.Value = "1": .CODFilial.Value = "CFCC": .ComboBox_VEÍCULOFilial.Value = "FICTICIO A": .PLACAFilial.Value = "CFCC001": .Combo_ANO_MODFilial.Value = "2026": .ComboBox_CORFilial.Value = "BRANCO": .txt_ValorFilial.Value = "246,1357": .DataFilial.Value = "03/11/2026": .ComboBox_OutroEnvioFilial.Value = "TESTE": .TextBoxChaveDeAcessoFilial.Value = KEY_CF
    End With
    Cad_NFE.CommandButton_GravarFilial.Value = True
    Set f = FSO10.CreateTextFile(ThisWorkbook.Path & "\liberar_concorrencia.flag", True, True): f.WriteLine C1_UltimoResultado: f.Close
    Application.Wait Now + TimeSerial(0, 0, 3)
    Set c = C3_Conectar()
    CR "CX_ESCRITOR_EXTERNO_RECUSADO", Val10(c, "SELECT Count(*) FROM Projections WHERE DocNumber='12061'") = 0 And Val10(c, "SELECT Count(*) FROM Vehicles WHERE Plate='CFCC001'") = 0 And InStr(C1_UltimoResultado, "RECUSADA") > 0 And Audit10(c) = "OK", "Processo ACE concorrente segura registro; evento real recusa sem identidade parcial", C1_UltimoResultado, "VBA_EVENTO_REAL_CONTRA_LOCK_ACE_EM_PROCESSO_EXTERNO"
    c.Close: F2_Snapshot "CX_concorrencia": Exit Sub
Bad:
    Set f = FSO10.CreateTextFile(ThisWorkbook.Path & "\liberar_concorrencia.flag", True, True): f.WriteLine Err.Description: f.Close
    CR "CX_CONCORRENCIA_ERRO", False, "Recusar escrita concorrente sem efeitos", Err.Number & ":" & Err.Description
End Sub
