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
                If Len(key) > 0 Then ok = ok And Len(key) = 44 And VarType(sh.Cells(row, 60).Value2) = vbString And sh.Cells(row, 60).NumberFormat = "@" And sh.Cells(row, 60).Errors(xlNumberAsText).Ignore
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
