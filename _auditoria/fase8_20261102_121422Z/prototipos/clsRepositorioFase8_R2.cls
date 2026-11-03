VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsRepositorioFase8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Repositorio SOMENTE experimental. Indice localiza; identidade e conteudo autorizam.
Private WithEvents mExcel As Excel.Application
Attribute mExcel.VB_VarHelpID = -1
Private mWB As Workbook, mDirty As Boolean, mRevision As String, mEpoch As Long
Private mRows As Variant, mEpisodes As Variant, mPID As Object, mEP As Object, mSeals As Object, mPending As Object, mActiveVID As Object
Private Function Dict8() As Object
    Set Dict8 = CreateObject("Scripting.Dictionary")
    Dict8.CompareMode = vbBinaryCompare
End Function
Public Sub Iniciar(ByVal wb As Workbook)
    Set mWB = wb: Set mExcel = wb.Application: mDirty = True: mEpoch = 0
    Recarregar
End Sub
Private Sub mExcel_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Sair
    If mWB Is Nothing Then Exit Sub
    If Not (Sh.Parent Is mWB) Then Exit Sub
    Select Case Sh.Name
        Case "_F8_Dados", "_F8_Episodios", "_F8_Movimentos", "_F8_Operacoes", "_F8_Selos", "_F8_Controle"
            mDirty = True: mEpoch = mEpoch + 1
    End Select
Sair:
End Sub
Public Property Get Revisao() As String
    Revisao = CStr(mWB.Worksheets("_F8_Controle").Range("B1").Value2) & ":" & CStr(mWB.Worksheets("_F8_Controle").Range("B2").Value2)
End Property
Public Property Get EpocaIndice() As Long
    EpocaIndice = mEpoch
End Property
Public Sub Invalidar()
    mDirty = True: mEpoch = mEpoch + 1
End Sub
Private Sub Indexar(ByVal d As Object, ByVal id As String, ByVal r As Long)
    If id = "" Then Exit Sub
    If d.Exists(id) Then d(id) = -1 Else d.Add id, r
End Sub
Public Sub Recarregar()
    Dim sh As Worksheet, n As Long, r As Long, a As Variant, estado As String, ep As String
    Set mPID = Dict8(): Set mEP = Dict8(): Set mSeals = Dict8(): Set mPending = Dict8(): Set mActiveVID = Dict8()
    Set sh = mWB.Worksheets("_F8_Dados"): n = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If n < 2 Then n = 2
    mRows = sh.Range("A1:R" & n).Value2
    For r = 2 To UBound(mRows, 1): Indexar mPID, CStr(mRows(r, 1)), r: Next r
    Set sh = mWB.Worksheets("_F8_Episodios"): n = Application.Max(2, sh.Cells(sh.Rows.Count, 1).End(xlUp).Row)
    mEpisodes = sh.Range("A1:E" & n).Value2
    For r = 2 To UBound(mEpisodes, 1)
        Indexar mEP, CStr(mEpisodes(r, 1)), r
        If CStr(mEpisodes(r, 3)) = "ATIVO" Then
            ep = CStr(mEpisodes(r, 2))
            If Not mActiveVID.Exists(ep) Then mActiveVID.Add ep, 0
            mActiveVID(ep) = CLng(mActiveVID(ep)) + 1
        End If
    Next r
    Set sh = mWB.Worksheets("_F8_Selos"): n = Application.Max(2, sh.Cells(sh.Rows.Count, 1).End(xlUp).Row)
    a = sh.Range("A1:B" & n).Value2
    For r = 2 To UBound(a, 1)
        If CStr(a(r, 1)) <> "" Then
            If mSeals.Exists(CStr(a(r, 1))) Then mSeals(CStr(a(r, 1))) = "ERRO:SELO_DUPLICADO" Else mSeals.Add CStr(a(r, 1)), CStr(a(r, 2))
        End If
    Next r
    Set sh = mWB.Worksheets("_F8_Operacoes"): n = Application.Max(2, sh.Cells(sh.Rows.Count, 1).End(xlUp).Row)
    a = sh.Range("A1:H" & n).Value2
    For r = 2 To UBound(a, 1)
        ep = CStr(a(r, 2)): estado = CStr(a(r, 4))
        If ep <> "" And estado <> "CONCLUIDA" And estado <> "COMPENSADA_VERIFICADA" Then
            If estado = "AGUARDANDO_CONCILIACAO" Then mPending(ep) = "AGUARDANDO_CONCILIACAO" Else mPending(ep) = "OPERACAO_PENDENTE"
        End If
    Next r
    mRevision = Revisao: mDirty = False
End Sub
Private Function EstadoLinha(ByVal r As Long, ByVal finalidade As String) As String
    Dim ep As String, e As Long, pid As String
    pid = CStr(mRows(r, 1)): ep = CStr(mRows(r, 2))
    If Not mPID.Exists(pid) Then EstadoLinha = "INEXISTENTE": Exit Function
    If CLng(mPID(pid)) = -1 Then EstadoLinha = "DUPLICADO": Exit Function
    If Not mEP.Exists(ep) Then EstadoLinha = "AGUARDANDO_CONCILIACAO": Exit Function
    e = CLng(mEP(ep)): If e < 2 Then EstadoLinha = "IDENTIDADE_AMBIGUA": Exit Function
    If CStr(mEpisodes(e, 2)) <> CStr(mRows(r, 3)) Then EstadoLinha = "IDENTIDADE_AMBIGUA": Exit Function
    If mPending.Exists(ep) Then EstadoLinha = CStr(mPending(ep)): Exit Function
    If finalidade = "OPERACIONAL" Then
        If Not mActiveVID.Exists(CStr(mRows(r, 3))) Then EstadoLinha = "IDENTIDADE_AMBIGUA": Exit Function
        If CLng(mActiveVID(CStr(mRows(r, 3)))) <> 1 Then EstadoLinha = "IDENTIDADE_AMBIGUA": Exit Function
        If CStr(mRows(r, 8)) <> "ATIVO" Or CStr(mEpisodes(e, 3)) <> "ATIVO" Then EstadoLinha = "BAIXADO": Exit Function
        If CStr(mEpisodes(e, 4)) <> pid Then EstadoLinha = "IDENTIDADE_AMBIGUA": Exit Function
        If mEpisodes(e, 5) <> mRows(r, 7) Then EstadoLinha = "VERSAO_OBSOLETA": Exit Function
    ElseIf finalidade <> "HISTORICO" Then
        EstadoLinha = "FINALIDADE_INVALIDA": Exit Function
    End If
    If Not mSeals.Exists(pid) Then EstadoLinha = "AGUARDANDO_CONCILIACAO": Exit Function
    If CStr(mSeals(pid)) <> F8_CanonRegistro(mWB, r) Then EstadoLinha = "CONTEUDO_DIVERGENTE": Exit Function
    EstadoLinha = "ENCONTRADO"
End Function
Public Function Resolver(ByVal pid As String, ByVal ep As String, ByVal vid As String, ByVal expectedVersion As Long, Optional ByVal finalidade As String = "OPERACIONAL", Optional ByVal expectedRevision As String = "") As Object
    Dim d As Object, sh As Worksheet, r As Long, n As Long, estado As String, a As Variant, e As Long
    Set d = Dict8(): d.Add "Status", "INEXISTENTE": Set Resolver = d
    If pid = "" Or ep = "" Or vid = "" Then Exit Function
    If mDirty Or mRevision <> Revisao Or Not mExcel.EnableEvents Then Recarregar
    If expectedRevision <> "" And expectedRevision <> Revisao Then d("Status") = "VERSAO_OBSOLETA": Exit Function
    Set sh = mWB.Worksheets("_F8_Dados"): n = Application.Max(2, sh.Cells(sh.Rows.Count, 1).End(xlUp).Row)
    r = Application.CountIf(sh.Range("A2:A" & n), pid)
    If r = 0 Then Exit Function
    If r <> 1 Then d("Status") = "DUPLICADO": Exit Function
    If Not mPID.Exists(pid) Then Recarregar
    If Not mPID.Exists(pid) Then Exit Function
    r = CLng(mPID(pid))
    If r < 2 Then d("Status") = "DUPLICADO": Exit Function
    If CStr(sh.Cells(r, 1).Value2) <> pid Then Recarregar: r = CLng(mPID(pid))
    a = sh.Cells(r, 1).Resize(1, 18).Value2
    If CStr(a(1, 2)) <> ep Or CStr(a(1, 3)) <> vid Then d("Status") = "IDENTIDADE_AMBIGUA": Exit Function
    If expectedVersion < 1 Or a(1, 7) <> expectedVersion Then d("Status") = "VERSAO_OBSOLETA": Exit Function
    ' Conferir duplicidade do episodio tambem quando eventos estiverem desligados.
    With mWB.Worksheets("_F8_Episodios")
        n = Application.Max(2, .Cells(.Rows.Count, 1).End(xlUp).Row)
        If Application.CountIf(.Range("A2:A" & n), ep) <> 1 Then d("Status") = "IDENTIDADE_AMBIGUA": Exit Function
    End With
    ' Revalidar metadados atuais, inclusive mudancas sem notificacao de evento.
    With mWB.Worksheets("_F8_Episodios")
        e = CLng(mEP(ep))
        If CStr(.Cells(e, 1).Value2) <> ep Then Recarregar: e = CLng(mEP(ep))
        For n = 1 To 5: mEpisodes(e, n) = .Cells(e, n).Value2: Next n
        n = Application.Max(2, .Cells(.Rows.Count, 1).End(xlUp).Row)
        mActiveVID(vid) = Application.CountIfs(.Range("B2:B" & n), vid, .Range("C2:C" & n), "ATIVO")
    End With
    With mWB.Worksheets("_F8_Operacoes")
        n = Application.Max(2, .Cells(.Rows.Count, 1).End(xlUp).Row)
        If Application.CountIfs(.Range("B2:B" & n), ep, .Range("D2:D" & n), "<>CONCLUIDA", .Range("D2:D" & n), "<>COMPENSADA_VERIFICADA") > 0 Then
            d("Status") = "OPERACAO_PENDENTE": Exit Function
        End If
    End With
    ' Atualizar o candidato com valores atuais; indice jamais e autoridade.
    For n = 1 To 18: mRows(r, n) = a(1, n): Next n
    estado = EstadoLinha(r, finalidade): d("Status") = estado
    If estado <> "ENCONTRADO" Then Exit Function
    d.Add "Row", r: d.Add "Data", a: d.Add "Revision", Revisao
End Function
Public Function Listar(ByVal query As String, ByVal finalidade As String, ByVal localEstoque As String, ByRef dados As Variant) As String
    Dim r As Long, n As Long, c As Long, estado As String, a() As Variant, saida() As Variant
    ' Uma leitura fresca em lote nao confia em notificacao de mudanca isolada.
    Recarregar
    ReDim a(0 To UBound(mRows, 1), 0 To 7)
    For r = 2 To UBound(mRows, 1)
        If CStr(mRows(r, 1)) <> "" Then
            If finalidade = "HISTORICO" Or CStr(mRows(r, 8)) = "ATIVO" Then
                estado = EstadoLinha(r, finalidade)
                If estado <> "ENCONTRADO" Then Listar = estado: dados = Empty: Exit Function
                If localEstoque = "TODOS" Or CStr(mRows(r, 9)) = localEstoque Then
                    If query = "" Or InStr(1, CStr(mRows(r, 10)) & ";" & CStr(mRows(r, 11)), query, vbTextCompare) > 0 Then
                        a(n, 0) = mRows(r, 1): a(n, 1) = mRows(r, 2): a(n, 2) = mRows(r, 3): a(n, 3) = mRows(r, 7)
                        a(n, 4) = mRows(r, 10): a(n, 5) = mRows(r, 11): a(n, 6) = mRows(r, 9): a(n, 7) = Revisao
                        n = n + 1
                    End If
                End If
            End If
        End If
    Next r
    If n = 0 Then dados = Empty: Listar = "SEM_RESULTADO": Exit Function
    ReDim saida(0 To n - 1, 0 To 7)
    For r = 0 To n - 1: For c = 0 To 7: saida(r, c) = a(r, c): Next c: Next r
    dados = saida: Listar = "ENCONTRADO"
End Function
Public Function LinhaIndice(ByVal pid As String) As Long
    ' Sonda de benchmark: NAO autoriza leitura/escrita sem Resolver.
    If mDirty Or mRevision <> Revisao Then Recarregar
    If mPID.Exists(pid) Then If CLng(mPID(pid)) > 1 Then LinhaIndice = CLng(mPID(pid))
End Function
Public Sub Finalizar()
    Set mExcel = Nothing: Set mWB = Nothing
End Sub
