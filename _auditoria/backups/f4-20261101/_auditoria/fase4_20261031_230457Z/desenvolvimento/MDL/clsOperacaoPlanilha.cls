VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsOperacaoPlanilha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Compensacao em memoria de destinos pequenos; nao e transacao ACID.
'Nao restaura exclusoes estruturais, nomes, objetos ou um Excel encerrado.
Private mLivro As Workbook
Private mItens As Collection
Private mAtiva As Boolean
Private mIniciada As Boolean
Private mConfirmada As Boolean
Private mBaixaIniciada As Boolean
Private mSalvamentoIniciado As Boolean
Private mEventos As Boolean
Private mTela As Boolean
Private mAlertas As Boolean
Private mCalculo As XlCalculation

Private Sub Class_Initialize()
    Set mItens = New Collection
End Sub

Public Sub Iniciar(ByVal livro As Workbook)
    If mIniciada Or mAtiva Or mItens.Count > 0 Then Err.Raise vbObjectError + 2100, , "Operacao ja iniciada."
    If livro Is Nothing Then Err.Raise vbObjectError + 2105, , "Pasta nao informada."
    If livro.ReadOnly Then Err.Raise vbObjectError + 2106, , "Pasta aberta somente para leitura."
    Set mLivro = livro
    mEventos = Application.EnableEvents
    mTela = Application.ScreenUpdating
    mAlertas = Application.DisplayAlerts
    mCalculo = Application.Calculation
    mIniciada = True
    mAtiva = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

Public Sub RegistrarDestino(ByVal alvo As Range)
    Dim celula As Range, item As Collection
    Dim temSpill As Variant
    If Not mAtiva Then Err.Raise vbObjectError + 2101, , "Operacao nao iniciada."
    If alvo Is Nothing Then Err.Raise vbObjectError + 2107, , "Destino nao informado."
    If Not alvo.Parent.Parent Is mLivro Then
        Err.Raise vbObjectError + 2102, , "Destino de outra pasta."
    End If
    VerificarDestinoGravacao alvo
    For Each celula In alvo.Cells
        If celula.HasArray Then Err.Raise vbObjectError + 2103, , "Destino com formula matricial."
        temSpill = False
        On Error Resume Next
        temSpill = CallByName(celula, "HasSpill", VbGet)
        Err.Clear
        On Error GoTo 0
        If Not IsNull(temSpill) Then
            If temSpill Then Err.Raise vbObjectError + 2104, , "Destino com formula derramada."
        End If
        Set item = New Collection
        item.Add celula
        item.Add CBool(celula.HasFormula)
        item.Add celula.Formula
        item.Add celula.Value2
        item.Add CStr(celula.NumberFormat)
        item.Add CBool(celula.Errors(xlNumberAsText).Ignore)
        mItens.Add item
    Next celula
End Sub

Public Sub MarcarBaixaIniciada()
    'A partir daqui a estrutura de referencias pode ter mudado fora dos destinos.
    mBaixaIniciada = True
End Sub

Public Sub MarcarSalvamentoIniciado()
    mSalvamentoIniciado = True
End Sub

Public Sub Confirmar()
    mConfirmada = True
    Finalizar
End Sub

Public Function Recuperar(ByRef diagnostico As String) As Boolean
    Dim i As Long, falhas As Long
    diagnostico = vbNullString
    If mConfirmada Then
        diagnostico = "Dados ja confirmados; falha posterior exige conferir a interface."
        Exit Function
    End If
    If mSalvamentoIniciado Then
        diagnostico = "Salvamento iniciado: conferir arquivo em disco antes de nova tentativa."
        Exit Function
    End If
    If mBaixaIniciada Then
        diagnostico = "Baixa estrutural iniciada: conciliacao manual obrigatoria antes de nova tentativa."
        Exit Function
    End If
    If Not mAtiva Then
        diagnostico = "Operacao inativa; nenhuma recuperacao executada."
        Exit Function
    End If
    For i = mItens.Count To 1 Step -1
        If Not RestaurarCelula(mItens(i)) Then falhas = falhas + 1
    Next i
    Recuperar = (falhas = 0)
    If Recuperar Then
        diagnostico = "Destinos restaurados em memoria antes de qualquer Delete."
    Else
        diagnostico = "Falhas na compensacao: " & falhas & ". Conciliacao manual obrigatoria."
    End If
End Function

Private Function RestaurarCelula(ByVal item As Collection) As Boolean
    Dim celula As Range, valor As Variant
    On Error GoTo Falhou
    Set celula = item(1)
    valor = item(4)
    celula.ClearContents
    If CBool(item(2)) Then
        celula.NumberFormat = CStr(item(5))
        celula.Formula = item(3)
    ElseIf VarType(valor) = vbString Then
        'Texto que comeca com = nao pode virar formula na recuperacao.
        celula.NumberFormat = "@"
        celula.Value2 = CStr(valor)
        celula.NumberFormat = CStr(item(5))
    Else
        celula.NumberFormat = CStr(item(5))
        celula.Value2 = valor
    End If
    celula.Errors(xlNumberAsText).Ignore = CBool(item(6))
    RestaurarCelula = True
    Exit Function
Falhou:
    RestaurarCelula = False
End Function

Public Sub Finalizar()
    If Not mAtiva Then Exit Sub
    'Tenta cada estado; falhas ficam no diagnostico da Janela Imediata.
    On Error Resume Next
    Err.Clear
    Application.Calculation = mCalculo
    If Err.Number <> 0 Then Debug.Print "Restaurar Calculation: "; Err.Number; Err.Description
    Err.Clear
    Application.DisplayAlerts = mAlertas
    If Err.Number <> 0 Then Debug.Print "Restaurar DisplayAlerts: "; Err.Number; Err.Description
    Err.Clear
    Application.ScreenUpdating = mTela
    If Err.Number <> 0 Then Debug.Print "Restaurar ScreenUpdating: "; Err.Number; Err.Description
    Err.Clear
    Application.EnableEvents = mEventos
    If Err.Number <> 0 Then Debug.Print "Restaurar EnableEvents: "; Err.Number; Err.Description
    mAtiva = False
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Finalizar
End Sub
