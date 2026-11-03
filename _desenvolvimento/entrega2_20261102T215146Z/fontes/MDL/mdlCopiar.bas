Attribute VB_Name = "mdlCopiar"
Option Explicit

'========================
' API Clipboard (Windows) - opcional, caso use o botão Copiar
'========================
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal source As LongPtr, ByVal length As LongPtr)
#End If

Private Const GMEM_MOVEABLE As Long = &H2
Private Const CF_UNICODETEXT As Long = 13


'========================
' API Pública (chame do UserForm)
'========================

' Navega e atualiza LabelTextoCelula + LabelConteudoCélula
Public Sub UF_NavegarCelula(ByVal frm As Object, ByVal dRow As Long, ByVal dCol As Long)
    MoveActiveCell dRow, dCol
    UF_AtualizaLabels frm, "Navegação"
End Sub

' Atualiza apenas os labels (útil no Initialize)
Public Sub UF_AtualizarPainel(ByVal frm As Object)
    UF_AtualizaLabels frm, "Pronto"
End Sub

' Copia o conteúdo da célula ativa e atualiza o status
Public Sub UF_CopiarCelula(ByVal frm As Object, Optional ByVal usarTextoFormatado As Boolean = True)
    On Error GoTo Trata

    Dim txt As String
    If usarTextoFormatado Then
        txt = CStr(ActiveCell.text)   'como aparece (com formatação)
    Else
        txt = CStr(ActiveCell.value)  'valor "puro"
    End If

    Clipboard_SetText txt
    UF_AtualizaLabels frm, "Conteúdo copiado"

    Exit Sub
Trata:
    UF_SetStatus frm, "Não foi possível copiar o conteúdo da célula."
End Sub


'========================
' Internos
'========================

Private Sub UF_AtualizaLabels(ByVal frm As Object, ByVal acao As String)
    On Error Resume Next

    If ActiveCell Is Nothing Then
        UF_SetStatus frm, "Nenhuma célula ativa."
        Exit Sub
    End If

    Dim addr As String
    Dim conteudo As String

    addr = ActiveCell.Address(False, False)
    conteudo = CStr(ActiveCell.text) 'como aparece (com formatação)

    If Len(conteudo) = 0 Then conteudo = "(vazio)"

    ' Evita status gigante (opcional)
    If Len(conteudo) > 200 Then conteudo = Left$(conteudo, 200) & "..."

    ' Atualiza LabelTextoCelula (se existir)
    If ControleExiste(frm, "LabelTextoCelula") Then
        frm.Controls("LabelTextoCelula").CAPTION = CStr(ActiveCell.text)
    End If

    ' Atualiza LabelConteudoCélula com o conteúdo da célula
    ' 1) Em duas linhas (recomendado para leitura):
    UF_SetStatus frm, conteudo

    ' 2) Se preferir tudo em uma linha, use esta no lugar da de cima:
    'UF_SetStatus frm, acao & " - " & addr & " => " & conteudo
End Sub

Private Sub UF_SetStatus(ByVal frm As Object, ByVal msg As String)
    On Error Resume Next
    If ControleExiste(frm, "LabelConteudoCélula") Then
        frm.Controls("LabelConteudoCélula").CAPTION = msg
    End If
End Sub

Private Function ControleExiste(ByVal frm As Object, ByVal nome As String) As Boolean
    On Error GoTo FIM
    Dim tmp As Object
    Set tmp = frm.Controls(nome)
    ControleExiste = True
    Exit Function
FIM:
    ControleExiste = False
End Function

Private Sub MoveActiveCell(ByVal dRow As Long, ByVal dCol As Long)
    On Error GoTo Trata

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If ActiveCell Is Nothing Then Exit Sub

    Dim r As Long, c As Long
    r = ActiveCell.row + dRow
    c = ActiveCell.Column + dCol

    If r < 1 Then r = 1
    If c < 1 Then c = 1
    If r > ActiveSheet.rows.count Then r = ActiveSheet.rows.count
    If c > ActiveSheet.Columns.count Then c = ActiveSheet.Columns.count

    ActiveSheet.Cells(r, c).Select
    Exit Sub

Trata:
    ' proteção contra planilha protegida/seleção inválida
End Sub

Private Sub Clipboard_SetText(ByVal s As String)
    Dim hMem As LongPtr, pMem As LongPtr, cb As LongPtr
    Dim clipboardAberto As Boolean
    Dim numeroErro As Long, descricaoErro As String
    Dim textoTerminado As String
    On Error GoTo TrataErro
    textoTerminado = s & vbNullChar
    cb = LenB(textoTerminado)
    hMem = GlobalAlloc(GMEM_MOVEABLE, cb)
    If hMem = 0 Then Err.Raise vbObjectError + 520, , "Memória indisponível para copiar."
    pMem = GlobalLock(hMem)
    If pMem = 0 Then Err.Raise vbObjectError + 521, , "Falha ao acessar memória para copiar."
    CopyMemory pMem, StrPtr(textoTerminado), cb
    GlobalUnlock hMem
    pMem = 0
    If OpenClipboard(CLngPtr(Application.hWnd)) = 0 Then Err.Raise vbObjectError + 522, , "Área de transferência ocupada."
    clipboardAberto = True
    If EmptyClipboard() = 0 Then Err.Raise vbObjectError + 523, , "Falha ao preparar área de transferência."
    If SetClipboardData(CF_UNICODETEXT, hMem) = 0 Then Err.Raise vbObjectError + 524, , "Falha ao copiar texto."
    hMem = 0 ' O Windows assumiu a posse do bloco.
saida:
    If clipboardAberto Then CloseClipboard
    If pMem <> 0 Then GlobalUnlock hMem
    If hMem <> 0 Then GlobalFree hMem
    On Error GoTo 0
    If numeroErro <> 0 Then Err.Raise numeroErro, "Clipboard_SetText", descricaoErro
    Exit Sub
TrataErro:
    numeroErro = Err.Number
    descricaoErro = Err.Description
    Resume saida
End Sub


