Attribute VB_Name = "mdlCriaBackup"
Option Explicit

Sub SaveWorkbookBackup()
    Dim caminho As String
    On Error GoTo Falhou
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Salve a pasta antes de criar o backup.", vbExclamation, "Backup"
        Exit Sub
    End If
    caminho = CriarBackupVersionado(ThisWorkbook, ThisWorkbook.Path & Application.PathSeparator & "Backups")
    MsgBox "Backup criado sem sobrescrever a pasta original:" & vbCrLf & caminho, vbInformation, "Backup"
    Exit Sub
Falhou:
    MsgBox "Backup nao confirmado. Erro " & Err.Number & ": " & Err.Description, vbCritical, "Backup"
End Sub

Sub SaveWorkbookBackupToFloppy()
    Dim seletor As Object, caminho As String
    On Error GoTo Falhou
    'Nome publico legado preservado; nenhum destino D: e presumido.
    Set seletor = Application.FileDialog(4)
    seletor.title = "Escolha a pasta de destino do backup versionado"
    If seletor.show <> -1 Then Exit Sub
    caminho = CriarBackupVersionado(ThisWorkbook, CStr(seletor.SelectedItems(1)))
    MsgBox "Backup criado:" & vbCrLf & caminho, vbInformation, "Backup"
    Exit Sub
Falhou:
    MsgBox "Backup nao confirmado. Erro " & Err.Number & ": " & Err.Description, vbCritical, "Backup"
End Sub

Public Function CriarBackupVersionado(ByVal livro As Workbook, ByVal pastaRaiz As String) As String
    Dim fso As Object
    Dim pastaVersao As String, destino As String, prefixo As String
    Dim tentativa As Long, numeroErro As Long, descricaoErro As String

    On Error GoTo Falhou
    If livro Is Nothing Then Err.Raise vbObjectError + 2200, , "Pasta de origem ausente."
    If Len(livro.Path) = 0 Then Err.Raise vbObjectError + 2201, , "A pasta ainda nao foi salva."
    If Len(pastaRaiz) = 0 Then Err.Raise vbObjectError + 2202, , "Destino nao informado."
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pastaRaiz) Then fso.CreateFolder pastaRaiz
    prefixo = fso.GetBaseName(livro.name) & "_" & Format$(Now, "yyyymmdd_hhnnss")
    For tentativa = 1 To 999
        pastaVersao = fso.BuildPath(pastaRaiz, prefixo & "_" & Format$(tentativa, "000"))
        'CreateFolder reserva o diretorio; nao reutiliza um backup existente.
        On Error Resume Next
        Err.Clear
        fso.CreateFolder pastaVersao
        numeroErro = Err.Number
        descricaoErro = Err.Description
        On Error GoTo Falhou
        If numeroErro = 0 Then Exit For
        If numeroErro <> 58 Then Err.Raise numeroErro, , descricaoErro
    Next tentativa
    If tentativa > 999 Then Err.Raise vbObjectError + 2203, , "Sem nome disponivel para o backup."
    destino = fso.BuildPath(pastaVersao, livro.name)
    If fso.FileExists(destino) Then Err.Raise vbObjectError + 2204, , "Destino ja existe."
    'SaveCopyAs inclui o estado em memoria sem gravar o arquivo original.
    livro.SaveCopyAs destino
    VerificarArquivoBackup destino
    CriarBackupVersionado = destino
    Exit Function
Falhou:
    numeroErro = Err.Number
    descricaoErro = Err.Description
    'Nao apaga copias, mesmo incompletas; preserva evidencia para diagnostico.
    Err.Raise numeroErro, "CriarBackupVersionado", descricaoErro
End Function

Private Sub VerificarArquivoBackup(ByVal caminho As String)
    Dim canal As Integer, assinatura(0 To 1) As Byte
    Dim numeroErro As Long, descricaoErro As String
    On Error GoTo Falhou
    If FileLen(caminho) < 2 Then Err.Raise vbObjectError + 2205, , "Backup vazio ou ausente."
    canal = FreeFile
    Open caminho For Binary Access Read As #canal
    Get #canal, , assinatura
    Close #canal
    canal = 0
    If Not ((assinatura(0) = &H50 And assinatura(1) = &H4B) Or _
            (assinatura(0) = &HD0 And assinatura(1) = &HCF)) Then
        Err.Raise vbObjectError + 2206, , "Cabecalho de arquivo Office nao reconhecido."
    End If
    'Tamanho e cabecalho nao substituem hash SHA-256 nem teste de restauracao.
    'Retencao e limpeza automatica nao sao executadas por esta rotina.
    Exit Sub
Falhou:
    numeroErro = Err.Number
    descricaoErro = Err.Description
    On Error Resume Next
    If canal > 0 Then Close #canal
    On Error GoTo 0
    Err.Raise numeroErro, "VerificarArquivoBackup", descricaoErro
End Sub
