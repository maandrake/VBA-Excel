
'Fase 4: casos reais ausentes na suite anterior; exclusivamente fixture ficticia.
Public Function AuditoriaFase4_IntegridadeComplementar() As String
    Dim ws As Worksheet, op As clsOperacaoPlanilha, diagnostico As String
    Dim chave As String, n As Long, d As String, protegidaPeloTeste As Boolean
    Dim antes As Long
    If InStr(1, CreateObject("Scripting.FileSystemObject").GetFile(ThisWorkbook.FullName).ShortPath, CreateObject("Scripting.FileSystemObject").GetFolder(Environ$("TEMP")).ShortPath & "\", vbTextCompare) <> 1 Then Err.Raise vbObjectError + 2240, , "Exige TEMP."
    If Left$(ThisWorkbook.Name, 18) <> "Fixture_Auditoria_" Then Err.Raise vbObjectError + 2241, , "Exige fixture."
    If ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then Err.Raise vbObjectError + 2242, , "Exige marcador ficticio."
    Set ws = ThisWorkbook.Worksheets("Entradas")
    If ws.ProtectContents Then Err.Raise vbObjectError + 2243, , "Fixture ja protegida; teste recusado."
    On Error GoTo Falhou
    antes = mVerificacoesF2
    ws.Range("A20:D20").ClearContents
    ws.Range("A20").Value2 = "ANTES"
    Set op = New clsOperacaoPlanilha
    op.Iniciar ThisWorkbook
    op.RegistrarDestino ws.Range("A20")
    ws.Range("A20").Value2 = "INTERMEDIARIO"
    op.RegistrarDestino ws.Range("A20")
    ws.Range("A20").Value2 = "DEPOIS"
    ExigirF2 op.Recuperar(diagnostico), "Recuperacao sobreposta"
    ExigirF2 ws.Range("A20").Value2 = "ANTES", "Ordem inversa dos snapshots"
    op.Finalizar
    chave = "00000000000000000000000000000000000000000001"
    ws.Range("B20").NumberFormat = "0.00"
    ws.Range("B20").Value2 = 12.5
    ws.Range("C20").NumberFormat = "@"
    ws.Range("C20").Value2 = chave
    ws.Range("C20").Errors(xlNumberAsText).Ignore = True
    ws.Range("D20").Formula = "=1+2"
    Set op = New clsOperacaoPlanilha
    op.Iniciar ThisWorkbook
    op.RegistrarDestino ws.Range("A20:D20")
    ws.Range("A20").Value2 = "ALTERADO"
    ws.Range("B20").Value2 = 99
    ws.Range("B20").NumberFormat = "0.0000"
    ws.Range("C20").Value2 = "ALTERADO"
    ws.Range("D20").Value2 = 99
    ws.Range("A20:D20").Locked = False
    ws.Range("B20").Locked = True
    ws.Protect Password:="F4QA"
    protegidaPeloTeste = True
    ExigirF2 Not op.Recuperar(diagnostico), "Falha individual reportada"
    ExigirF2 InStr(1, diagnostico, "Falhas na compensacao: 1.", vbBinaryCompare) = 1, "Diagnostico individual"
    ExigirF2 ws.Range("A20").Value2 = "ANTES", "Demais celulas continuam restauradas"
    ExigirF2 ws.Range("C20").Value2 = chave And ws.Range("D20").Formula = "=1+2", "Chave e formula apos falha individual"
    ExigirF2 ws.Range("B20").Value2 = 99, "Nao anunciar recuperacao do destino bloqueado"
    ws.Unprotect Password:="F4QA"
    protegidaPeloTeste = False
    ExigirF2 op.Recuperar(diagnostico), "Nova recuperacao antes de Delete Save"
    ExigirF2 ws.Range("B20").Value2 = 12.5 And ws.Range("B20").NumberFormat = "0.00", "Valor e NumberFormat restaurados"
    ExigirF2 VarType(ws.Range("C20").Value2) = vbString And Len(ws.Range("C20").Value2) = 44, "Tipo e comprimento da chave"
    ExigirF2 ws.Range("C20").NumberFormat = "@" And ws.Range("C20").Errors(xlNumberAsText).Ignore, "Formato e Ignore restaurados"
    op.Finalizar
    ws.Range("A20:D20").Locked = True
    AuditoriaFase4_IntegridadeComplementar = "PASS: " & (mVerificacoesF2 - antes) & " verificacoes VBA complementares; falha individual reconhecida sem atomicidade."
    Exit Function
Falhou:
    n = Err.Number: d = Err.Description
    On Error Resume Next
    If protegidaPeloTeste Then ws.Unprotect Password:="F4QA"
    If Not op Is Nothing Then op.Finalizar
    On Error GoTo 0
    Err.Raise n, "AuditoriaFase4_IntegridadeComplementar", d
End Function
