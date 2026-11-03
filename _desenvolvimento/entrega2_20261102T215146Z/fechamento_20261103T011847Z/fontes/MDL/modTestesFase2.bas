Attribute VB_Name = "modTestesFase2"
Option Explicit

'PREPARADA, NAO EXECUTADA. Somente fixture sem dados reais; nao chama eventos de gravacao.
Private mVerificacoesF2 As Long

Public Function AuditoriaFase2_Executar() As String
    Dim fluxos As Variant, fluxo As Variant, nomes As Variant
    Dim Etapa As Long, WS As Worksheet, operacao As clsOperacaoPlanilha
    Dim diagnostico As String, n As Long, d As String
    If InStr(1, CreateObject("Scripting.FileSystemObject").GetFile(ThisWorkbook.FullName).ShortPath, CreateObject("Scripting.FileSystemObject").GetFolder(Environ$("TEMP")).ShortPath & "\", vbTextCompare) <> 1 Then
        Err.Raise vbObjectError + 2200, , "Exige fixture em TEMP."
    End If
    If Left$(ThisWorkbook.name, 18) <> "Fixture_Auditoria_" Then
        Err.Raise vbObjectError + 2201, , "Nome da fixture invalido."
    End If
    If ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then
        Err.Raise vbObjectError + 2202, , "Marcador ficticio ausente."
    End If
    On Error GoTo Falhou
    mVerificacoesF2 = 0
    fluxos = Array( _
        Array("Entradas", "Entradas Anual"), _
        Array("Entradas Filial", "Entradas Anual Filial"), _
        Array("Saidas", "Todas Saidas"), _
        Array("Saidas", "Entradas Filial", "Entradas Anual Filial", "Todas Saidas"), _
        Array("Saidas Filial", "Todas Saidas Filial", "Entradas", "Entradas Anual"), _
        Array("Saidas", "Todas Saidas"))
    For Each fluxo In fluxos
        nomes = fluxo
        For Etapa = LBound(nomes) To UBound(nomes)
            TestarEtapaF2 nomes, Etapa
        Next Etapa
    Next fluxo

    Set WS = ThisWorkbook.Worksheets("Saidas")
    PrepararCelulasF2 WS
    Set operacao = New clsOperacaoPlanilha
    operacao.Iniciar ThisWorkbook
    operacao.RegistrarDestino WS.Range("A4:D4")
    operacao.MarcarBaixaIniciada
    ExigirF2 Not operacao.Recuperar(diagnostico), "Recusa apos iniciar baixa"
    operacao.Finalizar
    Set operacao = New clsOperacaoPlanilha
    operacao.Iniciar ThisWorkbook
    operacao.RegistrarDestino WS.Range("A4:D4")
    operacao.MarcarSalvamentoIniciado
    ExigirF2 Not operacao.Recuperar(diagnostico), "Recusa apos iniciar Save"
    operacao.Finalizar

    ExigirF2 GetStrippedValue("=5") = "5", "Limite numerico"
    ExigirF2 GetStrippedValue("=""=ABC""") = "=ABC", "Texto literal com igual"
    WS.Range("F4").FormatConditions.Delete
    WS.Range("F4").FormatConditions.Add xlCellValue, xlNotBetween, "5", "10"
    For Each fluxo In Array(4, 5, 7, 10, 11)
        WS.Range("F4").Value2 = fluxo
        Call ExigirF2((ActiveCondition(WS.Range("F4")) > 0) = (fluxo < 5 Or fluxo > 10), "xlNotBetween")
    Next fluxo
    AuditoriaFase2_Executar = "PASS: " & mVerificacoesF2 & " verificacoes VBA dos auxiliares; eventos dos formularios nao executados."
    Exit Function
Falhou:
    n = Err.Number: d = Err.Description
    If Not operacao Is Nothing Then operacao.Finalizar
    Err.Raise n, "AuditoriaFase2_Executar", d
End Function

Private Sub PrepararCelulasF2(ByVal WS As Worksheet)
    Dim formatoPadrao As String
    WS.Range("A4:D4").ClearContents
    WS.Range("A4:D4").ClearFormats
    formatoPadrao = CStr(WS.Range("B4").NumberFormat)
    WS.Range("B4").NumberFormat = "@"
    WS.Range("B4").Value2 = "=TEXTO_FICTICIO"
    WS.Range("B4").NumberFormat = formatoPadrao
    WS.Range("C4").Formula = "=1+2"
    WS.Range("D4").NumberFormat = "@"
    WS.Range("D4").Value2 = "00000000000000000000000000000000000000000001"
    WS.Range("D4").Errors(xlNumberAsText).Ignore = True
End Sub

Private Sub TestarEtapaF2(ByVal nomes As Variant, ByVal falha As Long)
    Dim i As Long, WS As Worksheet, operacao As clsOperacaoPlanilha
    Dim diagnostico As String, eventos As Boolean, tela As Boolean
    Dim alertas As Boolean, calculo As XlCalculation, n As Long, d As String
    eventos = Application.EnableEvents: tela = Application.ScreenUpdating
    alertas = Application.DisplayAlerts: calculo = Application.Calculation
    On Error GoTo Falhou
    Set operacao = New clsOperacaoPlanilha
    operacao.Iniciar ThisWorkbook
    For i = LBound(nomes) To UBound(nomes)
        Set WS = ThisWorkbook.Worksheets(CStr(nomes(i)))
        PrepararCelulasF2 WS
        operacao.RegistrarDestino WS.Range("A4:D4")
    Next i
    For i = LBound(nomes) To UBound(nomes)
        Set WS = ThisWorkbook.Worksheets(CStr(nomes(i)))
        WS.Range("A4").Value2 = "NF_FICTICIA_1"
        WS.Range("B4").Value2 = "ALTERADO"
        WS.Range("C4").Value2 = 99
        WS.Range("D4").Value2 = "ALTERADO"
        If i = falha Then Err.Raise vbObjectError + 2210, , "Falha artificial controlada."
    Next i
    Err.Raise vbObjectError + 2211, , "Falha artificial nao atingida."
Falhou:
    n = Err.Number: d = Err.Description
    If n <> vbObjectError + 2210 Then
        If Not operacao Is Nothing Then operacao.Finalizar
        Err.Raise n, "TestarEtapaF2", d
    End If
    'Resume desativa o tratador antes de verificar a compensacao.
    Resume Conferir
Conferir:
    On Error GoTo FalhaConferencia
    ExigirF2 operacao.Recuperar(diagnostico), "Compensacao por etapa"
    operacao.Finalizar
    ExigirF2 Application.EnableEvents = eventos, "Restaurar eventos"
    ExigirF2 Application.ScreenUpdating = tela, "Restaurar tela"
    ExigirF2 Application.DisplayAlerts = alertas, "Restaurar alertas"
    ExigirF2 Application.Calculation = calculo, "Restaurar calculo"
    For i = LBound(nomes) To UBound(nomes)
        Set WS = ThisWorkbook.Worksheets(CStr(nomes(i)))
        ExigirF2 IsEmpty(WS.Range("A4").Value2), "Sem gravacao parcial"
        ExigirF2 WS.Range("B4").Value2 = "=TEXTO_FICTICIO" And Not WS.Range("B4").HasFormula, "Texto nao vira formula"
        ExigirF2 WS.Range("C4").Formula = "=1+2", "Formula restaurada"
        ExigirF2 WS.Range("D4").Value2 = "00000000000000000000000000000000000000000001", "Chave restaurada"
        ExigirF2 WS.Range("D4").Errors(xlNumberAsText).Ignore, "Metadado numero como texto restaurado"
    Next i
    Set operacao = New clsOperacaoPlanilha
    operacao.Iniciar ThisWorkbook
    For i = LBound(nomes) To UBound(nomes)
        Set WS = ThisWorkbook.Worksheets(CStr(nomes(i)))
        operacao.RegistrarDestino WS.Range("A4:D4")
        WS.Range("A4").Value2 = "NF_FICTICIA_1"
    Next i
    operacao.Confirmar
    For i = LBound(nomes) To UBound(nomes)
        Set WS = ThisWorkbook.Worksheets(CStr(nomes(i)))
        ExigirF2 Application.CountIf(WS.Range("A4:A10"), "NF_FICTICIA_1") = 1, "Nova tentativa sem duplicacao"
        PrepararCelulasF2 WS
    Next i
    Exit Sub
FalhaConferencia:
    n = Err.Number: d = Err.Description
    If Not operacao Is Nothing Then operacao.Finalizar
    Err.Raise n, "TestarEtapaF2", d
End Sub

Private Sub ExigirF2(ByVal condicao As Boolean, ByVal nome As String)
    If Not condicao Then Err.Raise vbObjectError + 2220, , "FAIL: " & nome
    mVerificacoesF2 = mVerificacoesF2 + 1
End Sub

'Teste de regressao do ciclo de vida; PREPARADO, NAO EXECUTADO.
Public Function AuditoriaFase3_CicloVida() As String
    Dim op As clsOperacaoPlanilha, diagnostico As String
    If InStr(1, CreateObject("Scripting.FileSystemObject").GetFile(ThisWorkbook.FullName).ShortPath, CreateObject("Scripting.FileSystemObject").GetFolder(Environ$("TEMP")).ShortPath & "\", vbTextCompare) <> 1 Then
        Err.Raise vbObjectError + 2230, , "Exige fixture em TEMP."
    End If
    If Left$(ThisWorkbook.name, 18) <> "Fixture_Auditoria_" Then
        Err.Raise vbObjectError + 2231, , "Nome da fixture invalido."
    End If
    If ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then
        Err.Raise vbObjectError + 2232, , "Marcador ficticio ausente."
    End If
    Set op = New clsOperacaoPlanilha
    ExigirF2 Not op.Recuperar(diagnostico), "Nao recuperar contexto nunca iniciado"
    op.Iniciar ThisWorkbook
    op.Finalizar
    ExigirF2 Not op.Recuperar(diagnostico), "Nao recuperar contexto finalizado"
    ExigirF2 InicioRepetidoRecusadoF3(op), "Nao reutilizar contexto vazio finalizado"
    Set op = New clsOperacaoPlanilha
    op.Iniciar ThisWorkbook
    op.Confirmar
    ExigirF2 InicioRepetidoRecusadoF3(op), "Nao reutilizar contexto vazio confirmado"
    AuditoriaFase3_CicloVida = "PASS: ciclo de vida verificado; eventos fiscais nao executados."
End Function

Private Function InicioRepetidoRecusadoF3(ByVal op As clsOperacaoPlanilha) As Boolean
    On Error GoTo Recusado
    op.Iniciar ThisWorkbook
    op.Finalizar
    Exit Function
Recusado:
    InicioRepetidoRecusadoF3 = (Err.Number = vbObjectError + 2100)
End Function

'Fase 4: casos reais ausentes na suite anterior; exclusivamente fixture ficticia.
Public Function AuditoriaFase4_IntegridadeComplementar() As String
    Dim WS As Worksheet, wsBloqueada As Worksheet, op As clsOperacaoPlanilha, diagnostico As String
    Dim CHAVE As String, n As Long, d As String, protegidaPeloTeste As Boolean
    Dim antes As Long
    If InStr(1, CreateObject("Scripting.FileSystemObject").GetFile(ThisWorkbook.FullName).ShortPath, CreateObject("Scripting.FileSystemObject").GetFolder(Environ$("TEMP")).ShortPath & "\", vbTextCompare) <> 1 Then Err.Raise vbObjectError + 2240, , "Exige TEMP."
    If Left$(ThisWorkbook.name, 18) <> "Fixture_Auditoria_" Then Err.Raise vbObjectError + 2241, , "Exige fixture."
    If ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then Err.Raise vbObjectError + 2242, , "Exige marcador ficticio."
    Set WS = ThisWorkbook.Worksheets("Saidas")
    Set wsBloqueada = ThisWorkbook.Worksheets("Entradas")
    If wsBloqueada.ProtectContents Then Err.Raise vbObjectError + 2243, , "Fixture ja protegida; teste recusado."
    On Error GoTo Falhou
    antes = mVerificacoesF2
    WS.Range("A20:D20").ClearContents
    WS.Range("A20").Value2 = "ANTES"
    Set op = New clsOperacaoPlanilha
    op.Iniciar ThisWorkbook
    op.RegistrarDestino WS.Range("A20")
    WS.Range("A20").Value2 = "INTERMEDIARIO"
    op.RegistrarDestino WS.Range("A20")
    WS.Range("A20").Value2 = "DEPOIS"
    ExigirF2 op.Recuperar(diagnostico), "Recuperacao sobreposta"
    ExigirF2 WS.Range("A20").Value2 = "ANTES", "Ordem inversa dos snapshots"
    op.Finalizar
    CHAVE = "00000000000000000000000000000000000000000001"
    wsBloqueada.Range("B20").NumberFormat = "0.00"
    wsBloqueada.Range("B20").Value2 = 12.5
    WS.Range("C20").NumberFormat = "@"
    WS.Range("C20").Value2 = CHAVE
    WS.Range("C20").Errors(xlNumberAsText).Ignore = True
    WS.Range("D20").Formula = "=1+2"
    Set op = New clsOperacaoPlanilha
    op.Iniciar ThisWorkbook
    op.RegistrarDestino WS.Range("A20")
    op.RegistrarDestino wsBloqueada.Range("B20")
    op.RegistrarDestino WS.Range("C20:D20")
    WS.Range("A20").Value2 = "ALTERADO"
    wsBloqueada.Range("B20").Value2 = 99
    wsBloqueada.Range("B20").NumberFormat = "0.0000"
    WS.Range("C20").Value2 = "ALTERADO"
    WS.Range("D20").Value2 = 99
    wsBloqueada.Range("B20").Locked = False
    wsBloqueada.Range("B20").Locked = True
    wsBloqueada.Protect Password:="F4QA", AllowFormattingCells:=True
    protegidaPeloTeste = True
    ExigirF2 Not op.Recuperar(diagnostico), "Falha individual reportada"
    ExigirF2 InStr(1, diagnostico, "Falhas na compensacao: 1.", vbBinaryCompare) = 1, "Diagnostico individual"
    ExigirF2 WS.Range("A20").Value2 = "ANTES", "Demais celulas continuam restauradas"
    ExigirF2 WS.Range("C20").Value2 = CHAVE And WS.Range("D20").Formula = "=1+2", "Chave e formula apos falha individual"
    ExigirF2 wsBloqueada.Range("B20").Value2 = 99, "Nao anunciar recuperacao do destino bloqueado"
    wsBloqueada.Unprotect Password:="F4QA"
    protegidaPeloTeste = False
    ExigirF2 op.Recuperar(diagnostico), "Nova recuperacao antes de Delete Save"
    ExigirF2 wsBloqueada.Range("B20").Value2 = 12.5 And wsBloqueada.Range("B20").NumberFormat = "0.00", "Valor e NumberFormat restaurados"
    ExigirF2 VarType(WS.Range("C20").Value2) = vbString And Len(WS.Range("C20").Value2) = 44, "Tipo e comprimento da chave"
    ExigirF2 WS.Range("C20").NumberFormat = "@" And WS.Range("C20").Errors(xlNumberAsText).Ignore, "Formato e Ignore restaurados"
    op.Finalizar
    wsBloqueada.Range("B20").Locked = True
    AuditoriaFase4_IntegridadeComplementar = "PASS: " & (mVerificacoesF2 - antes) & " verificacoes VBA complementares; falha individual reconhecida sem atomicidade."
    Exit Function
Falhou:
    n = Err.Number: d = Err.Description
    On Error Resume Next
    If protegidaPeloTeste Then wsBloqueada.Unprotect Password:="F4QA"
    If Not op Is Nothing Then op.Finalizar
    On Error GoTo 0
    Err.Raise n, "AuditoriaFase4_IntegridadeComplementar", d
End Function
