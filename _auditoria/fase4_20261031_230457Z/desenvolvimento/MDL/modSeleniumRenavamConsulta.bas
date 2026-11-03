Attribute VB_Name = "modSeleniumRenavamConsulta"
Option Explicit

' Modulo: modSeleniumRenavamConsulta
' Objetivo: navegar no menu RENAVAM e preencher UF/placa.

Private g_drvRenavam As Object
Private Const LOGIN_TIMEOUT_SECONDS As Long = 300
Private Const PAGE_TIMEOUT_SECONDS As Long = 60
Private Const READY_INTERVAL_MS As Long = 250

' Rotina principal (chame a partir do CommandButtonDetranNetBIN)
Public Sub AbrirConsultaRenavam(ByVal uf As String, ByVal PLACA As String)
    Dim drv As Object
    Dim ufNorm As String
    Dim placaNorm As String
    Dim sessaoAtiva As Boolean
    Dim loginRealizado As Boolean

    ufNorm = UCase$(Trim$(uf))
    placaNorm = UCase$(Trim$(PLACA))

    If Len(ufNorm) = 0 Or Len(placaNorm) = 0 Then
        MsgBox "Informe UF e placa antes de continuar.", vbExclamation, "Consulta RENAVAM"
        Exit Sub
    End If

    On Error GoTo FAIL
    Set drv = GetDetranDriverRenavam()

    sessaoAtiva = SessaoDetranProvavelmenteAutenticadaRenavam(drv)

    If sessaoAtiva Then
        WriteLog "AbrirConsultaRenavam: sessao existente detectada; pulando fluxo de login."
    Else
        WriteLog "AbrirConsultaRenavam: nenhuma sessao ativa; navegando para tela de login."
        drv.Get URL_LOGIN
        WaitForUrlRenavam drv, URL_HOME, LOGIN_TIMEOUT_SECONDS
        If UrlAtualCorrespondeRenavam(drv, URL_HOME) Then
            loginRealizado = True
        ElseIf Not EstaPaginaDeLoginRenavam(drv) Then
            loginRealizado = True
        End If
    End If

    If Not NavegarMenuRenavam(drv) Then
        If EstaPaginaDeLoginRenavam(drv) Then
            WriteLog "AbrirConsultaRenavam: redirecionado para login; aguardando autenticacao."
            drv.Get URL_LOGIN
            WaitForUrlRenavam drv, URL_HOME, LOGIN_TIMEOUT_SECONDS
            If UrlAtualCorrespondeRenavam(drv, URL_HOME) Then
                loginRealizado = True
            ElseIf Not EstaPaginaDeLoginRenavam(drv) Then
                loginRealizado = True
            End If

            If Not NavegarMenuRenavam(drv) Then
                MsgBox "Nao foi possivel abrir o menu RENAVAM.", vbExclamation, "Consulta RENAVAM"
                Exit Sub
            End If
        Else
            MsgBox "Nao foi possivel abrir o menu RENAVAM.", vbExclamation, "Consulta RENAVAM"
            Exit Sub
        End If
    End If

    If Not SelecionarUfRenavam(drv, ufNorm) Then
        MsgBox "Nao foi possivel selecionar a UF.", vbExclamation, "Consulta RENAVAM"
        Exit Sub
    End If

    If Not PreencherPlacaRenavam(drv, placaNorm) Then
        MsgBox "Nao foi possivel preencher a placa.", vbExclamation, "Consulta RENAVAM"
        Exit Sub
    End If

    Exit Sub

FAIL:
    WriteLog "AbrirConsultaRenavam: erro " & Err.Number & " - " & Err.Description
    EncerrarSessaoRenavam
End Sub

' Handler para o botao. Chame este metodo no click do CommandButtonDetranNetBIN.
Public Sub CommandButtonDetranNetBIN_Handler()
    Dim uf As String
    Dim PLACA As String

    uf = InputBox("Informe a UF:", "Consulta RENAVAM")
    If Len(Trim$(uf)) = 0 Then Exit Sub

    PLACA = InputBox("Informe a placa:", "Consulta RENAVAM")
    If Len(Trim$(PLACA)) = 0 Then Exit Sub

    AbrirConsultaRenavam uf, PLACA
End Sub

Public Sub EncerrarSessaoRenavam()
    On Error Resume Next
    If Not g_drvRenavam Is Nothing Then g_drvRenavam.Quit
    Set g_drvRenavam = Nothing
    On Error GoTo 0
End Sub

Private Function GetDetranDriverRenavam() As Object
    Dim drv As Object

    If TryReuseDetranDriverRenavam(drv) Then
        Set GetDetranDriverRenavam = drv
        Exit Function
    End If

    On Error GoTo FAIL
    EncerrarSessaoRenavam
    WriteLog "GetDetranDriverRenavam: criando nova instancia do ChromeDriver."
    Set g_drvRenavam = CreateObject(PROGID_CHROME)
    ConfigureChromedriverExecutable g_drvRenavam
    ConfigureAndStartChrome g_drvRenavam, CHROME_BINARY_PATH
    Set GetDetranDriverRenavam = g_drvRenavam
    Exit Function

FAIL:
    Dim errMsg As String
    errMsg = Err.Description
    EncerrarSessaoRenavam
    On Error GoTo 0
    Err.Clear
    Err.Raise vbObjectError + 514, "GetDetranDriverRenavam", "Falha ao iniciar ChromeDriver: " & errMsg
End Function

Private Function TryReuseDetranDriverRenavam(ByRef drv As Object) As Boolean
    Dim ignoreUrl As String

    On Error GoTo CLEANUP

    If g_drvRenavam Is Nothing Then Exit Function

    ignoreUrl = CStr(g_drvRenavam.url)
    Set drv = g_drvRenavam
    TryReuseDetranDriverRenavam = True
    WriteLog "TryReuseDetranDriverRenavam: reaproveitando instancia ativa do ChromeDriver."
    Exit Function

CLEANUP:
    WriteLog "TryReuseDetranDriverRenavam: instancia existente invalida; sera descartada."
    Err.Clear
    EncerrarSessaoRenavam
End Function

Private Function SessaoDetranProvavelmenteAutenticadaRenavam(ByVal drv As Object) As Boolean
    On Error GoTo CLEANUP

    Dim currentUrl As String
    currentUrl = LCase$(CStr(drv.url))

    If Len(currentUrl) = 0 Then Exit Function
    If InStr(currentUrl, "controleacesso/login") > 0 Then Exit Function
    If InStr(currentUrl, "logout") > 0 Then Exit Function

    If InStr(currentUrl, "arearestrita") > 0 Then
        SessaoDetranProvavelmenteAutenticadaRenavam = True
        Exit Function
    End If

    If InStr(currentUrl, "consultaveiculo") > 0 Then
        SessaoDetranProvavelmenteAutenticadaRenavam = True
        Exit Function
    End If

    Exit Function

CLEANUP:
    SessaoDetranProvavelmenteAutenticadaRenavam = False
End Function

Private Sub WaitForUrlRenavam(ByVal drv As Object, ByVal targetUrl As String, ByVal timeoutSeconds As Long)
    Dim startedAt As Single
    startedAt = Timer

    Do
        Dim currentUrl As String
        On Error Resume Next
        currentUrl = CStr(drv.url)
        On Error GoTo 0

        If Len(targetUrl) = 0 Then Exit Do
        If LCase$(Left$(currentUrl, Len(targetUrl))) = LCase$(targetUrl) Then Exit Do
        If LCase$(currentUrl) = LCase$(targetUrl) Then Exit Do
        If SegundosDecorridos(startedAt) > timeoutSeconds Then Exit Do

        SleepMs READY_INTERVAL_MS
        DoEvents
    Loop
End Sub

Private Function UrlAtualCorrespondeRenavam(ByVal drv As Object, ByVal targetUrl As String) As Boolean
    Dim currentUrl As String

    On Error Resume Next
    currentUrl = CStr(drv.url)
    On Error GoTo 0

    If Len(targetUrl) = 0 Then Exit Function
    If Len(currentUrl) = 0 Then Exit Function

    If LCase$(Left$(currentUrl, Len(targetUrl))) = LCase$(targetUrl) Then
        UrlAtualCorrespondeRenavam = True
    ElseIf LCase$(currentUrl) = LCase$(targetUrl) Then
        UrlAtualCorrespondeRenavam = True
    End If
End Function

Private Function EstaPaginaDeLoginRenavam(ByVal drv As Object) As Boolean
    On Error GoTo CLEANUP

    Dim currentUrl As String
    currentUrl = LCase$(CStr(drv.url))

    If InStr(currentUrl, "controleacesso/login") > 0 Then
        EstaPaginaDeLoginRenavam = True
    ElseIf InStr(currentUrl, "login.asp") > 0 Then
        EstaPaginaDeLoginRenavam = True
    End If
    Exit Function

CLEANUP:
    EstaPaginaDeLoginRenavam = False
End Function

Private Function NavegarMenuRenavam(ByVal drv As Object) As Boolean
    On Error GoTo CLEANUP

    SwitchToDefaultContentSafe drv

    drv.Get URL_HOME
    WaitDocumentReadyRenavam drv, PAGE_TIMEOUT_SECONDS

    If SwitchToFrameByNameSafe(drv, "top") Then
        If Not ClickByIdSafe(drv, "RENAVAM") Then GoTo CLEANUP
    End If

    SwitchToDefaultContentSafe drv
    If Not SwitchToFrameByNameSafe(drv, "contents") Then GoTo CLEANUP

    SleepMs 300
    If Not ClickByIdSafe(drv, "HGRENAVAMRENAVAM") Then GoTo CLEANUP
    SleepMs 300
    If Not AbrirConsultaOutrasUfs(drv) Then GoTo CLEANUP

    WaitDocumentReadyRenavam drv, PAGE_TIMEOUT_SECONDS
    NavegarMenuRenavam = True
    Exit Function

CLEANUP:
    WriteLog "NavegarMenuRenavam: falha ao navegar no menu RENAVAM."
    NavegarMenuRenavam = False
End Function

Private Function AbrirConsultaOutrasUfs(ByVal drv As Object) As Boolean
    On Error GoTo CLEANUP

    If Not SwitchToFrameByNameSafe(drv, "contents") Then GoTo CLEANUP

    drv.ExecuteScript "if (typeof checkExpandNew === 'function') { checkExpandNew('RENAVAM','RENAVAM','consultaufs'); }"
    If Err.Number = 0 Then
        WriteLog "AbrirConsultaOutrasUfs: checkExpandNew executado."
    Else
        Err.Clear
    End If

    If Not ClickByIdSafe(drv, "HGRENAVAMRENAVAM") Then
        WriteLog "AbrirConsultaOutrasUfs: falha ao clicar no menu RENAVAM."
    End If
    SleepMs 300

    If Not ClickByIdSafe(drv, "inRENAVAMRENAVAMconsultaufs") Then
        WriteLog "AbrirConsultaOutrasUfs: click direto falhou, tentando via JS no link."
        drv.ExecuteScript "var el=document.getElementById('inRENAVAMRENAVAMconsultaufs'); if(el){el.click();}"
        If Err.Number <> 0 Then Err.Clear
    End If

    drv.ExecuteScript "if (typeof funcaodestino === 'function') { " & _
                      "funcaodestino('RENAVAM/RENAVAM/consultaufs.asp?DetranNet_Sistema=RENAVAM&DetranNet_Modulo=RENAVAM'); }"
    If Err.Number = 0 Then
        WriteLog "AbrirConsultaOutrasUfs: funcaodestino executado no frame contents."
    Else
        WriteLog "AbrirConsultaOutrasUfs: falha ao executar funcaodestino (" & Err.Description & ")."
        Err.Clear
    End If

    ' Aguarda a pagina de consulta estar pronta (campo UF)
    Dim startedAt As Single
    startedAt = Timer

    Do
        SwitchToDefaultContentSafe drv
        If SwitchToFrameByNameSafe(drv, "contents") Then
            WriteLog "AbrirConsultaOutrasUfs: etapa executada; conteúdo de dados omitido do log."
        End If
        SwitchToDefaultContentSafe drv
        If SwitchToFrameByNameSafe(drv, "form") Then
            WriteLog "AbrirConsultaOutrasUfs: etapa executada; conteúdo de dados omitido do log."
            If InStr(1, LCase$(CStr(drv.url)), "consultaufs.asp", vbTextCompare) > 0 Then
                SwitchToDefaultContentSafe drv
                If EnsureConsultaFrameRenavam(drv) Then
                    AbrirConsultaOutrasUfs = True
                    Exit Function
                End If
            End If
        End If
        SwitchToDefaultContentSafe drv
        If EnsureConsultaFrameRenavam(drv) Then
            AbrirConsultaOutrasUfs = True
            Exit Function
        End If

        If SegundosDecorridos(startedAt) > PAGE_TIMEOUT_SECONDS Then Exit Do
        SleepMs READY_INTERVAL_MS
        DoEvents
    Loop

CLEANUP:
    AbrirConsultaOutrasUfs = False
End Function

Private Function WaitDocumentReadyRenavam(ByVal drv As Object, Optional ByVal timeoutSeconds As Long = PAGE_TIMEOUT_SECONDS) As Boolean
    WaitDocumentReadyRenavam = WaitForReadyStateComplete(drv, timeoutSeconds, READY_INTERVAL_MS)
End Function

Private Function ClickByIdSafe(ByVal drv As Object, ByVal elementId As String) As Boolean
    On Error GoTo CLEANUP

    Dim el As Object
    Set el = drv.FindElementById(elementId)
    If el Is Nothing Then GoTo CLEANUP

    el.Click
    ClickByIdSafe = True
    Exit Function

CLEANUP:
    On Error Resume Next
    If Not drv Is Nothing Then
        drv.ExecuteScript "var el=document.getElementById('" & elementId & "'); if(el){el.click();}"
        If Err.Number = 0 Then
            ClickByIdSafe = True
            Exit Function
        End If
    End If
    On Error GoTo 0
    ClickByIdSafe = False
End Function

Private Function SelecionarUfRenavam(ByVal drv As Object, ByVal uf As String) As Boolean
    On Error GoTo CLEANUP

    If Not EnsureConsultaFrameRenavam(drv) Then GoTo CLEANUP

    Dim sel As Object
    Dim opts As Object
    Dim opt As Object

    Set sel = drv.FindElementByName("UF")
    If sel Is Nothing Then GoTo CLEANUP

    Set opts = sel.FindElementsByTag("option")
    If opts Is Nothing Then GoTo CLEANUP

    For Each opt In opts
        If UCase$(Trim$(opt.GetAttribute("value"))) = uf Or UCase$(Trim$(opt.text)) = uf Then
            opt.Click
            SelecionarUfRenavam = True
            Exit Function
        End If
    Next opt

CLEANUP:
    SelecionarUfRenavam = False
End Function

Private Function PreencherPlacaRenavam(ByVal drv As Object, ByVal PLACA As String) As Boolean
    On Error GoTo CLEANUP

    If Not EnsureConsultaFrameRenavam(drv) Then GoTo CLEANUP

    Dim el As Object
    Set el = FindElementByNameSafe(drv, "Chave")
    If el Is Nothing Then
        Set el = FindElementByNameSafe(drv, "chave")
    End If
    If el Is Nothing Then GoTo CLEANUP

    On Error Resume Next
    el.Clear
    On Error GoTo CLEANUP

    el.SendKeys PLACA
    PreencherPlacaRenavam = True
    Exit Function

CLEANUP:
    PreencherPlacaRenavam = False
End Function

Private Function EnsureConsultaFrameRenavam(ByVal drv As Object) As Boolean
    Dim el As Object
    On Error GoTo CLEANUP

    SwitchToDefaultContentSafe drv

    If SwitchToFrameByNameSafe(drv, "form") Then
        Set el = FindElementByNameSafe(drv, "UF")
        If Not el Is Nothing Then
            WriteLog "EnsureConsultaFrameRenavam: UF encontrado no frame 'form'."
            EnsureConsultaFrameRenavam = True
            Exit Function
        End If
    End If

    SwitchToDefaultContentSafe drv
    If SwitchToFrameByNameSafe(drv, "Consulta") Then
        Dim elConsulta As Object
        Set elConsulta = FindElementByNameSafe(drv, "UF")
        If Not elConsulta Is Nothing Then
            WriteLog "EnsureConsultaFrameRenavam: UF encontrado no frame 'Consulta'."
            EnsureConsultaFrameRenavam = True
            Exit Function
        End If
    End If
    Set el = FindElementByNameSafe(drv, "UF")
    If Not el Is Nothing Then
        WriteLog "EnsureConsultaFrameRenavam: UF encontrado no documento principal."
        EnsureConsultaFrameRenavam = True
        Exit Function
    End If

    Dim frameById As Object
    Set frameById = FindElementByIdSafe(drv, "frame1")
    If Not frameById Is Nothing Then
        drv.SwitchToFrame frameById
        Set el = FindElementByNameSafe(drv, "UF")
        If Not el Is Nothing Then
            EnsureConsultaFrameRenavam = True
            Exit Function
        End If
        SwitchToDefaultContentSafe drv
    End If

    Dim frames As Object
    Dim i As Long
    Dim frameEl As Object

    Set frames = drv.FindElementsByTag("iframe")
    If Not frames Is Nothing Then
        For i = 0 To frames.count - 1
            Set frameEl = FrameItemSafe(frames, i)
            If Not frameEl Is Nothing Then
                drv.SwitchToFrame frameEl
                Set el = FindElementByNameSafe(drv, "UF")
                If Not el Is Nothing Then
                    WriteLog "EnsureConsultaFrameRenavam: UF encontrado em iframe (index " & i & ")."
                    EnsureConsultaFrameRenavam = True
                    Exit Function
                End If
            End If
            SwitchToDefaultContentSafe drv
        Next i
    End If

    Set frames = drv.FindElementsByTag("frame")
    If Not frames Is Nothing Then
        For i = 0 To frames.count - 1
            Set frameEl = FrameItemSafe(frames, i)
            If Not frameEl Is Nothing Then
                drv.SwitchToFrame frameEl
                Set el = FindElementByNameSafe(drv, "UF")
                If Not el Is Nothing Then
                    WriteLog "EnsureConsultaFrameRenavam: UF encontrado em frame (index " & i & ")."
                    EnsureConsultaFrameRenavam = True
                    Exit Function
                End If
            End If
            SwitchToDefaultContentSafe drv
        Next i
    End If

CLEANUP:
    EnsureConsultaFrameRenavam = False
End Function

Private Function FrameItemSafe(ByVal coll As Object, ByVal zeroIdx As Long) As Object
    On Error GoTo TRY_ONE_BASE
    Set FrameItemSafe = coll.item(zeroIdx)
    Exit Function
TRY_ONE_BASE:
    On Error GoTo FAIL
    Set FrameItemSafe = coll.item(zeroIdx + 1)
    Exit Function
FAIL:
    Set FrameItemSafe = Nothing
End Function

Private Sub SwitchToDefaultContentSafe(ByVal drv As Object)
    On Error Resume Next
    drv.SwitchToDefaultContent
    On Error GoTo 0
End Sub

Private Function FindElementByNameSafe(ByVal drv As Object, ByVal nameValue As String) As Object
    On Error Resume Next
    Set FindElementByNameSafe = drv.FindElementByName(nameValue)
    On Error GoTo 0
End Function

Private Function FindElementByIdSafe(ByVal drv As Object, ByVal idValue As String) As Object
    On Error Resume Next
    Set FindElementByIdSafe = drv.FindElementById(idValue)
    On Error GoTo 0
End Function

Private Function FindElementByXPathSafe(ByVal drv As Object, ByVal xp As String) As Object
    On Error Resume Next
    Set FindElementByXPathSafe = drv.FindElementByXPath(xp)
    On Error GoTo 0
End Function

Private Function SwitchToFrameByNameSafe(ByVal drv As Object, ByVal frameName As String) As Boolean
    On Error GoTo CLEANUP

    Dim frameEl As Object
    Set frameEl = FindElementByXPathSafe(drv, "//frame[@name='" & frameName & "']")
    If frameEl Is Nothing Then
        Set frameEl = FindElementByXPathSafe(drv, "//iframe[@name='" & frameName & "']")
    End If
    If frameEl Is Nothing Then GoTo CLEANUP

    drv.SwitchToFrame frameEl
    SwitchToFrameByNameSafe = True
    Exit Function

CLEANUP:
    SwitchToFrameByNameSafe = False
End Function

' Configura o binario e inicia o Chrome pelo SeleniumBasic
Private Sub ConfigureAndStartChrome(ByVal drv As Object, ByVal binaryPath As String)
    On Error Resume Next
    If Len(binaryPath) > 0 Then
        If Len(Dir$(binaryPath)) > 0 Then drv.SetBinary binaryPath
    End If
    drv.AddArgument "--no-first-run"
    drv.AddArgument "--no-default-browser-check"
    drv.AddArgument "--disable-gpu"
    drv.AddArgument "--disable-extensions"
    drv.AddArgument "--lang=pt-BR"
    Err.Clear
    drv.Start "chrome"
    ' Nao suprimir falha do Start: sessao invalida gera erros crípticos (ex.: KeyNotFoundError "status") em comandos posteriores
    If Err.Number <> 0 Then
        Dim startErrNumber As Long, startErrDesc As String
        startErrNumber = Err.Number
        startErrDesc = Err.Description
        On Error GoTo 0
        Err.Raise IIf(startErrNumber <> 0, startErrNumber, vbObjectError + 516), "ConfigureAndStartChrome", "Falha ao iniciar sessao do ChromeDriver: " & startErrDesc
    End If
    On Error GoTo 0
End Sub




