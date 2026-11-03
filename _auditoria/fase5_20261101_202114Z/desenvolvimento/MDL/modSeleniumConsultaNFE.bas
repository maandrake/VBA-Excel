Attribute VB_Name = "modSeleniumConsultaNFE"
Option Explicit

Private Const SW_RESTORE As Long = 9
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long

Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetWindowTextLengthW Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowTextW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal cch As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long

Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal lpDestination As LongPtr, ByVal lpSource As LongPtr, ByVal cbCopy As LongPtr)

#If Win64 Then
Private m_foundWindowNFE As LongPtr
#Else
Private m_foundWindowNFE As Long
#End If
Private m_targetCaptionNFE As String

' ==============================================================
'  Módulo especializado em obter resumo para Nota Fiscal Eletrônica
'  Formato esperado na TextBoxDadosNFE:
'    "Marca/Modelo <valor> Placa <valor> Renavam <valor> Chassi <valor> Ano/Mod <valor> Cor <valor>"
' ==============================================================

Private Const PROGID_CHROME As String = "Selenium.ChromeDriver"
Private Const CHROME_BINARY_PATH As String = "C:\Program Files\Google\Chrome\Application\chrome.exe"
Private Const CHROMEDRIVER_EXE As String = "C:\Program Files\SeleniumBasic\chromedriver.exe"
Private Const CHROMEDRIVER_DIR As String = "C:\Program Files\SeleniumBasic"
Private Const URL_LOGIN As String = "https://sistema.detrannet.sc.gov.br/controleacesso/login.asp"
Private Const URL_HOME As String = "https://sistema.detrannet.sc.gov.br/arearestrita/tela_principal.asp"
Private Const URL_CONSULTA_BASE As String = "https://sistema.detrannet.sc.gov.br/arearestrita/restrito/aguarde.asp?destino=consultaveiculo2.asp?placa="
Private Const LOGIN_TIMEOUT_SECONDS As Long = 300
Private Const PAGE_TIMEOUT_SECONDS As Long = 60
Private Const READY_INTERVAL_MS As Long = 250

Private g_drvNFE As Object
Private g_ultimaMensagemErroNFE As String

Public Function ConsultarDetranRetorno(ByVal PLACA As String) As Collection
    Dim drv As Object
    Dim sessaoAtiva As Boolean
    Dim dossie As Object
    Dim dados As Collection
    Dim loginRealizado As Boolean
    Dim placaSanitizada As String
    Dim consultaUrl As String

    placaSanitizada = SanitizePlacaNFE(PLACA)
    If Len(placaSanitizada) = 0 Then
        LogNFE "ConsultarDetranRetorno: placa informada inválida após sanitização."
        Exit Function
    End If

    If StrComp(placaSanitizada, UCase$(Trim$(PLACA)), vbBinaryCompare) <> 0 Then
        LogNFE "ConsultarDetranRetorno: etapa executada; conteúdo de dados omitido do log."
    End If

    If Len(placaSanitizada) <> 7 Then
        LogNFE "ConsultarDetranRetorno: etapa executada; conteúdo de dados omitido do log."
    End If

    consultaUrl = BuildConsultaUrlNFE(placaSanitizada)
    If Len(consultaUrl) = 0 Then Exit Function
    PLACA = placaSanitizada

    g_ultimaMensagemErroNFE = ""

    On Error GoTo FAIL

    Set drv = GetDetranDriverNFE()

    sessaoAtiva = SessaoDetranProvavelmenteAutenticadaNFE(drv)

    If sessaoAtiva Then
        LogNFE "ConsultarDetranRetorno: sessão existente detectada; pulando fluxo de login."
    Else
        LogNFE "ConsultarDetranRetorno: nenhuma sessão ativa; navegando para tela de login."
        drv.Get URL_LOGIN
        WaitForUrlNFE drv, URL_HOME, LOGIN_TIMEOUT_SECONDS
        If UrlAtualCorrespondeNFE(drv, URL_HOME) Then
            loginRealizado = True
        ElseIf Not EstaPaginaDeLoginNFE(drv) Then
            loginRealizado = True
        End If
    End If

    LogNFE "ConsultarDetranRetorno: etapa executada; conteúdo de dados omitido do log."
    drv.Get consultaUrl
    WaitDocumentReadyNFE drv, PAGE_TIMEOUT_SECONDS

    If EstaPaginaDeLoginNFE(drv) Then
        LogNFE "ConsultarDetranRetorno: redirecionado para tela de login; aguardando autenticação."
        drv.Get URL_LOGIN
        WaitForUrlNFE drv, URL_HOME, LOGIN_TIMEOUT_SECONDS
        If UrlAtualCorrespondeNFE(drv, URL_HOME) Then
            loginRealizado = True
        ElseIf Not EstaPaginaDeLoginNFE(drv) Then
            loginRealizado = True
        End If
        LogNFE "ConsultarDetranRetorno: tentando novamente a URL de consulta."
        drv.Get consultaUrl
        WaitDocumentReadyNFE drv, PAGE_TIMEOUT_SECONDS

        If EstaPaginaDeLoginNFE(drv) Then
            LogNFE "ConsultarDetranRetorno: permaneceu na tela de login após tentativa adicional."
            GoTo FAIL
        End If
    End If

    Dim msgNaoCadastrado As String
    msgNaoCadastrado = DetectarPlacaNaoCadastradaNFE(drv)
    If Len(msgNaoCadastrado) > 0 Then
        LogNFE "ConsultarDetranRetorno: " & msgNaoCadastrado & "; tentando consulta BIN."
        Dim dossieBIN As Object
        Set dossieBIN = TentarConsultaBINPorPlaca(drv, PLACA)
        If Not dossieBIN Is Nothing Then
            If Len(NzTexto(CStr(dossieBIN("Placa")))) > 0 Or _
               Len(NzTexto(CStr(dossieBIN("Renavam")))) > 0 Then
                LogNFE "ConsultarDetranRetorno: dados obtidos via BIN."
                Set dados = BuildCollectionFromDossieNFE(dossieBIN)
                Set ConsultarDetranRetorno = dados
                Exit Function
            End If
        End If
        g_ultimaMensagemErroNFE = msgNaoCadastrado
        LogNFE "ConsultarDetranRetorno: BIN também não retornou dados para a placa."
        Exit Function
    End If

    Set dossie = RasparDossieVeiculoNFE(drv)
    Set dados = BuildCollectionFromDossieNFE(dossie)

    Set ConsultarDetranRetorno = dados
    Exit Function

FAIL:
    LogNFE "ConsultarDetranRetorno: erro " & Err.Number & " - " & Err.Description
    EncerrarSessaoDriverNFE
End Function

Private Function GetDetranDriverNFE() As Object
    Dim drv As Object

    If TryReuseDetranDriverNFE(drv) Then
        Set GetDetranDriverNFE = drv
        Exit Function
    End If

    On Error GoTo FAIL
    EncerrarSessaoDriverNFE
    LogNFE "GetDetranDriverNFE: criando nova instância do ChromeDriver."
    Set g_drvNFE = CreateObject(PROGID_CHROME)
    ConfigureChromedriverExecutableNFE g_drvNFE
    ConfigureAndStartChromeNFE g_drvNFE, CHROME_BINARY_PATH
    Set GetDetranDriverNFE = g_drvNFE
    Exit Function

FAIL:
    Dim errMsg As String
    errMsg = Err.Description
    EncerrarSessaoDriverNFE
    On Error GoTo 0
    Err.Clear
    Err.Raise vbObjectError + 601, "GetDetranDriverNFE", "Falha ao iniciar ChromeDriver: " & errMsg
End Function

Private Sub EncerrarSessaoDriverNFE()
    On Error Resume Next
    If Not g_drvNFE Is Nothing Then
        g_drvNFE.Quit
    End If
    Set g_drvNFE = Nothing
    On Error GoTo 0
End Sub

Public Sub EncerrarSessaoDetranNFE()
    EncerrarSessaoDriverNFE
End Sub

Private Function TryReuseDetranDriverNFE(ByRef drv As Object) As Boolean
    Dim ignoreUrl As String

    On Error GoTo CLEANUP

    If g_drvNFE Is Nothing Then Exit Function

    ignoreUrl = CStr(g_drvNFE.url)
    Set drv = g_drvNFE
    TryReuseDetranDriverNFE = True
    LogNFE "TryReuseDetranDriverNFE: reaproveitando instância ativa do ChromeDriver."
    Exit Function

CLEANUP:
    LogNFE "TryReuseDetranDriverNFE: instância existente inválida; será descartada."
    Err.Clear
    EncerrarSessaoDriverNFE
End Function

Private Function SanitizePlacaNFE(ByVal rawPlaca As String) As String
    Dim cleaned As String
    Dim buffer As String
    Dim idx As Long
    Dim ch As String

    cleaned = UCase$(Trim$(rawPlaca))
    cleaned = Replace(cleaned, "-", "")
    cleaned = Replace(cleaned, " ", "")

    For idx = 1 To Len(cleaned)
        ch = Mid$(cleaned, idx, 1)
        If ch Like "[A-Z0-9]" Then
            buffer = buffer & ch
        End If
    Next idx

    SanitizePlacaNFE = buffer
End Function

Private Function BuildConsultaUrlNFE(ByVal PLACA As String) As String
    Dim sanitized As String

    sanitized = SanitizePlacaNFE(PLACA)
    If Len(sanitized) = 0 Then Exit Function

    BuildConsultaUrlNFE = URL_CONSULTA_BASE & sanitized
End Function

Private Sub ConfigureChromedriverExecutableNFE(ByVal drv As Object, Optional ByVal driverExe As String = CHROMEDRIVER_EXE, Optional ByVal driverDir As String = CHROMEDRIVER_DIR)
    On Error Resume Next
    If Not drv Is Nothing Then
        If Len(Dir$(driverExe)) > 0 Then
            drv.SetDriverPath driverExe
        ElseIf Len(driverDir) > 0 Then
            drv.SetDriverPath driverDir
        End If
    End If
    EnsureDriverPathNFE driverDir
    On Error GoTo 0
End Sub

Private Sub EnsureDriverPathNFE(ByVal driverDir As String)
    Dim trimmedDir As String
    trimmedDir = Trim$(driverDir)
    If Len(trimmedDir) = 0 Then Exit Sub

    On Error Resume Next
    Dim sh As Object
    Dim env As Object
    Set sh = CreateObject("WScript.Shell")
    Set env = sh.Environment("PROCESS")
    If InStr(1, env("PATH"), trimmedDir, vbTextCompare) = 0 Then
        env("PATH") = trimmedDir & ";" & env("PATH")
    End If
    On Error GoTo 0
End Sub

Private Sub ConfigureAndStartChromeNFE(ByVal drv As Object, ByVal binaryPath As String)
    On Error Resume Next
    If Len(binaryPath) > 0 And Len(Dir$(binaryPath)) > 0 Then
        drv.SetBinary binaryPath
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
        Err.Raise IIf(startErrNumber <> 0, startErrNumber, vbObjectError + 602), "ConfigureAndStartChromeNFE", "Falha ao iniciar sessao do ChromeDriver: " & startErrDesc
    End If
    On Error GoTo 0
End Sub

Private Function SessaoDetranProvavelmenteAutenticadaNFE(ByVal drv As Object) As Boolean
    On Error GoTo CLEANUP

    Dim currentUrl As String
    currentUrl = LCase$(CStr(drv.url))

    If Len(currentUrl) = 0 Then Exit Function
    If InStr(currentUrl, "controleacesso/login") > 0 Then Exit Function
    If InStr(currentUrl, "logout") > 0 Then Exit Function

    If InStr(currentUrl, "arearestrita") > 0 Then
        SessaoDetranProvavelmenteAutenticadaNFE = True
        Exit Function
    End If

    If InStr(currentUrl, "consultaveiculo") > 0 Then
        SessaoDetranProvavelmenteAutenticadaNFE = True
        Exit Function
    End If

    Exit Function

CLEANUP:
    SessaoDetranProvavelmenteAutenticadaNFE = False
End Function

Private Function EstaPaginaDeLoginNFE(ByVal drv As Object) As Boolean
    On Error GoTo CLEANUP

    Dim currentUrl As String
    currentUrl = LCase$(CStr(drv.url))

    If InStr(currentUrl, "controleacesso/login") > 0 Then
        EstaPaginaDeLoginNFE = True
    ElseIf InStr(currentUrl, "login.asp") > 0 Then
        EstaPaginaDeLoginNFE = True
    End If
    Exit Function

CLEANUP:
    EstaPaginaDeLoginNFE = False
End Function

Private Sub WaitForUrlNFE(ByVal drv As Object, ByVal targetUrl As String, ByVal timeoutSeconds As Long)
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
        If NFE_ElapsedSeconds(startedAt) > timeoutSeconds Then Exit Do

        SleepMsNFE 200
        DoEvents
    Loop
End Sub

Private Function UrlAtualCorrespondeNFE(ByVal drv As Object, ByVal targetUrl As String) As Boolean
    Dim currentUrl As String

    On Error Resume Next
    currentUrl = CStr(drv.url)
    On Error GoTo 0

    If Len(targetUrl) = 0 Then Exit Function
    If Len(currentUrl) = 0 Then Exit Function

    If LCase$(Left$(currentUrl, Len(targetUrl))) = LCase$(targetUrl) Then
        UrlAtualCorrespondeNFE = True
    ElseIf LCase$(currentUrl) = LCase$(targetUrl) Then
        UrlAtualCorrespondeNFE = True
    End If
End Function

Private Function ObterJanelaChromeHandleNFE(ByVal titulo As String) As LongPtr
    m_targetCaptionNFE = Trim$(titulo)
    m_foundWindowNFE = 0

    If Len(m_targetCaptionNFE) = 0 Then Exit Function

#If Win64 Then
    EnumWindows AddressOf EnumWindowsProcNFE, 0
#Else
    EnumWindows AddressOf EnumWindowsProcNFE, 0
#End If

    ObterJanelaChromeHandleNFE = m_foundWindowNFE
End Function

#If Win64 Then
Private Function EnumWindowsProcNFE(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Private Function EnumWindowsProcNFE(ByVal hWnd As Long, ByVal lParam As Long) As Long
#End If
    Dim length As Long
    Dim CAPTION As String

    If IsWindowVisible(hWnd) = 0 Then
        EnumWindowsProcNFE = 1
        Exit Function
    End If

    length = GetWindowTextLengthW(hWnd)
    If length = 0 Then
        EnumWindowsProcNFE = 1
        Exit Function
    End If

    CAPTION = String$(length + 1, vbNullChar)
    GetWindowTextW hWnd, StrPtr(CAPTION), length + 1
    CAPTION = Left$(CAPTION, length)

    If InStr(1, CAPTION, m_targetCaptionNFE, vbTextCompare) > 0 Then
        m_foundWindowNFE = hWnd
        EnumWindowsProcNFE = 0
    Else
        EnumWindowsProcNFE = 1
    End If
End Function

Private Function WaitDocumentReadyNFE(ByVal drv As Object, Optional ByVal timeoutSeconds As Long = PAGE_TIMEOUT_SECONDS) As Boolean
    WaitDocumentReadyNFE = WaitForReadyStateCompleteNFE(drv, timeoutSeconds, READY_INTERVAL_MS)
End Function

Private Function WaitForReadyStateCompleteNFE(ByVal drv As Object, Optional ByVal timeoutSeconds As Long = 30, Optional ByVal sleepIntervalMs As Long = 250) As Boolean
    On Error GoTo CLEANUP

    Dim startedAt As Single
    startedAt = Timer

    Do
        Dim state As Variant
        state = drv.ExecuteScript("return document.readyState")
        If CStr(state) = "complete" Then
            WaitForReadyStateCompleteNFE = True
            Exit Function
        End If

        If NFE_ElapsedSeconds(startedAt) > timeoutSeconds Then Exit Do

        If sleepIntervalMs > 0 Then SleepMsNFE sleepIntervalMs
        DoEvents
    Loop

    Exit Function

CLEANUP:
    WaitForReadyStateCompleteNFE = False
End Function

Private Sub MostrarDriverWindowNFE(ByVal drv As Object)
    Dim titulo As String
    Dim hwndChrome As LongPtr
    Dim shell As Object

    On Error Resume Next
    If drv Is Nothing Then Exit Sub

    titulo = NzTexto(CStr(drv.Title))

    drv.Window.Maximize
    If Err.Number <> 0 Then Err.Clear
    drv.Window.SetPosition 100, 100
    If Err.Number <> 0 Then Err.Clear

    hwndChrome = ObterJanelaChromeHandleNFE(titulo)
    If hwndChrome <> 0 Then
        Call ShowWindow(hwndChrome, SW_RESTORE)
    End If

    Set shell = CreateObject("WScript.Shell")
    If Not shell Is Nothing Then
        If Len(titulo) > 0 Then shell.AppActivate titulo
    End If

    drv.ExecuteScript "window.focus();"
    If Err.Number <> 0 Then Err.Clear

    Set shell = Nothing
    On Error GoTo 0
End Sub

Private Sub SleepMsNFE(ByVal ms As Long)
    'Centraliza a espera; as chamadas publicas e privadas permanecem iguais.
    modSeleniumCommon.SleepMs ms
End Sub

Private Function RasparDossieVeiculoNFE(ByVal driver As Object) As Object
    Dim dados As Object
    Set dados = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    dados.CompareMode = vbTextCompare
    On Error GoTo 0
    dados.CompareMode = vbTextCompare

    LogNFE "RasparDossieVeiculoNFE: iniciando raspagem do dossiê."

    On Error Resume Next
    driver.ExecuteScript "ConsultarCaracteristicasVeiculo();"
    On Error GoTo 0

    dados("Placa") = LimpaNBSP(GetValorPorRotuloNFE(driver, "Placa", "divDadosVeiculo"))
    dados("Renavam") = LimpaNBSP(GetValorPorRotuloNFE(driver, "Renavam", "divDadosVeiculo"))
    dados("MarcaModelo") = LimpaNBSP(GetValorPorRotuloNFE(driver, "Marca/Modelo", "divDadosVeiculo"))
    dados("FabricacaoModelo") = LimpaNBSP(GetValorPorRotuloNFE(driver, "Fabricação/Modelo", "divDadosVeiculo"))
    dados("Chassi") = LimpaNBSP(GetValorPorRotuloNFE(driver, "Chassi", "divCaracteristicasVeiculo"))
    dados("Cor") = LimpaNBSP(GetValorPorRotuloNFE(driver, "Cor", "divDadosVeiculo"))

    Set RasparDossieVeiculoNFE = dados
End Function

Private Function GetValorPorRotuloNFE(ByVal driver As Object, ByVal rotulo As String, Optional ByVal divId As String = "") As String
    Dim xp As String
    Dim el As Object
    Dim base As String

    If Len(divId) > 0 Then
        base = "//div[@id='" & divId & "']"
    End If

    xp = base & "//td[contains(@class,'titcel')][normalize-space(text()[1])='" & rotulo & "']/span[1]"

    On Error Resume Next
    Set el = driver.FindElementByXPath(xp)
    If el Is Nothing Then
        xp = base & "//td[contains(@class,'titcel')][contains(normalize-space(text()[1]), '" & rotulo & "')]/span[1]"
        Set el = driver.FindElementByXPath(xp)
    End If
    On Error GoTo 0

    If Not el Is Nothing Then
        GetValorPorRotuloNFE = Trim$(CStr(el.text))
    Else
        GetValorPorRotuloNFE = ""
    End If
End Function

Private Function BuildCollectionFromDossieNFE(ByVal dossie As Object) As Collection
    Dim resultado As New Collection
    Dim keys As Variant
    Dim key As Variant
    Dim valor As String

    keys = Array("MarcaModelo", "Placa", "Renavam", "Chassi", "FabricacaoModelo", "Cor")

    For Each key In keys
        resultado.Add CStr(key)
        If Not dossie Is Nothing Then
            If dossie.Exists(CStr(key)) Then
                valor = NzTexto(CStr(dossie(CStr(key))))
            Else
                valor = ""
            End If
        Else
            valor = ""
        End If
        resultado.Add valor
    Next key

    Set BuildCollectionFromDossieNFE = resultado
End Function

Public Function GetValueNFE(ByVal col As Collection, ByVal key As String) As String
    Dim idx As Long
    If col Is Nothing Then Exit Function

    For idx = 1 To col.count Step 2
        If StrComp(CStr(col(idx)), key, vbTextCompare) = 0 Then
            GetValueNFE = NzTexto(CStr(col(idx + 1)))
            Exit Function
        End If
    Next idx
End Function

' ---------------------------------------------------------------
'  Fallback BIN: navegar pelo menu RENAVAM ? BIN e raspar dados
' ---------------------------------------------------------------

Private Function TentarConsultaBINPorPlaca(ByVal drv As Object, ByVal PLACA As String) As Object
    LogNFE "TentarConsultaBINPorPlaca: etapa executada; conteúdo de dados omitido do log."

    If Not NavegarParaBINNFE(drv) Then
        LogNFE "TentarConsultaBINPorPlaca: falha ao navegar para a página BIN."
        Exit Function
    End If

    If Not PreencherESubmeterPlacaBINNFE(drv, PLACA) Then
        LogNFE "TentarConsultaBINPorPlaca: falha ao preencher/submeter placa no BIN."
        Exit Function
    End If

    WaitDocumentReadyNFE drv, PAGE_TIMEOUT_SECONDS
    SleepMsNFE 800

    Set TentarConsultaBINPorPlaca = RasparDossieBINNFE(drv)
End Function

Private Function NavegarParaBINNFE(ByVal drv As Object) As Boolean
    On Error GoTo CLEANUP

    SwitchToDefaultContentNFE drv
    drv.Get URL_HOME
    WaitDocumentReadyNFE drv, PAGE_TIMEOUT_SECONDS

    ' Clica na aba RENAVAM no frame top
    If SwitchToFrameByNameNFE(drv, "top") Then
        ClickByIdNFE drv, "RENAVAM"
        SleepMsNFE 400
    End If

    ' Expande o menu RENAVAM no frame contents
    SwitchToDefaultContentNFE drv
    If Not SwitchToFrameByNameNFE(drv, "contents") Then GoTo CLEANUP
    SleepMsNFE 300
    ClickByIdNFE drv, "HGRENAVAMRENAVAM"
    SleepMsNFE 400

    ' Clica no submenu BIN
    If Not ClickByIdNFE(drv, "inRENAVAMRENAVAMconsultabin") Then
        ' Fallback via JS predestino
        On Error Resume Next
        drv.ExecuteScript "predestino('RENAVAM/RENAVAM/consultabin.asp?" & _
                          "DetranNet_Sistema=RENAVAM&DetranNet_Modulo=RENAVAM'," & _
                          "'RENAVAM','RENAVAM');"
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo CLEANUP
    End If

    WaitDocumentReadyNFE drv, PAGE_TIMEOUT_SECONDS
    SleepMsNFE 600

    NavegarParaBINNFE = True
    Exit Function

CLEANUP:
    LogNFE "NavegarParaBINNFE: falha na navegação."
    NavegarParaBINNFE = False
End Function

Private Function PreencherESubmeterPlacaBINNFE(ByVal drv As Object, ByVal PLACA As String) As Boolean
    Dim el As Object

    On Error GoTo CLEANUP

    ' txtDado pode estar no frame form, botoes ou no documento atual
    SwitchToDefaultContentNFE drv
    Set el = TentarEncontrarCampoBINNFE(drv)
    If el Is Nothing Then GoTo CLEANUP

    On Error Resume Next
    el.Clear
    el.SendKeys PLACA
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo CLEANUP

    LogNFE "PreencherESubmeterPlacaBINNFE: etapa executada; conteúdo de dados omitido do log."

    ' O botão Placa está no frame botoes
    SwitchToDefaultContentNFE drv
    If SwitchToFrameByNameNFE(drv, "botoes") Then
        If Not ClickByNameNFE(drv, "botaoC_BINPlaca") Then
            ' Fallback pelo id genérico "botao"
            ClickByIdNFE drv, "botao"
        End If
    Else
        ' Sem frame específico
        ClickByNameNFE drv, "botaoC_BINPlaca"
    End If

    PreencherESubmeterPlacaBINNFE = True
    Exit Function

CLEANUP:
    LogNFE "PreencherESubmeterPlacaBINNFE: falha."
    PreencherESubmeterPlacaBINNFE = False
End Function

Private Function TentarEncontrarCampoBINNFE(ByVal drv As Object) As Object
    Dim el As Object
    Dim frameNames As Variant
    Dim i As Long

    ' txtDado fica no frame "form" (mesma página do resultado BIN)
    frameNames = Array("form", "", "botoes", "contents", "Consulta")

    For i = 0 To UBound(frameNames)
        SwitchToDefaultContentNFE drv
        Dim frameOk As Boolean
        If Len(CStr(frameNames(i))) = 0 Then
            frameOk = True
        Else
            frameOk = SwitchToFrameByNameNFE(drv, CStr(frameNames(i)))
        End If

        If frameOk Then
            Set el = FindElementByNameNFE(drv, "txtDado")
            If Not el Is Nothing Then
                Set TentarEncontrarCampoBINNFE = el
                Exit Function
            End If
        End If
    Next i

    SwitchToDefaultContentNFE drv
End Function

Private Function RasparDossieBINNFE(ByVal driver As Object) As Object
    Dim dados As Object
    Set dados = CreateObject("Scripting.Dictionary")
    dados.CompareMode = vbTextCompare

    LogNFE "RasparDossieBINNFE: iniciando raspagem dos resultados BIN."

    ' Resultado do BIN fica no frame "form" — tenta ele primeiro, depois demais
    Dim frameNames As Variant
    frameNames = Array("form", "", "Consulta", "contents")
    Dim i As Long

    For i = 0 To UBound(frameNames)
        SwitchToDefaultContentNFE driver

        Dim frameOk As Boolean
        If Len(CStr(frameNames(i))) = 0 Then
            frameOk = True
        Else
            frameOk = SwitchToFrameByNameNFE(driver, CStr(frameNames(i)))
        End If

        If frameOk Then
            If Len(dados("Placa")) = 0 Then _
                dados("Placa") = LimparPlacaBINNFE(LimpaNBSP(GetValorBINPorRotuloNFE(driver, "Placa")))
            If Len(dados("Renavam")) = 0 Then _
                dados("Renavam") = LimpaNBSP(GetValorBINPorRotuloNFE(driver, "Renavam"))
            If Len(dados("MarcaModelo")) = 0 Then _
                dados("MarcaModelo") = LimpaNBSP(GetValorBINPorRotuloNFE(driver, "Marca/Modelo"))
            If Len(dados("Chassi")) = 0 Then _
                dados("Chassi") = LimpaNBSP(GetValorBINPorRotuloNFE(driver, "Chassi"))
            If Len(dados("FabricacaoModelo")) = 0 Then _
                dados("FabricacaoModelo") = LimpaNBSP(GetValorBINPorRotuloNFE(driver, _
                    "Fabrica" & Chr(231) & Chr(227) & "o/Modelo"))
            If Len(dados("Cor")) = 0 Then _
                dados("Cor") = LimpaNBSP(GetValorBINPorRotuloNFE(driver, "Cor"))

            If Len(dados("Placa")) > 0 Or Len(dados("Renavam")) > 0 Then
                LogNFE "RasparDossieBINNFE: dados encontrados (frame='" & CStr(frameNames(i)) & "')."
                SwitchToDefaultContentNFE driver
                Set RasparDossieBINNFE = dados
                Exit Function
            End If
        End If
    Next i

    SwitchToDefaultContentNFE driver
    LogNFE "RasparDossieBINNFE: nenhum dado encontrado após varredura de frames."
    Set RasparDossieBINNFE = dados
End Function

Private Function GetValorBINPorRotuloNFE(ByVal driver As Object, ByVal rotulo As String) As String
    ' Padrão real da página BIN:
    ' <font class="titcel">RÓTULO<br><span class="celbold">VALOR</span></font>
    Dim xp As String
    Dim el As Object

    ' Correspondência exata do rótulo
    xp = "//font[@class='titcel'][normalize-space(text()[1])='" & rotulo & "']/span[@class='celbold']"
    On Error Resume Next
    Set el = driver.FindElementByXPath(xp)
    On Error GoTo 0
    If Not el Is Nothing Then
        GetValorBINPorRotuloNFE = Trim$(CStr(el.text))
        Exit Function
    End If

    ' Fallback com contains (rótulos com acentos ligeiramente diferentes)
    xp = "//font[@class='titcel'][contains(normalize-space(text()[1]),'" & rotulo & "')]/span"
    On Error Resume Next
    Set el = driver.FindElementByXPath(xp)
    On Error GoTo 0
    If Not el Is Nothing Then
        GetValorBINPorRotuloNFE = Trim$(CStr(el.text))
    End If
End Function

Private Function LimparPlacaBINNFE(ByVal rawPlaca As String) As String
    ' BIN retorna "TDG3G54/MG (CIRCULACAO)" ? extrai só a placa "TDG3G54"
    Dim s As String
    Dim pos As Long

    s = Trim$(rawPlaca)

    ' Remove sufixo /UF e tudo que vier depois
    pos = InStr(s, "/")
    If pos > 0 Then s = Trim$(Left$(s, pos - 1))

    ' Remove sufixo com espaço (ex.: placa com status sem barra)
    pos = InStr(s, " ")
    If pos > 0 Then s = Trim$(Left$(s, pos - 1))

    LimparPlacaBINNFE = s
End Function

' --- Helpers de frame/elemento (escopo local NFE) ---------------

Private Sub SwitchToDefaultContentNFE(ByVal drv As Object)
    On Error Resume Next
    drv.SwitchToDefaultContent
    On Error GoTo 0
End Sub

Private Function SwitchToFrameByNameNFE(ByVal drv As Object, ByVal frameName As String) As Boolean
    On Error GoTo CLEANUP
    Dim frameEl As Object

    On Error Resume Next
    Set frameEl = drv.FindElementByXPath("//frame[@name='" & frameName & "']")
    If frameEl Is Nothing Then _
        Set frameEl = drv.FindElementByXPath("//iframe[@name='" & frameName & "']")
    On Error GoTo CLEANUP

    If frameEl Is Nothing Then GoTo CLEANUP
    drv.SwitchToFrame frameEl
    SwitchToFrameByNameNFE = True
    Exit Function

CLEANUP:
    SwitchToFrameByNameNFE = False
End Function

Private Function ClickByIdNFE(ByVal drv As Object, ByVal elementId As String) As Boolean
    On Error GoTo fallback
    Dim el As Object
    Set el = drv.FindElementById(elementId)
    If el Is Nothing Then GoTo fallback
    el.Click
    ClickByIdNFE = True
    Exit Function

fallback:
    On Error Resume Next
    drv.ExecuteScript "var e=document.getElementById('" & elementId & "');if(e)e.click();"
    ClickByIdNFE = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function ClickByNameNFE(ByVal drv As Object, ByVal nameValue As String) As Boolean
    On Error GoTo fallback
    Dim el As Object
    Set el = drv.FindElementByName(nameValue)
    If el Is Nothing Then GoTo fallback
    el.Click
    ClickByNameNFE = True
    Exit Function

fallback:
    On Error Resume Next
    drv.ExecuteScript "var e=document.getElementsByName('" & nameValue & "')[0];if(e)e.click();"
    ClickByNameNFE = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function FindElementByNameNFE(ByVal drv As Object, ByVal nameValue As String) As Object
    On Error Resume Next
    Set FindElementByNameNFE = drv.FindElementByName(nameValue)
    On Error GoTo 0
End Function

' ---------------------------------------------------------------

Private Function DetectarPlacaNaoCadastradaNFE(ByVal driver As Object) As String
    Dim bodyText As String
    Dim linhas() As String
    Dim i As Long

    On Error Resume Next
    bodyText = CStr(driver.ExecuteScript("return document.body ? document.body.innerText : '';"))
    On Error GoTo 0

    If InStr(UCase$(bodyText), "N" & Chr(195) & "O CADASTRADO NO DETRAN") > 0 Then
        linhas = Split(bodyText, vbLf)
        For i = 0 To UBound(linhas)
            If InStr(UCase$(linhas(i)), "N" & Chr(195) & "O CADASTRADO NO DETRAN") > 0 Then
                DetectarPlacaNaoCadastradaNFE = Trim$(linhas(i))
                Exit Function
            End If
        Next i
        DetectarPlacaNaoCadastradaNFE = "Ve" & Chr(237) & "culo n" & Chr(227) & "o cadastrado no DETRAN/SC."
    End If
End Function

Private Function LimpaNBSP(ByVal s As String) As String
    LimpaNBSP = Trim$(Replace(s, ChrW(160), " "))
End Function

Private Function NFE_ElapsedSeconds(ByVal startedAt As Single) As Double
    Dim current As Single
    current = Timer
    If current >= startedAt Then
        NFE_ElapsedSeconds = current - startedAt
    Else
        NFE_ElapsedSeconds = (86400# - startedAt) + current
    End If
End Function

Public Function ConsultarDadosNFE(ByVal PLACA As String) As String
    Dim dados As Collection
    Dim resumo As String

    PLACA = UCase$(Trim$(PLACA))
    If Len(PLACA) = 0 Then Exit Function

    LogNFE "ConsultarDadosNFE: etapa executada; conteúdo de dados omitido do log."

    Set dados = ConsultarDetranRetorno(PLACA)
    If dados Is Nothing Then
    LogNFE "ConsultarDadosNFE: etapa executada; conteúdo de dados omitido do log."
        Exit Function
    End If

    resumo = FormatDadosNFE(dados)
    ConsultarDadosNFE = resumo
End Function

Public Function FormatDadosNFE(ByVal dados As Collection) As String
    Dim marcaModeloRaw As String
    Dim PLACA As String
    Dim renavam As String
    Dim chassi As String
    Dim fabricacaoModelo As String
    Dim corRaw As String
    Dim anoModelo As String

    If dados Is Nothing Then Exit Function

    marcaModeloRaw = GetValueNFE(dados, "MarcaModelo")
    PLACA = GetValueNFE(dados, "Placa")
    renavam = GetValueNFE(dados, "Renavam")
    chassi = GetValueNFE(dados, "Chassi")
    fabricacaoModelo = GetValueNFE(dados, "FabricacaoModelo")
    corRaw = GetValueNFE(dados, "Cor")

    anoModelo = ExtrairAnoModelo(fabricacaoModelo)

    FormatDadosNFE = BuildResumoTexto( _
        LimparMarcaModelo(marcaModeloRaw), _
        PLACA, _
        renavam, _
    LimparChassi(chassi), _
        anoModelo, _
        LimparCor(corRaw) _
    )
End Function

Public Sub CopiarDadosDetranNFE(ByVal PLACA As String)
    Dim texto As String

    texto = ConsultarDadosNFE(PLACA)

    If Len(Trim$(texto)) = 0 Then
        LogNFE "CopiarDadosDetranNFE: texto vazio; nenhuma consulta realizada."
        If Len(g_ultimaMensagemErroNFE) > 0 Then
            MsgBox g_ultimaMensagemErroNFE, vbExclamation, "Consulta DETRANNET"
        Else
            MsgBox "Não foi possível consultar os dados do veículo.", vbExclamation, "Consulta DETRANNET"
        End If
        Exit Sub
    End If

    CopiarParaClipboardNFE texto
    LogNFE "CopiarDadosDetranNFE: dados copiados para área de transferência."
    MsgBox "Consulta concluída com sucesso!" & vbCrLf & vbCrLf & _
           "Os dados foram copiados para a área de transferência:" & vbCrLf & vbCrLf & _
           texto, vbInformation, "Consulta DETRANNET"
End Sub

Private Function BuildResumoTexto(ByVal marcaModelo As String, _
                                  ByVal PLACA As String, _
                                  ByVal renavam As String, _
                                  ByVal chassi As String, _
                                  ByVal anoModelo As String, _
                                  ByVal cor As String) As String
    Dim partes As Collection
    Dim resultado As String
    Dim parte As Variant

    Set partes = New Collection

    partes.Add "Marca/Modelo " & NzTexto(marcaModelo)
    partes.Add "Placa " & NzTexto(PLACA)
    partes.Add "Renavam " & NzTexto(renavam)
    partes.Add "Chassi " & NzTexto(chassi)
    partes.Add "Ano/Mod " & NzTexto(anoModelo)
    partes.Add "Cor " & NzTexto(cor)

    For Each parte In partes
        If Len(resultado) = 0 Then
            resultado = parte
        Else
            resultado = resultado & " " & parte
        End If
    Next parte

    BuildResumoTexto = Trim$(resultado)
End Function

Private Function LimparMarcaModelo(ByVal rawValue As String) As String
    Dim s As String
    Dim posCodigo As Long
    Dim posParenteses As Long

    s = Trim$(rawValue)
    If Len(s) = 0 Then Exit Function

    posCodigo = InStr(s, " - ")
    If posCodigo > 0 Then
        s = Mid$(s, posCodigo + 3)
    End If

    posParenteses = InStr(s, "(")
    If posParenteses > 0 Then
        s = Trim$(Left$(s, posParenteses - 1))
    End If

    LimparMarcaModelo = Trim$(s)
End Function

Private Function LimparChassi(ByVal rawValue As String) As String
    Dim s As String
    Dim startPos As Long
    Dim endPos As Long

    s = Trim$(rawValue)
    If Len(s) = 0 Then Exit Function

    Do
        startPos = InStr(s, "(")
        If startPos = 0 Then Exit Do

        endPos = InStr(startPos, s, ")")
        If endPos = 0 Then Exit Do

        s = Trim$(Left$(s, startPos - 1) & Mid$(s, endPos + 1))
    Loop

    LimparChassi = Trim$(s)
End Function

Private Function ExtrairAnoModelo(ByVal fabricacaoModelo As String) As String
    Dim partes As Variant
    Dim anoFabricacao As String
    Dim anoModelo As String

    fabricacaoModelo = Trim$(fabricacaoModelo)
    If Len(fabricacaoModelo) = 0 Then Exit Function

    partes = Split(fabricacaoModelo, "/")
    If UBound(partes) < 1 Then
        ExtrairAnoModelo = fabricacaoModelo
        Exit Function
    End If

    anoFabricacao = Trim$(CStr(partes(0)))
    anoModelo = Trim$(CStr(partes(1)))

    If Len(anoFabricacao) >= 2 And Len(anoModelo) >= 2 Then
        ExtrairAnoModelo = Right$(anoFabricacao, 2) & "/" & Right$(anoModelo, 2)
    Else
        ExtrairAnoModelo = fabricacaoModelo
    End If
End Function

Private Function LimparCor(ByVal rawValue As String) As String
    Dim s As String
    Dim posSeparador As Long

    s = Trim$(rawValue)
    If Len(s) = 0 Then Exit Function

    posSeparador = InStr(s, "-")
    If posSeparador > 0 Then
        s = Mid$(s, posSeparador + 1)
    End If

    LimparCor = Trim$(s)
End Function

Private Function NzTexto(ByVal value As String) As String
    If Len(Trim$(value)) = 0 Then
        NzTexto = ""
    Else
        NzTexto = Trim$(value)
    End If
End Function

Public Sub CommandButtonConsultaDetran()
    Dim PLACA As String

    PLACA = UCase$(Trim$(InputBox( _
        "Informe a placa do veículo (formato AAA0A00):", _
        "Consulta DETRANNET SC")))
    If Len(PLACA) = 0 Then
        LogNFE "CommandButtonConsultaDetran: placa não informada; operação cancelada."
        Exit Sub
    End If

    CopiarDadosDetranNFE PLACA
End Sub

Public Sub CommandButton_Click()
    CommandButtonConsultaDetran
End Sub

Public Sub CommandButtonReexibirChrome()
    Dim drv As Object
    Dim ignoreUrl As String

    On Error GoTo HANDLE_ERROR

    If g_drvNFE Is Nothing Then
        MsgBox "Nenhuma sessão do ChromeDriver está ativa.", vbExclamation, "Consulta DETRANNET"
        Exit Sub
    End If

    Set drv = g_drvNFE

    On Error Resume Next
    ignoreUrl = CStr(drv.url)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo HANDLE_ERROR
        EncerrarSessaoDriverNFE
        MsgBox "A sessão do ChromeDriver foi encerrada.", vbExclamation, "Consulta DETRANNET"
        Exit Sub
    End If
    On Error GoTo HANDLE_ERROR

    LogNFE "CommandButtonReexibirChrome: solicitando reexibição da janela do Chrome."
    MostrarDriverWindowNFE drv
    Exit Sub

HANDLE_ERROR:
    LogNFE "CommandButtonReexibirChrome: erro " & Err.Number & " - " & Err.Description
    MsgBox "Não foi possível reexibir o ChromeDriver.", vbExclamation, "Consulta DETRANNET"
End Sub

Private Sub CopiarParaClipboardNFE(ByVal texto As String)
    Dim copiado As Boolean
    Dim erroMsForms As String
    Dim erroWinApi As String

    texto = NzTexto(texto)
    If Len(texto) = 0 Then
        LogNFE "CopiarParaClipboardNFE: texto vazio; nada a copiar."
        Exit Sub
    End If

    copiado = CopiarViaMsFormsDataObject(texto, erroMsForms)

    If Not copiado Then
        If Len(erroMsForms) > 0 Then
            LogNFE "CopiarParaClipboardNFE: tentativa via MSForms falhou - " & erroMsForms & ". Tentando fallback WinAPI."
        Else
            LogNFE "CopiarParaClipboardNFE: MSForms.DataObject indisponível; tentando fallback WinAPI."
        End If

        copiado = CopiarViaClipboardWinAPI(texto, erroWinApi)

        If Not copiado And Len(erroWinApi) > 0 Then
            LogNFE "CopiarParaClipboardNFE: fallback WinAPI falhou - " & erroWinApi
        End If
    End If

    If copiado Then
        LogNFE "CopiarParaClipboardNFE: texto copiado para a área de transferência."
    Else
        LogNFE "CopiarParaClipboardNFE: falha ao colocar texto na área de transferência."
    End If
End Sub

Private Function CopiarViaMsFormsDataObject(ByVal texto As String, Optional ByRef erroDetalhado As String) As Boolean
    Dim clipboard As Object

    On Error Resume Next
    Set clipboard = CreateObject("MSForms.DataObject")
    If clipboard Is Nothing Then
        Err.Clear
        Set clipboard = CreateObject("Forms.DataObject")
    End If

    If clipboard Is Nothing Then
        erroDetalhado = "MSForms.DataObject indisponível"
        Exit Function
    End If

    clipboard.SetText CStr(texto)
    If Err.Number <> 0 Then
        erroDetalhado = Err.Description
        Err.Clear
        GoTo CLEANUP
    End If

    clipboard.PutInClipboard
    If Err.Number = 0 Then
        CopiarViaMsFormsDataObject = True
        GoTo CLEANUP
    End If

    erroDetalhado = Err.Description
    Err.Clear
    CallByName clipboard, "CopyToClipboard", VbMethod
    If Err.Number = 0 Then
        CopiarViaMsFormsDataObject = True
    ElseIf Len(erroDetalhado) = 0 Then
        erroDetalhado = Err.Description
    End If
    Err.Clear

CLEANUP:
    Set clipboard = Nothing
    On Error GoTo 0
End Function

Private Function CopiarViaClipboardWinAPI(ByVal texto As String, Optional ByRef erroDetalhado As String) As Boolean
    Dim dados As String
    Dim bytesNecessarios As LongPtr
    Dim hMem As LongPtr
    Dim ponteiroMemoria As LongPtr
    Dim tentativa As Long
    Dim clipboardAberto As Boolean

    On Error GoTo HANDLE_ERROR

    dados = CStr(texto) & vbNullChar & vbNullChar
    bytesNecessarios = LenB(dados)
    If bytesNecessarios = 0 Then
        erroDetalhado = "Nenhum dado para copiar"
        GoTo CLEANUP
    End If

    For tentativa = 1 To 5
        If OpenClipboard(CLngPtr(Application.hWnd)) <> 0 Then
            clipboardAberto = True
            Exit For
        End If
        SleepMsNFE 25
    Next tentativa

    If Not clipboardAberto Then
        erroDetalhado = "OpenClipboard falhou"
        GoTo CLEANUP
    End If

    If EmptyClipboard() = 0 Then
        erroDetalhado = "EmptyClipboard falhou"
        GoTo CLEANUP
    End If

    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, bytesNecessarios)
    If hMem = 0 Then
        erroDetalhado = "GlobalAlloc falhou"
        GoTo CLEANUP
    End If

    ponteiroMemoria = GlobalLock(hMem)
    If ponteiroMemoria = 0 Then
        erroDetalhado = "GlobalLock falhou"
        GoTo CLEANUP
    End If

    RtlMoveMemory ByVal ponteiroMemoria, ByVal StrPtr(dados), bytesNecessarios
    GlobalUnlock hMem
    ponteiroMemoria = 0

    If SetClipboardData(CF_UNICODETEXT, hMem) = 0 Then
        erroDetalhado = "SetClipboardData falhou"
        GoTo CLEANUP
    End If

    CopiarViaClipboardWinAPI = True
    hMem = 0 ' Propriedade do sistema após sucesso

    GoTo CLEANUP

HANDLE_ERROR:
    erroDetalhado = "Erro " & Err.Number & " - " & Err.Description
    Err.Clear

CLEANUP:
    On Error Resume Next
    If ponteiroMemoria <> 0 Then
        GlobalUnlock hMem
        ponteiroMemoria = 0
    End If
    If clipboardAberto Then CloseClipboard
    If hMem <> 0 Then GlobalFree hMem
    On Error GoTo 0
End Function

Private Sub LogNFE(ByVal message As String)
    On Error Resume Next
    modSeleniumCommon.WriteLog message
    If Err.Number <> 0 Then
        Debug.Print message
        Err.Clear
    End If
    On Error GoTo 0
End Sub



