Attribute VB_Name = "modSelTest"
' Módulo: modSelTest
' Objetivo: Verificar requisitos para funcionamento do SeleniumBasic com Word/VBA
' - Verifica registro COM (ProgID Selenium.ChromeDriver)
' - Sobe o ChromeDriver, executa scripts e valida métodos básicos de localização
' - Checa acesso ao URL de login do DETRANNET
' - Verifica compatibilidade de versões Chrome x ChromeDriver (major)
'
Option Explicit


Private Const SELENIUMBASIC_DEFAULT_DIR As String = "C:\Program Files\SeleniumBasic"
Private Const SELENIUMBASIC_DEFAULT_DIR_X86 As String = "C:\Program Files (x86)\SeleniumBasic"
Private Const DOTNET_MIN_RELEASE As Long = 378389 ' .NET Framework 4.5
Private Const DOTNET_MIN_VERSION_LABEL As String = "4.5"
Private Const STATUS_ICON_PASS As String = "??"
Private Const STATUS_ICON_FAIL As String = "??"
Private Const STATUS_LABEL_PASS As String = "TESTE PASSOU"
Private Const STATUS_LABEL_FAIL As String = "TESTE FALHOU"

Private Function BuildStatusBadge(ByVal passed As Boolean) As String
    If passed Then
        BuildStatusBadge = STATUS_ICON_PASS & " " & STATUS_LABEL_PASS
    Else
        BuildStatusBadge = STATUS_ICON_FAIL & " " & STATUS_LABEL_FAIL
    End If
End Function

Private Function BuildStatusHeader(ByVal passed As Boolean, Optional ByVal context As String = "") As String
    Dim badge As String
    badge = BuildStatusBadge(passed)
    If Len(context) > 0 Then
        BuildStatusHeader = badge & " - " & context
    Else
        BuildStatusHeader = badge
    End If
End Function

Public Sub SelTest_Run()
    Dim log As String
    Dim okAll As Boolean: okAll = True

    log = log & "[SeleniumBasic - Verificação de ambiente]" & vbCrLf

    ' -1) Validação completa de ambiente (paths, versões, compatibilidade)
    Dim envDiag As String, envOk As Boolean
    envOk = ValidateSeleniumEnvironment(envDiag)
    log = log & envDiag & vbCrLf
    If Not envOk Then okAll = False

    ' 0) Checar existência do chromedriver.exe
    If Len(Dir$(CHROMEDRIVER_EXE)) = 0 Then
        log = log & "- ERRO: ChromeDriver não encontrado em: " & CHROMEDRIVER_EXE & vbCrLf
        okAll = False
    Else
        log = log & "- ChromeDriver encontrado: " & CHROMEDRIVER_EXE & vbCrLf
    End If

    ' 1) Checar ProgID
    Dim drv As Object, errMsg As String, errCode As Long
    If TryCreateDriver(drv, errMsg, errCode) Then
        log = log & "- ProgID OK: " & PROGID_CHROME & vbCrLf
    Else
        log = log & "- ERRO: Não foi possível criar objeto '" & PROGID_CHROME & "'." & vbCrLf & _
                    "  Dica: Instale/Reinstale o SeleniumBasic e tente novamente." & vbCrLf & _
                    "  Detalhe: (" & errCode & " / 0x" & Hex$(errCode) & ") " & errMsg & vbCrLf & _
                    BuildSeleniumComDiagnostic()
        okAll = False
        GoTo FINISH ' não há o que testar sem driver
    End If

    On Error GoTo CLEANUP

    ' 1.1) Configurar caminho do driver no PATH do processo
    ConfigureChromedriverExecutable drv

    ' 2) Iniciar o navegador replicando fluxo padrão (Start sem parâmetro)
    Dim startVariant As String
    Dim startError As String

    On Error Resume Next
    drv.Start
    If Err.Number = 0 Then
        startVariant = "driver.Start"
        log = log & "- Chrome iniciado usando driver.Start (sem parâmetros)." & vbCrLf
    Else
        startError = Err.Description
        Err.Clear
        drv.Start "chrome"
        If Err.Number = 0 Then
            startVariant = "driver.Start ""chrome"""
            log = log & "- Chrome iniciado com fallback driver.Start ""chrome"" (Start sem parâmetro falhou: " & startError & ")." & vbCrLf
        Else
            startError = IIf(Len(startError) > 0, startError & " | ", vbNullString) & Err.Description
            log = log & "- ERRO: Chrome não pôde ser iniciado (driver.Start / driver.Start ""chrome"")." & vbCrLf & _
                      "  Detalhe: " & startError & vbCrLf
            okAll = False
            GoTo CLEAN
        End If
    End If
    On Error GoTo CLEAN

    drv.Get "about:blank"
    If Not WaitForReadyStateComplete(drv, 10, 250) Then
        log = log & "- ERRO: Chrome não ficou pronto em 'about:blank' (iniciado via " & startVariant & ")." & vbCrLf
        okAll = False
    Else
        log = log & "- Chrome iniciado e DOM pronto (" & startVariant & ")." & vbCrLf
    End If

    ' 3) Capturar user-agent e verificar versão do Chrome
    Dim ua As String, chromeMajor As Long
    ua = CStr(drv.ExecuteScript("return navigator.userAgent"))
    chromeMajor = ParseChromeMajorFromUA(ua)
    If chromeMajor > 0 Then
        log = log & "- Chrome (major): " & chromeMajor & vbCrLf
    Else
        log = log & "- AVISO: Não foi possível identificar versão do Chrome no user-agent." & vbCrLf
    End If

    ' 4) Validar métodos de localização (ById / ByCss)
    drv.ExecuteScript "document.body.innerHTML='<div id=""divDadosVeiculo""><table><tbody><tr><td><span>AAA0A00</span></td><td><span>REN123</span></td><td><span>XYZ</span></td></tr></tbody></table></div>'"
    Dim root As Object
    On Error Resume Next
    Set root = drv.FindElementById("divDadosVeiculo")
    On Error GoTo CLEAN
    If root Is Nothing Then
        log = log & "- ERRO: FindElementById falhou (divDadosVeiculo)." & vbCrLf
        okAll = False
    Else
        Dim rows As Object
        Set rows = root.FindElementsByCss("table > tbody > tr")
        If rows Is Nothing Or rows.count = 0 Then
            log = log & "- ERRO: FindElementsByCss falhou (linhas da tabela)." & vbCrLf
            okAll = False
        Else
            Dim tds As Object
            Set tds = rows.item(0).FindElementsByCss("td > span")
            If tds Is Nothing Or tds.count = 0 Then
                log = log & "- ERRO: FindElementsByCss falhou (td > span)." & vbCrLf
                okAll = False
            Else
                log = log & "- Localização de elementos OK (ById/ByCss)." & vbCrLf
            End If
        End If
    End If

    ' 5) Verificar compatibilidade ChromeDriver x Chrome
    Dim driverMajor As Long
    driverMajor = GetChromeDriverMajor()
    If driverMajor > 0 Then
        log = log & "- ChromeDriver (major): " & driverMajor & vbCrLf
        If chromeMajor > 0 And chromeMajor <> driverMajor Then
            log = log & "- AVISO: Versões divergentes (Chrome=" & chromeMajor & ", ChromeDriver=" & driverMajor & ")." & vbCrLf & _
                          "  Dica: Atualize o chromedriver.exe correspondente ao seu Chrome." & vbCrLf
        ElseIf chromeMajor > 0 Then
            log = log & "- Compatibilidade OK (major coincidem)." & vbCrLf
        End If
    Else
        log = log & "- AVISO: Não foi possível obter a versão do ChromeDriver." & vbCrLf
    End If

    ' 6) Acesso ao site de login (conectividade)
    If UrlReachable(drv, URL_LOGIN, 20) Then
        log = log & "- Acesso ao URL de login OK: " & URL_LOGIN & vbCrLf
    Else
        log = log & "- AVISO: Não foi possível validar acesso ao URL de login (rede/proxy?)." & vbCrLf
    End If

CLEAN:
    ' ok

CLEANUP:
    On Error Resume Next
    If Not drv Is Nothing Then drv.Quit

FINISH:
    Dim statusHeader As String
    statusHeader = BuildStatusHeader(okAll, "SelTest")

    Debug.Print statusHeader
    Debug.Print log

    AppendSelfTestLog statusHeader & vbCrLf & log

    If okAll Then
        MsgBox statusHeader & vbCrLf & vbCrLf & log, vbInformation, "SelTest"
    Else
        MsgBox statusHeader & vbCrLf & vbCrLf & log, vbCritical, "SelTest"
    End If
End Sub

' Executa uma consulta simples e verifica rapidamente os campos essenciais retornados.
Public Sub SmokeTestConsultaDetran()
    Const CAPTION As String = "Smoke Test DETRANNET"

    Dim placaPadrao As String
    placaPadrao = GetSetting("ConsultaDetranSC", "SmokeTest", "Placa", "AAA0A00")

    Dim PLACA As String
    PLACA = InputBox("Informe a placa para o smoke test (use uma placa de homologacao).", CAPTION, placaPadrao)
    PLACA = UCase$(Trim$(PLACA))
    If Len(PLACA) = 0 Then Exit Sub

    On Error Resume Next
    SaveSetting "ConsultaDetranSC", "SmokeTest", "Placa", PLACA
    On Error GoTo 0

    WriteLog "SmokeTestConsultaDetran: etapa executada; conteúdo de dados omitido do log."

    Dim dados As Collection
    Set dados = ConsultarDetranRetorno(PLACA)

    If dados Is Nothing Or dados.count = 0 Then
        WriteLog "SmokeTestConsultaDetran: nenhum dado retornado pela consulta."
        MsgBox "Falha no smoke test: nenhuma informacao foi retornada.", vbExclamation, CAPTION
        Exit Sub
    End If

    Dim camposObrigatorios As Collection
    Set camposObrigatorios = BuildSmokeTestFieldSpecs()

    Dim faltando As String
    Dim relatorio As String
    Dim campo As Variant
    Dim campoChave As String
    Dim campoRotulo As String
    Dim valor As String

    For Each campo In camposObrigatorios
        campoChave = CStr(campo(0))
        campoRotulo = CStr(campo(1))
        valor = modSeleniumConsultaNFE.GetValueNFE(dados, campoChave)
        WriteLog "SmokeTestConsultaDetran: etapa executada; conteúdo de dados omitido do log."

        If Len(Trim$(valor)) = 0 Then
            If Len(faltando) > 0 Then faltando = faltando & ", "
            faltando = faltando & campoRotulo
        Else
            relatorio = relatorio & campoRotulo & ": " & valor & vbCrLf
        End If
    Next campo

    Dim smokePassed As Boolean
    smokePassed = (Len(faltando) = 0)

    Dim statusHeader As String
    statusHeader = BuildStatusHeader(smokePassed, "SmokeTest")

    WriteLog "SmokeTestConsultaDetran: status => " & BuildStatusBadge(smokePassed)

    If smokePassed Then
        WriteLog "SmokeTestConsultaDetran: todos os campos principais foram preenchidos."
        MsgBox statusHeader & vbCrLf & vbCrLf & "Campos validados:" & vbCrLf & relatorio, vbInformation, CAPTION
    Else
        WriteLog "SmokeTestConsultaDetran: campos ausentes => " & faltando
        MsgBox statusHeader & vbCrLf & vbCrLf & "Campos com pendências: " & faltando & vbCrLf & "Verifique o fluxo antes de prosseguir.", vbCritical, CAPTION
    End If
End Sub

Private Function BuildSmokeTestFieldSpecs() As Collection
    Dim specs As New Collection

    AddFieldSpec specs, "Placa", "Placa"
    AddFieldSpec specs, "Renavam", "Renavam"
    AddFieldSpec specs, "PlacaAnterior", "Placa Anterior"
    AddFieldSpec specs, "Tipo", "Tipo"
    AddFieldSpec specs, "Categoria", "Categoria"
    AddFieldSpec specs, "Lugares", "Lugares"
    AddFieldSpec specs, "MarcaModelo", "Marca/Modelo"
    AddFieldSpec specs, "FabricacaoModelo", "Fabricação/Modelo"
    AddFieldSpec specs, "Combustivel", "Combustível"
    AddFieldSpec specs, "Cor", "Cor"
    AddFieldSpec specs, "Carroceria", "Carroceria"
    AddFieldSpec specs, "CategoriaDpvat", "Categoria DPVAT"
    AddFieldSpec specs, "NomeProprietarioAtual", "Nome do Proprietário Atual"
    AddFieldSpec specs, "NomePrincipalCondutor", "Nome do Principal Condutor"
    AddFieldSpec specs, "RecadastroDetran", "Recadastrado DETRAN"
    AddFieldSpec specs, "NomeProprietarioAnterior", "Nome do Proprietário Anterior"
    AddFieldSpec specs, "MunicipioEmplacamento", "Município de Emplacamento"
    AddFieldSpec specs, "Licenciamento", "Licenciado"
    AddFieldSpec specs, "DataDeAquisicao", "Data de aquisição"
    AddFieldSpec specs, "Situacao", "Situação"
    AddFieldSpec specs, "Observacao", "Observação"
    AddFieldSpec specs, "RestricaoAVenda", "Restrição à Venda"
    AddFieldSpec specs, "InformacoesPendentes", "Informações PENDENTES originadas das financeiras via SNG - Sistema Nacional de Gravame"
    AddFieldSpec specs, "Restricoes", "Restrições"
    AddFieldSpec specs, "Chassi", "Chassi"
    AddFieldSpec specs, "NumeroDoMotor", "Número do Motor"
    AddFieldSpec specs, "NumeroDoCambio", "Número do Câmbio"
    AddFieldSpec specs, "DataInclusaoNoCadastro", "Data Inclusão no Cadastro"
    AddFieldSpec specs, "QuantidadeDeEixos", "Quantidade de Eixos"
    AddFieldSpec specs, "NumeroDoEixoTraseiro", "Número do Eixo Traseiro"
    AddFieldSpec specs, "NumeroDoEixoAuxiliar", "Número do Eixo Auxiliar"
    AddFieldSpec specs, "NumeroDaCarroceria", "Número da Carroceria"
    AddFieldSpec specs, "Potencia", "Potência"
    AddFieldSpec specs, "Cilindradas", "Cilindradas"
    AddFieldSpec specs, "CapacidadeDeCargaTonelada", "Capacidade de Carga (Ton.)"
    AddFieldSpec specs, "PesoBrutoTotal", "Peso Bruto Total (PBT)"
    AddFieldSpec specs, "CapacidadeMaximaTracao", "Capacidade Máxima de Tração (CMT)"

    Set BuildSmokeTestFieldSpecs = specs
End Function

Private Sub AddFieldSpec(ByRef specs As Collection, ByVal key As String, ByVal label As String)
    Dim pair As Variant
    pair = Array(key, label)
    specs.Add pair
End Sub

Private Function ValidateSeleniumEnvironment(ByRef diagMessage As String) As Boolean
    Dim issues As String
    Dim warnings As String

    If Len(Dir$(CHROME_BINARY_PATH)) = 0 Then
        issues = issues & "  - chrome.exe não encontrado em: " & CHROME_BINARY_PATH & vbCrLf
    End If

    If Len(Dir$(CHROMEDRIVER_EXE)) = 0 Then
        issues = issues & "  - chromedriver.exe não encontrado em: " & CHROMEDRIVER_EXE & vbCrLf
    End If

    Dim chromeVerRaw As String, driverVerRaw As String
    Dim chromeVersionSource As String, driverVersionSource As String

    chromeVerRaw = GetChromeVersionString(chromeVersionSource)
    driverVerRaw = GetChromeDriverVersionString(driverVersionSource)

    Dim chromeMajor As Long, driverMajor As Long
    chromeMajor = ExtractMajorVersion(chromeVerRaw)
    driverMajor = ExtractMajorVersion(driverVerRaw)

    Dim dotNet45Diag As String
    Dim dotNet45Ok As Boolean
    dotNet45Ok = CheckDotNetRequirement(dotNet45Diag)

    Dim dotNet35Diag As String
    Dim dotNet35Ok As Boolean
    dotNet35Ok = CheckDotNet35Feature(dotNet35Diag)

    If chromeMajor = 0 Then
        issues = issues & "  - Não foi possível obter versão do Chrome (arquivo/registro)." & vbCrLf
    End If

    If driverMajor = 0 Then
        issues = issues & "  - Não foi possível obter versão do ChromeDriver (arquivo / --version)." & vbCrLf
    End If

    If Not dotNet45Ok Then
        issues = issues & "  - .NET Framework " & DOTNET_MIN_VERSION_LABEL & " ou superior não encontrado/insuficiente." & vbCrLf
    End If

    If Not dotNet35Ok Then
        issues = issues & "  - .NET Framework 3.5 (NetFx3) não habilitado. Ative em 'Ativar/Desativar Recursos do Windows'." & vbCrLf
    End If

    If chromeMajor > 0 And driverMajor > 0 Then
        If chromeMajor <> driverMajor Then
            issues = issues & "  - Versões divergentes: Chrome=" & chromeMajor & " x ChromeDriver=" & driverMajor & vbCrLf & _
                      "    Baixe o chromedriver correspondente à mesma versão major do Chrome." & vbCrLf
        End If
    End If

    Dim seleniumSummary As String
    seleniumSummary = BuildSeleniumInstallationSummary(warnings)

    warnings = TrimTrailingCrLf(warnings)

    If Len(issues) = 0 Then
        ValidateSeleniumEnvironment = True
        diagMessage = "- Ambiente Selenium: OK." & vbCrLf
    Else
        ValidateSeleniumEnvironment = False
        diagMessage = "- ERRO: Ambiente Selenium com problemas:" & vbCrLf & issues
    End If

    If Len(warnings) > 0 Then
        If ValidateSeleniumEnvironment Then
            diagMessage = diagMessage & "- Avisos:" & vbCrLf & warnings & vbCrLf
        Else
            diagMessage = diagMessage & "  Avisos:" & vbCrLf & warnings & vbCrLf
        End If
    End If

    If Len(chromeVerRaw) > 0 Then
        diagMessage = diagMessage & "  Chrome versão: " & chromeVerRaw
        If Len(chromeVersionSource) > 0 Then diagMessage = diagMessage & " (fonte: " & chromeVersionSource & ")"
        diagMessage = diagMessage & vbCrLf
    Else
        diagMessage = diagMessage & "  Chrome versão: (não localizado)"
        If Len(chromeVersionSource) > 0 Then diagMessage = diagMessage & " | tentativas: " & chromeVersionSource
        diagMessage = diagMessage & vbCrLf
    End If

    If Len(driverVerRaw) > 0 Then
        diagMessage = diagMessage & "  ChromeDriver versão: " & driverVerRaw
        If Len(driverVersionSource) > 0 Then diagMessage = diagMessage & " (fonte: " & driverVersionSource & ")"
        diagMessage = diagMessage & vbCrLf
    Else
        diagMessage = diagMessage & "  ChromeDriver versão: (não localizado)"
        If Len(driverVersionSource) > 0 Then diagMessage = diagMessage & " | tentativas: " & driverVersionSource
        diagMessage = diagMessage & vbCrLf
    End If

    If Len(dotNet45Diag) > 0 Then
        diagMessage = diagMessage & dotNet45Diag & vbCrLf
    End If

    If Len(dotNet35Diag) > 0 Then
        diagMessage = diagMessage & dotNet35Diag & vbCrLf
    End If

    diagMessage = diagMessage & "  Caminho Chrome: " & CHROME_BINARY_PATH & vbCrLf
    diagMessage = diagMessage & "  Caminho ChromeDriver: " & CHROMEDRIVER_EXE & vbCrLf

    If Len(seleniumSummary) > 0 Then
        diagMessage = diagMessage & seleniumSummary
    End If

    diagMessage = TrimTrailingCrLf(diagMessage)
End Function

Private Function RunAndCapture(ByVal exePath As String, ByVal args As String) As String
    Const LIMITE_SEGUNDOS As Double = 10#
    Dim sh As Object, execObj As Object
    Dim inicio As Single
    On Error GoTo FIM

    If Len(exePath) = 0 Then Exit Function
    If Len(Dir$(exePath)) = 0 Then Exit Function
    Set sh = CreateObject("WScript.Shell")
    Set execObj = sh.exec("""" & exePath & """ " & args)
    inicio = Timer
    Do While execObj.status = 0
        If SegundosDecorridos(inicio) >= LIMITE_SEGUNDOS Then
            'Somente o processo criado nesta chamada pode ser encerrado.
            execObj.Terminate
            Debug.Print "[RunAndCapture] Tempo limite de 10 segundos."
            Exit Function
        End If
        SleepMs 25
        DoEvents
    Loop
    RunAndCapture = Trim$(execObj.StdOut.ReadAll)
    Exit Function
FIM:
    On Error Resume Next
    If Not execObj Is Nothing Then
        If execObj.status = 0 Then execObj.Terminate
    End If
    RunAndCapture = vbNullString
End Function

Private Function ExtractMajorVersion(ByVal versionText As String) As Long
    Dim s As String
    s = Trim$(versionText)
    If Len(s) = 0 Then Exit Function

    Dim i As Long, startPos As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then
            startPos = i
            Do While i <= Len(s) And Mid$(s, i, 1) Like "#"
                i = i + 1
            Loop
            ExtractMajorVersion = CLng(Mid$(s, startPos, i - startPos))
            Exit Function
        End If
    Next i
End Function

Private Function TrimTrailingCrLf(ByVal text As String) As String
    Do While Len(text) > 0 And (Right$(text, 1) = vbCr Or Right$(text, 1) = vbLf)
        text = Left$(text, Len(text) - 1)
    Loop
    TrimTrailingCrLf = text
End Function

Private Function GetChromeVersionString(ByRef source As String) As String
    Dim version As String
    Dim attemptHistory As String

    attemptHistory = "arquivo"
    version = GetFileVersionSafe(CHROME_BINARY_PATH)
    If Len(version) > 0 Then
        source = "arquivo"
        GetChromeVersionString = version
        Exit Function
    End If

    attemptHistory = attemptHistory & ", registro HKCU"
    version = TryRegRead("HKCU\Software\Google\Chrome\BLBeacon\version")
    If Len(version) > 0 Then
        source = "registro HKCU"
        GetChromeVersionString = version
        Exit Function
    End If

    attemptHistory = attemptHistory & ", registro HKLM"
    version = TryRegRead("HKLM\SOFTWARE\Google\Chrome\BLBeacon\version")
    If Len(version) > 0 Then
        source = "registro HKLM"
        GetChromeVersionString = version
        Exit Function
    End If

    attemptHistory = attemptHistory & ", registro HKLM Wow6432Node"
    version = TryRegRead("HKLM\SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon\version")
    If Len(version) > 0 Then
        source = "registro HKLM Wow6432Node"
        GetChromeVersionString = version
        Exit Function
    End If

    source = "tentativas: " & attemptHistory
    GetChromeVersionString = ""
End Function

Private Function GetChromeDriverVersionString(ByRef source As String) As String
    Dim version As String
    version = GetFileVersionSafe(CHROMEDRIVER_EXE)
    If Len(version) > 0 Then
        source = "arquivo"
        GetChromeDriverVersionString = version
        Exit Function
    End If

    version = RunAndCapture(CHROMEDRIVER_EXE, "--version")
    If Len(version) > 0 Then
        source = "--version"
        GetChromeDriverVersionString = version
        Exit Function
    End If

    source = "arquivo, --version"
    GetChromeDriverVersionString = ""
End Function

Private Function GetFileVersionSafe(ByVal exePath As String) As String
    On Error GoTo EH
    If Len(exePath) = 0 Then Exit Function
    If Len(Dir$(exePath)) = 0 Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileVersionSafe = Trim$(CStr(fso.GetFileVersion(exePath)))
    Exit Function

EH:
    GetFileVersionSafe = ""
End Function

Private Function TryRegRead(ByVal regPath As String) As String
    On Error GoTo EH
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    TryRegRead = Trim$(CStr(wsh.RegRead(regPath)))
    Exit Function
EH:
    TryRegRead = ""
End Function

Private Function CheckDotNetRequirement(ByRef diag As String) As Boolean
    Dim releaseValue As Long
    Dim versionLabel As String

    releaseValue = GetDotNetReleaseValue()

    If releaseValue = 0 Then
        diag = "  .NET Framework: não encontrado (mínimo exigido " & DOTNET_MIN_VERSION_LABEL & ")"
        CheckDotNetRequirement = False
        Exit Function
    End If

    versionLabel = MapDotNetReleaseToVersion(releaseValue)
    If Len(versionLabel) = 0 Then versionLabel = "Release " & CStr(releaseValue)

    diag = "  .NET Framework detectado: " & versionLabel & " (Release " & releaseValue & ")"

    If releaseValue >= DOTNET_MIN_RELEASE Then
        CheckDotNetRequirement = True
    Else
        diag = diag & " - ABAIXO do mínimo requerido (" & DOTNET_MIN_VERSION_LABEL & ")"
        CheckDotNetRequirement = False
    End If
End Function

Private Function CheckDotNet35Feature(ByRef diag As String) As Boolean
    Dim svc As Object
    Dim col As Object
    Dim item As Object

    On Error Resume Next
    Set svc = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number = 0 Then
        Set col = svc.ExecQuery("SELECT InstallState, Version FROM Win32_OptionalFeature WHERE Name='NetFx3'")
        If Err.Number = 0 Then
            For Each item In col
                Select Case CLng(item.InstallState)
                    Case 1 ' Enabled
                        diag = "  .NET Framework 3.5 (NetFx3): habilitado" & _
                               IIf(Len(Trim$(CStr(item.version))) > 0, " (" & Trim$(CStr(item.version)) & ")", "")
                        CheckDotNet35Feature = True
                        Exit Function
                    Case 2 ' Disabled
                        diag = "  .NET Framework 3.5 (NetFx3): DESABILITADO (ativar em Recursos do Windows)"
                        Exit Function
                    Case Else
                        diag = "  .NET Framework 3.5 (NetFx3): estado desconhecido (InstallState=" & CStr(item.InstallState) & ")"
                        Exit Function
                End Select
            Next item
        End If
    End If
    On Error GoTo 0

    Dim basePaths As Variant
    Dim versionText As String
    Dim installText As String
    Dim p As Variant

    basePaths = Array( _
        "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5", _
        "HKLM\SOFTWARE\WOW6432Node\Microsoft\NET Framework Setup\NDP\v3.5")

    For Each p In basePaths
        installText = TryRegRead(p & "\Install")
        If Len(installText) > 0 Then
            On Error Resume Next
            If CLng(Val(installText)) = 1 Then
                versionText = TryRegRead(p & "\Version")
                If Len(versionText) = 0 Then versionText = "(versão não informada)"
                diag = "  .NET Framework 3.5 (registro): habilitado em " & p & " - versão " & versionText
                CheckDotNet35Feature = True
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
    Next p

    diag = "  .NET Framework 3.5 (NetFx3): não habilitado. Ative em 'Ativar/Desativar Recursos do Windows'."
End Function

Private Function GetDotNetReleaseValue() As Long
    Dim paths As Variant
    Dim i As Long
    Dim valueText As String
    Dim candidate As Long

    paths = Array( _
        "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\Release", _
        "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Client\Release", _
        "HKLM\SOFTWARE\WOW6432Node\Microsoft\NET Framework Setup\NDP\v4\Full\Release", _
        "HKLM\SOFTWARE\WOW6432Node\Microsoft\NET Framework Setup\NDP\v4\Client\Release" _
    )

    For i = LBound(paths) To UBound(paths)
        valueText = TryRegRead(CStr(paths(i)))
        If Len(valueText) > 0 Then
            On Error Resume Next
            candidate = CLng(valueText)
            If Err.Number = 0 And candidate > 0 Then
                GetDotNetReleaseValue = candidate
                Exit Function
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next i
End Function

Private Function MapDotNetReleaseToVersion(ByVal releaseValue As Long) As String
    Select Case True
        Case releaseValue >= 533320
            MapDotNetReleaseToVersion = ".NET Framework 4.8.1"
        Case releaseValue >= 528040
            MapDotNetReleaseToVersion = ".NET Framework 4.8"
        Case releaseValue >= 461808
            MapDotNetReleaseToVersion = ".NET Framework 4.7.2"
        Case releaseValue >= 461308
            MapDotNetReleaseToVersion = ".NET Framework 4.7.1"
        Case releaseValue >= 460798
            MapDotNetReleaseToVersion = ".NET Framework 4.7"
        Case releaseValue >= 394802
            MapDotNetReleaseToVersion = ".NET Framework 4.6.2"
        Case releaseValue >= 394254
            MapDotNetReleaseToVersion = ".NET Framework 4.6.1"
        Case releaseValue >= 393295
            MapDotNetReleaseToVersion = ".NET Framework 4.6"
        Case releaseValue >= 379893
            MapDotNetReleaseToVersion = ".NET Framework 4.5.2"
        Case releaseValue >= 378675
            MapDotNetReleaseToVersion = ".NET Framework 4.5.1"
        Case releaseValue >= 378389
            MapDotNetReleaseToVersion = ".NET Framework 4.5"
    End Select
End Function

Private Function BuildSeleniumInstallationSummary(ByRef warnings As String) As String
    Dim summary As String
    Dim comCommandRaw As String
    Dim comCommandNormalized As String

    comCommandRaw = GetSeleniumComServerCommand()
    If Len(comCommandRaw) > 0 Then
        comCommandNormalized = NormalizeCommandPath(comCommandRaw)
        summary = summary & "  SeleniumBasic COM (LocalServer32/Inproc): " & comCommandRaw & vbCrLf
        If Len(comCommandNormalized) > 0 Then
            If StrComp(comCommandNormalized, comCommandRaw, vbTextCompare) <> 0 Then
                summary = summary & "  SeleniumBasic caminho normalizado: " & comCommandNormalized & vbCrLf
            End If
        End If
    Else
        warnings = warnings & "  - Registro COM do SeleniumBasic (Selenium.ChromeDriver) não retornou caminho para LocalServer32/InprocServer32." & vbCrLf
    End If

    Dim installPath As String
    Dim exePath As String
    Dim typeLibPath As String
    Dim detectionSource As String

    Call DetectSeleniumInstall(installPath, exePath, typeLibPath, detectionSource, comCommandRaw)

    If Len(installPath) > 0 Then
        summary = summary & "  SeleniumBasic pasta detectada: " & installPath
        If Len(detectionSource) > 0 Then summary = summary & " (fonte: " & detectionSource & ")"
        summary = summary & vbCrLf
    Else
        summary = summary & "  SeleniumBasic pasta detectada: (não localizada automaticamente)" & vbCrLf
        warnings = warnings & "  - SeleniumBasic: pasta de instalação não localizada automaticamente. Ajuste caminhos ou execute o reparo se necessário." & vbCrLf
    End If

    If Len(exePath) > 0 Then
        summary = summary & "  SeleniumBasic executável: " & exePath & vbCrLf
    ElseIf Len(installPath) > 0 Then
        summary = summary & "  SeleniumBasic executável: não localizado (instalações >= 2.0.9.0 podem expor apenas SeleniumBasic.dll)." & vbCrLf
    End If

    If Len(typeLibPath) > 0 Then
        summary = summary & "  SeleniumBasic biblioteca (.tlb): " & typeLibPath & vbCrLf
    ElseIf Len(installPath) > 0 Then
        summary = summary & "  SeleniumBasic biblioteca (.tlb): não localizada (verifique referências VBA somente se necessário)." & vbCrLf
    End If

    BuildSeleniumInstallationSummary = summary
End Function

Private Function DetectSeleniumInstall(ByRef installPath As String, ByRef exePath As String, ByRef typeLibPath As String, ByRef detectionSource As String, Optional ByVal comCommandRaw As String = vbNullString) As Boolean
    Dim commandText As String
    commandText = comCommandRaw
    If Len(commandText) = 0 Then
        commandText = GetSeleniumComServerCommand()
    End If

    Dim normalizedCom As String
    normalizedCom = NormalizeCommandPath(commandText)

    If Len(normalizedCom) > 0 Then
        If FileExists(normalizedCom) Then
            exePath = normalizedCom
            detectionSource = "registro COM"
        End If

        Dim parentFolder As String
        parentFolder = GetParentFolderSafe(normalizedCom)
        If Len(parentFolder) > 0 Then
            installPath = parentFolder
            If Len(detectionSource) = 0 Then detectionSource = "registro COM (pasta)"
        End If
    End If

    Dim candidates As Variant
    candidates = GetSeleniumCandidateFolders(installPath)

    Dim i As Long
    Dim candidate As String
    For i = LBound(candidates) To UBound(candidates)
        candidate = CStr(candidates(i))
        If FolderExists(candidate) Then
            If Len(installPath) = 0 Then
                installPath = candidate
                If Len(detectionSource) = 0 Then detectionSource = "caminho padrão"
            End If

            If Len(exePath) = 0 Then
                exePath = FirstExistingFile(candidate, Array("SeleniumBasic.exe", "SeleniumBasicV2.exe", "SeleniumBasic64.exe", "SeleniumWrapper.exe"))
            End If

            If Len(typeLibPath) = 0 Then
                typeLibPath = FirstExistingFile(candidate, Array("SeleniumWrapper.tlb", "SeleniumBasic.tlb", "SeleniumWrapper64.tlb"))
            End If

            If Len(exePath) > 0 And Len(typeLibPath) > 0 And Len(installPath) > 0 Then Exit For
        End If
    Next i

    If Len(typeLibPath) = 0 And Len(installPath) > 0 Then
        Dim auxFolders As Variant
        auxFolders = Array(AppendPath(installPath, "Lib"), AppendPath(installPath, "TypeLib"), AppendPath(installPath, "tlb"))

        For i = LBound(auxFolders) To UBound(auxFolders)
            candidate = CStr(auxFolders(i))
            If FolderExists(candidate) Then
                typeLibPath = FirstExistingFile(candidate, Array("SeleniumWrapper.tlb", "SeleniumBasic.tlb", "SeleniumWrapper64.tlb"))
                If Len(typeLibPath) > 0 Then Exit For
            End If
        Next i
    End If

    DetectSeleniumInstall = (Len(installPath) > 0) Or (Len(exePath) > 0) Or (Len(typeLibPath) > 0)
End Function

Private Function GetSeleniumComServerCommand() As String
    Dim clsid As String
    clsid = TryRegRead("HKCR\Selenium.ChromeDriver\CLSID\")
    If Len(clsid) = 0 Then Exit Function

    Dim probePaths As Variant
    probePaths = Array( _
        "HKCR\CLSID\" & clsid & "\LocalServer32\", _
        "HKCR\CLSID\" & clsid & "\InprocServer32\", _
        "HKCR\WOW6432Node\CLSID\" & clsid & "\LocalServer32\", _
        "HKCR\WOW6432Node\CLSID\" & clsid & "\InprocServer32\" _
    )

    Dim p As Variant
    For Each p In probePaths
        Dim valueText As String
        valueText = TryRegRead(CStr(p))
        If Len(valueText) > 0 Then
            GetSeleniumComServerCommand = valueText
            Exit Function
        End If
    Next p
End Function

Private Function NormalizeCommandPath(ByVal commandText As String) As String
    Dim s As String
    s = Trim$(commandText)
    If Len(s) = 0 Then Exit Function

    If Left$(s, 1) = Chr$(34) Then
        Dim closing As Long
        closing = InStr(2, s, Chr$(34))
        If closing > 1 Then
            NormalizeCommandPath = Mid$(s, 2, closing - 2)
            Exit Function
        End If
    End If

    Dim ext As Variant
    For Each ext In Array(".exe", ".dll")
        Dim pos As Long
        pos = InStrRev(LCase$(s), CStr(ext))
        If pos > 0 Then
            NormalizeCommandPath = Left$(s, pos + Len(ext) - 1)
            Exit Function
        End If
    Next ext

    Dim spacePos As Long
    spacePos = InStr(s, " ")
    If spacePos > 0 Then
        NormalizeCommandPath = Left$(s, spacePos - 1)
    Else
        NormalizeCommandPath = s
    End If
End Function

Private Function GetParentFolderSafe(ByVal filePath As String) As String
    On Error Resume Next
    If Len(filePath) = 0 Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolderSafe = fso.GetParentFolderName(filePath)
    On Error GoTo 0
End Function

Private Function GetSeleniumCandidateFolders(Optional ByVal firstCandidate As String = "") As Variant
    Dim arr() As String
    Dim count As Long

    AddCandidateFolder arr, count, firstCandidate
    AddCandidateFolder arr, count, CHROMEDRIVER_DIR
    AddCandidateFolder arr, count, SELENIUMBASIC_DEFAULT_DIR
    AddCandidateFolder arr, count, SELENIUMBASIC_DEFAULT_DIR_X86
    AddCandidateFolder arr, count, ExpandEnvPath("LOCALAPPDATA", "\SeleniumBasic")
    AddCandidateFolder arr, count, ExpandEnvPath("ProgramData", "\SeleniumBasic")

    If count = 0 Then
        ReDim arr(1 To 1)
        arr(1) = ""
    End If

    GetSeleniumCandidateFolders = arr
End Function

Private Sub AddCandidateFolder(ByRef arr() As String, ByRef count As Long, ByVal folderPath As String)
    folderPath = Trim$(folderPath)
    If Len(folderPath) = 0 Then Exit Sub

    If Right$(folderPath, 1) = "\" Then
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    End If

    Dim i As Long
    For i = 1 To count
        If StrComp(arr(i), folderPath, vbTextCompare) = 0 Then Exit Sub
    Next i

    count = count + 1
    ReDim Preserve arr(1 To count)
    arr(count) = folderPath
End Sub

Private Function ExpandEnvPath(ByVal envName As String, ByVal suffix As String) As String
    Dim basePath As String
    basePath = Trim$(Environ$(envName))
    If Len(basePath) = 0 Then Exit Function

    If Len(suffix) > 0 Then
        If Right$(basePath, 1) = "\" And Left$(suffix, 1) = "\" Then
            ExpandEnvPath = basePath & Mid$(suffix, 2)
        ElseIf Right$(basePath, 1) = "\" Or Left$(suffix, 1) = "\" Then
            ExpandEnvPath = basePath & suffix
        Else
            ExpandEnvPath = basePath & "\" & suffix
        End If
    Else
        ExpandEnvPath = basePath
    End If
End Function

Private Function FirstExistingFile(ByVal baseFolder As String, ByVal fileNames As Variant) As String
    Dim i As Long
    For i = LBound(fileNames) To UBound(fileNames)
        Dim candidate As String
        candidate = AppendPath(baseFolder, CStr(fileNames(i)))
        If FileExists(candidate) Then
            FirstExistingFile = candidate
            Exit Function
        End If
    Next i
End Function

Private Function AppendPath(ByVal baseFolder As String, ByVal relativeName As String) As String
    If Len(baseFolder) = 0 Then
        AppendPath = relativeName
    ElseIf Len(relativeName) = 0 Then
        AppendPath = baseFolder
    ElseIf Right$(baseFolder, 1) = "\" Then
        AppendPath = baseFolder & relativeName
    Else
        AppendPath = baseFolder & "\" & relativeName
    End If
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    If Len(folderPath) = 0 Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
    On Error GoTo 0
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    If Len(filePath) = 0 Then Exit Function
    FileExists = Len(Dir$(filePath)) > 0
    On Error GoTo 0
End Function

Private Sub AppendSelfTestLog(ByVal content As String)
    If Len(Trim$(content)) = 0 Then Exit Sub

    On Error GoTo fallback

    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    lines = Split(content, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If Len(lineText) > 0 Then
            WriteLog "[SelTest] : etapa executada; conteúdo de dados omitido do log."
        End If
    Next i
    On Error GoTo 0
    Exit Sub

fallback:
    On Error Resume Next
    Dim fso As Object, folderPath As String, logFile As Object
    folderPath = "C:\Temp"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
        Set logFile = fso.OpenTextFile(folderPath & "\detran_log.txt", 8, True)
        If Not logFile Is Nothing Then
            logFile.WriteLine Now & " - [SelTest] " & Replace(content, vbCrLf, " | ")
            logFile.Close
        End If
    End If
    On Error GoTo 0
End Sub

Private Function TryCreateDriver(ByRef drv As Object, ByRef errMsg As String, ByRef errCode As Long) As Boolean
    errCode = 0
    On Error GoTo EH
    Set drv = CreateObject(PROGID_CHROME)
    TryCreateDriver = Not drv Is Nothing
    Exit Function
EH:
    errCode = Err.Number
    errMsg = Err.Description
    TryCreateDriver = False
End Function

Private Function BuildSeleniumComDiagnostic() As String
    Dim report As String
    Dim clsid As String
    Dim serverPath As String
    Dim probePaths As Variant
    Dim p As Variant

    report = "  Diagnóstico SeleniumBasic COM:" & vbCrLf

    clsid = TryRegRead("HKCR\Selenium.ChromeDriver\CLSID\")
    If Len(clsid) = 0 Then
        report = report & "    - CLSID não encontrado em HKCR\Selenium.ChromeDriver\CLSID" & vbCrLf
    Else
        report = report & "    - CLSID referenciado: " & clsid & vbCrLf

        probePaths = Array( _
            "HKCR\CLSID\" & clsid & "\InprocServer32\", _
            "HKCR\CLSID\" & clsid & "\LocalServer32\", _
            "HKCR\WOW6432Node\CLSID\" & clsid & "\InprocServer32\", _
            "HKCR\WOW6432Node\CLSID\" & clsid & "\LocalServer32\" _
        )

        For Each p In probePaths
            serverPath = TryRegRead(CStr(p))
            If Len(serverPath) > 0 Then
                report = report & "    - Entrada encontrada em " & CStr(p) & ": " & serverPath & vbCrLf
                Exit For
            End If
        Next p

        If Len(serverPath) = 0 Then
            report = report & "    - Servidor COM (InprocServer32/LocalServer32) não encontrado para o CLSID." & vbCrLf
        Else
            Dim normalized As String
            normalized = NormalizeCommandPath(serverPath)
            If Len(normalized) > 0 And StrComp(normalized, serverPath, vbTextCompare) <> 0 Then
                report = report & "    - Caminho normalizado: " & normalized & vbCrLf
            End If
        End If
    End If

    Dim installPath As String
    Dim exePath As String
    Dim typeLibPath As String
    Dim detectionSource As String

    Call DetectSeleniumInstall(installPath, exePath, typeLibPath, detectionSource, serverPath)

    If Len(installPath) > 0 Then
        report = report & "    - Pasta detectada: " & installPath
        If Len(detectionSource) > 0 Then report = report & " (fonte: " & detectionSource & ")"
        report = report & vbCrLf
    Else
        report = report & "    - Pasta de instalação do SeleniumBasic não localizada automaticamente." & vbCrLf
    End If

    If Len(exePath) > 0 Then
        report = report & "    - Executável detectado: " & exePath & vbCrLf
    ElseIf Len(installPath) > 0 Then
        report = report & "    - Executável SeleniumBasic*.exe não localizado; versões recentes podem registrar apenas SeleniumBasic.dll." & vbCrLf
    End If

    If Len(typeLibPath) > 0 Then
        report = report & "    - Biblioteca detectada: " & typeLibPath & vbCrLf
    ElseIf Len(installPath) > 0 Then
        report = report & "    - Biblioteca (.tlb) não localizada; confirme referências no VBA apenas se necessário." & vbCrLf
    End If

    report = report & "    - Sugestão: execute o reparo do SeleniumBasic (Repair/Register) ou reinstale o pacote para restaurar o registro COM." & vbCrLf

    BuildSeleniumComDiagnostic = report
End Function

Private Function UrlReachable(ByVal drv As Object, ByVal url As String, ByVal timeoutSeconds As Long) As Boolean
    On Error GoTo EH
    drv.Get url
    UrlReachable = WaitForReadyStateComplete(drv, timeoutSeconds, 250)
    Exit Function
EH:
    UrlReachable = False
End Function

Private Function ParseChromeMajorFromUA(ByVal ua As String) As Long
    Dim i As Long, pos As Long, endPos As Long, ver As String
    pos = InStr(1, ua, "Chrome/")
    If pos > 0 Then
        pos = pos + 7
        endPos = InStr(pos, ua, ".")
        If endPos > pos Then
            ver = Mid$(ua, pos, endPos - pos)
            If IsNumeric(ver) Then ParseChromeMajorFromUA = CLng(ver)
        End If
    End If
End Function

Private Function GetChromeDriverMajor() As Long
    Dim source As String
    Dim rawVersion As String
    rawVersion = GetChromeDriverVersionString(source)
    GetChromeDriverMajor = ExtractMajorVersion(rawVersion)
End Function




