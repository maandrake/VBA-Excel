Attribute VB_Name = "modSeleniumCommon"
Option Explicit

Public Const PROGID_CHROME As String = "Selenium.ChromeDriver"
Public Const CHROME_BINARY_PATH As String = "C:\Program Files\Google\Chrome\Application\chrome.exe"
Public Const CHROMEDRIVER_EXE As String = "C:\Program Files\SeleniumBasic\chromedriver.exe"
Public Const CHROMEDRIVER_DIR As String = "C:\Program Files\SeleniumBasic"
Public Const URL_LOGIN As String = "https://sistema.detrannet.sc.gov.br/controleacesso/login.asp"
Public Const URL_HOME As String = "https://sistema.detrannet.sc.gov.br/arearestrita/tela_principal.asp"
Public Const URL_CONSULTA_BASE As String = "https://sistema.detrannet.sc.gov.br/arearestrita/restrito/aguarde.asp?destino=consultaveiculo2.asp?placa="

Public Sub ConfigureChromedriverExecutable(ByVal drv As Object, Optional ByVal driverExe As String = CHROMEDRIVER_EXE, Optional ByVal driverDir As String = CHROMEDRIVER_DIR)
    On Error Resume Next
    If Not drv Is Nothing Then
        drv.SetDriverPath driverExe
    End If
    EnsureDriverPath driverDir
    On Error GoTo 0
End Sub

Public Sub EnsureDriverPath(ByVal driverDir As String)
    On Error Resume Next
    Dim trimmedDir As String
    trimmedDir = Trim$(driverDir)
    If Len(trimmedDir) = 0 Then Exit Sub

    Dim sh As Object, env As Object
    Set sh = CreateObject("WScript.Shell")
    Set env = sh.Environment("PROCESS")
    If InStr(1, env("PATH"), trimmedDir, vbTextCompare) = 0 Then
        env("PATH") = trimmedDir & ";" & env("PATH")
    End If
    On Error GoTo 0
End Sub

Public Function WaitForReadyStateComplete(ByVal drv As Object, Optional ByVal timeoutSeconds As Long = 30, Optional ByVal sleepIntervalMs As Long = 250) As Boolean
    On Error GoTo CLEANUP

    Dim startedAt As Single
    startedAt = Timer

    Do
        Dim state As Variant
        state = drv.ExecuteScript("return document.readyState")
        If CStr(state) = "complete" Then
            WaitForReadyStateComplete = True
            Exit Function
        End If

        If SegundosDecorridos(startedAt) > timeoutSeconds Then Exit Do

        If sleepIntervalMs > 0 Then SleepMs sleepIntervalMs
        DoEvents
    Loop

    Exit Function

CLEANUP:
    WaitForReadyStateComplete = False
End Function

Public Sub SleepMs(ByVal ms As Long)
    Dim t As Single
    t = Timer
    Do While SegundosDecorridos(t) * 1000 < ms
        DoEvents
    Loop
End Sub

Public Function TruncateForLog(ByVal value As String, Optional ByVal maxLen As Long = 120) As String
    Dim s As String
    s = Trim$(value)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    If maxLen <= 3 Then maxLen = 3

    If Len(s) > maxLen Then
        TruncateForLog = Left$(s, maxLen - 3) & "..."
    Else
        TruncateForLog = s
    End If
End Function

' Log compartilhado por todos os módulos; grava em C:\Temp\detran_log.txt
Public Sub WriteLog(ByVal message As String)
    Const LOG_FOLDER As String = "C:\Temp"
    Const LOG_FILE_NAME As String = "detran_log.txt"

    Dim fso As Object
    Dim logFile As Object
    Dim logPath As String

    On Error GoTo HANDLE_ERROR

    message = Replace(message, vbCr, " ")
    message = Replace(message, vbLf, " ")
    message = TruncateForLog(message, 500)

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then GoTo HANDLE_ERROR

    If Not fso.FolderExists(LOG_FOLDER) Then
        fso.CreateFolder LOG_FOLDER
    End If

    logPath = LOG_FOLDER & "\" & LOG_FILE_NAME
    Set logFile = fso.OpenTextFile(logPath, 8, True)
    logFile.WriteLine Format$(Now, "dd-mm-yyyy hh:nn:ss") & " - " & message
    logFile.Close

    GoTo CLEANUP

HANDLE_ERROR:
    Debug.Print "[WriteLog] Falha ao escrever log: " & Err.Number & " - " & Err.Description

CLEANUP:
    On Error Resume Next
    If Not logFile Is Nothing Then logFile.Close
    On Error GoTo 0
End Sub


' Timer reinicia em zero a cada meia-noite; estas esperas duram menos de um dia.
Public Function SegundosDecorridos(ByVal inicio As Single) As Double
    Dim agora As Double
    agora = Timer
    If agora < inicio Then agora = agora + 86400#
    SegundosDecorridos = agora - inicio
End Function

