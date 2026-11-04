Attribute VB_Name = "modTestesAuditoria"
Option Explicit

'Executar somente na fixture nova criada por Preparar-TestesVBA.ps1.
'Esta suite foi preparada, mas nao executada: VBOM bloqueado nesta auditoria.
Public Function Auditoria_ExecutarTestesSeguros() As String
    Dim lista As AuditoriaListaFake
    Dim WS As Worksheet
    Dim dados As Variant
    Dim CHAVE As String
    Dim telaAnterior As Boolean
    Dim numeroErro As Long, descricaoErro As String

    If InStr(1, CreateObject("Scripting.FileSystemObject").GetFile(ThisWorkbook.FullName).ShortPath, CreateObject("Scripting.FileSystemObject").GetFolder(Environ$("TEMP")).ShortPath & "\", vbTextCompare) <> 1 Then
        Err.Raise vbObjectError + 1700, , "A suite exige uma fixture na pasta temporaria."
    End If
    If Left$(ThisWorkbook.name, 18) <> "Fixture_Auditoria_" Then
        Err.Raise vbObjectError + 1701, , "Esta pasta nao e uma fixture de auditoria."
    End If
    If ThisWorkbook.Worksheets("Menu Principal").Range("Z1").Value2 <> "AUDITORIA_VBA_FICTICIO" Then
        Err.Raise vbObjectError + 1702, , "Marcador de dados ficticios ausente."
    End If

    telaAnterior = Application.ScreenUpdating
    On Error GoTo Falhou
    Application.ScreenUpdating = False

    Set lista = New AuditoriaListaFake
    Set WS = ThisWorkbook.Worksheets("Entradas")
    WS.Range("A1:K10").ClearContents
    WS.Range("A1").Value2 = "CABECALHO"
    WS.Range("A2").Value2 = "TESTE_1"
    WS.Range("D2").Value2 = "VEICULO FICTICIO"
    WS.Range("A3").Value2 = "TESTE_2"
    WS.Range("D3").Value2 = "VEICULO FICTICIO"
    WS.Range("E3").value = CVErr(2042)
    Set WS = ThisWorkbook.Worksheets("Entradas Filial")
    WS.Range("A1:L10").ClearContents
    WS.Range("A1").Value2 = "CABECALHO"
    WS.Range("A2").Value2 = "FILIAL_1"
    CHAVE = "00000000000000000000000000000000000000000001"
    WS.Range("L2").NumberFormat = "@"
    WS.Range("L2").Value2 = CHAVE

    PesquisarEntradasRapido "", lista
    Exigir lista.ListCount = 0, "Pesquisa vazia"
    PesquisarEntradasRapido "veiculo ficticio", lista
    Exigir lista.ListCount = 2, "Duplicados, caixa e celula com erro"
    dados = lista.list
    Exigir dados(0, 0) = "Entradas", "Origem do resultado"
    Exigir dados(0, 12) = 2, "Linha original"
    PesquisarEntradasRapido CHAVE, lista
    Exigir lista.ListCount = 1, "Pesquisa pela chave"
    dados = lista.list
    Exigir dados(0, 13) = CHAVE, "Chave integral no resultado"
    PesquisarEntradasRapido "INEXISTENTE", lista
    Exigir lista.ListCount = 0, "Nenhum resultado"
    Exigir TruncateForLog("  a" & vbCrLf & "b  ") = "a  b", "Limpeza do log"
    Exigir TruncateForLog("abcdef", 5) = "ab...", "Truncamento do log"

    Application.ScreenUpdating = telaAnterior
    Auditoria_ExecutarTestesSeguros = "PASS: 9 verificacoes; pesquisa e log; dados ficticios"
    Exit Function
Falhou:
    numeroErro = Err.Number
    descricaoErro = Err.Description
    Application.ScreenUpdating = telaAnterior
    Err.Raise numeroErro, "Auditoria_ExecutarTestesSeguros", descricaoErro
End Function

Private Sub Exigir(ByVal condicao As Boolean, ByVal caso As String)
    If Not condicao Then Err.Raise vbObjectError + 1703, , "FAIL: " & caso
End Sub
