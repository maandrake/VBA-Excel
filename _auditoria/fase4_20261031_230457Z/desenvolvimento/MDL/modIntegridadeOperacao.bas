Attribute VB_Name = "modIntegridadeOperacao"
Option Explicit

Public Function PrimeiraVagaGravacao(ByVal ws As Worksheet, ByVal primeiraLinha As Long) As Range
    Dim linha As Long
    If primeiraLinha < 1 Or primeiraLinha > ws.Rows.Count Then
        Err.Raise vbObjectError + 2120, , "Linha inicial invalida."
    End If
    For linha = primeiraLinha To ws.Rows.Count
        'Preserva a regra existente: primeira celula realmente Empty, nao formula vazia.
        If IsEmpty(ws.Cells(linha, 1).Value2) Then
            Set PrimeiraVagaGravacao = ws.Cells(linha, 1)
            Exit Function
        End If
    Next linha
    Err.Raise vbObjectError + 2121, , "Nao existe vaga na planilha " & ws.Name
End Function

Public Function ProximaLinhaGravacao(ByVal ws As Worksheet) As Range
    Dim linha As Long
    linha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If linha < 2 Then linha = 2
    If linha > ws.Rows.Count Then Err.Raise vbObjectError + 2122, , "Planilha sem capacidade."
    Set ProximaLinhaGravacao = ws.Cells(linha, 1)
End Function

Public Sub VerificarDestinoGravacao(ByVal alvo As Range)
    Dim celula As Range
    If alvo.Parent.ProtectContents Then
        Err.Raise vbObjectError + 2123, , "Planilha protegida: " & alvo.Parent.Name
    End If
    For Each celula In alvo.Cells
        If celula.MergeCells Then Err.Raise vbObjectError + 2124, , "Destino com celula mesclada."
    Next celula
End Sub

Public Sub ValidarConversoesGravacao(ByVal valores As Variant, ByVal dataTexto As Variant)
    Dim valor As Variant, convertido As Double, dataConvertida As Date
    'Usa as mesmas conversoes locais do codigo original, antes de qualquer escrita.
    For Each valor In valores
        convertido = CDbl(valor)
    Next valor
    dataConvertida = CDate(dataTexto)
End Sub

Public Sub ValidarChaveTexto(ByVal chave As String)
    If Len(chave) = 0 Then Exit Sub
    If Len(chave) <> 44 Or chave Like "*[!0-9]*" Then
        Err.Raise vbObjectError + 2125, , "A chave informada deve conter 44 digitos."
    End If
End Sub
