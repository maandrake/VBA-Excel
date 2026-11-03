Attribute VB_Name = "modIntegridadeOperacao"
Option Explicit

Public Function PrimeiraVagaGravacao(ByVal WS As Worksheet, ByVal primeiraLinha As Long) As Range
    Dim linha As Long
    If primeiraLinha < 1 Or primeiraLinha > WS.rows.count Then
        Err.Raise vbObjectError + 2120, , "Linha inicial invalida."
    End If
    For linha = primeiraLinha To WS.rows.count
        'Preserva a regra existente: primeira celula realmente Empty, nao formula vazia.
        If IsEmpty(WS.Cells(linha, 1).Value2) Then
            Set PrimeiraVagaGravacao = WS.Cells(linha, 1)
            Exit Function
        End If
    Next linha
    Err.Raise vbObjectError + 2121, , "Nao existe vaga na planilha " & WS.name
End Function

Public Function ProximaLinhaGravacao(ByVal WS As Worksheet) As Range
    Dim linha As Long
    linha = WS.Cells(WS.rows.count, 1).End(xlUp).row + 1
    If linha < 2 Then linha = 2
    If linha > WS.rows.count Then Err.Raise vbObjectError + 2122, , "Planilha sem capacidade."
    Set ProximaLinhaGravacao = WS.Cells(linha, 1)
End Function

Public Sub VerificarDestinoGravacao(ByVal alvo As Range)
    Dim celula As Range
    If alvo.Parent.ProtectContents Then
        Err.Raise vbObjectError + 2123, , "Planilha protegida: " & alvo.Parent.name
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

Public Sub ValidarChaveTexto(ByVal CHAVE As String)
    If Len(CHAVE) = 0 Then Exit Sub
    If Len(CHAVE) <> 44 Or CHAVE Like "*[!0-9]*" Then
        Err.Raise vbObjectError + 2125, , "A chave informada deve conter 44 digitos."
    End If
End Sub
