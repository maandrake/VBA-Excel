Attribute VB_Name = "mdlMacroConexão"
Option Explicit
Sub MacroConexão()
    Dim livro As Workbook
    On Error GoTo Indisponivel
    'Integração legada: conserva a pasta ativa; proprietário ainda não confirmado.
    Set livro = ActiveWorkbook
    If livro Is Nothing Then Exit Sub
    livro.Connections("Consulta - Tutores").Refresh
    Exit Sub
Indisponivel:
    MsgBox "Não foi possível atualizar a conexão 'Consulta - Tutores'." & _
        vbCrLf & "Confira a pasta ativa e a configuração da integração legada." & _
        vbCrLf & "Erro " & Err.Number & ": " & Err.Description, _
        vbExclamation, "Conexão indisponível"
End Sub

