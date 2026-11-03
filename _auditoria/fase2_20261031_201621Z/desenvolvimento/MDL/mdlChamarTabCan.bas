Attribute VB_Name = "mdlChamarTabCan"
Option Explicit
Sub ChamarTabCan()
    On Error GoTo Indisponivel
    ThisWorkbook.Worksheets("Cancelado, Estorno e Devolução").Select
    Exit Sub
Indisponivel:
    MsgBox "Não foi possível abrir a aba 'Cancelado, Estorno e Devolução'." & _
        vbCrLf & "Confira o vínculo legado antes de definir uma substituição." & _
        vbCrLf & "Erro " & Err.Number & ": " & Err.Description, _
        vbExclamation, "Navegação indisponível"
End Sub

