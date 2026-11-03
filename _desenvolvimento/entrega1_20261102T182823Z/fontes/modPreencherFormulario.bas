Attribute VB_Name = "modPreencherFormulario"
Option Explicit

Public Sub PreencherDadosEntradas( _
    ByVal frm As Object, _
    ByVal lista As Object, _
    ByVal linha As Long)
    C1_SelecionarEntrada frm, lista, False
End Sub

Public Sub PreencherDadosEntradasFilial( _
    ByVal frm As Object, _
    ByVal lista As Object, _
    ByVal linha As Long)
    C1_SelecionarEntrada frm, lista, True
End Sub
