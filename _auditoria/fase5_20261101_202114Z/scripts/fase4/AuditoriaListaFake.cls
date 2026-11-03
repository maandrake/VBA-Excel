VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "AuditoriaListaFake"
Option Explicit

'Objeto de teste em memoria. Nao substitui controles dos formularios reais.
Public ColumnCount As Long
Public ColumnWidths As String
Public ListIndex As Long
Private mLista As Variant

Public Sub Clear()
    mLista = Empty
    ListIndex = -1
End Sub

Public Property Let List(ByVal dados As Variant)
    mLista = dados
End Property

Public Property Get List() As Variant
    List = mLista
End Property

Public Property Get ListCount() As Long
    If IsArray(mLista) Then ListCount = UBound(mLista, 1) - LBound(mLista, 1) + 1
End Property
