VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "AuditoriaListaFake"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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

Public Property Let list(ByVal dados As Variant)
    mLista = dados
End Property

Public Property Get list() As Variant
    list = mLista
End Property

Public Property Get ListCount() As Long
    If IsArray(mLista) Then ListCount = UBound(mLista, 1) - LBound(mLista, 1) + 1
End Property
