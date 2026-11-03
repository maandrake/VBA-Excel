Attribute VB_Name = "mdlOcultarTítulo"

Option Explicit
Option Private Module

Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
Public Declare PtrSafe Function GetWindowLong _
                       Lib "user32" Alias "GetWindowLongA" ( _
                       ByVal hWnd As LongPtr, _
                       ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong _
                       Lib "user32" Alias "SetWindowLongA" ( _
                       ByVal hWnd As LongPtr, _
                       ByVal nIndex As Long, _
                       ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function DrawMenuBar _
                       Lib "user32" ( _
                       ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function FindWindowA _
                       Lib "user32" (ByVal lpClassName As String, _
                       ByVal lpWindowName As String) As LongPtr

Sub HideTitleBar(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As LongPtr
    lFrmHdl = FindWindowA(vbNullString, frm.CAPTION)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub

'COLOCAR NO SEU CÓDIGO DE USUÁRIO
'Private Sub UserForm_Initialize()
'    HideTitleBar Me
'End Sub


