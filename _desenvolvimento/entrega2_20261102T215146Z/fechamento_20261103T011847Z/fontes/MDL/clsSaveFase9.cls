VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsSaveFase9"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private WithEvents app As Excel.Application
Attribute app.VB_VarHelpID = -1
Private target As Workbook
Private armed As Boolean, seen As Boolean
Public Sub Iniciar(ByVal wb As Workbook)
    Set target = wb: Set app = wb.Application
End Sub
Public Sub RecusarUmaVez()
    armed = True: seen = False
End Sub
Public Property Get Observado() As Boolean
    Observado = seen
End Property
Private Sub app_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If armed And wb Is target Then Cancel = True: seen = True: armed = False
End Sub
