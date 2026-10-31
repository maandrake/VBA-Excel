Attribute VB_Name = "MóduloOfficeRibbonEditor"
Option Explicit
Public Sub ChamarFrmEntrada(ByRef Control As Office.IRibbonControl)
    
    Cad_NFE.Show
End Sub

Public Sub ChamarFrmSaida(ByRef Control As Office.IRibbonControl)
    
    Cad_NFE_Saida.Show
End Sub

Public Sub ChamarSalvar(ByRef Control As Office.IRibbonControl)
    
    ActiveWorkbook.Save
End Sub


