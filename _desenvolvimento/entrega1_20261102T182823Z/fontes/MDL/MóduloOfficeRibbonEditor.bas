Attribute VB_Name = "MóduloOfficeRibbonEditor"
Option Explicit
Public Sub ChamarFrmEntrada(ByRef Control As Office.IRibbonControl)
    
    Cad_NFE.show
End Sub

Public Sub ChamarFrmSaida(ByRef Control As Office.IRibbonControl)
    
    Cad_NFE_Saida.show
End Sub

Public Sub ChamarSalvar(ByRef Control As Office.IRibbonControl)
    
    'O callback pertence a esta pasta, mesmo com outro arquivo ativo.
    ThisWorkbook.save
End Sub


