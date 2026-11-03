Attribute VB_Name = "Planilha15"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Evento órfão: a planilha Entradas Filial não contém CommandButton1 nem os
' controles de pesquisa referenciados pelo corpo legado, que não compilava.
Private Sub CommandButton1_Click()
    MsgBox "Esta pesquisa não está disponível nesta planilha.", vbInformation
End Sub


