VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsTesteCandidato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Public ArquivoXML As String
Public Etapa As String
Public Codigo As Long
Public Function Arquivo() As String
    Arquivo = ArquivoXML
End Function
Public Function Ponto(ByVal nome As String) As Long
    If nome = Etapa Then Ponto = Codigo
End Function
