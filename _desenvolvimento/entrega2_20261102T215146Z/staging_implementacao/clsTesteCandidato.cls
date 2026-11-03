VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsTesteCandidato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ArquivoXML As String
Public Etapa As String
Public Codigo As Long
Public Responsavel As String
Public Motivo As String
Public Evidencia As String
Public Function Arquivo() As String
    Arquivo = ArquivoXML
End Function
Public Function Ponto(ByVal nome As String) As Long
    If nome = Etapa Then Ponto = Codigo
End Function

Public Function Responder(ByVal titulo As String, ByVal prompt As String) As String
    Select Case titulo
        Case "Responsavel": Responder = Responsavel
        Case "Motivo": Responder = Motivo
        Case "Identidade nova": Responder = Evidencia
    End Select
End Function
