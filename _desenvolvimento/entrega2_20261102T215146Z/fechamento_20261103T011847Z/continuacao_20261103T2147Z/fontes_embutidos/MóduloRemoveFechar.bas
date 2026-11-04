Attribute VB_Name = "MóduloRemoveFechar"
Option Explicit
'Função que retornará um identificador para o UserForm que possui o nome da classe e o nome do UserForm
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

'Função que recupera as informações sobre o nome da classe e o estilo da janela do UserForm
Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long

'Função que altera o estilo da janela do UserForm
Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Função que redesenha a barra de título do UserForm
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long

Sub RemoveBotao(ObjForm As Object)

    'Código que remove o botão Fechar do UserForm
    SetWindowLongA FindWindowA("ThunderDFrame", ObjForm.CAPTION), -16, GetWindowLongA(FindWindowA("ThunderDFrame", ObjForm.CAPTION), -16) And &HFFF7FFFF

    'Código que atualiza a barra de título do UserForm sem o botão Fechar
    DrawMenuBar FindWindowA("ThunderDFrame", ObjForm.CAPTION)

End Sub

