VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormFiltro 
   Caption         =   "Filtro"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   3360
   OleObjectBlob   =   "UserFormFiltro.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserFormFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    
    HideTitleBar Me
End Sub

Private Sub CommandButtonFechar_Click()

    UserFormFiltro.Hide
End Sub

Private Sub CommandButtonFiltroCompra_Click()
    C1_Bloqueado "Filtro contextual legado; use pesquisa integrada por identidade"
End Sub

Private Sub CommandButtonFiltroCompra_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelBakupFiltroCompra.ForeColor = &H80000012
End Sub

Private Sub FrameMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    LabelBakupFiltroCompra.ForeColor = &H80000011
    LabelRefazer.Visible = False
    LabelFiltroAguardando.ForeColor = &H80000011
    LabelRepasse.ForeColor = &H80000011
    LabelVenda.ForeColor = &H80000011
    LabelTransferência.ForeColor = &H80000011
    LabelDevolução.ForeColor = &H80000011
    LabelMatriz.Visible = False
    LabelFilial.Visible = False
End Sub

Private Sub Command_Entrada_Click()
    C1_Bloqueado "Filtro contextual legado; use pesquisa integrada por identidade"
End Sub

Private Sub Command_Entrada_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelRefazer.Visible = True
End Sub

Private Sub CommandButtonFiltroAguardando_Click()
    C1_Bloqueado "Filtro contextual legado; use pesquisa integrada por identidade"
End Sub

Private Sub CommandButtonFiltroAguardando_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelFiltroAguardando.ForeColor = &H80000012
End Sub

Private Sub CommandButtonRepasse_Click()
    C1_Bloqueado "Filtro contextual legado; use pesquisa integrada por identidade"
End Sub

Private Sub CommandButtonRepasse_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelRepasse.ForeColor = &H80000012
End Sub

Private Sub CommandButtonVenda_Click()
    C1_Bloqueado "Filtro contextual legado; use pesquisa integrada por identidade"
End Sub

Private Sub CommandButtonVenda_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelVenda.ForeColor = &H80000012
End Sub

Private Sub CommandButtonTransferência_Click()
    C1_Bloqueado "Filtro contextual legado; use pesquisa integrada por identidade"
End Sub

Private Sub CommandButtonTransferência_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelTransferência.ForeColor = &H80000012
End Sub

Private Sub CommandButtonDevolução_Click()
    C1_Bloqueado "Filtro contextual legado; use pesquisa integrada por identidade"
End Sub

Private Sub CommandButtonDevolução_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelDevolução.ForeColor = &H80000012
End Sub

Private Sub CommandButtonMatriz_Click()

    ThisWorkbook.Worksheets("Entradas").Select
End Sub

Private Sub CommandButtonMatriz_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelMatriz.Visible = True
End Sub

Private Sub CommandButtonFilial_Click()

    ThisWorkbook.Worksheets("Entradas Filial").Select
End Sub

Private Sub CommandButtonFilial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

    MouseCursor (32649)
    LabelFilial.Visible = True
End Sub


