Attribute VB_Name = "MóduloPesquisa"

Option Explicit

Private Const LINHA_INICIAL_DADOS As Long = 2
Private Const COL_INICIO As Long = 1        ' Coluna A
Private Const COL_FIM As Long = 11         ' Coluna K - Entradas
Private Const COL_FIM_FILIAL As Long = 12  ' Coluna L - Entradas Filial

'==========================================================
' CONFIGURAR LISTBOX DE PESQUISA
'==========================================================
Public Sub ConfigurarListBoxPesquisa(ByVal Lst As Object)
    C1_ConfigLista Lst
End Sub

'==========================================================
' PESQUISAR ENTRADAS E ENTRADAS FILIAL
'==========================================================
Public Sub PesquisarEntradasRapido( _
    ByVal TextoPesquisa As String, _
    ByVal Lst As Object)
    C1_Pesquisar TextoPesquisa, Lst
End Sub

'==========================================================
' ABRIR REGISTRO SELECIONADO NA LISTBOX
'==========================================================
Public Sub AbrirRegistroDaLista(ByVal Lst As Object)
    C1_Navegar Lst
End Sub

'==========================================================
' IDENTIFICAR ÚLTIMA LINHA UTILIZADA
'==========================================================
Private Function UltimaLinhaUsada( _
    ByVal WS As Worksheet, _
    ByVal coluna As String) As Long

    UltimaLinhaUsada = WS.Cells( _
        WS.rows.count, coluna _
    ).End(xlUp).row

End Function

'==========================================================
' TRANSPOR MATRIZ PARA LISTBOX
'==========================================================
Private Function TransporParaListBox( _
    ByVal Matriz As Variant) As Variant

    Dim i As Long
    Dim j As Long

    Dim Temp() As Variant

    ReDim Temp( _
        0 To UBound(Matriz, 2), _
        0 To UBound(Matriz, 1) _
    )

    For i = 0 To UBound(Matriz, 1)

        For j = 0 To UBound(Matriz, 2)

            Temp(j, i) = Matriz(i, j)

        Next j

    Next i

    TransporParaListBox = Temp

End Function

'==========================================================
' IR PARA CÉLULA DA NFe
'==========================================================
Public Sub IrParaCelulaNFe( _
    ByVal NomePlanilha As String, _
    ByVal NumeroNFe As String)
    C1_Bloqueado "Navegacao somente por NF-e e ambigua; use pesquisa por identidade"
End Sub

