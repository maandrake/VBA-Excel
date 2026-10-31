Attribute VB_Name = "modPreencherFormulario"
Public Sub PreencherDadosEntradas( _
    ByVal Frm As Object, _
    ByVal Lista As Object, _
    ByVal linha As Long)

    With Frm

        .LabelTabela1.CAPTION = Lista.List(linha, 0)

        .LabelTabela1.BorderStyle = fmBorderStyleSingle
        .LabelTabela1.BorderColor = &HFF&
        .LabelTabela1.ForeColor = &HFF&

        .ComboBox_Operação.value = Lista.List(linha, 2)
        .N_Nfe.value = Lista.List(linha, 1)
        .UN.value = Lista.List(linha, 3)
        .COD.value = Lista.List(linha, 5)
        .ComboBox_VEÍCULO.value = Lista.List(linha, 4)
        .PLACA.value = Lista.List(linha, 6)
        .ComboBox_COR.value = Lista.List(linha, 8)
        .Combo_ANO_MOD.value = Lista.List(linha, 7)

        'Valor
        If IsNumeric(Lista.List(linha, 9)) Then

            .txt_Valor.value = Format( _
                CDbl(Lista.List(linha, 9)), _
                "Currency" _
            )

        Else

            .txt_Valor.value = Lista.List(linha, 9)

        End If

        'Data
        .Data.value = Lista.List(linha, 10)

        'Outro envio
        'Somente se houver a coluna 11
        If Lista.ColumnCount > 11 Then
            .ComboBox_OutroEnvio.value = Lista.List(linha, 11)
        End If

    End With

End Sub

Public Sub PreencherDadosEntradasFilial( _
    ByVal Frm As Object, _
    ByVal Lista As Object, _
    ByVal linha As Long)

    With Frm

        .LabelTabelaFilial1.CAPTION = Lista.List(linha, 0)

        .LabelTabelaFilial1.BorderStyle = fmBorderStyleSingle
        .LabelTabelaFilial1.BorderColor = &HFF&
        .LabelTabelaFilial1.ForeColor = &HFF&

        .ComboBox_OperaçãoFilial.value = Lista.List(linha, 2)
        .N_NfeFilial.value = Lista.List(linha, 1)
        .UNFilial.value = Lista.List(linha, 3)
        .CODFilial.value = Lista.List(linha, 5)
        .ComboBox_VEÍCULOFilial.value = Lista.List(linha, 4)
        .PLACAFilial.value = Lista.List(linha, 6)
        .ComboBox_CORFilial.value = Lista.List(linha, 8)
        .Combo_ANO_MODFilial.value = Lista.List(linha, 7)

        'Valor
        If IsNumeric(Lista.List(linha, 9)) Then

            .txt_ValorFilial.value = Format( _
                CDbl(Lista.List(linha, 9)), _
                "Currency" _
            )

        Else

            .txt_ValorFilial.value = Lista.List(linha, 9)

        End If

        'Data
        .DataFilial.value = Lista.List(linha, 10)

        '==================================================
        ' OUTRO ENVIO FILIAL
        '==================================================
        If Lista.ColumnCount > 11 Then

            .ComboBox_OutroEnvioFilial.value = _
                Lista.List(linha, 11)

        End If

    End With

End Sub
