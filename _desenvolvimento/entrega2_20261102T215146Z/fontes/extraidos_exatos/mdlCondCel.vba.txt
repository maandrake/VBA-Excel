Attribute VB_Name = "mdlCondCel"
Option Explicit

Function ActiveCondition(rng As Range) As Integer
Dim Ndx As Long
Dim FC As FormatCondition
Dim Temp As Variant
Dim Temp2 As Variant

If rng.FormatConditions.count = 0 Then
    ActiveCondition = 0
Else
    For Ndx = 1 To rng.FormatConditions.count
        Set FC = rng.FormatConditions(Ndx)
        Select Case FC.Type
            Case xlCellValue
            Select Case FC.Operator
                Case xlBetween
                    Temp = GetStrippedValue(FC.Formula1)
                    Temp2 = GetStrippedValue(FC.Formula2)
                    If IsNumeric(Temp) Then
                       If CDbl(rng.value) >= CDbl(Temp) And _
                           CDbl(rng.value) <= CDbl(Temp2) Then
                           ActiveCondition = Ndx
                           Exit Function
                       End If
                   Else
                      If rng.value >= Temp And _
                         rng.value <= Temp2 Then
                         ActiveCondition = Ndx
                         Exit Function
                      End If
                   End If

                Case xlGreater
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                       If CDbl(rng.value) > CDbl(Temp) Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    Else
                       If rng.value > Temp Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If

                Case xlEqual
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                       If CDbl(rng.value) = CDbl(Temp) Then
                           ActiveCondition = Ndx
                           Exit Function
                       End If
                    Else
                       If Temp = rng.value Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If


                Case xlGreaterEqual
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                       If CDbl(rng.value) >= CDbl(Temp) Then
                           ActiveCondition = Ndx
                           Exit Function
                       End If
                    Else
                       If rng.value >= Temp Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If

               
                Case xlLess
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                        If CDbl(rng.value) < CDbl(Temp) Then
                           ActiveCondition = Ndx
                           Exit Function
                        End If
                    Else
                        If rng.value < Temp Then
                           ActiveCondition = Ndx
                           Exit Function
                        End If
                    End If

                Case xlLessEqual
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                       If CDbl(rng.value) <= CDbl(Temp) Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    Else
                       If rng.value <= Temp Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If


                Case xlNotEqual
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                       If CDbl(rng.value) <> CDbl(Temp) Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    Else
                       If Temp <> rng.value Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If

               Case xlNotBetween
                    Temp = GetStrippedValue(FC.Formula1)
                    Temp2 = GetStrippedValue(FC.Formula2)
                    If IsNumeric(Temp) Then
                       If CDbl(rng.value) < CDbl(Temp) Or _
                          CDbl(rng.value) > CDbl(Temp2) Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    Else
                       If rng.value < Temp Or rng.value > Temp2 Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If
             
               Case Else
                    Debug.Print "UNKNOWN OPERATOR"
           End Select


        Case xlExpression
            If rng.Worksheet.Evaluate(FC.Formula1) Then
               ActiveCondition = Ndx
               Exit Function
            End If

        Case Else
            Debug.Print "UNKNOWN TYPE"
       End Select

    Next Ndx

End If

ActiveCondition = 0



End Function
 
Function ColorIndexOfCF(rng As Range, _
    Optional OfText As Boolean = False) As Integer

Dim AC As Integer
AC = ActiveCondition(rng)
If AC = 0 Then
    If OfText = True Then
       ColorIndexOfCF = rng.Font.ColorIndex
    Else
       ColorIndexOfCF = rng.Interior.ColorIndex
    End If
Else
    If OfText = True Then
       ColorIndexOfCF = rng.FormatConditions(AC).Font.ColorIndex
    Else
       ColorIndexOfCF = rng.FormatConditions(AC).Interior.ColorIndex
    End If
End If

End Function
 
Function ColorOfCF(rng As Range, Optional OfText As Boolean = False) As Long

Dim AC As Integer
AC = ActiveCondition(rng)
If AC = 0 Then
    If OfText = True Then
       ColorOfCF = rng.Font.color
    Else
       ColorOfCF = rng.Interior.color
    End If
Else
    If OfText = True Then
       ColorOfCF = rng.FormatConditions(AC).Font.color
    Else
       ColorOfCF = rng.FormatConditions(AC).Interior.color
    End If
End If

End Function

Function GetStrippedValue(CF As String) As String
    Dim Temp As String
    ' Formula1 pode devolver =5 ou ="TEXTO"; remova so o prefixo.
    Temp = Trim$(CF)
    If Left$(Temp, 1) = "=" Then Temp = mid$(Temp, 2)
    If Len(Temp) >= 2 Then
        If Left$(Temp, 1) = Chr$(34) And Right$(Temp, 1) = Chr$(34) Then
            Temp = mid$(Temp, 2, Len(Temp) - 2)
            Temp = Replace(Temp, Chr$(34) & Chr$(34), Chr$(34))
        End If
    End If
    GetStrippedValue = Temp
End Function
 
Function CountOfCF(InRange As Range, _
    Optional Condition As Integer = -1) As Long
    Dim count As Long
    Dim rng As Range
    Dim FCNum As Integer

    For Each rng In InRange.Cells
        FCNum = ActiveCondition(rng)
        If FCNum > 0 Then
            If Condition = -1 Or Condition = FCNum Then
                count = count + 1
            End If
        End If
    Next rng
    CountOfCF = count
End Function
Function SumByCFColorIndex(rng As Range, CI As Integer) As Double
    Dim r As Range
    Dim total As Double
    For Each r In rng.Cells
        If ColorIndexOfCF(r, False) = CI Then
            total = total + r.value
        End If
    Next r
    SumByCFColorIndex = total
End Function


