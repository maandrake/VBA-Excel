Attribute VB_Name = "mdlCondCel"
Option Explicit

Function ActiveCondition(Rng As Range) As Integer
Dim Ndx As Long
Dim FC As FormatCondition
Dim Temp As Variant
Dim Temp2 As Variant

If Rng.FormatConditions.count = 0 Then
    ActiveCondition = 0
Else
    For Ndx = 1 To Rng.FormatConditions.count
        Set FC = Rng.FormatConditions(Ndx)
        Select Case FC.Type
            Case xlCellValue
            Select Case FC.Operator
                Case xlBetween
                    Temp = GetStrippedValue(FC.Formula1)
                    Temp2 = GetStrippedValue(FC.Formula2)
                    If IsNumeric(Temp) Then
                       If CDbl(Rng.value) >= CDbl(FC.Formula1) And _
                           CDbl(Rng.value) <= CDbl(FC.Formula2) Then
                           ActiveCondition = Ndx
                           Exit Function
                       End If
                   Else
                      If Rng.value >= Temp And _
                         Rng.value <= Temp2 Then
                         ActiveCondition = Ndx
                         Exit Function
                      End If
                   End If

                Case xlGreater
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                       If CDbl(Rng.value) > CDbl(FC.Formula1) Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    Else
                       If Rng.value > Temp Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If

                Case xlEqual
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                       If CDbl(Rng.value) = CDbl(FC.Formula1) Then
                           ActiveCondition = Ndx
                           Exit Function
                       End If
                    Else
                       If Temp = Rng.value Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If


                Case xlGreaterEqual
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                       If CDbl(Rng.value) >= CDbl(FC.Formula1) Then
                           ActiveCondition = Ndx
                           Exit Function
                       End If
                    Else
                       If Rng.value >= Temp Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If

               
                Case xlLess
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                        If CDbl(Rng.value) < CDbl(FC.Formula1) Then
                           ActiveCondition = Ndx
                           Exit Function
                        End If
                    Else
                        If Rng.value < Temp Then
                           ActiveCondition = Ndx
                           Exit Function
                        End If
                    End If

                Case xlLessEqual
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                       If CDbl(Rng.value) <= CDbl(FC.Formula1) Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    Else
                       If Rng.value <= Temp Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If


                Case xlNotEqual
                    Temp = GetStrippedValue(FC.Formula1)
                    If IsNumeric(Temp) Then
                       If CDbl(Rng.value) <> CDbl(FC.Formula1) Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    Else
                       If Temp <> Rng.value Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If

               Case xlNotBetween
                    Temp = GetStrippedValue(FC.Formula1)
                    Temp2 = GetStrippedValue(FC.Formula2)
                    If IsNumeric(Temp) Then
                       If Not (CDbl(Rng.value) <= CDbl(FC.Formula1)) And _
                          (CDbl(Rng.value) >= CDbl(FC.Formula2)) Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    Else
                       If Not Rng.value <= Temp And _
                          Rng.value >= Temp2 Then
                          ActiveCondition = Ndx
                          Exit Function
                       End If
                    End If
             
               Case Else
                    Debug.Print "UNKNOWN OPERATOR"
           End Select


        Case xlExpression
            If Application.Evaluate(FC.Formula1) Then
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
 
Function ColorIndexOfCF(Rng As Range, _
    Optional OfText As Boolean = False) As Integer

Dim AC As Integer
AC = ActiveCondition(Rng)
If AC = 0 Then
    If OfText = True Then
       ColorIndexOfCF = Rng.Font.ColorIndex
    Else
       ColorIndexOfCF = Rng.Interior.ColorIndex
    End If
Else
    If OfText = True Then
       ColorIndexOfCF = Rng.FormatConditions(AC).Font.ColorIndex
    Else
       ColorIndexOfCF = Rng.FormatConditions(AC).Interior.ColorIndex
    End If
End If

End Function
 
Function ColorOfCF(Rng As Range, Optional OfText As Boolean = False) As Long

Dim AC As Integer
AC = ActiveCondition(Rng)
If AC = 0 Then
    If OfText = True Then
       ColorOfCF = Rng.Font.color
    Else
       ColorOfCF = Rng.Interior.color
    End If
Else
    If OfText = True Then
       ColorOfCF = Rng.FormatConditions(AC).Font.color
    Else
       ColorOfCF = Rng.FormatConditions(AC).Interior.color
    End If
End If

End Function

Function GetStrippedValue(CF As String) As String
    Dim Temp As String
    If InStr(1, CF, "=", vbTextCompare) Then
       Temp = Mid(CF, 3, Len(CF) - 3)
       If Left(Temp, 1) = "=" Then
           Temp = Mid(Temp, 2)
       End If
    Else
       Temp = CF
    End If
    GetStrippedValue = Temp
End Function
 
Function CountOfCF(InRange As Range, _
    Optional Condition As Integer = -1) As Long
    Dim count As Long
    Dim Rng As Range
    Dim FCNum As Integer

    For Each Rng In InRange.Cells
        FCNum = ActiveCondition(Rng)
        If FCNum > 0 Then
            If Condition = -1 Or Condition = FCNum Then
                count = count + 1
            End If
        End If
    Next Rng
    CountOfCF = count
End Function
Function SumByCFColorIndex(Rng As Range, CI As Integer) As Double
    Dim r As Range
    Dim total As Double
    For Each r In Rng.Cells
        If ColorIndexOfCF(r, False) = CI Then
            total = total + r.value
        End If
    Next r
    SumByCFColorIndex = total
End Function


