Attribute VB_Name = "Module5"
'Creates the conditional formatting to show the red highlighting over the tracking signal
'Darker red with values further from 0, and white between -4 and 4
Sub CondFormat()
Attribute CondFormat.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CondFormat Macro
'

'
    Dim rng As Range
    Set rng = Range("H30:N30")
    rng.FormatConditions.AddColorScale ColorScaleType:=3
    rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
    rng.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    rng.FormatConditions(1).ColorScaleCriteria(1).Value = -5
    With rng.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 255
        .TintAndShade = 0
    End With
    rng.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueNumber
    rng.FormatConditions(1).ColorScaleCriteria(2).Value = 0
    With rng.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = RGB(255, 255, 255)
        .TintAndShade = 0
    End With
    rng.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueNumber
    rng.FormatConditions(1).ColorScaleCriteria(3).Value = 5
    With rng.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 255
        .TintAndShade = 0
    End With
    Application.CutCopyMode = False
End Sub

'Function to simulate confidence intervals for winter's multiplicative
'Uses given data to forecast one point outwards, then randomly selects an error from the data to add to this point.
'Then repeats this cycle until the steps outward is complete. Do this 1000 times. At each time set the confidence
'intervals to be the value at 950 and 50.
Function bootStrapPI(data() As Variant, calc() As Variant, steps As Long, a As Double, b As Double, g As Double) As Variant()
    Dim res() As Variant
    Dim mydata() As Variant
    ReDim res(1 To UBound(data))
    For i = 1 To UBound(data)
        res(i) = data(i) - calc(i)
    Next i
    Dim final() As Variant
    ReDim final(1 To steps, 1 To 1000)
    Dim level As Double
    Dim trend As Double
    Dim seasons() As Double
    Dim seasons2() As Double
    Dim val As Double
    Dim index As Long
    level = Module4.getLevel(data)
    trend = Module4.getTrend(data)
    seasons = Module4.getSeasonal(data, Sheet4.TextBox3.Value)
    For i = 1 To 1000
        mydata = data
        For j = 1 To steps
            seasons2 = seasons
            val = Module4.wintersM(a, b, g, level, trend, seasons2, mydata, 1)(UBound(mydata) + 1)
            ReDim Preserve mydata(1 To UBound(mydata) + 1)
            index = -Int(-Rnd * UBound(res))
            final(j, i) = val + res(index)
            mydata(UBound(mydata)) = final(j, i)
        Next j
    Next i
    Dim last() As Variant
    Dim temp(1 To 1000) As Variant
    ReDim last(1 To UBound(data) + steps, 1 To 2)
    For i = 1 To steps
        For j = 1 To 1000
            temp(j) = final(i, j)
        Next j
        Call QuickSort(temp, 1, 1000)
        last(i + UBound(data), 1) = temp(950)
        last(i + UBound(data), 2) = temp(50)
    Next i
    bootStrapPI = last
End Function

'Generic quicksort algorithm used to sort the data at the end of generating the confidence intervals
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub
