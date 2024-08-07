Attribute VB_Name = "Module4"
'Uses the built in polynomial regression to generate a linear regression for the data
Function polyRegression(data() As Variant, degree As Integer) As Variant()
    Worksheets("data").Range("U1:U" & UBound(data)).Value = Application.Transpose(data)
    For i = 1 To UBound(data)
        Worksheets("data").Range("V" & i).Value = i
    Next i
    Dim inp As String
    inp = "=linest(data!U1:U" & UBound(data) & ",data!V1:V" & UBound(data) & "^{"
    For i = 1 To degree - 1
        inp = inp & i & ","
    Next i
    inp = inp & degree & "})"
    polyRegression = Application.Evaluate(inp)
End Function

'Calls polyregression to find the intercept of the best fit line
Function getLevel(data() As Variant) As Double
    getLevel = polyRegression(data, 1)(2)
End Function

'Calls polyregression to find the slope of the best fit line
Function getTrend(data() As Variant) As Double
    getTrend = polyRegression(data, 1)(1)
End Function

'Computes seasonal indices for winters multiplicative
'Values will be ~1
Function getSeasonal(data() As Variant, periods As Long) As Double()
    Dim res() As Double
    ReDim res(1 To periods)
    Dim avg As Double
    avg = 0
    For i = 1 To periods
        avg = avg + data(i)
    Next i
    avg = avg / periods
    For i = 1 To periods
        res(i) = data(i) / avg
    Next i
    getSeasonal = res
End Function

'Computes seasonal indices for winters additive
'Values will sum to ~0
Function getSeasonal2(data() As Variant, periods As Long) As Double()
    Dim res() As Double
    ReDim res(1 To periods)
    Dim level As Double
    level = getLevel(data)
    For i = 1 To periods:
        res(i) = data(i) - level
    Next i
    getSeasonal2 = res
End Function

'Computes the forecast for a n-period moving average
Function nMonthAvg(n As Variant, data() As Variant) As Variant()
    Dim avg As Double
    Dim res() As Variant
    ReDim res(1 To UBound(data) + 1)
    For i = n + 1 To UBound(data) + 1
        avg = 0
        For j = 0 To n - 1
            avg = avg + data(i - n + j)
        Next j
        res(i) = avg / n
    Next i
    nMonthAvg = res
End Function

'Computes the forecast for a given alpha value and given data
Function ExpAvg(a As Double, level As Double, data() As Variant) As Variant()
    Dim res() As Variant
    ReDim res(1 To UBound(data))
    res(1) = level
    For i = 2 To UBound(data)
        res(i) = a * data(i - 1) + (1 - a) * res(i - 1)
    Next i
    ExpAvg = res
End Function

'Computes the Holts trend method forecast for given alpha, beta, and data
Function Holts(a As Variant, beta As Variant, level As Double, trend As Double, data() As Variant, steps As Long) As Variant()
    Dim res() As Variant
    ReDim res(1 To UBound(data) + steps)
    Dim l, l2, b, b2 As Double
    res(1) = level + trend
    l2 = level
    b2 = trend
    For i = 2 To UBound(data)
        l = a * data(i - 1) + (1 - a) * (l2 + b2)
        b = beta * (l - l2) + (1 - beta) * b2
        l2 = l
        b2 = b
        res(i) = l + b
    Next i
    For i = UBound(data) + 1 To UBound(data) + steps
        res(i) = l + (i - UBound(data)) * b
    Next i
    Holts = res
End Function

'Computes the winters additive method given alpha, beta, gamma, and data
Function wintersA(a As Variant, beta As Variant, gamma As Variant, level As Double, trend As Double, seasons() As Double, data() As Variant, steps As Long) As Variant()
    Dim res() As Variant
    Dim periods As Long
    periods = UBound(seasons)
    ReDim res(1 To UBound(data) + steps)
    Dim l, l2, b, b2 As Double
    res(1) = level + trend + seasons(1)
    l2 = level
    b2 = trend
    For i = 2 To UBound(data)
        l = a * (data(i - 1) - seasons((i - 1) Mod periods + 1)) + (1 - a) * (l2 + b2)
        b = beta * (l - l2) + (1 - beta) * b2
        seasons((i - 1) Mod periods + 1) = gamma * (data(i - 1) - l2 - b2) + (1 - gamma) * seasons((i - 1) Mod periods + 1)
        l2 = l
        b2 = b
        res(i) = l + b + seasons((i - 1) Mod periods + 1)
    Next i
    For i = UBound(data) + 1 To UBound(data) + steps
        res(i) = l + (i - UBound(data)) * b + seasons((i - 1) Mod periods + 1)
    Next i
    wintersA = res
End Function

'Same as prior, except with winters multiplicative
Function wintersM(a As Variant, beta As Variant, gamma As Variant, level As Double, trend As Double, seasons() As Double, data() As Variant, steps As Long) As Variant()
    Dim res() As Variant
    Dim periods As Long
    periods = UBound(seasons)
    ReDim res(1 To UBound(data) + steps)
    Dim l, l2, b, b2 As Double
    res(1) = (level + trend) * seasons(1)
    l2 = level
    b2 = trend
    For i = 2 To UBound(data)
        l = a * data(i - 1) / seasons((i - 1) Mod periods + 1) + (1 - a) * (l2 + b2)
        b = beta * (l - l2) + (1 - beta) * b2
        seasons((i - 1) Mod periods + 1) = gamma * data(i - 1) / (l + b) + (1 - gamma) * seasons((i - 1) Mod periods + 1)
        l2 = l
        b2 = b
        res(i) = (l + b) * seasons((i - 1) Mod periods + 1)
    Next i
    For i = UBound(data) + 1 To UBound(data) + steps
        res(i) = (l + (i - UBound(data)) * b) * seasons((i - 1) Mod periods + 1)
    Next i
    wintersM = res
    Exit Function
End Function

'Computes the Root Mean Squared Error given actual data and calculated forecast
Function RMSE(data() As Variant, calc() As Variant) As Double
    Dim res As Double
    res = 0
    For i = 1 To UBound(data)
        res = res + (data(i) - calc(i)) ^ 2
    Next i
    RMSE = (res / UBound(data)) ^ 0.5
End Function

'Same as previous but with Mean Absolute Error
Function MAE(data() As Variant, calc() As Variant) As Double
    Dim res As Double
    res = 0
    For i = 1 To UBound(data)
        res = res + Abs(data(i) - calc(i))
    Next i
    MAE = res / UBound(data)
End Function

'Computes tracking signal
Function Tracking(data() As Variant, calc() As Variant) As Double
    Dim res As Double
    res = 0
    For i = 1 To UBound(data)
        res = res + data(i) - calc(i)
    Next i
    Tracking = res / MAE(data, calc)
End Function

'Finds the optimal alpha value for the data, num is number of values to test, and refinement is cycles of precision
Function calcAlpha(data(), num As Long, refinement As Long) As Variant()
    Dim bestA As Double
    Dim bestRMSE As Double
    Dim pt(0 To 1) As Double
    Dim level As Double
    level = getLevel(data)
    bestRMSE = -1
    bestA = 0.5
    Dim count As Long
    Dim tempa As Double
    Dim radius As Double
    radius = 0.5
    count = 0
    For i = 1 To refinement
        tempa = bestA
        pt(0) = IIf(tempa - radius < 0, 0, tempa - radius)
        'Checks for float underflow
        While pt(0) < tempa + radius And pt(0) <= 1 And radius > 0.00000000000001
            If count Mod 100 = 0 Then
                DoEvents
                Application.StatusBar = "EXP: " & Int(100 * count / (refinement * num)) & "%"
            End If
            count = count + 1
            pt(1) = RMSE(data, ExpAvg(pt(0), level, data))
            If pt(1) < bestRMSE Or bestRMSE = -1 Then
                bestRMSE = pt(1)
                bestA = pt(0)
            End If
            pt(0) = pt(0) + 2 * radius / num
        Wend
        radius = 2 * radius / num
    Next i
    Worksheets("FORECAST").Range("H4").Value = bestA
    Worksheets("FORECAST").Range("I4").Value = bestRMSE
    calcAlpha = ExpAvg(bestA, level, data)
End Function

'Same as previous function but with holt's trend method
'Steps is the number of data points to extrapolate outwards
Function calcHolts(data(), num As Long, refinement As Long, steps As Long) As Variant()
    Dim bestA As Double
    Dim bestB As Double
    Dim bestRMSE As Double
    Dim level As Double
    Dim trend As Double
    level = getLevel(data)
    trend = getTrend(data)
    bestRMSE = -1
    bestA = 0.5
    bestB = 0.5
    Dim inp(0 To 2) As Variant
    Dim count As Long
    count = 0
    Dim radius As Double
    Dim tempa As Double
    Dim tempb As Double
    radius = 0.5
    For i = 1 To refinement
        tempa = bestA
        tempb = bestB
        inp(0) = IIf(tempa - radius < 0, 0, tempa - radius)
        While inp(0) < tempa + radius And inp(0) < 1
            inp(1) = IIf(tempb - radius < 0, 0, tempb - radius)
            While inp(1) < bestB + radius And inp(1) < 1
                inp(2) = RMSE(data, Holts(inp(0), inp(1), level, trend, data, 0))
                If count Mod 100 = 0 Then
                    DoEvents
                    Application.StatusBar = "Holts: " & Int(100 * count / (num * (refinement + 1))) & "%"
                End If
                count = count + 1
                If inp(2) < bestRMSE Or bestRMSE = -1 Then
                    bestRMSE = inp(2)
                    bestA = inp(0)
                    bestB = inp(1)
                End If
                inp(1) = inp(1) + 2 * radius / (num ^ (1 / 2))
            Wend
            inp(0) = inp(0) + 2 * radius / (num ^ (1 / 2))
        Wend
        radius = radius / num ^ (1 / 2)
    Next i
    Debug.Print "Holts: " & bestA, bestB, bestRMSE
    Worksheets("FORECAST").Range("H4").Value = bestA
    Worksheets("FORECAST").Range("I4").Value = bestB
    Worksheets("FORECAST").Range("J4").Value = bestRMSE
    calcHolts = Holts(bestA, bestB, level, trend, data, steps)
End Function

'Same as previous except handles both additive and multiplicative holt winter's
Function calcWinters(data() As Variant, num As Long, refinement As Long, periods As Long, mul As Boolean, steps As Long) As Variant()
    Dim bestA As Double
    Dim bestB As Double
    Dim bestG As Double
    Dim bestRMSE As Double
    Dim maxVal As Double
    Dim level As Double
    Dim trend As Double
    Dim seasons() As Double
    Dim seasons2() As Double
    level = getLevel(data)
    trend = getTrend(data)
    periods = IIf(periods > UBound(data), UBound(data), periods)
    If mul Then
        seasons = getSeasonal(data, periods)
    Else
        seasons = getSeasonal2(data, periods)
    End If
    bestRMSE = -1
    maxVal = 0
    bestA = 0.5
    bestB = 0.5
    bestG = 0.5
    Dim inp(0 To 3) As Variant
    Dim radius As Double
    Dim count As Long
    count = 0
    radius = 0.5
    For i = 1 To refinement
        tempa = bestA
        tempb = bestB
        tempg = bestG
        inp(0) = IIf(tempa - radius < 0, 0, tempa - radius)
        While inp(0) < tempa + radius And inp(0) <= 1
            inp(1) = IIf(tempb - radius < 0, 0, tempb - radius)
            While inp(1) < tempb + radius And inp(1) <= 1
                inp(2) = IIf(tempg - radius * (1 - inp(0)) < 0, 0, tempg - radius * (1 - inp(0)))
                While inp(2) < tempg + radius * (1 - inp(0)) And inp(2) <= 1 - inp(0)
                    If count Mod 100 = 0 Then
                        DoEvents
                        Application.StatusBar = "Winters: " & Int(100 * count / (num * (refinement))) & "%"
                    End If
                    count = count + 1
                    seasons2 = seasons
                    If mul Then
                        inp(3) = RMSE(data, wintersM(inp(0), inp(1), inp(2), level, trend, seasons2, data, 0))
                    Else
                        inp(3) = RMSE(data, wintersA(inp(0), inp(1), inp(2), level, trend, seasons2, data, 0))
                    End If
                    If inp(3) < bestRMSE Or bestRMSE = -1 Then
                        bestRMSE = inp(3)
                        bestA = inp(0)
                        bestB = inp(1)
                        bestG = inp(2)
                    End If
                    inp(2) = inp(2) + 2 * radius * (1 - inp(0)) / (num ^ (1 / 3))
                Wend
                inp(1) = inp(1) + 2 * radius / (num ^ (1 / 3))
            Wend
            inp(0) = inp(0) + 2 * radius / (num ^ (1 / 3))
        Wend
        radius = 2 * radius / (num ^ (1 / 3))
    Next i
    Debug.Print "Winters: " & bestA, bestB, bestG, bestRMSE
    Worksheets("FORECAST").Range("H4").Value = bestA
    Worksheets("FORECAST").Range("I4").Value = bestB
    Worksheets("FORECAST").Range("J4").Value = bestG
    Worksheets("FORECAST").Range("K4").Value = bestRMSE
    If mul Then
        calcWinters = wintersM(bestA, bestB, bestG, level, trend, seasons, data, steps)
    Else
        calcWinters = wintersA(bestA, bestB, bestG, level, trend, seasons, data, steps)
    End If
End Function

'Runs all methods above and creates the table used to show error margins
Sub main()
    Dim vals() As Variant
    Dim res() As Variant
    Dim lr As Long
    lr = Worksheets("data").Range("J" & Rows.count).End(xlUp).Row
    vals = Application.Transpose(Worksheets("data").Range("K2:K" & lr).Value)
    For i = 1 To UBound(vals): vals(i) = IIf(vals(i) = 0, 0.001, vals(i)): Next i
    With Worksheets("data")
        .Range("L2", .Cells(.Rows.count, .Columns.count)).Clear
    End With
    With Worksheets("FORECAST")
        .Range("G27:N30").Clear
        .Range("G27").Value = "Error"
        .Range("H27").Value = "3-Month Avg"
        .Range("I27").Value = "6-Month Avg"
        .Range("J27").Value = "12-Month Avg"
        .Range("K27").Value = "Exp Smoothing"
        .Range("L27").Value = "Holt's"
        .Range("M27").Value = "Winter's ADD"
        .Range("N27").Value = "Winter's MUL"
        .Range("G28").Value = "RMSE"
        .Range("G29").Value = "MAE"
        .Range("G30").Value = "Tracking"
        .ListObjects.Add(xlSrcRange, .Range("G27:N30"), , xlYes).Name = "Errors"
        .ListObjects("Errors").TableStyle = "TableStyleMedium12"
    End With
    res = nMonthAvg(3, vals)
    Worksheets("data").Range("L2:L" & lr).Value = Application.Transpose(res)
    Worksheets("FORECAST").Range("H28").Value = RMSE(vals, res)
    Worksheets("FORECAST").Range("H29").Value = MAE(vals, res)
    Worksheets("FORECAST").Range("H30").Value = Tracking(vals, res)
    res = nMonthAvg(6, vals)
    Worksheets("data").Range("M2:M" & lr).Value = Application.Transpose(res)
    Worksheets("FORECAST").Range("I28").Value = RMSE(vals, res)
    Worksheets("FORECAST").Range("I29").Value = MAE(vals, res)
    Worksheets("FORECAST").Range("I30").Value = Tracking(vals, res)
    res = nMonthAvg(12, vals)
    Worksheets("data").Range("N2:N" & lr).Value = Application.Transpose(res)
    Worksheets("FORECAST").Range("J28").Value = RMSE(vals, res)
    Worksheets("FORECAST").Range("J29").Value = MAE(vals, res)
    Worksheets("FORECAST").Range("J30").Value = Tracking(vals, res)
    res = calcAlpha(vals, Sheet4.TextBox1.Value, Sheet4.TextBox2.Value)
    Worksheets("data").Range("O2:O" & lr).Value = Application.Transpose(res)
    Worksheets("FORECAST").Range("K28").Value = RMSE(vals, res)
    Worksheets("FORECAST").Range("K29").Value = MAE(vals, res)
    Worksheets("FORECAST").Range("K30").Value = Tracking(vals, res)
    res = calcHolts(vals, Sheet4.TextBox1.Value, Sheet4.TextBox2.Value, Sheet4.TextBox4.Value)
    Worksheets("data").Range("P2:P" & lr).Value = Application.Transpose(res)
    Worksheets("FORECAST").Range("L28").Value = RMSE(vals, res)
    Worksheets("FORECAST").Range("L29").Value = MAE(vals, res)
    Worksheets("FORECAST").Range("L30").Value = Tracking(vals, res)
    res = calcWinters(vals, Sheet4.TextBox1.Value, Sheet4.TextBox2.Value, Sheet4.TextBox3.Value, False, 0)
    Worksheets("data").Range("Q2:Q" & lr).Value = Application.Transpose(res)
    Worksheets("FORECAST").Range("M28").Value = RMSE(vals, res)
    Worksheets("FORECAST").Range("M29").Value = MAE(vals, res)
    Worksheets("FORECAST").Range("M30").Value = Tracking(vals, res)
    res = calcWinters(vals, Sheet4.TextBox1.Value, Sheet4.TextBox2.Value, Sheet4.TextBox3.Value, True, 0)
    Worksheets("data").Range("R2:R" & lr).Value = Application.Transpose(res)
    Worksheets("FORECAST").Range("N28").Value = RMSE(vals, res)
    Worksheets("FORECAST").Range("N29").Value = MAE(vals, res)
    Worksheets("FORECAST").Range("N30").Value = Tracking(vals, res)
    Worksheets("FORECAST").ChartObjects("Forecast").Activate
    lr = Worksheets("data").Range("J" & Rows.count).End(xlUp).Row
    ActiveChart.SetSourceData Source:=Worksheets("data").Range("J1:R" & lr)
    Application.StatusBar = "Ready"
End Sub



