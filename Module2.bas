Attribute VB_Name = "Module2"
'Performs the aggregation based on period chosen by user and updates the chart to show data
Function finalData()
    Dim lastVal As Variant
    lastVal = 0
    On Error Resume Next
    Dim lr As Long
    lr = Worksheets("data").Range("A" & Rows.count).End(xlUp).Row
    lastVal = CDate(Worksheets("data").Range("A" & lr).Value)
    On Error GoTo 0
    Dim firstVal As Variant
    firstVal = CDate(Worksheets("data").Range("A2").Value)
    lr = Worksheets("data").Range("J" & Rows.count).End(xlUp).Row
    lr = IIf(lr = 1, 2, lr)
    On Error Resume Next
    Worksheets("data").Range("J2:K" & lr).Rows.delete
    On Error GoTo 0
    If lastVal = 0 Then
        Exit Function
    End If
    If Sheet3.GroupBy.Value = "Monthly" Then
        Call Monthly(DateDiff("m", DateValue("1/1/2018"), firstVal), DateDiff("m", DateValue("1/1/2018"), lastVal) + 1)
    ElseIf Sheet3.GroupBy.Value = "Quarterly" Then
        Call Quarterly(DateDiff("q", DateValue("1/1/2018"), firstVal), DateDiff("q", DateValue("1/1/2018"), lastVal) + 1)
    Else
        Call Yearly(DateDiff("yyyy", DateValue("1/1/2018"), firstVal), DateDiff("yyyy", DateValue("1/1/2018"), lastVal) + 1)
    End If
    Worksheets("ANALYSIS").ChartObjects("MainChart").Activate
    lr = Worksheets("data").Range("J" & Rows.count).End(xlUp).Row
    ActiveChart.SetSourceData Source:=Worksheets("data").Range("J1:K" & lr)
End Function

'Performs monthly aggregation of data
Function Monthly(m1 As Integer, m2 As Integer)
    Dim val() As Variant
    ReDim val(1 To (m2 - m1), 1 To 2)
    Dim test As String
    y = 2018 + m1 \ 12
    Dim r As Long
    Dim lr As Long
    lr = Worksheets("data").Range("A" & Rows.count).End(xlUp).Row
    r = 2
    For i = m1 To m2 - 1
        test = (i Mod 12) + 1 & "/1/" & y
        test = CDate(test)
        val(i - m1 + 1, 1) = CDate(Excel.Application.WorksheetFunction.EoMonth(test, 0))
        val(i - m1 + 1, 2) = 0
        While Month(Worksheets("data").Range("A" & r).Value) = Month(test) And Year(Worksheets("data").Range("A" & r).Value) = Year(test) And r <= lr
            val(i - m1 + 1, 2) = val(i - m1 + 1, 2) + Worksheets("data").Range("B" & r).Value
            r = r + 1
        Wend
        If (i Mod 12 = 11) Then
            y = y + 1
        End If
    Next i
    Worksheets("data").Range("J2:K" & (m2 - m1) + 1).Value = val
End Function

'Quartlerly aggregation
Function Quarterly(q1 As Integer, q2 As Integer)
    Dim val() As Variant
    ReDim val(1 To (q2 - q1), 1 To 2)
    Dim test As String
    y = 2018 + q1 \ 4
    Dim r As Long
    Dim lr As Long
    lr = Worksheets("data").Range("a" & Rows.count).End(xlUp).Row
    r = 2
    For i = q1 To q2 - 1
        test = (i Mod 4 + 1) * 3 & "/1/" & y
        test = CDate(test)
        val(i - q1 + 1, 1) = CDate(Excel.Application.WorksheetFunction.EoMonth(test, 0))
        val(i - q1 + 1, 2) = 0
        While (Month(Worksheets("data").Range("A" & r).Value) + 2) \ 3 = (Month(test) + 2) \ 3 And Year(Worksheets("data").Range("A" & r).Value) = Year(test) And r <= lr
            DoEvents
            val(i - q1 + 1, 2) = val(i - q1 + 1, 2) + Worksheets("data").Range("B" & r).Value
            r = r + 1
        Wend
        If (i Mod 4 = 3) Then
            y = y + 1
        End If
    Next i
    Worksheets("data").Range("J2:K" & (q2 - q1) + 1).Value = val
End Function

'Yearly aggregation
Function Yearly(ye1 As Integer, ye2 As Integer)
    Dim val() As Variant
    ReDim val(1 To (ye2 - ye1), 1 To 2)
    y = 2018 + ye1
    Dim r As Long
    Dim lr As Long
    lr = Worksheets("data").Range("A" & Rows.count).End(xlUp).Row
    r = 2
    For i = ye1 To ye2 - 1
        val(i - ye1 + 1, 1) = CDate("12/31/" & y)
        val(i - ye1 + 1, 2) = 0
        While Year(Worksheets("data").Range("A" & r).Value) = y And r <= lr
            val(i - ye1 + 1, 2) = val(i - ye1 + 1, 2) + Worksheets("data").Range("B" & r).Value
            r = r + 1
        Wend
        y = y + 1
    Next i
    Worksheets("data").Range("J2:K" & (ye2 - ye1) + 1).Value = val
End Function

'Generates raw sales data based on filters applied to movement
'Aggregates over daily data
Public Function getsales()
    Dim lr As Long
    Worksheets("data").Range("K1").Value = "Sales"
    lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5, Criteria1:=Array("601", "602", "633"), Operator:=xlFilterValues
    Dim rng As Range
    On Error Resume Next
    Set rng = Worksheets("MOVEMENT").Range("D2:Q" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If rng Is Nothing Then
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
        Exit Function
    End If
    Dim count As Long
    Dim day As String
    Dim tot As Double
    Dim res() As Variant
    Dim sum As Long
    sum = 0
    For a = 1 To rng.Areas.count
        For r = 1 To rng.Areas(a).Rows.count
            sum = sum + 1
        Next r
    Next a
    ReDim res(1 To sum, 1 To 2)
    day = ""
    count = 2
    If Not rng Is Nothing Then
        For a = 1 To rng.Areas.count
            If a Mod 100 = 0 Then
                DoEvents
                Application.StatusBar = Int(100 * a / rng.Areas.count) & "%"
            End If
            For r = 1 To rng.Areas(a).Rows.count
                If day <> rng.Areas(a).Cells(r, 1).Value Then
                    day = rng.Areas(a).Cells(r, 1).Value
                    res(count - 1, 1) = day
                    If count <> 2 Then
                        res(count - 2, 2) = tot
                    End If
                    tot = 0
                    count = count + 1
                End If
                On Error Resume Next
                tot = tot + rng.Areas(a).Cells(r, 14).Value
                On Error GoTo 0
            Next r
        Next a
        res(count - 2, 2) = tot
        Worksheets("data").Range("A2:B" & count - 1).Value = res
    End If
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
    Application.StatusBar = "Done"
End Function

'Similar to sales function, except records usage values instead of sales value
Public Function getUnits()
    Dim lr As Long
    Worksheets("data").Range("K1").Value = "Units"
    lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5, Criteria1:=Array("601", "602", "633"), Operator:=xlFilterValues
    Dim rng As Range
    On Error Resume Next
    Set rng = Worksheets("MOVEMENT").Range("D2:F" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If rng Is Nothing Then
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
        Exit Function
    End If
    Dim count As Long
    Dim day As String
    Dim tot As Double
    Dim res() As Variant
    Dim sum As Long
    sum = 0
    For a = 1 To rng.Areas.count
        For r = 1 To rng.Areas(a).Rows.count
            sum = sum + 1
        Next r
    Next a
    ReDim res(1 To sum, 1 To 2)
    day = ""
    count = 2
    If Not rng Is Nothing Then
        For a = 1 To rng.Areas.count
            If a Mod 100 = 0 Then
                DoEvents
                Application.StatusBar = Int(100 * a / rng.Areas.count) & "%"
            End If
            For r = 1 To rng.Areas(a).Rows.count
                If day <> rng.Areas(a).Cells(r, 1).Value Then
                    day = rng.Areas(a).Cells(r, 1).Value
                    res(count - 1, 1) = day
                    If count <> 2 Then
                        res(count - 2, 2) = tot
                    End If
                    tot = 0
                    count = count + 1
                End If
                tot = tot - rng.Areas(a).Cells(r, 3).Value
            Next r
        Next a
        res(count - 2, 2) = tot
        Worksheets("data").Range("A2:B" & count - 1).Value = res
    End If
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
    Application.StatusBar = "Done"
End Function

'Updates values stored in "data" sheet to quickly clear on "analysis"
'Keeps a record of unique values of "product group", "product hierarchy", etc.
Function UpdateData()
    Dim lr As Long
    Worksheets("ZMMMATERIAL").Range("A1").AutoFilter
    Dim proGroups As Scripting.Dictionary: Set proGroups = New Scripting.Dictionary
    Dim hier As Scripting.Dictionary: Set hier = New Scripting.Dictionary
    Dim cats As Scripting.Dictionary: Set cats = New Scripting.Dictionary
    Dim matGroups As Scripting.Dictionary: Set matGroups = New Scripting.Dictionary
    Dim Class As Scripting.Dictionary: Set Class = New Scripting.Dictionary
    Dim mats As Scripting.Dictionary: Set mats = New Scripting.Dictionary
    Dim nums As Scripting.Dictionary: Set nums = New Scripting.Dictionary
    Dim rng As Range
    lr = Worksheets("ZMMMATERIAL").Range("A" & Rows.count).End(xlUp).Row
    On Error Resume Next
    Set rng = Worksheets("ZMMMATERIAL").Range("A2:AA" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If rng Is Nothing Then
        Exit Function
    End If
    For a = 1 To rng.Areas.count
        For r = 1 To rng.Areas(a).Rows.count
            proGroups(rng.Areas(a).Cells(r, "P").Value) = True
            hier(rng.Areas(a).Cells(r, "L").Value) = True
            cats(rng.Areas(a).Cells(r, "Y").Value) = True
            matGroups(rng.Areas(a).Cells(r, "N").Value) = True
            Class(rng.Areas(a).Cells(r, "U").Value) = True
            mats(rng.Areas(a).Cells(r, "C").Value) = True
            nums(rng.Areas(a).Cells(r, "B").Value) = True
        Next r
    Next a
    Worksheets("data").Range("C2:C" & proGroups.count + 1).Value = Application.Transpose(proGroups.Keys)
    Worksheets("data").Range("D2:D" & hier.count + 1).Value = Application.Transpose(hier.Keys)
    Worksheets("data").Range("E2:E" & cats.count + 1).Value = Application.Transpose(cats.Keys)
    Worksheets("data").Range("F2:F" & matGroups.count + 1).Value = Application.Transpose(matGroups.Keys)
    Worksheets("data").Range("G2:G" & Class.count + 1).Value = Application.Transpose(Class.Keys)
    Worksheets("data").Range("H2:H" & mats.count + 1).Value = Application.Transpose(mats.Keys)
    Worksheets("data").Range("I2:I" & nums.count + 1).Value = Application.Transpose(nums.Keys)
    Dim nr As Name
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Set nr = wb.Names.Item("ProGroup")
    nr.RefersTo = "=data!$C$2:$C$" & proGroups.count + 1
    Set nr = wb.Names.Item("ProHier")
    nr.RefersTo = "=data!$D$2:$D$" & hier.count + 1
    Set nr = wb.Names.Item("MatCat")
    nr.RefersTo = "=data!$E$2:$E$" & cats.count + 1
    Set nr = wb.Names.Item("MatGroup")
    nr.RefersTo = "=data!$F$2:$F$" & matGroups.count + 1
    Set nr = wb.Names.Item("ValClass")
    nr.RefersTo = "=data!$G$2:$G$" & Class.count + 1
    Set nr = wb.Names.Item("Mat")
    nr.RefersTo = "=data!$H$2:$H$" & mats.count + 1
    Set nr = wb.Names.Item("MatNum")
    nr.RefersTo = "=data!$I$2:$I$" & nums.count + 1
End Function

'Generate all raws/intermediates for selected data
'All raws are given (R) and intermediates given (I)
Function genRaws()
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    Dim dict2 As Scripting.Dictionary
    Set dict2 = New Scripting.Dictionary
    Dim rng As Range
    Dim lr As Long
    lr = Worksheets("ZMMMATERIAL").Range("A" & Rows.count).End(xlUp).Row
    Set rng = Worksheets("ZMMMATERIAL").Range("B2:B" & lr).Rows.SpecialCells(xlCellTypeVisible)
    For acnt = 1 To rng.Areas.count
        For rcnt = 1 To rng.Areas(acnt).Rows.count
            dict2(rng.Areas(acnt).Rows(rcnt).Value) = True
        Next
    Next
    Worksheets("ZPPBOM").Range("A1").AutoFilter
    Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=4, Criteria1:=dict2.Keys, Operator:=xlFilterValues
    lr = Worksheets("ZPPBOM").Range("A" & Rows.count).End(xlUp).Row
    Set rng = Nothing
    On Error Resume Next
    Set rng = Worksheets("ZPPBOM").Range("A2:H" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    Dim rm As String
    If Not rng Is Nothing Then
        For acnt = 1 To rng.Areas.count
            For rcnt = 1 To rng.Areas(acnt).Rows.count
                If rng.Areas(acnt).Cells(rcnt, 8) Then
                    rm = " (R)"
                Else
                    rm = " (I)"
                End If
                dict(rng.Areas(acnt).Cells(rcnt, 1).Value & rm) = True
            Next
        Next
    End If
    Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=4
    genRaws = dict.Keys
End Function

'Generates data showing how much a material is used to create the materials under the current filter
Function rawData(rm As Variant)
    Application.StatusBar = "Working..."
    Worksheets("ZPPBOM").Range("A1").AutoFilter
    Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=1, Criteria1:=rm
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5, Criteria1:=Array("601", "602", "603"), Operator:=xlFilterValues
    Dim lr As Long
    Dim prods2() As Variant
    ReDim prods2(0)
    Dim rng As Range
    lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    Set rng = Nothing
    On Error Resume Next
    Set rng = Worksheets("MOVEMENT").Range("B2:B" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If rng Is Nothing Then
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
        Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=1
        Exit Function
    End If
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=3
    Dim i As Long
    For Each cll In rng.Cells
        dict(CStr(cll.Value)) = 0
    Next cll
    'Filter by produced materials that have been sold in the current selection that are made with "rm" variable
    Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=4, Criteria1:=dict.Keys, Operator:=xlFilterValues
    lr = Worksheets("ZPPBOM").Range("A" & Rows.count).End(xlUp).Row
    Set rng = Nothing
    On Error Resume Next
    'Select ranges to refer to "unit from" and "product number"
    Set rng = Worksheets("ZPPBOM").Range("D2:D" & lr).Rows.SpecialCells(xlCellTypeVisible)
    Set rng2 = Worksheets("ZPPBOM").Range("C2:C" & lr).Rows.SpecialCells(xlCellTypeVisible)
    Dim unit As String
    On Error GoTo 0
    Application.StatusBar = "PreCalc Done"
    If Not rng Is Nothing Then
        unit = rng2.Areas(1).Rows(1).Value & "s"
        Worksheets("data").Range("K1").Value = unit
        i = 0
        For Each cll In rng
            prods2(i) = CStr(cll.Value)
            ReDim Preserve prods2(i + 1)
            i = i + 1
        Next cll
        Application.StatusBar = "ProductCount Done"
        'Filter to only show sales of materials that are made with "rm" variable
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=2, Criteria1:=prods2, Operator:=xlFilterValues
        Set rng = Worksheets("ZPPBOM").Range("A2:G" & lr).Rows.SpecialCells(xlCellTypeVisible)
        Dim num As Variant
        Dim fac As Double
        i = 2
        'Go through data and store dictionary associating material number to its conversion factor
        For acnt = 1 To rng.Areas.count
            For rcnt = 1 To rng.Areas(acnt).Rows.count
                If i Mod 100 = 0 Then
                    DoEvents
                End If
                num = CStr(rng.Areas(acnt).Cells(rcnt, 4).Value)
                fac = rng.Areas(acnt).Cells(rcnt, 7).Value
                dict(num) = fac
                i = i + 1
            Next
        Next
        Application.StatusBar = "Factors Done"
        lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
        Set rng = Worksheets("MOVEMENT").Range("B2:F" & lr).Rows.SpecialCells(xlCellTypeVisible)
        Dim sum As Long
        sum = 0
        For acnt = 1 To rng.Areas.count
            For rcnt = 1 To rng.Areas(acnt).Rows.count
                sum = sum + 1
            Next rcnt
        Next acnt
        i = 2
        Dim day As Variant
        Dim tot As Double
        Dim count As Long
        Dim res() As Variant
        ReDim res(1 To sum, 1 To 2)
        count = 0
        day = ""
        'Go through data and aggregate daily, converting qty of material to qty of raw/intermediate
        For acnt = 1 To rng.Areas.count
            For rcnt = 1 To rng.Areas(acnt).Rows.count
                If count Mod 100 = 0 Then
                    DoEvents
                    Application.StatusBar = Int(100 * count / sum) & "%"
                End If
                num = rng.Areas(acnt).Cells(rcnt, 1).Value
                If day <> rng.Areas(acnt).Cells(rcnt, 3).Value Then
                    day = rng.Areas(acnt).Cells(rcnt, 3).Value
                    res(i - 1, 1) = day
                    If i <> 2 Then
                        res(i - 2, 2) = tot
                    End If
                    tot = 0
                    i = i + 1
                End If
                count = count + 1
                tot = tot - rng.Areas(acnt).Cells(rcnt, 5).Value * dict(CStr(num))
            Next
        Next
        res(i - 2, 2) = tot
        Worksheets("data").Range("A2:B" & i - 1).Value = res
    End If
    'Reset filters
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
    Dim vals() As Variant
    ReDim vals(1 To Sheet3.MatNum.ListCount)
    For i = 1 To UBound(vals): vals(i) = CStr(Sheet3.MatNum.List(i - 1)): Next i
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=2, Criteria1:=vals, Operator:=xlFilterValues
    Worksheets("ZMMMATERIAL").Range("A1").AutoFilter Field:=3
    Worksheets("ZMMMATERIAL").Range("A1").AutoFilter Field:=2, Criteria1:=vals, Operator:=xlFilterValues
    Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=1
    Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=4
    Application.StatusBar = "Done"
End Function

'Same as getSales, except recording weight used
Function getUsage()
    Dim lr As Long
    Worksheets("data").Range("K1").Value = "Usage - LBs"
    lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5, Criteria1:=Array("261", "262"), Operator:=xlFilterValues
    Dim rng As Range
    On Error Resume Next
    Set rng = Worksheets("MOVEMENT").Range("D2:G" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If rng Is Nothing Then
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
        Exit Function
    End If
    Dim count As Long
    Dim day As String
    Dim tot As Double
    Dim res() As Variant
    Dim sum As Long
    sum = 0
    For a = 1 To rng.Areas.count
        For r = 1 To rng.Areas(a).Rows.count
            sum = sum + 1
        Next r
    Next a
    ReDim res(1 To sum, 1 To 2)
    day = ""
    count = 2
    If Not rng Is Nothing Then
        For a = 1 To rng.Areas.count
            If a Mod 100 = 0 Then
                DoEvents
                Application.StatusBar = Int(100 * a / rng.Areas.count) & "%"
            End If
            For r = 1 To rng.Areas(a).Rows.count
                If day <> rng.Areas(a).Cells(r, 1).Value Then
                    day = rng.Areas(a).Cells(r, 1).Value
                    res(count - 1, 1) = day
                    If count <> 2 Then
                        res(count - 2, 2) = tot
                    End If
                    tot = 0
                    count = count + 1
                End If
                tot = tot - rng.Areas(a).Cells(r, 15).Value
            Next r
        Next a
        res(count - 2, 2) = tot
        Worksheets("data").Range("A2:B" & count - 1).Value = res
    End If
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
    Application.StatusBar = "Done"
End Function

'Same as getUsage except recording weight sold
Function getWeight()
    Dim lr As Long
    Worksheets("data").Range("K1").Value = "Sales - LBs"
    lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5, Criteria1:=Array("601", "602", "633"), Operator:=xlFilterValues
    Dim rng As Range
    On Error Resume Next
    Set rng = Worksheets("MOVEMENT").Range("D2:G" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If rng Is Nothing Then
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
        Exit Function
    End If
    Dim count As Long
    Dim day As String
    Dim tot As Double
    Dim res() As Variant
    Dim sum As Long
    sum = 0
    For a = 1 To rng.Areas.count
        For r = 1 To rng.Areas(a).Rows.count
            sum = sum + 1
        Next r
    Next a
    ReDim res(1 To sum, 1 To 2)
    day = ""
    count = 2
    If Not rng Is Nothing Then
        For a = 1 To rng.Areas.count
            If a Mod 100 = 0 Then
                DoEvents
                Application.StatusBar = Int(100 * a / rng.Areas.count) & "%"
            End If
            For r = 1 To rng.Areas(a).Rows.count
                If day <> rng.Areas(a).Cells(r, 1).Value Then
                    day = rng.Areas(a).Cells(r, 1).Value
                    res(count - 1, 1) = day
                    If count <> 2 Then
                        res(count - 2, 2) = tot
                    End If
                    tot = 0
                    count = count + 1
                End If
                tot = tot - rng.Areas(a).Cells(r, 15).Value
            Next r
        Next a
        res(count - 2, 2) = tot
        Worksheets("data").Range("A2:B" & count - 1).Value = res
    End If
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
    Application.StatusBar = "Done"
End Function
