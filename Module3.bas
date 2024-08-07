Attribute VB_Name = "Module3"
'Returns the number of columns required to show data based on aggregation
Function NumCols() As Long
    If Sheet3.GroupBy.Value = "Monthly" Then
        NumCols = DateDiff("m", Sheet3.TextBox1.Value, Sheet3.TextBox2.Value) + 1
    ElseIf Sheet3.GroupBy.Value = "Quarterly" Then
        NumCols = DateDiff("q", Sheet3.TextBox1.Value, Sheet3.TextBox2.Value) + 1
    Else
        NumCols = Year(Sheet3.TextBox2.Value) - Year(Sheet3.TextBox1.Value) + 1
    End If
End Function

'Creates the table displaying the quantity and sales of all finished goods in current filter
Function getFinishedGoods()
    Dim lr As Long
    Dim lc As Long
    Dim rng As Range
    lc = Columns.count
    lc = IIf(lc < 7, 7, lc)
    lr = Worksheets("ANALYSIS").Range("G" & Rows.count).End(xlUp).Row
    lr = IIf(lr < 31, 31, lr)
    Worksheets("ANALYSIS").Range("A31", Worksheets("ANALYSIS").Cells(lr + 4, lc)).Clear
    Dim dict As Scripting.Dictionary: Set dict = New Scripting.Dictionary
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5, Criteria1:=Array("601", "602", "633"), Operator:=xlFilterValues
    Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=1, Criteria1:=Sheet3.MatNum.Value
    lr = Worksheets("ZPPBOM").Range("A" & Rows.count).End(xlUp).Row
    On Error Resume Next
    Set rng = Worksheets("ZPPBOM").Range("D2:D" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    'Checking if material is raw, if it is then change data to be all finished goods relating to it
    'Else data will be all finished goods related to filter
    If Not rng Is Nothing Then
        Dim prods() As Variant
        prods = Application.Transpose(rng.Value)
        For j = 1 To UBound(prods): prods(j) = CStr(prods(j)): Next j
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=2, Criteria1:=prods, Operator:=xlFilterValues
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=3
    End If
    lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
    Set rng = Nothing
    On Error Resume Next
    Set rng = Worksheets("MOVEMENT").Range("A2:P" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not rng Is Nothing Then
        Dim group As Variant
        Dim hier As Variant
        lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
        Dim sum As Long
        sum = 0
        For a = 1 To rng.Areas.count
            For r = 1 To rng.Areas(a).Rows.count
                sum = sum + 1
            Next r
        Next a
        Dim count As Long
        count = 1
        'Record information for all finished goods in filter and store in dictionary
        For a = 1 To rng.Areas.count
            For r = 1 To rng.Areas(a).Rows.count
                If count Mod 1000 = 0 Then
                    DoEvents
                    Application.StatusBar = "Finished Goods: " & Int(100 * count / sum) & "%"
                End If
                count = count + 1
                If Not dict.Exists(CStr(rng.Areas(a).Cells(r, 2).Value)) Then
                    group = Application.VLookup(CStr(rng.Areas(a).Cells(r, 2)), Worksheets("ZMMMATERIAL").Range("B2:P" & lr), 15, False)
                    hier = Application.VLookup(CStr(rng.Areas(a).Cells(r, 2)), Worksheets("ZMMMATERIAL").Range("B2:L" & lr), 11, False)
                    dict(CStr(rng.Areas(a).Cells(r, 2).Value)) = Array(rng.Areas(a).Cells(r, 1).Value, rng.Areas(a).Cells(r, 3).Value, rng.Areas(a).Cells(r, 8).Value, group, hier)
                End If
            Next r
        Next a
        Dim res() As Variant
        'Declare and initialize resulting 2D array
        ReDim res(1 To dict.count + 1, 1 To 6 + 2 * NumCols())
        Dim i As Long
        i = 2
        For Each key In dict.Keys
            res(i, 1) = dict(key)(0)
            res(i, 2) = key
            For j = 3 To 6: res(i, j) = dict(key)(j - 2): Next j
            dict(key) = i
            i = i + 1
        Next key
        Dim start As Date
        start = Sheet3.TextBox1.Value
        If Sheet3.GroupBy.Value = "Quarterly" Then
            start = Excel.Application.WorksheetFunction.EoMonth(((Month(start) + 2) \ 3) + 2 & "/28/" & Year(start), 0)
        ElseIf Sheet3.GroupBy.Value = "Yearly" Then
            start = "12/31/" & Year(start)
        End If
        'Set up table headers
        res(1, 1) = "Plant"
        res(1, 2) = "Part Number"
        res(1, 3) = "Description"
        res(1, 4) = "Unit"
        res(1, 5) = "Product Group"
        res(1, 6) = "Product Hierarchy"
        For j = 7 To UBound(res, 2) Step 2
            If Sheet3.GroupBy.Value = "Monthly" Then
                res(1, j) = "Sales " & CDate(Excel.Application.WorksheetFunction.EoMonth(start, 0))
                res(1, j + 1) = "Qty " & CDate(Excel.Application.WorksheetFunction.EoMonth(start, 0))
                start = DateAdd("m", 1, start)
            ElseIf Sheet3.GroupBy.Value = "Quarterly" Then
                res(1, j) = "Sales " & CDate(Excel.Application.WorksheetFunction.EoMonth(start, 0))
                res(1, j + 1) = "Qty " & CDate(Excel.Application.WorksheetFunction.EoMonth(start, 0))
                start = DateAdd("q", 1, start)
            Else
                res(1, j) = "Sales " & start
                res(1, j + 1) = "Qty " & start
                start = "12/31/" & Year(start) + 1
            End If
        Next j
        For j = 2 To UBound(res, 1)
            For k = 7 To UBound(res, 2)
                res(j, k) = 0
            Next k
        Next j
        lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
        Set rng = Nothing
        On Error Resume Next
        Set rng = Worksheets("MOVEMENT").Range("A2:Q" & lr).Rows.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        Dim index As Long
        Dim index2 As Long
        If Not rng Is Nothing Then
            count = 1
            sum = 0
            For a = 1 To rng.Areas.count
                For r = 1 To rng.Areas(a).Rows.count
                    sum = sum + 1
                Next r
            Next a
            For a = 1 To rng.Areas.count
                For r = 1 To rng.Areas(a).Rows.count
                'Main loop to go through movement and add units together
                    If count Mod 1000 = 0 Then
                        DoEvents
                        Application.StatusBar = "Finished Goods: " & Int(100 * count / sum) & "%"
                    End If
                    count = count + 1
                    If Sheet3.GroupBy.Value = "Monthly" Then
                        index = Month(rng.Areas(a).Cells(r, 4)) + Year(rng.Areas(a).Cells(r, 4)) * 12 - Month(Sheet3.TextBox1.Value) - Year(Sheet3.TextBox1.Value) * 12
                    ElseIf Sheet3.GroupBy.Value = "Quarterly" Then
                        index = (Month(rng.Areas(a).Cells(r, 4)) + 2) \ 3 + Year(rng.Areas(a).Cells(r, 4)) * 4 - (Month(Sheet3.TextBox1.Value) + 2) \ 3 - Year(Sheet3.TextBox1.Value) * 4
                    Else
                        index = Year(rng.Areas(a).Cells(r, 4)) - Year(Sheet3.TextBox1.Value)
                    End If
                    index = 2 * index + 7
                    index2 = dict(CStr(rng.Areas(a).Cells(r, 2).Value))
                    On Error Resume Next
                    res(index2, index) = res(index2, index) + rng.Areas(a).Cells(r, 17).Value
                    On Error GoTo 0
                    'If unit is not entered with base unit then look for alternate unit
                    If rng.Areas(a).Cells(r, 7).Value <> rng.Areas(a).Cells(r, 8).Value Then
                        If Not dict.Exists(rng.Areas(a).Cells(r, 2) & rng.Areas(a).Cells(r, 7)) Then
                            dict(rng.Areas(a).Cells(r, 2) & rng.Areas(a).Cells(r, 7)) = findVal(rng.Areas(a).Cells(r, 7).Value, rng.Areas(a).Cells(r, 2).Value)
                        End If
                        res(index2, index + 1) = res(index2, index + 1) - rng.Areas(a).Cells(r, 6).Value * dict(rng.Areas(a).Cells(r, 2) & rng.Areas(a).Cells(r, 7))
                    Else
                        res(index2, index + 1) = res(index2, index + 1) - rng.Areas(a).Cells(r, 6).Value
                    End If
                Next r
            Next a
        End If
        'Define and edit style of table
        With Worksheets("ANALYSIS")
            .Range(.Cells(31, 7), .Cells(30 + UBound(res, 1), UBound(res, 2) + 6)).Value = res
            .ListObjects.Add(xlSrcRange, .Range(.Cells(31, 7), .Cells(30 + UBound(res, 1), UBound(res, 2) + 6)), , xlYes).Name = "FinishedGoods"
            .ListObjects("FinishedGoods").TableStyle = "TableStyleMedium12"
        End With
    End If
    Worksheets("ANALYSIS").Range("F31").Value = "Finished Goods:"
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
    If Sheet3.MatNum.Value <> "" Then
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=2, Criteria1:=Sheet3.MatNum.Value
    End If
End Function

'Similar to previous function, except for "raw" variable to change behavior
Function getIntermediates(Optional raw As Boolean = True)
    Dim lr As Long
    Dim rng As Range
    Dim dict As Scripting.Dictionary: Set dict = New Scripting.Dictionary
    Dim dict2 As Scripting.Dictionary: Set dict2 = New Scripting.Dictionary
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=3
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5, Criteria1:=Array("261", "262"), Operator:=xlFilterValues
    Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=1, Criteria1:=Sheet3.MatNum.Value
    lr = Worksheets("ZPPBOM").Range("A" & Rows.count).End(xlUp).Row
    On Error Resume Next
    Set rng = Worksheets("ZPPBOM").Range("D2:D" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    Dim prods() As Variant
    'If material entered is a raw/intermediate and we are not looking at the raw materials
    'Then produce all higher assembies related to material
    'Else produce all intermediates related to all materials in the filter
    If Not rng Is Nothing And Not raw Then
        lr = Worksheets("ZPPBOM").Range("A" & Rows.count).End(xlUp).Row
        prods = Application.Transpose(Worksheets("ZPPBOM").Range("D2:D" & lr).Rows.SpecialCells(xlCellTypeVisible).Value)
        For j = 1 To UBound(prods): prods(j) = CStr(prods(j)): Next j
        Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=1, Criteria1:=prods, Operator:=xlFilterValues
        lr = Worksheets("ZPPBOM").Range("A" & Rows.count).End(xlUp).Row
        Set rng = Nothing
        On Error Resume Next
        Set rng = Worksheets("ZPPBOM").Range("A2:C" & lr).Rows.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=1
    ElseIf rng Is Nothing Then
        Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=1
        Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=8, Criteria1:=raw
        lr = Worksheets("ZMMMATERIAL").Range("A" & Rows.count).End(xlUp).Row
        Set rng = Worksheets("ZMMMATERIAL").Range("B2:B" & lr).Rows.SpecialCells(xlCellTypeVisible)
        Dim count As Long
        count = 0
        For a = 1 To rng.Areas.count: For r = 1 To rng.Areas(a).Rows.count: count = count + 1: ReDim Preserve prods(1 To count): prods(count) = rng.Areas(a).Rows(r).Value: Next r: Next a
        Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=4, Criteria1:=prods, Operator:=xlFilterValues
        Set rng = Nothing
        lr = Worksheets("ZPPBOM").Range("A" & Rows.count).End(xlUp).Row
        On Error Resume Next
        Set rng = Worksheets("ZPPBOM").Range("A2:C" & lr).Rows.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=4
    Else
        Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=1
        Set rng = Nothing
    End If
    If Not rng Is Nothing Then
        Dim group As Variant
        Dim hier As Variant
        Dim Plant As Variant
        Dim sum As Long
        Dim count2 As Long
        sum = 0
        For a = 1 To rng.Areas.count
            For r = 1 To rng.Areas(a).Rows.count
                sum = sum + 1
            Next r
        Next a
        count2 = 1
        For a = 1 To rng.Areas.count
            For r = 1 To rng.Areas(a).Rows.count
                If count2 Mod 1000 = 0 Then
                    DoEvents
                    If Not raw Then
                        Application.StatusBar = "Intermediates: " & Int(100 * count2 / sum) & "%"
                    Else
                        Application.StatusBar = "Raws: " & Int(100 * count2 / sum) & "%"
                    End If
                End If
                count2 = count2 + 1
                If Not dict.Exists(CStr(rng.Areas(a).Cells(r, 1).Value)) Then
                    group = Application.VLookup(CStr(rng.Areas(a).Cells(r, 1)), Worksheets("ZMMMATERIAL").Range("B2:P" & lr), 15, False)
                    hier = Application.VLookup(CStr(rng.Areas(a).Cells(r, 1)), Worksheets("ZMMMATERIAL").Range("B2:L" & lr), 11, False)
                    Plant = Application.XLookup(CStr(rng.Areas(a).Cells(r, 1)), Worksheets("ZMMMATERIAL").Range("B2:B" & lr), Worksheets("ZMMMATERIAL").Range("A2:A" & lr))
                    dict(CStr(rng.Areas(a).Cells(r, 1).Value)) = Array(Plant, rng.Areas(a).Cells(r, 2).Value, rng.Areas(a).Cells(r, 3).Value, group, hier)
                End If
                If Not dict2.Exists(CStr(rng.Areas(a).Cells(r, 1).Value) & rng.Areas(a).Cells(r, 4).Value) Then
                    dict2(CStr(rng.Areas(a).Cells(r, 1).Value) & rng.Areas(a).Cells(r, 4).Value) = Array(rng.Areas(a).Cells(r, 7).Value, rng.Areas(a).Cells(r, 6).Value)
                End If
            Next r
        Next a
        Dim res() As Variant
        'Declare and initialize final 2D array
        ReDim res(1 To dict.count + 1, 1 To 6 + 2 * NumCols())
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=2, Criteria1:=dict.Keys, Operator:=xlFilterValues
        Dim i As Long
        i = 2
        For Each key In dict.Keys
            res(i, 1) = dict(key)(0)
            res(i, 2) = key
            For j = 3 To 6: res(i, j) = dict(key)(j - 2): Next j
            dict(key) = i
            i = i + 1
        Next key
        Dim start As Date
        start = Sheet3.TextBox1.Value
        If Sheet3.GroupBy.Value = "Quarterly" Then
            start = Excel.Application.WorksheetFunction.EoMonth(((Month(start) + 2) \ 3) + 2 & "/28/" & Year(start), 0)
        ElseIf Sheet3.GroupBy.Value = "Yearly" Then
            start = "12/31/" & Year(start)
        End If
        'Create headers for table
        res(1, 1) = "Plant"
        res(1, 2) = "Part Number"
        res(1, 3) = "Description"
        res(1, 4) = "Unit"
        res(1, 5) = "Product Group"
        res(1, 6) = "Product Hierarchy"
        For j = 7 To UBound(res, 2) Step 2
            If Sheet3.GroupBy.Value = "Monthly" Then
                res(1, j) = "Qty " & CDate(Excel.Application.WorksheetFunction.EoMonth(start, 0))
                res(1, j + 1) = "Usage " & CDate(Excel.Application.WorksheetFunction.EoMonth(start, 0))
                start = DateAdd("m", 1, start)
            ElseIf Sheet3.GroupBy.Value = "Quarterly" Then
                res(1, j) = "Qty " & CDate(Excel.Application.WorksheetFunction.EoMonth(start, 0))
                res(1, j + 1) = "Usage " & CDate(Excel.Application.WorksheetFunction.EoMonth(start, 0))
                start = DateAdd("q", 1, start)
            Else
                res(1, j) = "Qty " & start
                res(1, j + 1) = "Usage " & start
                start = "12/31/" & Year(start) + 1
            End If
        Next j
        For j = 2 To UBound(res, 1)
            For k = 7 To UBound(res, 2)
                res(j, k) = 0
            Next k
        Next j
        lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
        Set rng = Nothing
        On Error Resume Next
        Set rng = Worksheets("MOVEMENT").Range("A2:I" & lr).Rows.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        Dim index As Long
        Dim index2 As Long
        Dim conv As Double
        If Not rng Is Nothing Then
            sum = 0
            count2 = 1
            For a = 1 To rng.Areas.count
                For r = 1 To rng.Areas(a).Rows.count
                    sum = sum + 1
                Next r
            Next a
            For a = 1 To rng.Areas.count
                For r = 1 To rng.Areas(a).Rows.count
                'Go through data and aggregate the usage of the material in its base unit
                    If count2 Mod 1000 = 0 Then
                        DoEvents
                        If Not raw Then
                            Application.StatusBar = "Intermediates: " & Int(100 * count2 / sum) & "%"
                        Else
                            Application.StatusBar = "Raws: " & Int(100 * count2 / sum) & "%"
                        End If
                    End If
                    count2 = count2 + 1
                    If Sheet3.GroupBy.Value = "Monthly" Then
                        index = Month(rng.Areas(a).Cells(r, 4)) + Year(rng.Areas(a).Cells(r, 4)) * 12 - Month(Sheet3.TextBox1.Value) - Year(Sheet3.TextBox1.Value) * 12
                    ElseIf Sheet3.GroupBy.Value = "Quarterly" Then
                        index = (Month(rng.Areas(a).Cells(r, 4)) + 2) \ 3 + Year(rng.Areas(a).Cells(r, 4)) * 4 - (Month(Sheet3.TextBox1.Value) + 2) \ 3 - Year(Sheet3.TextBox1.Value) * 4
                    Else
                        index = Year(rng.Areas(a).Cells(r, 4)) - Year(Sheet3.TextBox1.Value)
                    End If
                    index = 2 * index + 7
                    index2 = dict(CStr(rng.Areas(a).Cells(r, 2).Value))
                    'If not stored in base unit convert first
                    If rng.Areas(a).Cells(r, 7).Value <> rng.Areas(a).Cells(r, 8).Value Then
                        If Not dict.Exists(rng.Areas(a).Cells(r, 2) & rng.Areas(a).Cells(r, 7)) Then
                            dict(rng.Areas(a).Cells(r, 2) & rng.Areas(a).Cells(r, 7)) = findVal(rng.Areas(a).Cells(r, 7).Value, rng.Areas(a).Cells(r, 2).Value)
                        End If
                        res(index2, index + 1) = res(index2, index + 1) - rng.Areas(a).Cells(r, 6).Value * dict(rng.Areas(a).Cells(r, 2) & rng.Areas(a).Cells(r, 7))
                    Else
                        res(index2, index + 1) = res(index2, index + 1) - rng.Areas(a).Cells(r, 6).Value
                    End If
                Next r
            Next a
        End If
        'Call function to fill the qty used by the finished goods of each intermediate/raw
        Call fillUsage(res, dict2)
        lr = Worksheets("ANALYSIS").Range("G" & Rows.count).End(xlUp).Row
        If raw Then
            lr = IIf(lr < 31, 35, lr)
        Else
            lr = IIf(lr < 31, 33, lr)
        End If
        'Declare tables and style accordingly
        With Worksheets("ANALYSIS")
            If raw And Worksheets("ANALYSIS").Range("F" & lr + 2).Value = "Intermediates:" Then
                lr = lr + 2
            End If
            .Range(.Cells(lr + 2, 7), .Cells(lr + 1 + UBound(res, 1), UBound(res, 2) + 6)).Value = res
            If raw Then
                .ListObjects.Add(xlSrcRange, .Range(.Cells(lr + 2, 7), .Cells(lr + 1 + UBound(res, 1), UBound(res, 2) + 6)), , xlYes).Name = "Raws"
                .ListObjects("Raws").TableStyle = "TableStyleMedium12"
            Else
                .ListObjects.Add(xlSrcRange, .Range(.Cells(lr + 2, 7), .Cells(lr + 1 + UBound(res, 1), UBound(res, 2) + 6)), , xlYes).Name = "Intermediates"
                .ListObjects("Intermediates").TableStyle = "TableStyleMedium12"
            End If
        End With
    Else
        lr = Worksheets("ANALYSIS").Range("G" & Rows.count).End(xlUp).Row
    End If
    lr = IIf(lr < 31, 31, lr)
    If raw Then
        If Worksheets("ANALYSIS").Range("F" & lr + 2).Value = "Intermediates:" Then
            lr = lr + 2
        End If
        Worksheets("ANALYSIS").Range("F" & lr + 2).Value = "Raws:"
    Else
        Worksheets("ANALYSIS").Range("F" & lr + 2).Value = "Intermediates:"
    End If
    'Reset filters
    Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=5
    Worksheets("ZPPBOM").Range("A1").AutoFilter Field:=8
    If Sheet3.MatNum.Value <> "" Then
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=2, Criteria1:=Sheet3.MatNum.Value
    Else
        Dim vals() As Variant
        ReDim vals(1 To Sheet3.MatNum.ListCount)
        For j = 1 To UBound(vals): vals(j) = CStr(Sheet3.MatNum.List(j - 1)): Next j
        Worksheets("MOVEMENT").Range("A1").AutoFilter Field:=2, Criteria1:=vals, Operator:=xlFilterValues
        Worksheets("ZMMMATERIAL").Range("A1").AutoFilter Field:=2, Criteria1:=vals, Operator:=xlFilterValues
    End If
End Function

'Fills data of how much of a material was used to create each finished good
Function fillUsage(data() As Variant, dict As Scripting.Dictionary)
    Dim fgs As Long
    Dim factor As Double
    fgs = Worksheets("ANALYSIS").Range("G31").End(xlDown).Row - 31
    'Go through each raw material
    For i = 2 To UBound(data, 1)
        DoEvents
        Application.StatusBar = Int(100 * i / UBound(data, 1)) & "%"
        'Go through each finished good
        For j = 1 To fgs
            If dict.Exists(data(i, 2) & Worksheets("ANALYSIS").Range("H" & j + 31).Value) Then
                factor = dict(data(i, 2) & Worksheets("ANALYSIS").Range("H" & j + 31).Value)(0)
                'Go through each period of aggregation
                For k = 7 To UBound(data, 2) Step 2
                    data(i, k) = data(i, k) + Worksheets("ANALYSIS").Cells(31 + j, k + 7).Value * factor
                Next k
            End If
        Next j
    Next i
End Function

'Function to find conversion factor from alternate unit to base
Function findVal(unit As String, Material As Variant) As Double
    Dim lr As Long
    lr = Worksheets("ALTUNIT").Range("B" & Rows.count).End(xlUp).Row
    Worksheets("ALTUNIT").Range("A1").AutoFilter Field:=1, Criteria1:=Material
    Worksheets("ALTUNIT").Range("A1").AutoFilter Field:=9, Criteria1:=unit
    Dim rng As Range
    On Error Resume Next
    Set rng = Worksheets("ALTUNIT").Range("F2:H" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not rng Is Nothing Then
        findVal = rng.Areas(1).Cells(1, 1).Value / rng.Areas(1).Cells(1, 3).Value
    Else
        If unit = "M" Then
            findVal = findVal("YD", Material) * 1.09361
        ElseIf unit = "YD" Then
            findVal = findVal("M", Material) / 1.09361
        ElseIf unit = "YD2" Then
            findVal = findVal("M2", Material) / 1.19599
        ElseIf unit = "M2" Then
            findVal = findVal("YD2", Material) * 1.19599
        ElseIf unit = "PT" Then
            findVal = findVal("QT", Material) / 2
        ElseIf unit = "QT" Then
            findVal = findVal("PT", Material) * 2
        ElseIf unit = "GAL" Then
            If Material = 198497 Then
                findVal = findVal("PC", Material) / 5
            End If
        ElseIf unit = "OZ" Then
            findVal = findVal("LB", Material) / 16
        End If
    End If
End Function

'Creates each table with a call
Sub main()
    Call getFinishedGoods
    Call getIntermediates(False)
    Call getIntermediates(True)
    Application.StatusBar = "Ready"
End Sub
