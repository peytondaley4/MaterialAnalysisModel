Attribute VB_Name = "Module1"
'Combines data of same movement type, same material, same date, and same supplier/customer
Function fixMovement()
    Application.StatusBar = "0%"
    Dim lr As Long
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
    'Sort data first by date, and then by material
    With Worksheets("MOVEMENT").Sort
        .SortFields.Add key:=Worksheets("MOVEMENT").Range("D1"), Order:=xlAscending
        .SortFields.Add key:=Worksheets("MOVEMENT").Range("B1"), Order:=xlAscending
        .SetRange Worksheets("MOVEMENT").Range("A1:P" & lr)
        .Header = xlYes
        .Apply
    End With
    Dim res As Variant
    ReDim res(1 To lr - 1, 1 To 18)
    Dim rng As Range
    Dim val As Variant
    lr = Worksheets("ZMMMATERIAL").Range("A" & Rows.count).End(xlUp).Row
    Set rng = Worksheets("ZMMMATERIAL").Range("B2:I" & lr)
    'Go through all materials and store a dictionary associating materials and base units to weight
    For a = 1 To rng.Areas.count
        For r = 1 To rng.Areas(a).Rows.count
            If Not (rng.Areas(a).Cells(r, 3).Value = "LB" Or rng.Areas(a).Cells(r, 3).Value = "KG" Or rng.Areas(a).Cells(r, 3).Value = "G" Or rng.Areas(a).Cells(r, 3).Value = "OZ") Then
                If rng.Areas(a).Cells(r, 8).Value = "KG" Then
                    val = rng.Areas(a).Cells(r, 7).Value * 2.2
                ElseIf rng.Areas(a).Cells(r, 8).Value = "G" Then
                    val = rng.Areas(a).Cells(r, 7).Value * 2.2 / 1000
                Else
                    val = rng.Areas(a).Cells(r, 7)
                End If
                dict(rng.Areas(a).Cells(r, 1).Value & rng.Areas(a).Cells(r, 3).Value) = val
            End If
        Next r
    Next a
    lr = Worksheets("MOVEMENT").Range("A" & Rows.count).End(xlUp).Row
    Set rng = Worksheets("MOVEMENT").Range("A2:R" & lr)
    Dim key As Variant
    Dim pkey As Variant
    pkey = ""
    Dim count As Long
    count = 1
    'Go through movement and combine all matching rows
    For i = 1 To rng.Rows.count
        key = rng.Cells(i, 1) & rng.Cells(i, 2) & rng.Cells(i, 4) & rng.Cells(i, 5) & rng.Cells(i, 11) & rng.Cells(i, 13) & rng.Cells(i, 14)
        If key = pkey Then
            res(count - 1, 6) = res(count - 1, 6) + rng.Cells(i, 6).Value
            res(count - 1, 9) = res(count - 1, 9) + rng.Cells(i, 9).Value
        Else
            For j = 1 To 17
                res(count, j) = rng.Cells(i, j).Value
            Next j
            count = count + 1
        End If
        'Record weight of entry, directly if listed in mass units, through findVal if not
        If rng.Cells(i, 7).Value = "LB" Then
            res(count - 1, 18) = res(count - 1, 6)
        ElseIf rng.Cells(i, 7).Value = "KG" Then
            res(count - 1, 18) = res(count - 1, 6) * 2.2
        ElseIf rng.Cells(i, 7).Value = "G" Then
            res(count - 1, 18) = res(count - 1, 6) * 2.2 / 1000
        ElseIf rng.Cells(i, 7).Value = "OZ" Then
            res(count - 1, 18) = res(count - 1, 6) / 16
        Else
            'Record entries in dictionary, reduces redundant calculation
            If Not dict.Exists(res(count - 1, 2) & res(count - 1, 7)) Then
                dict(res(count - 1, 2) & res(count - 1, 7)) = findVal(CStr(res(count - 1, 7)), res(count - 1, 2), dict)
            End If
            res(count - 1, 18) = dict(res(count - 1, 2) & res(count - 1, 7)) * res(count - 1, 6)
        End If
        If i Mod 1000 = 0 Then
            DoEvents
            Application.StatusBar = Int(100 * i / rng.Rows.count) & "%"
            Debug.Print i
        End If
        pkey = key
    Next i
    Application.StatusBar = "Ready"
    Worksheets("MOVEMENT").Range("A2:R" & lr).delete
    Worksheets("MOVEMENT").Range("A2:R" & count) = res
    Sheet3.Clear_Click
End Function

'Given a unit and a material, return the conversion factor from this unit to the base
Function findVal(unit As String, Material As Variant, dict As Scripting.Dictionary) As Double
    Dim lr As Long
    lr = Worksheets("ALTUNIT").Range("B" & Rows.count).End(xlUp).Row
    Worksheets("ALTUNIT").Range("A1").AutoFilter Field:=1, Criteria1:=Material
    Worksheets("ALTUNIT").Range("A1").AutoFilter Field:=9, Criteria1:=unit
    Dim rng As Range
    On Error Resume Next
    Set rng = Worksheets("ALTUNIT").Range("F2:H" & lr).Rows.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    'If alternate unit exists return the value in pounds
    If Not rng Is Nothing Then
        If rng.Areas(1).Cells(1, 2).Value = "KG" Then
            findVal = 2.2 * rng.Areas(1).Cells(1, 1).Value / rng.Areas(1).Cells(1, 3).Value
        ElseIf rng.Areas(1).Cells(1, 2).Value = "LB" Then
            findVal = rng.Areas(1).Cells(1, 1).Value / rng.Areas(1).Cells(1, 3).Value
        ElseIf rng.Areas(1).Cells(1, 2).Value = "G" Then
            findVal = (2.2 / 1000) * rng.Areas(1).Cells(1, 1).Value / rng.Areas(1).Cells(1, 3).Value
        ElseIf rng.Areas(1).Cells(1, 2).Value = "OZ" Then
            findVal = (1 / 16) * rng.Areas(1).Cells(1, 1).Value / rng.Areas(1).Cells(1, 3).Value
        Else
            If dict.Exists(Material & unit) Then
                findVal = dict(Material & unit)
                Exit Function
            End If
            If rng.Areas(1).Cells(1, 2).Value <> unit Then
                findVal = findVal(rng.Areas(1).Cells(1, 2).Value, Material, dict) * rng.Areas(1).Cells(1, 1).Value / rng.Areas(1).Cells(1, 3).Value
            Else
                findVal = 1
            End If
        End If
    Else
    'If no alternate is listed, convert to other unit and try again
        If unit = "M" Then
            findVal = findVal("YD", Material, dict) * 1.09361
        ElseIf unit = "YD" Then
            findVal = findVal("M", Material, dict) / 1.09361
        ElseIf unit = "YD2" Then
            findVal = findVal("M2", Material, dict) / 1.19599
        ElseIf unit = "M2" Then
            findVal = findVal("YD2", Material, dict) * 1.19599
        ElseIf unit = "PT" Then
            findVal = findVal("QT", Material, dict) / 2
        ElseIf unit = "QT" Then
            findVal = findVal("PT", Material, dict) * 2
        ElseIf unit = "GAL" Then
            If Material = 198497 Then
                findVal = findVal("PC", Material, dict) / 5
            End If
        Else
            Stop
        End If
    End If
End Function

'Open sales spreadsheet and perform similar actions to "fixmovement" but then
'Match the data to the movement table and store sales value in place
Function fixsalesdata()
    Dim FileToOpen As Variant
    Dim OpenBook As Workbook
    Dim wks As Worksheet
    FileToOpen = Application.GetOpenFilename(Title:="Browse for your File & Import Range", FileFilter:="Excel Files (*.xls*),*xls*")
    Set OpenBook = Application.Workbooks.Open(FileToOpen)
    Set wks = OpenBook.Worksheets(1)
    Dim lr As Long
    lr = wks.Range("A" & Rows.count).End(xlUp).Row
    With wks.Sort
        .SortFields.Add key:=wks.Range("G1"), Order:=xlAscending
        .SortFields.Add key:=wks.Range("I1"), Order:=xlAscending
        .SetRange wks.Range("A1:U" & lr)
        .Header = xlYes
        .Apply
    End With
    Dim res As Variant
    ReDim res(1 To lr - 1, 1 To 21)
    Dim rng As Range
    Set rng = wks.Range("A2:P" & lr)
    Dim key As Variant
    Dim pkey As Variant
    Dim dict As Scripting.Dictionary: Set dict = New Scripting.Dictionary
    pkey = ""
    Dim count As Long
    count = 1
    For i = 1 To rng.Rows.count
        key = rng.Cells(i, 1).Value & rng.Cells(i, 7).Value & rng.Cells(i, 9).Value & rng.Cells(i, 18).Value
        If count Mod 1000 = 0 Then
            DoEvents
            Application.StatusBar = Int(100 * count / lr) & "%"
        End If
        If key = pkey Then
            dict(key) = dict(key) + rng.Cells(i, 8).Value
            res(count - 1, 11) = res(count - 1, 11) + rng.Cells(i, 11).Value
            res(count - 1, 8) = res(count - 1, 8) + rng.Cells(i, 8).Value
        Else
            dict(key) = rng.Cells(i, 8).Value
            For j = 1 To 21
                res(count, j) = rng.Cells(i, j).Value
            Next j
            count = count + 1
        End If
        pkey = key
    Next i
    wks.Range("A2:U" & lr).delete
    wks.Range("A2:U" & count) = res
    Set wks = ThisWorkbook.Worksheets("MOVEMENT")
    lr = wks.Range("A" & Rows.count).End(xlUp).Row
    wks.Range("A1").AutoFilter Field:=5, Criteria1:=Array("601", "602", "633"), Operator:=xlFilterValues
    Set rng = wks.Range("A2:O" & lr).Rows.SpecialCells(xlCellTypeVisible)
    count = 0
    For a = 1 To rng.Areas.count
        For r = 1 To rng.Areas(a).Rows.count
            If count Mod 1000 = 0 Then
                Application.StatusBar = "Row: " & count
                DoEvents
            End If
            count = count + 1
            key = rng.Areas(a).Cells(r, 1).Value & rng.Areas(a).Cells(r, 4).Value & rng.Areas(a).Cells(r, 2).Value & rng.Areas(a).Cells(r, 15).Value
            If dict.Exists(key) Then
                rng.Areas(a).Cells(r, 17) = dict(key)
            End If
        Next r
    Next a
    Application.StatusBar = "Ready"
    wks.Range("A1").AutoFilter Field:=5
End Function
