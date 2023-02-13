' vba macro to detect duplicate records in a table


Sub test()

Dim i As Long
Dim j As Long

Dim ws As Worksheet
Dim tbl As ListObject
Dim rng As Range

Set ws = ThisWorkbook.Worksheets("Sheet1")
Set tbl = ws.ListObjects("Table1")
Set rng = tbl.DataBodyRange

For i = 1 To rng.Rows.Count
    For j = i + 1 To rng.Rows.Count
        If rng.Cells(i, 1).Value = rng.Cells(j, 1).Value Then
            rng.Cells(j, 1).Interior.Color = vbRed
        End If
    Next j
Next i
    
    End Sub

' this work
' vba macro to detect equal records in two sheets


Sub test2()

Dim i As Long
Dim j As Long

Dim ws As Worksheet
Dim tbl As ListObject
Dim rng As Range

Dim ws2 As Worksheet
Dim tbl2 As ListObject
Dim rng2 As Range

Set ws = ThisWorkbook.Worksheets("Sheet1")
Set tbl = ws.ListObjects("Table1")
Set rng = tbl.DataBodyRange

Set ws2 = ThisWorkbook.Worksheets("Sheet2")
Set tbl2 = ws2.ListObjects("Table2")
Set rng2 = tbl2.DataBodyRange

For i = 1 To rng.Rows.Count
    For j = 1 To rng2.Rows.Count
        If rng.Cells(i, 1).Value = rng2.Cells(j, 1).Value Then
            rng2.Cells(j, 1).Interior.Color = vbRed
        End If
    Next j
Next i
    
    End Sub


' vba macro to detect equal records in two sheets  and move them to a third sheet



Sub test3()

Dim i As Long
Dim j As Long

Dim ws As Worksheet
Dim tbl As ListObject
Dim rng As Range

Dim ws2 As Worksheet
Dim tbl2 As ListObject
Dim rng2 As Range

Dim ws3 As Worksheet
Dim tbl3 As ListObject
Dim rng3 As Range

Set ws = ThisWorkbook.Worksheets("Sheet3")
Set tbl = ws.ListObjects("Table3")
Set rng = tbl.DataBodyRange

Set ws2 = ThisWorkbook.Worksheets("Sheet4")
Set tbl2 = ws2.ListObjects("Table4")
Set rng2 = tbl2.DataBodyRange

Set ws3 = ThisWorkbook.Worksheets("Sheet5")
Set tbl3 = ws3.ListObjects("Table6")
Set rng3 = tbl3.DataBodyRange

For i = 1 To rng.Rows.Count
    For j = 1 To rng2.Rows.Count
        If rng.Cells(i, 1).Value = rng2.Cells(j, 1).Value Then
            rng2.Cells(j, 1).Interior.Color = vbRed
            rng2.Cells(j, 1).Copy
            rng3.Cells(rng3.Rows.Count, 1).End(xlUp).Offset(i).PasteSpecial
        End If
    Next j
Next i
    
    End Sub






















