Option Explicit

Sub CleanLocationDevices()
'
' Cleans and formats location-device query results
'
    Dim lastRow As Long
    Dim i As Long

    Rows("1:3").Select
    Selection.Delete Shift:=xlUp

    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "C").End(xlUp).Row
    ActiveSheet.Range("$B$1:$F$" & lastRow).RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlNo

    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "C").End(xlUp).Row

    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Columns("B:B").NumberFormat = "0000"

    For i = 1 To lastRow
        Cells(i, "B").Value = Mid(Cells(i, "C").Value, 4, 4)
    Next i

    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("B1:B" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1:F" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    For i = 1 To lastRow
        Cells(i, "A").Value = i
    Next i
End Sub
