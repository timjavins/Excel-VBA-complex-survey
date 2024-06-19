Sub CleanRegionalCams()
'
' CleanRegionalCams Macro
' Clean and format the regional camera query results
'
' Keyboard Shortcut: Ctrl+Shift+A
'
    Dim lastRow As Long
    Dim i As Long
    ' Delete top 3 rows
    Rows("1:3").Select
    Selection.Delete Shift:=xlUp
    ' Find the last row
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "C").End(xlUp).Row
    ' Remove duplicates, including blanks
    ActiveSheet.Range("$B$1:$F$" & lastRow).RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlNo
    ' Update the lastRow variable
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "C").End(xlUp).Row
    ' Insert a new column to the left of the selected column
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ' Format column B as "0000"
    Columns("B:B").NumberFormat = "0000"
    ' Populate column B with characters 4 through 7 of each cell of column C
    For i = 1 To lastRow
        Cells(i, "B").Value = Mid(Cells(i, "C").Value, 4, 4)
    Next i
    ' Sort the sheet by column B
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

    ' Populate each cell of column A with its row number
    For i = 1 To lastRow
        Cells(i, "A").Value = i
    Next i
End Sub
