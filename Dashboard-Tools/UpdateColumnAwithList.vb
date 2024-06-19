Sub UpdateColumnA()
    ' This sub routine takes a list in column F and updates column A by removing
    ' any matches from column F and appending any new values to the end of column A.
    ' This is to update the column A list so that it includes all target values.
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRowA As Long
    Dim lastRowF As Long
    Dim cellValue As Variant
    
    ' Set the worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Find the last row in column A and F
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    ' Loop through each cell in column F
    For Each cell In ws.Range("F1:F" & lastRowF)
        cellValue = Format(cell.Value, "0000")
        ' Check if the cell's value is in column A
        Set rng = ws.Range("A1:A" & lastRowA).Find(cellValue)
        ' If the cell's value is in column A
        If Not rng Is Nothing Then
            ' Copy the cell to the next cell to the right
            cell.Offset(0, 1).Value = cellValue
            ' Clear the cell
            cell.Value = ""
        ' If the cell's value is not in column A
        ElseIf cellValue <> "" Then
            ' Append the cell to the end of column A
            lastRowA = lastRowA + 1
            ws.Cells(lastRowA, "A").Value = cellValue
            ' Copy the cell to the next cell to the right
            cell.Offset(0, 1).Value = cellValue
            cell.value = ""
        End If
    Next cell
End Sub