Sub RegionalList()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRowB As Long
    Dim lastRowG As Long
    
    ' Set the worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Find the last row in column B
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    
    ' Loop through each cell in column B
    For Each cell In ws.Range("B1:B" & lastRowB)
        ' Find the last row in column G
        lastRowG = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
        ' Check if the cell's value is in column G
        Set rng = ws.Range("G1:G" & lastRowG).Find(cell.Value)
        ' If the cell's value is not in column G, add it
        If rng Is Nothing Then
            lastRowG = lastRowG + 1
            ws.Cells(lastRowG, "G").Value = cell.Value
            ' Set value to the right to the region name
            ws.Cells(lastRowG, "H").Value = "[INSERT REGION HERE]"
        End If
    Next cell
End Sub
