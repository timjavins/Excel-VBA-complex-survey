Sub DeleteRows()
    Dim ws As Worksheet
    Dim rng As Range
    Dim i As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("POSregisters") ' Change to your sheet name

    ' Start from the last row and go up
    For i = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row To 1 Step -1
        ' Check if the cell value matches the format "13xx"
        If ws.Cells(i, "F").Value Like "13??" Then
            ' If it does, delete the row
            ws.Rows(i).Delete
        End If
    Next i
End Sub