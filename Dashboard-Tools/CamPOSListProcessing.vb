Sub CamPOSListProcessing()
    
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRowA As Long
    Dim lastRowE As Long
    Dim lastRowG As Long
    Dim cellValue As Variant
    
    ' Set the worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Find the last row in column A, E and G
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    lastRowG = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' NOTE: This code assumes that there is a header row
    ' and thus loops from row 2 to the last row.
    ' Column A and B are store number and region from the camera list.
    ' Column C and D are blank.
    ' Column E and F are the store number and region from the POS list.
    ' Column G and H are blank.

    ' Loop through each cell in column E
    For Each cell In ws.Range("E1:E" & lastRowE)
        cellValue = cell.Value
        ' Check if the cell's value is in column A
        Set rng = ws.Range("A1:A" & lastRowA).Find(cellValue)
        ' If the cell's value is in column A
        If Not rng Is Nothing Then
            ' Set corresponding column C value to the active column E value
            rng.Offset(0, 2).Value = cellValue
            ' Set the corresponding column D value to the column F value that is next to the active column E cell
            rng.Offset(0, 3).Value = cell.Offset(0, 1).Value
            ' Clear the column E and F values
            cell.Offset(0, 1).Value = ""
            cell.Value = ""
        ' If the cell's value is not in column A
        Else
            ' Increment the last row in column G
            lastRowG = lastRowG + 1
            ' Add the active column E value and the corresponding column G value to the first available cell in columns F and G
            ws.Cells(lastRowG, "G").Value = cellValue
            ws.Cells(lastRowG, "H").Value = cell.Offset(0, 1).Value
            ' Clear the active column D and E values
            cell.Value = ""
            cell.Offset(0, 1).Value = ""
        End If
    Next cell
End Sub
