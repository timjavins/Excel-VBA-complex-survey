Sub MoveAndDeleteRows()
    ' This sub routine checks for blank cells in a join of the camera and POS lists.
    ' If there is a row with blank spaces, it moves the camera list entry to
    ' a new list. It then deletes the row from the original list.
    ' The camera list is in columns A and B, and the POS list is in columns C and D.
    ' Columns J and K must be empty. The first several rows of the join must
    ' have no blanks in columns C and D in order for the new list (J and K) to
    ' to stay behind the deletions.
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRowA As Long
    Dim lastRowJ As Long
    Dim i As Long
    
    ' Set the worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Find the last row in column A and J
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    ' Loop through each cell in column A
    i = 1
    While i <= lastRowA
        ' Check if the corresponding cell in column C is blank
        If ws.Cells(i, "C").Value = "" Then
            ' Move the values from columns A and B to columns J and K
            lastRowJ = lastRowJ + 1
            ws.Cells(lastRowJ, "J").Value = ws.Cells(i, "A").Value
            ws.Cells(lastRowJ, "K").Value = ws.Cells(i, "B").Value
            ' Delete the row
            ws.Rows(i).Delete
            ' Decrease the last row count as we have deleted a row
            lastRowA = lastRowA - 1
        Else
            ' Only increment the counter if we didn't delete a row
            i = i + 1
        End If
    Wend
End Sub