Option Explicit

Sub MoveAndDeleteRows()
    ' Checks for blank asset-mapping rows and moves unmatched device rows.
    Dim ws As Worksheet
    Dim lastRowA As Long
    Dim lastRowJ As Long
    Dim i As Long

    Set ws = ThisWorkbook.ActiveSheet

    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    i = 1
    While i <= lastRowA
        If ws.Cells(i, "C").Value = "" Then
            lastRowJ = lastRowJ + 1
            ws.Cells(lastRowJ, "J").Value = ws.Cells(i, "A").Value
            ws.Cells(lastRowJ, "K").Value = ws.Cells(i, "B").Value
            ws.Rows(i).Delete
            lastRowA = lastRowA - 1
        Else
            i = i + 1
        End If
    Wend
End Sub
