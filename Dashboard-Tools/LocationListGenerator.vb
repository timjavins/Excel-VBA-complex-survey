Option Explicit

Sub LocationList()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRowB As Long
    Dim lastRowG As Long

    Set ws = ThisWorkbook.ActiveSheet

    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For Each cell In ws.Range("B1:B" & lastRowB)
        lastRowG = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
        Set rng = ws.Range("G1:G" & lastRowG).Find(cell.Value)
        If rng Is Nothing Then
            lastRowG = lastRowG + 1
            ws.Cells(lastRowG, "G").Value = cell.Value
            ws.Cells(lastRowG, "H").Value = "[INSERT LOCATION GROUP HERE]"
        End If
    Next cell
End Sub
