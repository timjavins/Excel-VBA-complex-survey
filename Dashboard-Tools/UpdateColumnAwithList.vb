Option Explicit

Sub UpdateColumnA()
    ' Updates column A by removing matched values from column F and appending new values.
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRowA As Long
    Dim lastRowF As Long
    Dim cellValue As Variant

    Set ws = ThisWorkbook.ActiveSheet

    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    For Each cell In ws.Range("F1:F" & lastRowF)
        cellValue = Format(cell.Value, "0000")
        Set rng = ws.Range("A1:A" & lastRowA).Find(cellValue)

        If Not rng Is Nothing Then
            cell.Offset(0, 1).Value = cellValue
            cell.Value = ""
        ElseIf cellValue <> "" Then
            lastRowA = lastRowA + 1
            ws.Cells(lastRowA, "A").Value = cellValue
            cell.Offset(0, 1).Value = cellValue
            cell.Value = ""
        End If
    Next cell
End Sub
