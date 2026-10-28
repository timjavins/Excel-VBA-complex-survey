Option Explicit

Sub DeviceAssetListProcessing()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRowA As Long
    Dim lastRowE As Long
    Dim lastRowG As Long
    Dim cellValue As Variant

    Set ws = ThisWorkbook.ActiveSheet

    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    lastRowG = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    ' Columns A/B: device location list.
    ' Columns E/F: asset location list.
    For Each cell In ws.Range("E1:E" & lastRowE)
        cellValue = cell.Value
        Set rng = ws.Range("A1:A" & lastRowA).Find(cellValue)

        If Not rng Is Nothing Then
            rng.Offset(0, 2).Value = cellValue
            rng.Offset(0, 3).Value = cell.Offset(0, 1).Value
            cell.Offset(0, 1).Value = ""
            cell.Value = ""
        Else
            lastRowG = lastRowG + 1
            ws.Cells(lastRowG, "G").Value = cellValue
            ws.Cells(lastRowG, "H").Value = cell.Offset(0, 1).Value
            cell.Value = ""
            cell.Offset(0, 1).Value = ""
        End If
    Next cell
End Sub
