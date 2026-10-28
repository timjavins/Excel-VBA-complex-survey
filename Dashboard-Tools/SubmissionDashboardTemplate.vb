Option Explicit

' Checks if the middle four numbers in column A on ws2 are present in column A of ws1.
' If found, it updates the adjacent cell in ws1 to "Done".
' If not found, it appends the value to ws1 and updates the adjacent cell to "Done".
Sub CheckAndAddValues()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws2RowNum As Long
    Dim rng As Range
    Dim cell As Range
    Dim val As String
    Dim found As Range

    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sheet2")

    For Each cell In ws2.Range("A1:A" & ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row)
        ws2RowNum = cell.Row
        val = Mid(cell.Value, 4, 4)
        Set found = ws1.Columns("A:A").Find(What:=val, LookIn:=xlValues, LookAt:=xlWhole)

        If Not found Is Nothing Then
            found.Offset(0, 1).Value = "Done"
            found.Offset(0, 2).Value = ws2.Cells(ws2RowNum, "D").Value
            found.Offset(0, 3).Value = ws2.Cells(ws2RowNum, "E").Value
        Else
            Set rng = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Offset(1, 0)
            rng.Value = val
            rng.Offset(0, 1).Value = "Done"
            rng.Offset(0, 2).Value = ws2.Cells(ws2RowNum, "D").Value
            rng.Offset(0, 3).Value = ws2.Cells(ws2RowNum, "E").Value
        End If
    Next cell
End Sub

Sub RankByFrequency()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim cell As Range
    Dim freq As Object
    Dim rank As Long
    Dim keys As Variant
    Dim items As Variant
    Dim i As Long
    Dim j As Long
    Dim tempKey As Variant
    Dim tempItem As Variant
    Dim firstAddress As String
    Dim foundCell As Range

    Set ws = ThisWorkbook.Sheets("Completed Locations - Copy")
    Set dataRange = ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)
    Set freq = CreateObject("Scripting.Dictionary")

    For Each cell In dataRange
        If Not freq.Exists(cell.Value) Then
            freq.Add cell.Value, 1
        Else
            freq(cell.Value) = freq(cell.Value) + 1
        End If
    Next cell

    keys = freq.Keys
    items = freq.Items

    For i = LBound(items) To UBound(items) - 1
        For j = i + 1 To UBound(items)
            If items(i) < items(j) Then
                tempItem = items(i)
                items(i) = items(j)
                items(j) = tempItem

                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
            End If
        Next j
    Next i

    rank = 1
    For i = LBound(keys) To UBound(keys)
        Set foundCell = ws.Columns("B").Find(What:=keys(i), LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            firstAddress = foundCell.Address
            Do
                ws.Range("D" & foundCell.Row).Value = rank
                Set foundCell = ws.Columns("B").FindNext(foundCell)
            Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
            rank = rank + 1
        End If
    Next i
End Sub
