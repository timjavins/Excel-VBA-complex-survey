Option Explicit

' Configure this pattern for your local identifier cleanup rule.
Private Const SAMPLE_IDENTIFIER_PATTERN As String = "13??"

Sub DeleteRowsByPattern()
    Dim ws As Worksheet
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Terminals")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("POSregisters")
    End If
    If ws Is Nothing Then Exit Sub

    For i = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row To 1 Step -1
        If CStr(ws.Cells(i, "F").Value) Like SAMPLE_IDENTIFIER_PATTERN Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
