' Description: FORM worksheet event handlers for the Asset Mapping Form.
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If theFormIsLoading Or terminalList Is Nothing Then
        Exit Sub
    End If

    If Not Intersect(Target, wsForm.Cells(11, 2).Resize(terminalList.Count)) Is Nothing And Target.Value <> "" Then
        For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            wsForm.Cells(11, 2).Borders(border).LineStyle = xlNone
        Next border
        Module3.CreateDeviceDropdown Target.Row, Target.Column
    ElseIf Not Intersect(Target, wsForm.Cells(11, 3).Resize(terminalList.Count)) Is Nothing And Target.Value <> "" Then
        For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            Target.Borders(border).LineStyle = xlNone
        Next border
        Module4.AddPortChannelLocation Target.Row, Target.Column
    End If
End Sub
