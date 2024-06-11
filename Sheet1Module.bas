Private Sub Worksheet_Change(ByVal Target As Range)
    ' Check if LoadForm is running
    If theLoadFormIsRunning Then
        Exit Sub
    End If
    ' Check if wsForm and POSList have been initialized
    If wsForm Is Nothing Or POSList Is Nothing Then
        Exit Sub
    End If
    ' Check if the changed cell is one of the Device cells
    If Not Intersect(Target, wsForm.Cells(9, 2).Resize(POSList.Count)) Is Nothing Then
        ' Check if the active cell has a non-empty value
        If Target.Value <> "" Then
            ' Call the CreateCameraDropdown function in the general module
            Module1.CreateCameraDropdown Target.Row, Target.Column
        End If
    End If
End Sub
