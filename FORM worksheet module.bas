' Description: This module contains the code for the FORM worksheet module.
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    ' Check if the form is loading or if the POSList is not set
    If theFormIsLoading Or POSList Is Nothing Then
        Exit Sub
    End If

    ' Check if the changed cell is in the NVR column and is not empty
    If Not Intersect(Target, wsForm.Cells(11, 2).Resize(POSList.Count)) Is Nothing And Target.Value <> "" Then
        ' Remove the border formatting from cell B11
        For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            wsForm.Cells(11, 2).Borders(border).LineStyle = xlNone
        Next
        ' Call the CreateCameraDropdown function in the general module
        Module3.CreateCameraDropdown Target.Row, Target.Column
    ' Check if the changed cell is in the Camera column and is not empty
    ElseIf Not Intersect(Target, wsForm.Cells(11, 3).Resize(POSList.Count)) Is Nothing And Target.Value <> "" Then
        ' Remove the border formatting from Target cell
        For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            Target.Borders(border).LineStyle = xlNone
        Next
        ' Call the AddPortChannelStoreNums function in the general module
        Module4.AddPortChannelStoreNums Target.Row, Target.Column
    End If
    On Error GoTo 0
End Sub
