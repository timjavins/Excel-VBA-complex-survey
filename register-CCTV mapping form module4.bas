' Description: Code module 4 for the Register-CCTV Mapping Form
Public Sub AddPortChannelStoreNums(activeRow As Integer, activeColumn As Integer)
    ' Unprotect the wsForm sheet
    wsForm.Unprotect Password:=FORM_PASSWORD

    ' Get the selected Camera
    Set camCell = wsForm.Cells(activeRow, activeColumn)
    selectedCam = camCell.Value
    
    ' unprotect the three cells to the right of the active cell
    camCell.Offset(0, 1).Locked = False
    camCell.Offset(0, 2).Locked = False
    camCell.Offset(0, 3).Locked = False
    'Make the content of the three cells to the right of the active cell invisible
    camCell.Offset(0, 1).Font.Color = RGB(255, 255, 255)
    camCell.Offset(0, 2).Font.Color = RGB(255, 255, 255)
    camCell.Offset(0, 3).Font.Color = RGB(255, 255, 255)

    On Error Resume Next
    ' Handle the case where the selected camera is "No camera"
    If selectedCam = "No camera" Or selectedCam = "" Then
        camCell.Offset(0, 1).Value = 0
        camCell.Offset(0, 2).Value = 0
        camCell.Offset(0, 3).Value = storeNumStr
        'Format the storeNumStr cell as 0000
        camCell.Offset(0, 3).NumberFormat = "0000"
    ' Use selectedCam to get the related port and channel numbers from the storePOSCamRows collection
    Else
        On Error Resume Next
        ' Clear all four borders of the modified cell
        For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            camCell.Borders(border).LineStyle = xlNone
        Next border
        ' Loop through the storePOSCamRows collection to find the related port & channel numbers for the selected camera
        For i = 1 To storePOSCamRows.Count
            Dim bRow As Range
            Set bRow = storePOSCamRows.Item(i)
            If bRow.Cells(5).Value = selectedCam Then
                ' Populate the three cells to the right of the active cell with the related port and channel numbers and format the number
                camCell.Offset(0, 1).Value = bRow.Cells(3).Value
                camCell.Offset(0, 1).NumberFormat = "00"
                camCell.Offset(0, 2).Value = bRow.Cells(4).Value
                camCell.Offset(0, 2).NumberFormat = "00"
                camCell.Offset(0, 3).Value = storeNumStr
                camCell.Offset(0, 3).NumberFormat = "0000"
            End If
        Next i
        On Error GoTo 0
    End If
    On Error GoTo 0
    ' Protect the wsForm sheet
    wsForm.Protect Password:=FORM_PASSWORD, UserInterfaceOnly:=True

    Module5.CheckAllFields
End Sub
