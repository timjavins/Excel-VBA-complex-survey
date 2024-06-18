' Description: Code module 5 for the Register-CCTV Mapping Form
Public Sub CheckAllFields()
    'Check if all POS registers have been assigned an NVR
    Dim allAssigned As Boolean
    allAssigned = True
    For i = 11 To wsForm.Cells(wsForm.Rows.Count, 1).End(xlUp).Row
        If wsForm.Cells(i, 2).Value = "" Then
            allAssigned = False
            ' Change font color of cell A3 to red and no strikethrough to indicate incomplete step 2
            wsForm.Cells(3, 1).Font.Color = RGB(255, 0, 0)
            wsForm.Cells(3, 1).Font.Strikethrough = False
            Exit For
        End If
    Next i
    If allAssigned = True Then
        ' Change font color of cell A3 to green and strikethrough to indicate completion of step 2
        wsForm.Cells(3, 1).Font.Color = RGB(0, 128, 0)
        wsForm.Cells(3, 1).Font.Strikethrough = True
    End If
    'Check if all POS registers have been assigned a camera
    For i = 11 To wsForm.Cells(wsForm.Rows.Count, 1).End(xlUp).Row
        If wsForm.Cells(i, 3).Value = "" Then
            allAssigned = False
            ' Change font color of cell A4 to red and no strikethrough to indicate incomplete step 3
            wsForm.Cells(4, 1).Font.Color = RGB(255, 0, 0)
            wsForm.Cells(4, 1).Font.Strikethrough = False
            Exit Sub
        End If
    Next i
    If allAssigned = True Then
        ' Clear the contents of cell C8
        wsForm.Cells(8, 3).Clear
        ' Change font color of cell A4 to green and strikethough to indicate completion of step 3
        wsForm.Cells(4, 1).Font.Color = RGB(0, 128, 0)
        wsForm.Cells(4, 1).Font.Strikethrough = True
        ' Change font color of cell A5 to red and no strikethrough to indicate active step 4
        wsForm.Cells(5, 1).Font.Color = RGB(255, 0, 0)
        wsForm.Cells(5, 1).Font.Strikethrough = False
        'Change all four border walls of cell D8 to red and thick to highlight the Submit button
        For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            wsForm.Cells(8, 4).Borders(border).Color = RGB(255, 0, 0)
            wsForm.Cells(8, 4).Borders(border).Weight = xlThick
        Next border
    End If
End Sub
