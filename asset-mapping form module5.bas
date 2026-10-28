' Description: Code module 5 for the Asset Mapping Form
Option Explicit

Public Sub CheckAllFields()
    Dim allAssigned As Boolean
    Dim i As Long

    allAssigned = True

    For i = 11 To wsForm.Cells(wsForm.Rows.Count, 1).End(xlUp).Row
        If wsForm.Cells(i, 2).Value = "" Then
            allAssigned = False
            wsForm.Cells(3, 1).Font.Color = RGB(255, 0, 0)
            wsForm.Cells(3, 1).Font.Strikethrough = False
            Exit For
        End If
    Next i

    If allAssigned = True Then
        wsForm.Cells(3, 1).Font.Color = RGB(0, 128, 0)
        wsForm.Cells(3, 1).Font.Strikethrough = True
    End If

    For i = 11 To wsForm.Cells(wsForm.Rows.Count, 1).End(xlUp).Row
        If wsForm.Cells(i, 3).Value = "" Then
            allAssigned = False
            wsForm.Cells(4, 1).Font.Color = RGB(255, 0, 0)
            wsForm.Cells(4, 1).Font.Strikethrough = False
            Exit Sub
        End If
    Next i

    If allAssigned = True Then
        wsForm.Cells(8, 3).Clear
        wsForm.Cells(4, 1).Font.Color = RGB(0, 128, 0)
        wsForm.Cells(4, 1).Font.Strikethrough = True
        wsForm.Cells(5, 1).Font.Color = RGB(255, 0, 0)
        wsForm.Cells(5, 1).Font.Strikethrough = False

        For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            wsForm.Cells(8, 4).Borders(border).Color = RGB(255, 0, 0)
            wsForm.Cells(8, 4).Borders(border).Weight = xlThick
        Next border
    End If
End Sub
