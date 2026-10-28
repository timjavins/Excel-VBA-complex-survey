' Description: Code module 3 for the Asset Mapping Form
Option Explicit

Public Sub CreateDeviceDropdown(activeRow As Integer, activeColumn As Integer)
    Dim selectedControllerValue As String
    Dim activeControllerCell As Range
    Dim relatedDeviceList As Collection
    Dim deviceNameList() As String
    Dim i As Long
    Dim deviceNameCsv As String
    Dim rowValues As Range

    Set activeControllerCell = wsForm.Cells(activeRow, activeColumn)
    selectedControllerValue = CStr(activeControllerCell.Value)
    Set relatedDeviceList = New Collection

    wsForm.Unprotect Password:=FORM_PASSWORD

    If selectedControllerValue = NO_DEVICE_LABEL Then
        activeControllerCell.Offset(0, 1).Clear
        activeControllerCell.Offset(0, 1).Value = NO_DEVICE_LABEL
    Else
        For i = 1 To locationDeviceRows.Count
            Set rowValues = locationDeviceRows.Item(i)
            If rowValues.Cells(2).Value = selectedControllerValue Then
                On Error Resume Next
                relatedDeviceList.Add rowValues.Cells(5).Value, CStr(rowValues.Cells(5).Value)
                On Error GoTo 0
            End If
        Next i

        If relatedDeviceList.Count > 0 Then
            ReDim deviceNameList(relatedDeviceList.Count - 1)
            For i = 1 To relatedDeviceList.Count
                deviceNameList(i - 1) = CStr(relatedDeviceList.Item(i))
            Next i
            deviceNameCsv = Join(deviceNameList, ",")
        Else
            deviceNameCsv = NO_DEVICE_LABEL
        End If

        activeControllerCell.Offset(0, 1).Locked = False
        activeControllerCell.Offset(0, 1).Clear
        With activeControllerCell.Offset(0, 1).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=deviceNameCsv
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "Select Device"
            .ErrorTitle = "Invalid Entry"
            .InputMessage = "Please select a device from the list."
            .ErrorMessage = "The value entered is not in the list."
            .ShowInput = True
            .ShowError = True
        End With

        For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            activeControllerCell.Offset(0, 1).Borders(border).Color = RGB(255, 0, 0)
        Next border
    End If

    activeControllerCell.Offset(0, 1).Locked = False
    wsForm.Cells(4, 1).Font.Strikethrough = False
    wsForm.Cells(4, 1).Font.Color = RGB(255, 0, 0)
    wsForm.Cells(5, 1).Font.Strikethrough = False
    wsForm.Cells(5, 1).Font.Color = RGB(0, 0, 0)

    For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
        wsForm.Cells(8, 4).Borders(border).LineStyle = xlNone
    Next border

    wsForm.Protect Password:=FORM_PASSWORD, UserInterfaceOnly:=True
    Module5.CheckAllFields
End Sub
