' Description: Code module 4 for the Asset Mapping Form
Option Explicit

Public Sub AddPortChannelLocation(activeRow As Integer, activeColumn As Integer)
    Dim i As Long
    Dim rowValues As Range

    wsForm.Unprotect Password:=FORM_PASSWORD

    Set deviceCell = wsForm.Cells(activeRow, activeColumn)
    selectedDevice = CStr(deviceCell.Value)

    deviceCell.Offset(0, 1).Locked = False
    deviceCell.Offset(0, 2).Locked = False
    deviceCell.Offset(0, 3).Locked = False
    deviceCell.Offset(0, 1).Font.Color = RGB(255, 255, 255)
    deviceCell.Offset(0, 2).Font.Color = RGB(255, 255, 255)
    deviceCell.Offset(0, 3).Font.Color = RGB(255, 255, 255)

    If selectedDevice = NO_DEVICE_LABEL Or selectedDevice = "" Then
        deviceCell.Offset(0, 1).Value = 0
        deviceCell.Offset(0, 2).Value = 0
        deviceCell.Offset(0, 3).Value = locationIdText
        deviceCell.Offset(0, 3).NumberFormat = "0000"
    Else
        For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            deviceCell.Borders(border).LineStyle = xlNone
        Next border

        For i = 1 To locationDeviceRows.Count
            Set rowValues = locationDeviceRows.Item(i)
            If rowValues.Cells(5).Value = selectedDevice Then
                deviceCell.Offset(0, 1).Value = rowValues.Cells(3).Value
                deviceCell.Offset(0, 1).NumberFormat = "00"
                deviceCell.Offset(0, 2).Value = rowValues.Cells(4).Value
                deviceCell.Offset(0, 2).NumberFormat = "00"
                deviceCell.Offset(0, 3).Value = locationIdText
                deviceCell.Offset(0, 3).NumberFormat = "0000"
                Exit For
            End If
        Next i
    End If

    wsForm.Protect Password:=FORM_PASSWORD, UserInterfaceOnly:=True
    Module5.CheckAllFields
End Sub
