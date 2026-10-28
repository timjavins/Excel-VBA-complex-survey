' Description: Code module 2 for the Asset Mapping Form
Option Explicit

Public Sub LoadForm()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim yesNoAnswer As VbMsgBoxResult
    Dim locationGroupCell As Range
    Dim rowValues As Range
    Dim controllerCsv As String
    Dim controllerItem As Variant

    theFormIsLoading = True

    Set controllerList = New Collection
    Set terminalList = New Collection
    Set locationDeviceRows = New Collection
    locationIdText = Format(locationIdCell.Value, "0000")

    yesNoAnswer = MsgBox("Location " & locationIdText & vbCrLf & _
           "Is this the correct location identifier?", vbYesNo, "Location Confirmation")
    If yesNoAnswer = vbNo Then
        theFormIsLoading = False
        Exit Sub
    End If

    Set locationGroupCell = deviceLocationIndex.Find(What:=locationIdCell.Value, LookIn:=xlFormulas, LookAt:=xlWhole)
    If locationGroupCell Is Nothing Then
        MsgBox "Location " & locationIdText & " does not match any configured devices.", vbExclamation, "Devices Not Found"
        theFormIsLoading = False
        Exit Sub
    End If

    On Error GoTo DeviceSheetError
    Set wsDevices = ResolveDeviceSheet(CStr(locationGroupCell.Offset(0, 1).Value))
    On Error GoTo 0

    wsForm.Unprotect Password:=FORM_PASSWORD

    On Error Resume Next
    Set submitButton = wsForm.Buttons("SubmitButton")
    If Not submitButton Is Nothing Then
        wsForm.Buttons(submitButton.Name).Delete
    End If
    Set submitButton = Nothing
    On Error GoTo 0

    wsForm.Range("C8").ClearContents
    wsForm.Range("A11:A" & wsForm.Rows.Count).Clear
    wsForm.Range("B11:B" & wsForm.Rows.Count).Clear
    wsForm.Range("C11:C" & wsForm.Rows.Count).Clear
    wsForm.Range("D8").Clear

    If locationIdCell = 0 Then
        MsgBox "Please verify the location identifier and try again.", vbExclamation, "No Terminal IDs Found"
        theFormIsLoading = False
        Exit Sub
    End If

    For i = 1 To terminalLocationColumn.Rows.Count
        If terminalLocationColumn.Cells(i, 1).Value = locationIdCell.Value Then
            On Error Resume Next
            terminalList.Add terminalIdColumn.Cells(i, 1).Value, CStr(terminalIdColumn.Cells(i, 1).Value)
            On Error GoTo 0
        End If
    Next i

    If terminalList.Count = 0 Then
        MsgBox "Please verify the location identifier and try again.", vbExclamation, "No Terminal IDs Found"
        theFormIsLoading = False
        Exit Sub
    End If

    For i = 1 To wsDevices.Columns("B").Cells.Count
        If wsDevices.Cells(i, 2).Value = locationIdCell.Value Then
            Set rowValues = wsDevices.Range(wsDevices.Cells(i, 2), wsDevices.Cells(i, 6))
            locationDeviceRows.Add rowValues, CStr(i)
        End If
    Next i

    For i = 1 To locationDeviceRows.Count
        Dim deviceRow As Range
        Dim controllerKey As String

        Set deviceRow = locationDeviceRows.Item(i)
        controllerKey = CStr(deviceRow.Cells(2).Value)
        On Error Resume Next
        controllerList.Add deviceRow.Cells(2).Value, controllerKey
        On Error GoTo 0
    Next i

    For j = 1 To terminalList.Count
        wsForm.Cells(j + 10, 1).Value = terminalList.Item(j)
        wsForm.Cells(j + 10, 1).Locked = True
    Next j

    controllerCsv = ""
    For Each controllerItem In controllerList
        If controllerCsv <> "" Then
            controllerCsv = controllerCsv & ","
        End If
        controllerCsv = controllerCsv & CStr(controllerItem)
    Next controllerItem

    If controllerCsv <> "" Then
        controllerCsv = controllerCsv & ","
    End If
    controllerCsv = controllerCsv & NO_DEVICE_LABEL

    For j = 1 To terminalList.Count
        wsForm.Cells(j + 10, 2).Locked = False
        With wsForm.Cells(j + 10, 2).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=controllerCsv
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "Select Controller"
            .ErrorTitle = "Invalid Entry"
            .InputMessage = "Please select a controller from the list."
            .ErrorMessage = "The value entered is not in the list."
            .ShowInput = True
            .ShowError = True
        End With
    Next j

    For k = terminalList.Count + 11 To wsForm.Cells(wsForm.Rows.Count, 1).End(xlUp).Row
        wsForm.Rows(k).Locked = True
    Next k

    wsForm.Cells(8, 3).Font.Name = "Arial"
    wsForm.Cells(8, 3).Font.Size = 12
    wsForm.Cells(8, 3).Font.Bold = True
    wsForm.Cells(8, 3).Font.Italic = True
    wsForm.Cells(8, 3).Font.Color = RGB(255, 0, 0)

    wsForm.Cells(8, 3).Value = "Choose '" & NO_DEVICE_LABEL & "' if no mapped device exists. No blanks."

    Set submitButton = wsForm.Buttons.Add(wsForm.Cells(8, 4).Left, wsForm.Cells(8, 4).Top, wsForm.Cells(8, 4).Width, wsForm.Cells(8, 4).Height)
    With submitButton
        .OnAction = "SubmitAnswers"
        .Caption = "Submit"
        .Name = "SubmitButton"
        .Locked = False
    End With

    wsForm.Cells(2, 1).Font.Color = RGB(0, 128, 0)
    wsForm.Cells(2, 1).Font.Strikethrough = True
    wsForm.Cells(3, 1).Font.Color = RGB(255, 0, 0)
    wsForm.Cells(3, 1).Font.Strikethrough = False
    wsForm.Cells(4, 1).Font.Color = RGB(0, 0, 0)
    wsForm.Cells(4, 1).Font.Strikethrough = False
    wsForm.Cells(5, 1).Font.Color = RGB(0, 0, 0)
    wsForm.Cells(5, 1).Font.Strikethrough = False

    For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
        wsForm.Cells(8, 1).Borders(border).Color = RGB(0, 0, 0)
        wsForm.Cells(8, 1).Borders(border).Weight = xlThin
    Next border

    For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
        wsForm.Cells(11, 2).Borders(border).Color = RGB(255, 0, 0)
    Next border

    wsForm.ScrollArea = "A1:C" & wsForm.Rows.Count
    wsForm.Protect Password:=FORM_PASSWORD, UserInterfaceOnly:=True
    theFormIsLoading = False
    Exit Sub

DeviceSheetError:
    MsgBox "Could not find a device sheet for location group '" & CStr(locationGroupCell.Offset(0, 1).Value) & "'.", vbExclamation, "Configuration Error"
    theFormIsLoading = False
End Sub
