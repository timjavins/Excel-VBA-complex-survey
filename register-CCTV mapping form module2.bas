' Description: Code module 2 for the Register-CCTV Mapping Form
Public Sub LoadForm()
' TO DO: Add a MsgBox to display the number of potential register numbers, perhaps with context and reassuring messages.
    ' Set theFormIsLoading to True
    theFormIsLoading = True
    
    ' Declare counting variables (might not be necesary but doesn't hurt)
    Dim i As Long
    Dim j As Long
    Dim k As Long

    ' Initialize variables
    Set NVRList = New Collection
    Set POSList = New Collection
    Set storePOSCamRows = New Collection
    storeNumStr = Format(storeNum.Value, "0000")

    ' Confirmation message box to show the selected store number and confirm if it's correct. If not, exit the script.
    Dim yesNoAnswer As VbMsgBoxResult
    yesNoAnswer = MsgBox("Store " & storeNumStr & vbCrLf & _
           "Is this the correct store number?", vbYesNo, "Store Number Confirmation")
    If yesNoAnswer = vbNo Then Exit Sub

    ' Unprotect the wsForm sheet
    wsForm.Unprotect Password:="Be Happe"

    ' Reset the submit button and delete it
    On Error Resume Next
    Set submitButton = wsForm.Buttons("SubmitButton")
    wsForm.Buttons(submitButton.Name).Delete
    On Error GoTo 0

    ' Clear the form (A11:A, B11:B, and C11:C and only the content of C8
    wsForm.Range("C8").ClearContents
    wsForm.Range("A11:A" & wsForm.Rows.Count).Clear
    wsForm.Range("B11:B" & wsForm.Rows.Count).Clear
    wsForm.Range("C11:C" & wsForm.Rows.Count).Clear
    wsForm.Range("D8").Clear

    ' If storeNum is equal to zero or POSList is empty, display a message box and exit the sub
    If storeNum = 0 Then
        MsgBox "Is there a typo in the store number? Please check and try again.", vbExclamation, "No Register Numbers Found"
        Exit Sub
    End If

    ' Create a progress bar, using regStore.Rows.Count as the maximum value
'    Dim progressForm As UserForm
'    Set progressForm = New UserForm
'    Dim progressBar As MSForms.Control
'    Set progressBar = progressForm.Controls.Add("Forms.ProgressBar.1")
'    ' Initialize the ProgressBar
'    progressBar.Min = 0
'    progressBar.Max = regStore.Rows.Count
'    progressBar.Value = 0
'    ' Display the UserForm in cell C8
'    progressForm.Show vbModeless
'    progressForm.Top = wsForm.Range("C8").Top
'    progressForm.Left = wsForm.Range("C8").Left

    ' Populate the POSList with the register numbers for the selected store
    On Error Resume Next
    For i = 1 To regStore.Rows.Count
        If regStore.Cells(i, 1).Value = storeNum.Value Then
            POSList.Add regNum.Cells(i, 1).Value, CStr(regNum.Cells(i, 1).Value)
        End If
    Next i
    On Error GoTo 0
    
    ' Delete the progress bar
    ' Unload progressForm
    
    ' If POSList is empty display a message box and exit the sub
    If POSList.Count = 0 Then
        MsgBox "Is there a typo in the store number? Please check and try again.", vbExclamation, "No Register Numbers Found"
        Exit Sub
    End If

    ' Match the storeNumStr in wsCameras column B and add the corresponding values for
    ' columns B, C, D, E, and F to the storePOSCamRows collection
    ' This will serve as a short-list of the camera rows for the selected store
    On Error Resume Next
    For i = 1 To wsCameras.Columns("B").Cells.Count
        If wsCameras.Cells(i, 2).Value = storeNumStr Then
            ' The Add method of a collection in VBA takes two arguments: the item to add, and
            ' a unique key to associate with the item. Here, the item is the range of cells, and
            ' the key is the string representation of i (CStr(i)). The CStr function is used to
            ' convert the integer i to a string.
            storePOSCamRows.Add wsCameras.Range(wsCameras.Cells(i, 2), wsCameras.Cells(i, 6)), CStr(i)
        End If
    Next i
    On Error GoTo 0

    On Error Resume Next
    For i = 1 To storePOSCamRows.Count
        Dim bRow As Range
        Set bRow = storePOSCamRows.Item(i)
        ' Check if bRow.Cells(2).Value is already in the NVRList.
        ' If it is, skip. If it isn't, add it.
        On Error Resume Next
        Dim check As Variant
        check = NVRList.Item(CStr(bRow.Cells(2).Value))
        If Err.Number <> 0 Then
            NVRList.Add bRow.Cells(2).Value, CStr(bRow.Cells(2).Value)
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    On Error GoTo 0

    ' Populate the form with the POS register numbers
    On Error Resume Next
    For j = 1 To POSList.Count
        wsForm.Cells(j + 10, 1).Value = POSList.Item(j)
        wsForm.Cells(j + 10, 1).Locked = True
    Next j
    On Error GoTo 0

    On Error Resume Next
    ' Create a string with comma-separated NVR list items
    Dim NVRString As String
    For Each NVRItem In NVRList
        NVRString = NVRString & NVRItem & ","
    Next NVRItem
    ' Add "No camera" to the NVRString
    NVRString = NVRString & "No camera"
    On Error GoTo 0

    On Error Resume Next
    ' Populate the form with the NVR numbers as dropdown menus
    For j = 1 To POSList.Count
        'unprotect the cell to allow data entry
        wsForm.Cells(j + 10, 2).Locked = False
        With wsForm.Cells(j + 10, 2).Validation
            .Delete ' remove any previous validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=NVRString
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "Select NVR"
            .ErrorTitle = "Invalid Entry"
            .InputMessage = "Please select a NVR from the list."
            .ErrorMessage = "The value you entered is not in the list. Please select a value from the list."
            .ShowInput = True
            .ShowError = True
        End With
    Next j
    'Make all the rest of the rows after the last POS register row blank and prevent any data entry
    For k = POSList.Count + 11 To wsForm.Cells(wsForm.Rows.Count, 1).End(xlUp).Row
        wsForm.Rows(k).Locked = True
    Next k
    On Error GoTo 0

    On Error Resume Next
    ' Format cell C8 with red Arial font, size 12, bold, and italic
    wsForm.Cells(8, 3).Font.Name = "Arial"
    wsForm.Cells(8, 3).Font.Size = 12
    wsForm.Cells(8, 3).Font.Bold = True
    wsForm.Cells(8, 3).Font.Italic = True
    wsForm.Cells(8, 3).Font.Color = RGB(255, 0, 0)

    ' Display additional instructions in wsForm cell C8
    wsForm.Cells(8, 3).Value = "Choose ""No camera"" for NVR if the register is not covered. NO BLANKS."
    ' Create a button to submit the form data, located at the left edge of wsForm cell C8
    Set submitButton = wsForm.Buttons.Add(wsForm.Cells(8, 4).Left, wsForm.Cells(8, 4).Top, wsForm.Cells(8, 4).Width, wsForm.Cells(8, 4).Height)
    With submitButton
        .OnAction = "SubmitAnswers"
        .Caption = "Submit"
        .Name = "SubmitButton"
        .Locked = False
    End With
    On Error GoTo 0

    ' Camera dropdown menus will be created in sub routine CreateCameraDropdown

    'Change font color of cell A2 to green and strikethrough to indicate completion of step 1
    wsForm.Cells(2, 1).Font.Color = RGB(0, 128, 0)
    wsForm.Cells(2, 1).Font.Strikethrough = True
    'Change font color of cell A3 to red and no strikethrough to indicate active step
    wsForm.Cells(3, 1).Font.Color = RGB(255, 0, 0)
    wsForm.Cells(3, 1).Font.Strikethrough = False
    'Change all four border walls of cell A8 to black with a thin border
    For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
        wsForm.Cells(8, 1).Borders(border).Color = RGB(0, 0, 0)
        wsForm.Cells(8, 1).Borders(border).Weight = xlThin
    Next border
    'Change all four border walls of cell B11 to red to highlight the NVR dropdown menus
    For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
        wsForm.Cells(11, 2).Borders(border).Color = RGB(255, 0, 0)
    Next border

    'limit the scroll area to columns A to C
    wsForm.ScrollArea = "A1:C" & wsForm.Rows.Count

    'Protect the sheet
    wsForm.Protect Password:="Be Happe", UserInterfaceOnly:=True
    ' Set theFormIsLoading back to False
    theFormIsLoading = False
End Sub