' Description: Code module 3 for the Register-CCTV Mapping Form
Public Sub CreateCameraDropdown(activeRow As Integer, activeColumn As Integer)
    ' Get the selected NVR
    Dim selectedNVR As String
    Dim NVRCell As Range
    Set NVRCell = wsForm.Cells(activeRow, activeColumn)
    selectedNVR = NVRCell.Value
    ' Use selectedNVR to get the related cameras from the storePOSCamRows collection
    ' and create a dropdown menu in the cell to the right of the active cell
    Dim relatedCams As Collection   ' Collection to store the related cameras
    Set relatedCams = New Collection    ' Initialize the collection for related cameras for the selected NVR
    Dim camNameString As String
    On Error Resume Next
    If selectedNVR = "No camera" Then
        ' Clear the cell to the right of the active cell and remove any previous validation
        NVRCell.Offset(0, 1).Clear
        ' Change the value in the cell to the right of the active cell to "No camera"
        NVRCell.Offset(0, 1).Value = "No camera"
    Else
        'unprotect the sheet
        wsForm.Unprotect Password:="Be Happe"
        On Error Resume Next
        ' Loop through the storePOSCamRows collection to find the related cameras for the selected NVR
'        Dim i As Long
        For i = 1 To storePOSCamRows.Count
            Dim bRow As Range
            Set bRow = storePOSCamRows.Item(i)
            If bRow.Cells(2).Value = selectedNVR Then
                relatedCams.Add bRow.Cells(5).Value, CStr(bRow.Cells(5).Value)
            End If
        Next i
        On Error GoTo 0
        
        On Error Resume Next
        ' Create a string with comma-separated camera names
        Dim camNames() As String
        ReDim camNames(relatedCams.Count - 1)
        For i = 1 To relatedCams.Count
            camNames(i - 1) = relatedCams.Item(i)
        Next i
        On Error GoTo 0
        'unprotect the cell to allow data entry
        NVRCell.Offset(0, 1).Locked = False
        On Error Resume Next
        camNameString = Join(camNames, ",")
        ' Clear the cell to the right of the active cell and remove any previous validation
        NVRCell.Offset(0, 1).Clear
        With NVRCell.Offset(0, 1).Validation
            .Delete ' remove any previous validation
        End With
        On Error GoTo 0

        On Error Resume Next
        ' Create a dropdown menu in the cell to the right with the camera names related to the NVR
        With NVRCell.Offset(0, 1).Validation
            .Delete ' remove any previous validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=camNameString
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "Select Camera"
            .ErrorTitle = "Invalid Entry"
            .InputMessage = "Please select a Camera from the list."
            .ErrorMessage = "The value you entered is not in the list. Please select a value from the list."
            .ShowInput = True
            .ShowError = True
        End With
        On Error GoTo 0
        ' Change all four border walls of the cell to the right of the active cell to red
        For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
            NVRCell.Offset(0, 1).Borders(border).Color = RGB(255, 0, 0)
        Next border
    End If
    On Error GoTo 0
    'unprotect the cell to allow data entry
    NVRCell.Offset(0, 1).Locked = False
    'Change Step 3 font color to red, no strikethrough
    wsForm.Cells(4, 1).Font.Strikethrough = False
    wsForm.Cells(4, 1).Font.Color = RGB(255, 0, 0)
    'Protect the sheet
    wsForm.Protect Password:="Be Happe", UserInterfaceOnly:=True
    Module5.CheckAllFields
End Sub