Public wsForm As Worksheet
Public wsCameras As Worksheet
Public wsPOSregisters As Worksheet
Public storeNum As Range
Public regStore As Range
Public regNum As Range
Public POSList As Collection
Public NVRList As Collection
Public storePOSCamRows As Collection
Public relatedCams() As Variant
Public theLoadFormIsRunning As Boolean
Public submitButton As Button
Public storeNumStr As String

Sub LoadForm()
' TO DO: Add a MsgBox to display the number of potential register numbers, perhaps with context and reassuring messages.
    ' Set isLoadFormRunning to True
    theLoadFormIsRunning = True
    
    ' Declare variables
    Dim i As Long
    Dim j As Long
    Dim k As Long

    ' Initialize variables
    Set wsForm = ThisWorkbook.Sheets("FORM")
    Set wsCameras = ThisWorkbook.Sheets("Cameras")
    Set wsPOSregisters = ThisWorkbook.Sheets("POSregisters")
    Set storeNum = wsForm.Range("A5")
    Set regStore = wsPOSregisters.Range("regStore")
    Set regNum = wsPOSregisters.Range("regNum")
    Set NVRList = New Collection
    Set POSList = New Collection
    Set storePOSCamRows = New Collection
    storeNumStr = Format(storeNum.Value, "0000")

    ' Confirmation message box to show the selected store number and confirm if it's correct. If not, exit the script.
    Dim yesNoAnswer As VbMsgBoxResult
    yesNoAnswer = MsgBox("Store " & storeNumStr & vbCrLf & _
           "Is this the correct store number?", vbYesNo, "Store Number Confirmation")
    If yesNoAnswer = vbNo Then Exit Sub

    ' Reset the submit button and delete it
    On Error Resume Next
    Set submitButton = wsForm.Buttons("SubmitButton")
    wsForm.Buttons(submitButton.Name).Delete
    On Error GoTo 0

    ' Clear the form (A9:A, B9:B, and C9:C and only the content of C7)
    wsForm.Range("C7").ClearContents
    wsForm.Range("A9:A" & wsForm.Rows.Count).Clear
    wsForm.Range("B9:B" & wsForm.Rows.Count).Clear
    wsForm.Range("C9:C" & wsForm.Rows.Count).Clear
    wsForm.Range("D7").Clear

    ' If storeNum is equal to zero or POSList is empty, display a message box and exit the sub
    If storeNum = 0 Then
        MsgBox "Is there a typo in the store number? Please check and try again.", vbExclamation, "No Register Numbers Found"
        Exit Sub
    End If

    ' Populate the POSList with the register numbers for the selected store
    On Error Resume Next
    For i = 1 To regStore.Rows.Count
        If regStore.Cells(i, 1).Value = storeNum.Value Then
            POSList.Add regNum.Cells(i, 1).Value, CStr(regNum.Cells(i, 1).Value)
        End If
    Next i
    On Error GoTo 0
    
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
        wsForm.Cells(j + 8, 1).Value = POSList.Item(j)
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
        With wsForm.Cells(j + 8, 2).Validation
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
    On Error GoTo 0

    On Error Resume Next
    ' Display additional instructions in wsForm cell C7
    wsForm.Cells(7, 3).Value = "Choose ""No camera"" for Device if the register is not covered. NO BLANKS."
    ' Create a button to submit the form data, located at the left edge of wsForm cell C7
    Set submitButton = wsForm.Buttons.Add(wsForm.Cells(7, 4).Left, wsForm.Cells(7, 4).Top, wsForm.Cells(7, 4).Width, wsForm.Cells(7, 4).Height)
    With submitButton
        .OnAction = "SubmitAnswers"
        .Caption = "Submit"
        .Name = "SubmitButton"
        .Locked = False
    End With
    On Error GoTo 0

    ' Camera dropdown menus will be created in another sub routine
    ' Protect the whole workbook except wsForm cell A5, the buttons, and the dropdown menus
    ThisWorkbook.Protect Password:="[INSERT PASSWORD HERE]", Structure:=True, Windows:=False
    wsForm.Unprotect Password:="[INSERT PASSWORD HERE]"
    wsForm.Cells.Locked = True
    wsForm.Cells(5, 1).Locked = False
    ' 
    ' Set theLoadFormIsRunning back to False
    theLoadFormIsRunning = False
End Sub

    ' Dim objXMLHTTP, strURL, strResponse
    '
    ' Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    '
    ' strURL = "http://example.com/api"
    '
    ' objXMLHTTP.Open "POST", strURL, False
    ' objXMLHTTP.setRequestHeader "Content-Type", "application/json"
    ' objXMLHTTP.send "{""data"": ""csv""}"
    '
    ' strResponse = objXMLHTTP.responseText
    '
    ' Set objXMLHTTP = Nothing
    '
    ' MsgBox strResponse

Public Sub CreateCameraDropdown(activeRow As Integer, activeColumn As Integer)
    ' Get the selected NVR
    Dim selectedNVR As String
    Dim deviceCell As Range
    Set deviceCell = wsForm.Cells(activeRow, activeColumn)
    selectedNVR = deviceCell.Value
    ' Use selectedNVR to get the related cameras from the storePOSCamRows collection
    ' and create a dropdown menu in the cell to the right of the active cell
    Dim relatedCams As Collection   ' Collection to store the related cameras
    Set relatedCams = New Collection    ' Initialize the collection for related cameras for the selected NVR
    Dim camNameString As String
    On Error Resume Next
    If selectedNVR = "No camera" Then
        ' Clear the cell to the right of the active cell and remove any previous validation
        deviceCell.Offset(0, 1).Clear
        ' Change the value in the cell to the right of the active cell to "No camera"
        deviceCell.Offset(0, 1).Value = "No camera"
    Else
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

        On Error Resume Next
        camNameString = Join(camNames, ",")
        ' Clear the cell to the right of the active cell and remove any previous validation
        deviceCell.Offset(0, 1).Clear
        With deviceCell.Offset(0, 1).Validation
            .Delete ' remove any previous validation
        End With
        On Error GoTo 0

        On Error Resume Next
        ' Create a dropdown menu in the cell to the right with the camera names related to the NVR
        With deviceCell.Offset(0, 1).Validation
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
    End If
    On Error GoTo 0
End Sub

Sub SubmitAnswers()
    ' TO DO: Add a MsgBox to confirm the submission
    ' TO DO: Add code to submit the form data to the API
    ' TO DO: Add code to handle the API response
    ' TO DO: Add code to clear the form after successful submission
    ' TO DO: Add code to display a success message after successful submission
    ' TO DO: Add code to handle any errors during submission
    ' TO DO: Add code to display an error message if submission fails
    ' TO DO: Add code to log the submission status and any errors
    ' TO DO: Add code to handle any other post-submission tasks
    ' TO DO: Add code to prevent multiple submissions
    ' TO DO: change the code to include the store number and camera slot & channel numbers in submission
    ' TO DO: change the code to submit all the rows at once instead of one at a time
    Dim answer As Range
    Dim answers As Range
    Dim fileSysObj As Object
    Dim filePath As String
    Dim txtFile As Object
    Dim cell As Range
    Dim rowValue As String
    Dim rowCount As Long
    Dim NetworkErrors As Boolean

    ' Set the file path for the CSV file
    filePath = "[INSERT FILEPATH HERE]"
    On Error Resume Next
    ' Set the range to append to CSV
    rowCount = POSList.Count + 8
    Set answers = wsForm.Range("A9:C" & rowCount)
    ' Check the range for any empty cells
    For Each answer In answers
        If answer.Value = "" Then
            MsgBox "Please fill in all fields before submitting.", vbExclamation, "Incomplete Form"
            Exit Sub
        End If
    Next answer
    On Error GoTo 0
    
    ' Create FileSystemObject
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    ' Open/create the file for appending
    Set txtFile = fileSysObj.OpenTextFile(filePath, 8, True)
    ' Error handling for problems accessing the file
    If Err.Number <> 0 Then
        MsgBox "Are you on the company network/VPN? There was an error accessing network location: " & Err.Description, vbExclamation, "Network Error"
        NetworkErrors = True
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    ' Loop through each row in the range
    On Error Resume Next
    For Each row In answers.Rows
        rowValue = ""
        
        ' Loop through each cell in the row
        For Each cell In row.Cells
            ' Append the cell's text to the rowValue string
            rowValue = rowValue & cell.Value & ","
        Next cell

        ' Append the Windows username and the current date and time (Universal Time Coordinated) to the row
        rowValue = rowValue & Environ("USERNAME") & ","
        rowValue = rowValue & Format(Now(), "yyyy-mm-dd hh:mm:ss")
        
        ' Write the rowValue to the CSV file
        txtFile.WriteLine rowValue
        If Err.Number <> 0 Then
            MsgBox "Are you on the company network/VPN? There was an error accessing network location: " & Err.Description, vbExclamation, "Network Error"
            NetworkErrors = True
            Err.Clear
            Exit Sub
        End If
    Next row
    On Error GoTo 0
    
    On Error Resume Next
    ' Close the file
    txtFile.Close
    ' Error handling for problems accessing the file
    If Err.Number <> 0 Then
        MsgBox "Are you on the company network/VPN? There was an error accessing network location: " & Err.Description, vbExclamation, "Network Error"
        NetworkErrors = True
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    NetworkErrors = False
End Sub

