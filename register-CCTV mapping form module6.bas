' Description: Code module 6 for the Register-CCTV Mapping Form
Public Sub SubmitAnswers()
    ' TO DO: Add code to submit the form data to the API
    ' TO DO: Add code to handle the API response
    ' TO DO: Add code to clear the form after successful submission
    ' TO DO: Add more code to handle any errors during submission
    ' TO DO: Add code to log the submission status and any errors
    ' TO DO: Add code to handle any other post-submission tasks
    ' TO DO: Add code to prevent multiple submissions
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
    filePath = "\\A0319P1116\file_repo\inbound\APREGUPDATE\" & Environ("USERNAME") & "_form_data.csv"
    On Error Resume Next
    ' Set the range to append to CSV
    rowCount = POSList.Count + 10
    Set answers = wsForm.Range("A11:F" & rowCount)
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
        NetworkErrors = True
        NetworkError
        Exit Sub
    End If
    On Error GoTo 0

    On Error Resume Next
    ' Loop through each row in the range
    For Each Row In answers.Rows
        rowValue = ""
        
        ' Loop through each cell in the row
        For Each cell In Row.Cells
            ' Check if the cell is one of the cells that need leading zeroes
            If cell.Column = 4 Or cell.Column = 5 Then
                ' Format the cell's value as a string with leading zeroes
                cellValue = Format(cell.Value, "00")
            ElseIf cell.Column = 6 Then
                ' Format the cell's value as a string with leading zeroes
                cellValue = Format(cell.Value, "0000")
            Else
                ' Convert the cell's value to a string
                cellValue = CStr(cell.Value)
            End If
            ' concatenate cell 2 and 4 with a dash between
            If cell.Column = 2 Then
                NVRPort = cellValue & "-" & Format(cell.Offset(0, 2).Value, "00")
            End If

            ' Append the cell's text to the rowValue string
            rowValue = rowValue & cellValue & ","
        Next cell

        ' Append the Windows username and the current date and time (Universal Time Coordinated) to the row
        rowValue = rowValue & NVRPort & ","
        rowValue = rowValue & Environ("USERNAME") & ","
        rowValue = rowValue & Format(Now(), "yyyy-mm-dd hh:mm:ss")
        
        ' Write the rowValue to the CSV file
        txtFile.WriteLine rowValue
        If Err.Number <> 0 Then
            NetworkErrors = True
            NetworkError
            Exit Sub
        End If
    Next Row
    On Error GoTo 0
    
    On Error Resume Next
    ' Set the file path for the "Completed Stores" file
    Dim completedStoresPath As String
    completedStoresPath = "\\A0319P1116\file_repo\inbound\APREGUPDATE\Completed Stores.csv"
    ' Open/create the "Completed Stores" file for appending
    Dim completedStores As Object
    Set completedStores = fileSysObj.OpenTextFile(completedStoresPath, 8, True)
    ' Write the storeNum, username, and date-time to the "Completed Stores" file
    completedStores.WriteLine storeNum & "," & Environ("USERNAME") & "," & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    ' Error handling for problems accessing the "Completed Stores" file
    If Err.Number <> 0 Then
        NetworkErrors = True
        NetworkError
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error Resume Next
    ' Close the submissions file
    txtFile.Close
    ' Close the "Completed Stores" file
    completedStores.Close
    ' Error handling for problems accessing the files
    If Err.Number <> 0 Then
        NetworkErrors = True
        NetworkError
        Exit Sub
    End If
    On Error GoTo 0
    ' Change font color of cell A5 to green and strikethrough to indicate completion of step 4
    wsForm.Cells(5, 1).Font.Color = RGB(0, 128, 0)
    wsForm.Cells(5, 1).Font.Strikethrough = True

    NetworkErrors = False
    MsgBox "You are here to win! Your form submission was successful. Thank you for owning this process for store " & storeNum & ".", vbExclamation, "Success!"
   'Change all four border walls of cell A8 to red to revert to the beginning of the guide
    For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
        wsForm.Cells(8, 1).Borders(border).Color = RGB(255, 0, 0)
    Next border
End Sub

Sub NetworkError()
    MsgBox "Are you on the company network/VPN? There was an error submitting on the network.", vbExclamation, "Network Error"
    Err.Clear
    Exit Sub
End Sub
