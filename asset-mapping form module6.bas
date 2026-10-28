' Description: Code module 6 for the Asset Mapping Form
Option Explicit

Public Sub SubmitAnswers()
    Dim answer As Range
    Dim answers As Range
    Dim fileSysObj As Object
    Dim filePath As String
    Dim txtFile As Object
    Dim cell As Range
    Dim rowValue As String
    Dim rowCount As Long
    Dim completedLocationsPath As String
    Dim completedLocationsFile As Object
    Dim currentRow As Range
    Dim cellValue As String
    Dim controllerPort As String
    Dim timestampValue As String

    rowCount = terminalList.Count + 10
    Set answers = wsForm.Range("A11:F" & rowCount)

    For Each answer In answers
        If Trim$(CStr(answer.Value)) = "" Then
            MsgBox "Please fill in all fields before submitting.", vbExclamation, "Incomplete Form"
            Exit Sub
        End If
    Next answer

    timestampValue = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    filePath = BASE_DATA_PATH & Environ("USERNAME") & "_asset_mapping_data.csv"
    completedLocationsPath = BASE_DATA_PATH & "Completed Locations.csv"

    Set fileSysObj = CreateObject("Scripting.FileSystemObject")

    On Error GoTo OpenSubmissionError
    Set txtFile = fileSysObj.OpenTextFile(filePath, 8, True)

    On Error GoTo WriteSubmissionError
    For Each currentRow In answers.Rows
        rowValue = ""
        controllerPort = ""

        For Each cell In currentRow.Cells
            Select Case cell.Column
                Case 4, 5
                    cellValue = Format(cell.Value, "00")
                Case 6
                    cellValue = Format(cell.Value, "0000")
                Case Else
                    cellValue = CStr(cell.Value)
            End Select

            If cell.Column = 2 Then
                controllerPort = cellValue & "-" & Format(cell.Offset(0, 2).Value, "00")
            End If

            If rowValue <> "" Then
                rowValue = rowValue & ","
            End If
            rowValue = rowValue & EscapeCsvValue(cellValue)
        Next cell

        rowValue = rowValue & "," & EscapeCsvValue(controllerPort)
        rowValue = rowValue & "," & EscapeCsvValue(Environ("USERNAME"))
        rowValue = rowValue & "," & EscapeCsvValue(timestampValue)

        txtFile.WriteLine rowValue
    Next currentRow

    On Error GoTo OpenCompletionError
    Set completedLocationsFile = fileSysObj.OpenTextFile(completedLocationsPath, 8, True)

    On Error GoTo WriteCompletionError
    completedLocationsFile.WriteLine EscapeCsvValue(Format(locationIdCell.Value, "0000")) & "," & _
                                     EscapeCsvValue(Environ("USERNAME")) & "," & _
                                     EscapeCsvValue(timestampValue)

    On Error Resume Next
    txtFile.Close
    completedLocationsFile.Close
    On Error GoTo 0

    wsForm.Cells(5, 1).Font.Color = RGB(0, 128, 0)
    wsForm.Cells(5, 1).Font.Strikethrough = True

    MsgBox "Submission successful for location " & Format(locationIdCell.Value, "0000") & ".", vbInformation, "Success"

    For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
        wsForm.Cells(8, 1).Borders(border).Color = RGB(255, 0, 0)
    Next border

    Exit Sub

OpenSubmissionError:
    NetworkError "Unable to open the submission file at the configured network path."
    Exit Sub

WriteSubmissionError:
    NetworkError "Unable to write submission rows to the output CSV file."
    On Error Resume Next
    If Not txtFile Is Nothing Then txtFile.Close
    On Error GoTo 0
    Exit Sub

OpenCompletionError:
    NetworkError "Unable to open the completion log file at the configured network path."
    On Error Resume Next
    If Not txtFile Is Nothing Then txtFile.Close
    On Error GoTo 0
    Exit Sub

WriteCompletionError:
    NetworkError "Unable to write to the completion log file."
    On Error Resume Next
    If Not txtFile Is Nothing Then txtFile.Close
    If Not completedLocationsFile Is Nothing Then completedLocationsFile.Close
    On Error GoTo 0
End Sub

Private Function EscapeCsvValue(ByVal rawValue As String) As String
    Dim escapedValue As String

    escapedValue = Replace(rawValue, Chr$(34), Chr$(34) & Chr$(34))
    If InStr(escapedValue, ",") > 0 Or InStr(escapedValue, Chr$(34)) > 0 Or InStr(escapedValue, vbCr) > 0 Or InStr(escapedValue, vbLf) > 0 Then
        EscapeCsvValue = Chr$(34) & escapedValue & Chr$(34)
    Else
        EscapeCsvValue = escapedValue
    End If
End Function

Public Sub NetworkError(ByVal details As String)
    MsgBox details & vbCrLf & "Check network/VPN access and folder permissions.", vbExclamation, "Network Error"
    Err.Clear
End Sub
