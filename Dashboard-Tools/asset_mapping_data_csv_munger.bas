' Description: VBA helpers for post-processing asset mapping CSV output
Option Explicit

Const TEST_USERNAME As String = "[LAN ID]" ' Optional test username to filter out
Const ONEDRIVE_ORG As String = "[COMPANY]"  ' OneDrive organization name
Const FILE_PATH As String = "[Folder path to the controller_data.csv file]" ' Relative path from OneDrive root
Const UNASSIGNED_LABEL As String = "No device"

Sub DeleteTesterRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("asset_mapping_data")
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    For i = lastRow To 1 Step -1
        If ws.Cells(i, "F").Value = TEST_USERNAME Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Sub DeleteUnassignedRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("asset_mapping_data")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = lastRow To 1 Step -1
        If ws.Cells(i, "A").Value = UNASSIGNED_LABEL Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Sub LookupEndpointIP()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim controllerId As String

    Set ws = ThisWorkbook.Sheets("asset_mapping_data")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 1 To lastRow
        controllerId = ws.Cells(i, "A").Value
        ws.Cells(i, "B").Value = LookupIPFromCSV(controllerId)
    Next i
End Sub

Function LookupIPFromCSV(controllerId As String) As String
    Dim filePath As String
    Dim fileSysObj As Object
    Dim txtFile As Object
    Dim line As String
    Dim parts() As String

    filePath = "C:\Users\" & Environ("USERNAME") & "\OneDrive - " & ONEDRIVE_ORG & FILE_PATH
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")

    On Error GoTo FileOpenError
    Set txtFile = fileSysObj.OpenTextFile(filePath, 1, False)

    Do While Not txtFile.AtEndOfStream
        line = txtFile.ReadLine
        parts = Split(line, ",")
        If UBound(parts) >= 1 Then
            If parts(0) = controllerId Then
                LookupIPFromCSV = parts(1)
                txtFile.Close
                Exit Function
            End If
        End If
    Loop

    txtFile.Close
    LookupIPFromCSV = "N/A"
    Exit Function

FileOpenError:
    LookupIPFromCSV = "N/A"
    MsgBox "Error opening file: " & Err.Description, vbExclamation, "File Error"
End Function

Sub LookupEndpointChannel()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim controllerId As String
    Dim endpointName As String

    Set ws = ThisWorkbook.Sheets("asset_mapping_data")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 1 To lastRow
        controllerId = ws.Cells(i, "A").Value
        endpointName = ws.Cells(i, "E").Value
        ws.Cells(i, "D").Value = LookupChannelFromCSV(controllerId, endpointName)
    Next i
End Sub

Function LookupChannelFromCSV(controllerId As String, endpointName As String) As String
    Dim filePath As String
    Dim fileSysObj As Object
    Dim txtFile As Object
    Dim line As String
    Dim parts() As String

    filePath = "C:\Users\" & Environ("USERNAME") & "\OneDrive - " & ONEDRIVE_ORG & FILE_PATH & "-cleaned+mapped.csv"
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")

    On Error GoTo FileOpenError
    Set txtFile = fileSysObj.OpenTextFile(filePath, 1, False)

    Do While Not txtFile.AtEndOfStream
        line = txtFile.ReadLine
        parts = Split(line, ",")
        If UBound(parts) >= 5 Then
            If parts(2) = controllerId And parts(5) = endpointName Then
                LookupChannelFromCSV = parts(4)
                txtFile.Close
                Exit Function
            End If
        End If
    Loop

    txtFile.Close
    LookupChannelFromCSV = "N/A"
    Exit Function

FileOpenError:
    LookupChannelFromCSV = "N/A"
    MsgBox "Error opening file: " & Err.Description, vbExclamation, "File Error"
End Function
