' Description: VBA for Excel

' Configuration constants
Const TEST_USERNAME As String = "[LAN ID]" ' Change this to match the test username you want to filter out
Const ONEDRIVE_ORG As String = "[COMPANY]"  ' Change this to match your organization's OneDrive folder name

' Iterate through all rows and delete rows where column F matches the test username
Sub DeleteTesterRows()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("form_data")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    Dim i As Long
    For i = lastRow To 1 Step -1
        If ws.Cells(i, "F").Value = TEST_USERNAME Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

' Delete all rows that have any value of "No camera"
Sub DeleteNoCameraRows()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("form_data")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    For i = lastRow To 1 Step -1
        If ws.Cells(i, "A").Value = "No camera" Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

' lookup the IP address for a given NVR in a separate CSV file
Sub LookupIP()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("form_data")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRow
        Dim nvr As String
        nvr = ws.Cells(i, "A").Value
        Dim ip As String
        ip = LookupIPFromCSV(nvr)
        ws.Cells(i, "B").Value = ip
    Next i
End Sub

' Helper function to lookup the IP address for a given NVR in a separate CSV file
Function LookupIPFromCSV(nvr As String) As String
    Dim filePath As String
    filePath = "C:\Users\" & Environ("USERNAME") & "\OneDrive - " & ONEDRIVE_ORG & "\Documents\Workflows\Register-Camera mapping\LP_Tech_NVR_nationwide_5_22_2024_.csv"
    Dim fileSysObj As Object
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    Dim txtFile As Object
    On Error Resume Next
    Set txtFile = fileSysObj.OpenTextFile(filePath, 1, False)
    If Err.Number <> 0 Then
        MsgBox "Error opening file: " & Err.Description, vbExclamation, "File Error"
        Exit Function
    End If
    On Error GoTo 0
    Dim line As String
    Do While Not txtFile.AtEndOfStream
        line = txtFile.ReadLine
        Dim parts() As String
        parts = Split(line, ",")
        If UBound(parts) >= 1 Then
            If parts(0) = nvr Then
                LookupIPFromCSV = parts(1)
                Exit Function
            End If
        End If
    Loop
    LookupIPFromCSV = "N/A"
End Function

' lookup the channel number for a given NVR and camera in a separate CSV file
Sub LookupChannel()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("form_data")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRow
        Dim nvr As String
        nvr = ws.Cells(i, "A").Value
        Dim camera As String
        camera = ws.Cells(i, "E").Value
        Dim channel As String
        channel = LookupChannelFromCSV(nvr, camera)
        ws.Cells(i, "D").Value = channel
    Next i
End Sub

' Helper function to lookup the channel number for a given NVR and camera in a separate CSV file
Function LookupChannelFromCSV(nvr As String, camera As String) As String
    Dim filePath As String
    filePath = "C:\Users\" & Environ("USERNAME") & "\OneDrive - " & ONEDRIVE_ORG & "\Documents\Workflows\Register-Camera mapping\NW Cameras-cleaned.csv"
    Dim fileSysObj As Object
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    Dim txtFile As Object
    On Error Resume Next
    Set txtFile = fileSysObj.OpenTextFile(filePath, 1, False)
    If Err.Number <> 0 Then
        MsgBox "Error opening file: " & Err.Description, vbExclamation, "File Error"
        Exit Function
    End If
    On Error GoTo 0
    Dim line As String
    Do While Not txtFile.AtEndOfStream
        line = txtFile.ReadLine
        Dim parts() As String
        parts = Split(line, ",")
        If UBound(parts) >= 5 Then ' Ensure there are at least 6 parts
            If parts(2) = nvr And parts(5) = camera Then
                LookupChannelFromCSV = parts(4)
                Exit Function
            End If
        End If
    Loop
    LookupChannelFromCSV = "N/A"
End Function


