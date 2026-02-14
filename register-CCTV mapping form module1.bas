' Description: Code module 1 for the Register-CCTV Mapping Form

' Configuration constants
Public Const FORM_PASSWORD As String = "[PASSWORD]" ' Update this to the desired password for protecting the form. Make sure to keep it secure and not share it with unauthorized users.
Public Const BASE_DATA_PATH As String = "[NETWORK PATH TO DATA FOLDER]" ' Update this to the actual network path where the data folder is located. Make sure to include the trailing backslash.

' Public variables
Public wsForm As Worksheet
Public wsCameras As Worksheet
Public wsPOSregisters As Worksheet
Public wsRegions As Worksheet
Public storeNum As Range
Public camRegions As Range
Public regStore As Range
Public regNum As Range
Public NVRCell As Range
Public POSList As Collection
Public NVRList As Collection
Public storePOSCamRows As Collection
Public relatedCams() As Variant
Public theFormIsLoading As Boolean
Public submitButton As Button
Public storeNumStr As String
Public selectedNVR As String
Public selectedCam As String
Public camCell As Range
Public border As Variant
Public versionNum as Double


Public Sub InitializeForm()
    ' The version number should be updated every time the code is updated.
    ' The version number is the date & time of update in the format YYYYMMDD.HHMM (24 hours, not "AM/PM").
    versionNum = 20240617.2359

    ' Initialize variables
    Set wsForm = ThisWorkbook.Sheets("FORM")
    Set wsPOSregisters = ThisWorkbook.Sheets("POSregisters")
    Set wsRegions = ThisWorkbook.Sheets("Regions")
    Set storeNum = wsForm.Range("A8")
    Set camRegions = wsRegions.Range("A:A")
    Set regStore = wsPOSregisters.Range("E:E")
    Set regNum = wsPOSregisters.Range("F:F")
    theFormIsLoading = True
    ' Unprotect the wsForm sheet
    wsForm.Unprotect Password:=FORM_PASSWORD
    'Change all four border walls of cell A8 to red
    For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
        wsForm.Cells(8, 1).Borders(border).Color = RGB(255, 0, 0)
        ' change the thickness of the border to normal
        wsForm.Cells(8, 1).Borders(border).Weight = xlThick
    Next border

    'Initialize font colors and no strikethrough for cells A2 to A5
    wsForm.Cells(2, 1).Font.Color = RGB(255, 0, 0)
    wsForm.Cells(2, 1).Font.Strikethrough = False
    wsForm.Cells(3, 1).Font.Color = RGB(0, 0, 0)
    wsForm.Cells(3, 1).Font.Strikethrough = False
    wsForm.Cells(4, 1).Font.Color = RGB(0, 0, 0)
    wsForm.Cells(4, 1).Font.Strikethrough = False
    wsForm.Cells(5, 1).Font.Color = RGB(0, 0, 0)
    wsForm.Cells(5, 1).Font.Strikethrough = False

    ' Clear the form (clear the contents of cells A8 and C8, plus clear D8 and rows 11 downward)
    wsForm.Range("A8").ClearContents
    wsForm.Range("C8").ClearContents
    wsForm.Range("D8").ClearFormats
    wsForm.Range("D8").Clear
    wsForm.Range("A11:A" & wsForm.Rows.Count).Clear
    wsForm.Range("B11:B" & wsForm.Rows.Count).Clear
    wsForm.Range("C11:C" & wsForm.Rows.Count).Clear
    wsForm.Range("D11:D" & wsForm.Rows.Count).Clear
    wsForm.Range("E:E").Clear
    wsForm.Range("F:F").Clear
    'limit the scroll area to columns A to C
    wsForm.ScrollArea = "A1:C" & wsForm.Rows.Count

    ' Delete the submit button if it exists
    On Error Resume Next
    Set submitButton = wsForm.Buttons("SubmitButton")
    wsForm.Buttons(submitButton.Name).Delete
    On Error GoTo 0

    theFormIsLoading = False

    ' Protect the whole workbook except wsForm cell A8, the buttons, and the dropdown menus
    ThisWorkbook.Protect Password:=FORM_PASSWORD, Structure:=True, Windows:=False
    ' Unprotect the wsForm sheet
    wsForm.Unprotect Password:=FORM_PASSWORD
    ' Lock all wsForm cells except A8 and the buttons
    wsForm.Cells.Locked = True
    wsForm.Cells(8, 1).Locked = False
    ' Macros can still run when the sheet is protected
    wsForm.Protect Password:=FORM_PASSWORD, UserInterfaceOnly:=True

    ' The rest of this sub routine will perform update enforcement by checking
    ' the version number in a remote location.

    ' Declare version variables
    Dim remoteVersionNum As Double
    Dim fileSysObj As Object
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    Dim txtStream As Object
    
    ' Open the env.ini file
    On Error Resume Next
    Set txtStream = fileSysObj.OpenTextFile(BASE_DATA_PATH & "env.ini", 1)
    If Err.Number <> 0 Then
        MsgBox "The version verification failed. Are you on the company network/VPN? If you are on the network, you may have an access problem.", vbExclamation, "Verification Error"
        ' MsgBox "Error opening file: " & Err.Description, vbExclamation, "File Error"
        ThisWorkbook.Close SaveChanges:=False
    End If
    On Error GoTo 0
    
    ' Read the version number from the env.ini file
    remoteVersionNum = CDbl(txtStream.ReadLine)
    
    ' Close the text stream and release the FileSystemObject
    txtStream.Close
    Set txtStream = Nothing
    Set fileSysObj = Nothing
    
    ' Compare the version numbers
    If versionNum < remoteVersionNum Then
        MsgBox "This form is outdated. Please update to the latest version.", vbExclamation, "Outdated Version"
        ThisWorkbook.Close SaveChanges:=False
    End If

End Sub
