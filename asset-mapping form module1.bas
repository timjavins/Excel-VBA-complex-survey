' Description: Code module 1 for the Asset Mapping Form
Option Explicit

' Configuration constants
Public Const FORM_PASSWORD As String = "[PASSWORD]" ' Keep this value private when configured.
Public Const BASE_DATA_PATH As String = "[NETWORK PATH TO DATA FOLDER]" ' Include trailing backslash when configured.
Public Const FORM_SHEET_NAME As String = "FORM"
Public Const TERMINALS_SHEET_NAME As String = "Terminals"
Public Const LEGACY_TERMINALS_SHEET_NAME As String = "POSregisters"
Public Const LOCATION_INDEX_SHEET_NAME As String = "Locations"
Public Const LEGACY_LOCATION_INDEX_SHEET_NAME As String = "Regions"
Public Const DEVICE_SHEET_SUFFIX As String = "Devices"
Public Const LEGACY_DEVICE_SHEET_SUFFIX As String = "Cameras"
Public Const NO_DEVICE_LABEL As String = "No device"

' Public variables
Public wsForm As Worksheet
Public wsDevices As Worksheet
Public wsTerminals As Worksheet
Public wsLocations As Worksheet
Public locationIdCell As Range
Public deviceLocationIndex As Range
Public terminalLocationColumn As Range
Public terminalIdColumn As Range
Public controllerCell As Range
Public terminalList As Collection
Public controllerList As Collection
Public locationDeviceRows As Collection
Public relatedDevices() As Variant
Public theFormIsLoading As Boolean
Public submitButton As Button
Public locationIdText As String
Public selectedController As String
Public selectedDevice As String
Public deviceCell As Range
Public border As Variant
Public versionNum As Double

Public Sub InitializeForm()
    Dim remoteVersionNum As Double
    Dim fileSysObj As Object
    Dim txtStream As Object

    ' Update with YYYYMMDD.HHMM when code changes.
    versionNum = 20261028.1

    Set wsForm = GetSheetWithFallback(FORM_SHEET_NAME, FORM_SHEET_NAME)
    Set wsTerminals = GetSheetWithFallback(TERMINALS_SHEET_NAME, LEGACY_TERMINALS_SHEET_NAME)
    Set wsLocations = GetSheetWithFallback(LOCATION_INDEX_SHEET_NAME, LEGACY_LOCATION_INDEX_SHEET_NAME)
    Set locationIdCell = wsForm.Range("A8")
    Set deviceLocationIndex = wsLocations.Range("A:A")
    Set terminalLocationColumn = wsTerminals.Range("E:E")
    Set terminalIdColumn = wsTerminals.Range("F:F")

    theFormIsLoading = True

    wsForm.Unprotect Password:=FORM_PASSWORD
    For Each border In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
        wsForm.Cells(8, 1).Borders(border).Color = RGB(255, 0, 0)
        wsForm.Cells(8, 1).Borders(border).Weight = xlThick
    Next border

    wsForm.Cells(2, 1).Font.Color = RGB(255, 0, 0)
    wsForm.Cells(2, 1).Font.Strikethrough = False
    wsForm.Cells(3, 1).Font.Color = RGB(0, 0, 0)
    wsForm.Cells(3, 1).Font.Strikethrough = False
    wsForm.Cells(4, 1).Font.Color = RGB(0, 0, 0)
    wsForm.Cells(4, 1).Font.Strikethrough = False
    wsForm.Cells(5, 1).Font.Color = RGB(0, 0, 0)
    wsForm.Cells(5, 1).Font.Strikethrough = False

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
    wsForm.ScrollArea = "A1:C" & wsForm.Rows.Count

    On Error Resume Next
    Set submitButton = wsForm.Buttons("SubmitButton")
    If Not submitButton Is Nothing Then
        wsForm.Buttons(submitButton.Name).Delete
    End If
    Set submitButton = Nothing
    On Error GoTo 0

    theFormIsLoading = False

    ThisWorkbook.Protect Password:=FORM_PASSWORD, Structure:=True, Windows:=False
    wsForm.Unprotect Password:=FORM_PASSWORD
    wsForm.Cells.Locked = True
    wsForm.Cells(8, 1).Locked = False
    wsForm.Protect Password:=FORM_PASSWORD, UserInterfaceOnly:=True

    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    On Error GoTo VersionFileError
    Set txtStream = fileSysObj.OpenTextFile(BASE_DATA_PATH & "env.ini", 1)
    remoteVersionNum = CDbl(txtStream.ReadLine)
    txtStream.Close
    Set txtStream = Nothing
    Set fileSysObj = Nothing
    On Error GoTo 0

    If versionNum < remoteVersionNum Then
        MsgBox "This workbook template is outdated. Please update to the latest version.", vbExclamation, "Outdated Version"
        ThisWorkbook.Close SaveChanges:=False
    End If

    Exit Sub

VersionFileError:
    MsgBox "Version verification failed. Confirm network/VPN access and data-folder permissions.", vbExclamation, "Verification Error"
    If Not txtStream Is Nothing Then
        txtStream.Close
    End If
    Set txtStream = Nothing
    Set fileSysObj = Nothing
    ThisWorkbook.Close SaveChanges:=False
End Sub

Public Function ResolveDeviceSheet(ByVal locationGroupName As String) As Worksheet
    Dim preferredSheetName As String
    Dim legacySheetName As String

    preferredSheetName = locationGroupName & DEVICE_SHEET_SUFFIX
    legacySheetName = locationGroupName & LEGACY_DEVICE_SHEET_SUFFIX
    Set ResolveDeviceSheet = GetSheetWithFallback(preferredSheetName, legacySheetName)
End Function

Private Function GetSheetWithFallback(ByVal preferredName As String, ByVal fallbackName As String) As Worksheet
    On Error Resume Next
    Set GetSheetWithFallback = ThisWorkbook.Sheets(preferredName)
    On Error GoTo 0

    If GetSheetWithFallback Is Nothing Then
        On Error Resume Next
        Set GetSheetWithFallback = ThisWorkbook.Sheets(fallbackName)
        On Error GoTo 0
    End If

    If GetSheetWithFallback Is Nothing Then
        Err.Raise vbObjectError + 513, "GetSheetWithFallback", "Missing required worksheet: " & preferredName
    End If
End Function
