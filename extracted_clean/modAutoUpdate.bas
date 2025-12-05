Attribute VB_Name = "modAutoUpdate"
Option Explicit

' GAFC Audit Helper - Auto Update Module
' Checks for updates from GitHub releases and auto-updates silently

Private Const MANIFEST_URL As String = "https://raw.githubusercontent.com/muaroi2002/gafc-audit-helper-releases/main/releases/audit_tool.json"
Private Const UPDATE_CHECK_INTERVAL_DAYS As Double = 1  ' Check daily

' Type must be declared at module level, before any procedures
Private Type UpdateState
    lastCheckDate As Date
    skipVersion As String
End Type

' Get version from XLAM custom properties (set during build)
Private Function CURRENT_VERSION() As String
    On Error Resume Next
    Dim wb As Workbook
    Set wb = ThisWorkbook
    CURRENT_VERSION = wb.CustomDocumentProperties("Version").Value

    ' Fallback if property not set
    If Err.Number <> 0 Or Len(CURRENT_VERSION) = 0 Then
        CURRENT_VERSION = "1.0.6"  ' Default fallback
    End If
    On Error GoTo 0
End Function

' Main entry point - call from Workbook_Open
Public Sub CheckForUpdates(Optional ByVal forceCheck As Boolean = False)
    On Error Resume Next
    Dim state As UpdateState, latestVersion As String, downloadUrl As String, releaseNotes As String

    ' Skip if checked recently (unless forced)
    If Not forceCheck Then
        state = LoadUpdateState()
        If state.lastCheckDate > 0 And (Now - state.lastCheckDate) < UPDATE_CHECK_INTERVAL_DAYS Then
            Exit Sub
        End If
    End If

    If GetLatestVersion(latestVersion, downloadUrl, releaseNotes) Then
        ' Save check time
        state.lastCheckDate = Now
        SaveUpdateState state

        ' Compare versions
        If CompareVersions(latestVersion, CURRENT_VERSION) > 0 Then
            ' New version available - auto update without prompting
            AutoUpdate latestVersion, releaseNotes, downloadUrl
        End If
    End If
End Sub

' Fetch latest version from manifest
Private Function GetLatestVersion(ByRef latestVersion As String, _
                                   ByRef downloadUrl As String, _
                                   ByRef releaseNotes As String) As Boolean
    On Error Resume Next

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    If http Is Nothing Then Exit Function

    http.Open "GET", MANIFEST_URL, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send

    If http.status <> 200 Then Exit Function

    Dim jsonText As String
    jsonText = http.responseText

    ' Parse JSON manually (simple parsing)
    latestVersion = ExtractJsonValue(jsonText, "latest")
    downloadUrl = ExtractJsonValue(jsonText, "download_url")
    releaseNotes = ExtractJsonValue(jsonText, "release_notes")

    GetLatestVersion = (Len(latestVersion) > 0 And Len(downloadUrl) > 0)
End Function

' Simple JSON value extractor
Private Function ExtractJsonValue(ByVal json As String, ByVal key As String) As String
    On Error Resume Next
    Dim startPos As Long, endPos As Long
    Dim searchStr As String

    searchStr = """" & key & """:"
    startPos = InStr(1, json, searchStr, vbTextCompare)
    If startPos = 0 Then Exit Function

    startPos = startPos + Len(searchStr)

    ' Skip whitespace and quotes
    Do While Mid(json, startPos, 1) = " " Or Mid(json, startPos, 1) = vbTab
        startPos = startPos + 1
    Loop

    If Mid(json, startPos, 1) = """" Then
        startPos = startPos + 1
        endPos = InStr(startPos, json, """")
    Else
        ' Number or boolean
        endPos = InStr(startPos, json, ",")
        If endPos = 0 Then endPos = InStr(startPos, json, "}")
    End If

    If endPos > startPos Then
        ExtractJsonValue = Mid(json, startPos, endPos - startPos)
        ExtractJsonValue = Trim(ExtractJsonValue)
    End If
End Function

' Compare version strings (e.g., "1.0.3" vs "1.0.2")
' Returns: 1 if v1 > v2, -1 if v1 < v2, 0 if equal
Private Function CompareVersions(ByVal v1 As String, ByVal v2 As String) As Integer
    On Error Resume Next

    Dim parts1() As String, parts2() As String
    Dim i As Integer, n1 As Long, n2 As Long

    parts1 = Split(v1, ".")
    parts2 = Split(v2, ".")

    Dim maxLen As Integer
    maxLen = IIf(UBound(parts1) > UBound(parts2), UBound(parts1), UBound(parts2))

    For i = 0 To maxLen
        n1 = 0: n2 = 0
        If i <= UBound(parts1) Then n1 = CLng(parts1(i))
        If i <= UBound(parts2) Then n2 = CLng(parts2(i))

        If n1 > n2 Then
            CompareVersions = 1
            Exit Function
        ElseIf n1 < n2 Then
            CompareVersions = -1
            Exit Function
        End If
    Next i

    CompareVersions = 0
End Function

' Auto update without user prompt
Private Sub AutoUpdate(ByVal newVersion As String, _
                       ByVal releaseNotes As String, _
                       ByVal downloadUrl As String)
    ' Show brief notification on status bar (ASCII-safe)
    Application.StatusBar = "Updating to version " & newVersion & "..."

    ' Download and install update silently
    DownloadAndInstall downloadUrl, newVersion
End Sub

' Download and install update
Private Sub DownloadAndInstall(ByVal downloadUrl As String, ByVal newVersion As String)
    On Error GoTo ErrorHandler

    ' Use PowerShell script for update
    Dim xlstartPath As String
    xlstartPath = Application.StartupPath

    Dim updateScript As String
    updateScript = xlstartPath & "\..\..\update_audit_helper.ps1"

    ' Check if script exists in common locations
    If Dir(updateScript) = "" Then
        updateScript = Environ("USERPROFILE") & "\Downloads\gafc_audit_helper_installer\scripts\update_audit_helper.ps1"
    End If

    If Dir(updateScript) <> "" Then
        ' Run PowerShell update script silently in background
        Dim cmd As String
        cmd = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & updateScript & """"

        ' Run update script in background
        shell cmd, vbHide

        ' Close Excel to allow update (after brief delay)
        Application.OnTime Now + TimeValue("00:00:02"), "CloseExcelForUpdate"
    Else
        ' Fallback: open download page
        MsgBox "Cannot find update script. Please download manually from GitHub.", vbExclamation

        shell "explorer.exe https://github.com/muaroi2002/gafc-audit-helper-releases/releases/latest", vbNormalFocus
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Update error: " & Err.Description & vbCrLf & vbCrLf & _
           "Please download manually from GitHub.", vbExclamation
End Sub

' Load update state from Registry
Private Function LoadUpdateState() As UpdateState
    On Error Resume Next
    Dim state As UpdateState

    Dim lastCheck As String
    lastCheck = GetSetting("GAFCAuditHelper", "AutoUpdate", "LastCheck", "")
    If Len(lastCheck) > 0 Then state.lastCheckDate = CDate(lastCheck)

    state.skipVersion = GetSetting("GAFCAuditHelper", "AutoUpdate", "SkipVersion", "")

    LoadUpdateState = state
End Function

' Save update state to Registry
Private Sub SaveUpdateState(ByRef state As UpdateState)
    On Error Resume Next

    If state.lastCheckDate > 0 Then
        SaveSetting "GAFCAuditHelper", "AutoUpdate", "LastCheck", CStr(state.lastCheckDate)
    End If

    SaveSetting "GAFCAuditHelper", "AutoUpdate", "SkipVersion", state.skipVersion
End Sub

' Get current version
Public Function GetCurrentVersion() As String
    GetCurrentVersion = CURRENT_VERSION
End Function

' Close Excel for update (called by OnTime)
Public Sub CloseExcelForUpdate()
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.Quit
End Sub


