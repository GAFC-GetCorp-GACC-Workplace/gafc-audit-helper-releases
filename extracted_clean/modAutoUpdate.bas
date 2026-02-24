Attribute VB_Name = "modAutoUpdate"
Option Explicit

' GAFC Audit Helper - Auto Update Module
' Checks for updates from GitHub releases and auto-updates silently

Private Const MANIFEST_URL As String = "https://raw.githubusercontent.com/muaroi2002/gafc-audit-helper-releases/main/releases/audit_tool.json"
Private Const UPDATE_CHECK_INTERVAL_DAYS As Double = 1  ' Check daily
Private Const REG_APP As String = "GAFCAuditHelper"
Private Const REG_SECTION As String = "AutoUpdate"
Private Const REG_PENDING_PATH As String = "PendingPath"
Private Const REG_PENDING_VERSION As String = "PendingVersion"
Private isForceCheck As Boolean

' Type must be declared at module level, before any procedures
Private Type UpdateState
    lastCheckDate As Date
    skipVersion As String
End Type

' Get version from XLAM custom properties (set during build)
Private Function CURRENT_VERSION() As String
    ' Always return hardcoded version since CustomDocumentProperties may be outdated
    CURRENT_VERSION = "1.1.2"
End Function

' Main entry point - call from Workbook_Open
Public Sub CheckForUpdates(Optional ByVal forceCheck As Boolean = False)
    isForceCheck = forceCheck
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
        Dim cmp As Integer
        cmp = CompareVersions(latestVersion, CURRENT_VERSION)

        If cmp > 0 Then
            ' New version available - auto update without prompting
            AutoUpdate latestVersion, releaseNotes, downloadUrl
        Else
            InfoToast "You are on the latest version."
        End If
    Else
        InfoToast "Could not check for updates."
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
    Application.StatusBar = "Downloading update " & newVersion & "..."

    ' Download and install update silently
    DownloadAndInstall downloadUrl, newVersion
End Sub

' Download and install update
Private Sub DownloadAndInstall(ByVal downloadUrl As String, ByVal newVersion As String)
    On Error GoTo ErrorHandler

    ' Download update script from GitHub to temp folder
    Dim tempFolder As String, scriptPath As String
    tempFolder = Environ("TEMP")
    scriptPath = tempFolder & "\gafc_update_" & newVersion & ".xlam"

    ' Download new XLAM to temp; apply sẽ đợi đến khi user đóng Excel
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    If http Is Nothing Then GoTo ErrorHandler

    http.Open "GET", downloadUrl, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send
    If http.status <> 200 Then GoTo ErrorHandler

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 'binary
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile scriptPath, 2 'overwrite
    stream.Close

    ' Lưu pending để áp dụng khi Excel thật sự tắt
    SetPendingUpdate scriptPath, newVersion
    Application.StatusBar = "Update downloaded (" & newVersion & "). Will apply after you close Excel."
    If isForceCheck Then
        InfoToast "Update ready. Please save work and close Excel to apply."
    End If

    Exit Sub

ErrorHandler:
    ' Im lặng theo yêu cầu, chỉ reset status bar
    Application.StatusBar = False
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

Public Sub ApplyPendingUpdateIfNeeded()
    On Error Resume Next
    Dim pendingPath As String
    pendingPath = GetSetting(REG_APP, REG_SECTION, REG_PENDING_PATH, "")
    If Len(Trim$(pendingPath)) = 0 Then Exit Sub
    If Dir$(pendingPath, vbNormal) = "" Then
        ClearPendingUpdate
        Exit Sub
    End If

    Dim dest As String
    dest = Environ("APPDATA") & "\Microsoft\Excel\XLSTART\gafc_audit_helper.xlam"

    ' Tạo script chờ Excel tắt rồi copy
    Dim tempScript As String
    tempScript = Environ("TEMP") & "\gafc_apply_update.ps1"
    Dim ps As String
    ps = "$src='" & Replace(pendingPath, "'", "''") & "';" & vbCrLf & _
         "$dest='" & Replace(dest, "'", "''") & "';" & vbCrLf & _
         "while(Get-Process excel -ErrorAction SilentlyContinue){Start-Sleep -Milliseconds 500};" & vbCrLf & _
         "try { Copy-Item -Path $src -Destination $dest -Force } catch {};" & vbCrLf & _
         "Remove-Item -Path $src -Force -ErrorAction SilentlyContinue;" & vbCrLf & _
         "Remove-ItemProperty -Path 'HKCU:\\Software\\VB and VBA Program Settings\\" & REG_APP & "\\" & REG_SECTION & "' -Name '" & REG_PENDING_PATH & "' -ErrorAction SilentlyContinue;" & vbCrLf & _
         "Remove-ItemProperty -Path 'HKCU:\\Software\\VB and VBA Program Settings\\" & REG_APP & "\\" & REG_SECTION & "' -Name '" & REG_PENDING_VERSION & "' -ErrorAction SilentlyContinue;"

    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(tempScript, True)
    ts.Write ps
    ts.Close

    Dim cmd As String
    cmd = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & tempScript & """"
    CreateObject("WScript.Shell").Run cmd, 0, False
End Sub

Private Sub SetPendingUpdate(ByVal path As String, ByVal version As String)
    On Error Resume Next
    SaveSetting REG_APP, REG_SECTION, REG_PENDING_PATH, path
    SaveSetting REG_APP, REG_SECTION, REG_PENDING_VERSION, version
End Sub

Private Sub ClearPendingUpdate()
    On Error Resume Next
    SaveSetting REG_APP, REG_SECTION, REG_PENDING_PATH, ""
    SaveSetting REG_APP, REG_SECTION, REG_PENDING_VERSION, ""
End Sub


