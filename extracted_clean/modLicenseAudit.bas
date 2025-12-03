Attribute VB_Name = "modLicenseAudit"
Option Explicit

' GAFC Audit Helper - License client (app_id = audit_tool)
' Flow:
'   - Activate once with license key (online)
'   - Auto-validate weekly (7 days) if license has > 7 days remaining
'   - If license expires within 7 days: check daily + show warning
'   - Offline grace: 7 days from last successful validation
'   - Auto-retry silently when network returns
' NOTE: Set DEV_ALLOW_BYPASS = True to avoid bị khóa cứng khi debug (bỏ chặn nếu check tampering lỗi)

Private Const SERVER_URL As String = "https://license-gafc-server.vercel.app"
Private Const APP_ID As String = "audit-tool"
Private Const CHECK_INTERVAL_DAYS As Double = 7   ' check every 7 days normally
Private Const URGENT_CHECK_DAYS As Double = 1     ' check daily if expiring soon
Private Const EXPIRY_WARNING_DAYS As Double = 7   ' warn if expires within 7 days
Private Const GRACE_DAYS As Double = 7            ' offline grace after last successful check
Private Const DEV_ALLOW_BYPASS As Boolean = True  ' set False khi build release

Private Type LicenseState
    licenseKey As String
    lastValidAt As Date
    expiresAt As Date
    graceUntil As Date
    lastReason As String
    bootCount As Long
    checksum As String
End Type

' ===== Public entry points =====

Public Function ActivateLicenseAudit(ByVal licenseKey As String, Optional ByRef errorMessage As String) As Boolean
    Dim hw As Object
    Set hw = CollectHardware()
    If hw Is Nothing Then
        If DEV_ALLOW_BYPASS Then
            Set hw = BuildFallbackHardware()
        Else
            errorMessage = "Khong the thu thap thong tin thiet bi."
            Exit Function
        End If
    End If

    Dim body As String
    body = BuildPayload(licenseKey, hw, True)

    Dim status As Long, resp As String
    resp = HttpPostJson(SERVER_URL & "/api/v2/licenses/activate", body, status)

    ' Handle network errors - allow grace immediately on first activation
    If status = -1 Then
        ' Network error on first activation - set grace period to allow offline work
        Dim st As LicenseState
        st.licenseKey = licenseKey
        st.lastValidAt = 0  ' Never validated online yet
        st.graceUntil = Now + GRACE_DAYS  ' Give grace period
        st.lastReason = "NETWORK_ERROR_PENDING_ACTIVATION"
        SaveState st
        errorMessage = "Khong the ket noi may chu. Ban co " & GRACE_DAYS & " ngay de hoan tat kich hoat khi co mang."
        ActivateLicenseAudit = True  ' Allow to proceed with grace
        Exit Function
    End If

    If status <> 200 Then
        errorMessage = BuildErrorMessage(status, resp)
        SaveErrorReason JsonString(resp, "reason")
        Exit Function
    End If

    If Not JsonBool(resp, "ok") Then
        Dim reasonCode As String
        reasonCode = JsonString(resp, "reason")
        errorMessage = FriendlyErrorOrFallback(reasonCode, "Kich hoat khong thanh cong. Kiem tra key/mang va thu lai.")
        SaveErrorReason reasonCode
        Exit Function
    End If

    Dim expStr As String
    expStr = JsonString(resp, "expires_at")

    st.licenseKey = licenseKey
    st.lastValidAt = Now
    st.graceUntil = Now + GRACE_DAYS
    st.bootCount = GetBootCount()
    st.lastReason = ""
    If Len(expStr) > 0 Then st.expiresAt = SafeParseDate(expStr)
    SaveState st

    ActivateLicenseAudit = True
End Function

' Validate silently; returns True if allowed to run (cache/grace aware).
Public Function ValidateLicenseAudit(Optional ByVal forceOnline As Boolean = False) As Boolean
    If Not ValidateCodeIntegrity() Then
        If Not DEV_ALLOW_BYPASS Then Exit Function
    End If

    Dim st As LicenseState
    st = LoadState()
    If Len(st.licenseKey) = 0 Then Exit Function

    ' Detect clock tampering
    If DetectClockTampering(st) Then
        forceOnline = True
    End If

    ' Determine check interval based on expiry date
    Dim checkInterval As Double
    Dim daysUntilExpiry As Double
    If st.expiresAt > 0 Then
        daysUntilExpiry = st.expiresAt - Now
        If daysUntilExpiry <= EXPIRY_WARNING_DAYS Then
            checkInterval = URGENT_CHECK_DAYS
        Else
            checkInterval = CHECK_INTERVAL_DAYS
        End If
    Else
        checkInterval = CHECK_INTERVAL_DAYS
    End If

    ' Recent cache: no need to call server
    If Not forceOnline Then
        If st.lastValidAt > 0 Then
            If (Now - st.lastValidAt) < checkInterval Then
                ' Check expiration date in cache mode to prevent expired licenses from working
                If st.expiresAt > 0 And Now > st.expiresAt Then
                    ValidateLicenseAudit = False
                    st.lastReason = "EXPIRED"
                    SaveState st
                    Exit Function
                End If

                ValidateLicenseAudit = True
                ShowExpiryWarningIfNeeded st
                Exit Function
            End If
        End If
    End If

    ' Try online check
    Dim hw As Object
    Set hw = CollectHardware()
    If hw Is Nothing Then
        ValidateLicenseAudit = AllowByGrace(st)
        Exit Function
    End If

    Dim body As String
    body = BuildPayload(st.licenseKey, hw, False)

    Dim status As Long, resp As String
    resp = HttpPostJson(SERVER_URL & "/api/v2/licenses/check", body, status)

    If status = 200 And JsonBool(resp, "valid") Then
        st.lastValidAt = Now
        st.graceUntil = Now + GRACE_DAYS
        st.lastReason = ""
        st.bootCount = GetBootCount()
        Dim expStr As String
        expStr = JsonString(resp, "expires_at")
        If Len(expStr) > 0 Then st.expiresAt = SafeParseDate(expStr)
        SaveState st
        ShowExpiryWarningIfNeeded st
        ValidateLicenseAudit = True
        Exit Function
    End If

    ' Network or server error: allow if still in grace
    If status = -1 Then
        ValidateLicenseAudit = AllowByGrace(st)
        Exit Function
    End If

    ' Invalid for other reasons: store reason and block
    st.lastReason = JsonString(resp, "reason")
    SaveState st
    ValidateLicenseAudit = False
End Function

' Ensure co key (hoi nguoi dung mot lan neu chua kich hoat) roi validate
Public Function EnsureLicenseAndValidate() As Boolean
    Dim st As LicenseState
    st = LoadState()

    ' Neu chua co key, hoi ngay
    If Len(Trim$(st.licenseKey)) = 0 Then
        ' Loop until activation succeeds or user cancels
        Do
            Dim key As String
            key = PromptForKeyUI()
            If Len(Trim$(key)) = 0 Then Exit Function

            Dim errMsg As String
            If ActivateLicenseAudit(Trim$(key), errMsg) Then
                ' Activation succeeded or allowed with grace, exit loop
                Exit Do
            Else
                ' Activation failed, show error and loop again
                Dim result As VbMsgBoxResult
                If Len(errMsg) = 0 Then errMsg = "Kich hoat khong thanh cong. Kiem tra key/mang va thu lai."
                result = MsgBox(errMsg & vbCrLf & vbCrLf & "Nhan OK de nhap lai, Cancel de thoat.", vbCritical + vbOKCancel)
                If result = vbCancel Then
                    Exit Function
                End If
            End If
        Loop
    End If

    ' Da co key => validate (cache/grace da xu ly ben trong)
    EnsureLicenseAndValidate = ValidateLicenseAudit()
End Function

' Tra ve ly do cuoi cung (neu co) tu state, dung de hien thi thong bao cho user
Public Function GetLastReason() As String
    Dim st As LicenseState
    st = LoadState()
    GetLastReason = st.lastReason
End Function

' Dung o dau moi macro/tab: dam bao co key va con hieu luc; tu hoi key neu chua co
Public Function GuardLicense() As Boolean
    ' Hoi key neu chua nhap
    Dim st As LicenseState
    st = LoadState()
    If Len(Trim$(st.licenseKey)) = 0 Then
        ' Loop until activation succeeds or user cancels
        Do
            Dim key As String
            key = PromptForKeyUI()
            If Len(Trim$(key)) = 0 Then
                ' User cancelled
                MsgBox "Ban can nhap license key de su dung.", vbExclamation
                Exit Function
            End If

            Dim errMsg As String
            If ActivateLicenseAudit(Trim$(key), errMsg) Then
                ' Activation succeeded or allowed with grace, exit loop
                Exit Do
            Else
                ' Activation failed, show error and loop again
                Dim result As VbMsgBoxResult
                If Len(errMsg) = 0 Then errMsg = "Kich hoat khong thanh cong. Kiem tra key/mang va thu lai."
                result = MsgBox(errMsg & vbCrLf & vbCrLf & "Nhan OK de nhap lai, Cancel de thoat.", vbCritical + vbOKCancel)
                If result = vbCancel Then
                    Exit Function
                End If
            End If
        Loop
    End If

    ' Validate (co cache/grace)
    If ValidateLicenseAudit() Then
        GuardLicense = True
        Exit Function
    End If

    ' Neu khong hop le, bao ly do
    Dim reason As String
    reason = GetLastReason()
    If Len(reason) = 0 Then
        MsgBox "Khong xac thuc duoc license. Vui long kiem tra mang hoac thu lai sau.", vbCritical
    ElseIf reason = "EXPIRED" Then
        MsgBox "License da het han. Vui long gia han de tiep tuc su dung.", vbCritical
    ElseIf reason = "REVOKED" Then
        MsgBox "License da bi thu hoi. Lien he ho tro.", vbCritical
    ElseIf reason Like "LICENSE_*" Then
        MsgBox "License khong dung ung dung hoac da dung o noi khac.", vbCritical
    ElseIf reason = "NETWORK_ERROR_PENDING_ACTIVATION" Then
        MsgBox "Chua kich hoat thanh cong do loi mang. Vui long ket noi mang va thu lai.", vbExclamation
    Else
        MsgBox "License khong hop le. Ly do: " & reason, vbCritical
    End If
End Function

' Hien thi UI nhap key (UserForm neu co, fallback InputBox)
Private Function PromptForKeyUI() As String
    On Error GoTo Fallback
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmLicensePrompt")
    frm.Show vbModal
    PromptForKeyUI = frm.EnteredKey
    Unload frm
    Exit Function
Fallback:
    PromptForKeyUI = InputBox("Nhap license key de su dung GAFC Audit Helper:", "Kich hoat license")
End Function

' ===== Helpers =====

Private Function AllowByGrace(ByRef st As LicenseState) As Boolean
    If st.graceUntil > 0 And Now <= st.graceUntil Then
        AllowByGrace = True
    End If
End Function

Private Sub ShowExpiryWarningIfNeeded(ByRef st As LicenseState)
    If st.expiresAt <= 0 Then Exit Sub
    Dim daysLeft As Long
    daysLeft = Int(st.expiresAt - Now)
    If daysLeft >= 0 And daysLeft <= EXPIRY_WARNING_DAYS Then
        Dim msg As String
        msg = "Canh bao: License se het han trong " & daysLeft & " ngay." & vbCrLf & _
              "Vui long gia han de tiep tuc su dung GAFC Audit Helper."
        MsgBox msg, vbExclamation, "License sap het han"
    End If
End Sub

Private Function DetectClockTampering(ByRef st As LicenseState) As Boolean
    ' Check 1: lastValidAt in future (clock rolled back)
    If st.lastValidAt > Now + 0.1 Then
        DetectClockTampering = True
        Exit Function
    End If

    ' Check 2: Boot count decreased (system reinstall or tampering)
    Dim currentBoot As Long
    currentBoot = GetBootCount()
    If st.bootCount > 0 And currentBoot < st.bootCount - 5 Then
        DetectClockTampering = True
        Exit Function
    End If

    ' Check 3: State file modification time in future
    If DetectFileTimeTampering() Then
        DetectClockTampering = True
        Exit Function
    End If

    DetectClockTampering = False
End Function

Private Function DetectFileTimeTampering() As Boolean
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim stateFile As String
    stateFile = StatePath()
    If fso.FileExists(stateFile) Then
        Dim f As Object
        Set f = fso.GetFile(stateFile)
        If f.DateLastModified > Now + 0.1 Then
            DetectFileTimeTampering = True
            Exit Function
        End If
    End If
    DetectFileTimeTampering = False
End Function

Private Function GetBootCount() As Long
    On Error Resume Next
    Dim wmi As Object, col As Object, item As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("Select LastBootUpTime from Win32_OperatingSystem")
    For Each item In col
        Dim bootTime As String
        bootTime = item.LastBootUpTime
        GetBootCount = CLng(DateDiff("h", CDate(Mid(bootTime, 1, 14)), Now))
        Exit For
    Next
    If GetBootCount = 0 Then GetBootCount = 1
End Function

Private Function CollectHardware() As Object
    On Error GoTo TryFallback
    Dim svc As Object
    Set svc = GetObject("winmgmts:\\.\root\cimv2")
    If svc Is Nothing Then GoTo TryFallback

    Dim cpu As String, mb As String, bios As String, disk As String, mac As String, uuid As String
    cpu = FirstProp(svc.ExecQuery("Select ProcessorId from Win32_Processor"), "ProcessorId")
    ' If ProcessorId is empty, try to get CPU Name instead
    If Len(Trim(cpu)) = 0 Then
        cpu = FirstProp(svc.ExecQuery("Select Name from Win32_Processor"), "Name")
    End If

    mb = FirstProp(svc.ExecQuery("Select SerialNumber from Win32_BaseBoard"), "SerialNumber")
    bios = FirstProp(svc.ExecQuery("Select SerialNumber from Win32_BIOS"), "SerialNumber")
    disk = FirstProp(svc.ExecQuery("Select SerialNumber from Win32_PhysicalMedia"), "SerialNumber")
    mac = FirstPropMac(svc)
    uuid = FirstProp(svc.ExecQuery("Select UUID from Win32_ComputerSystemProduct"), "UUID")

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("cpuId") = CleanHW(cpu)
    d("motherboardSerial") = CleanHW(mb)
    d("biosSerial") = CleanHW(bios)
    d("diskSerial") = CleanHW(disk)
    d("macAddress") = CleanHW(mac)
    d("systemUuid") = CleanHW(uuid)

    ' Verify we have at least 3/6 fields
    Dim validCount As Integer
    validCount = 0
    If Len(d("cpuId")) > 0 Then validCount = validCount + 1
    If Len(d("motherboardSerial")) > 0 Then validCount = validCount + 1
    If Len(d("biosSerial")) > 0 Then validCount = validCount + 1
    If Len(d("diskSerial")) > 0 Then validCount = validCount + 1
    If Len(d("macAddress")) > 0 Then validCount = validCount + 1
    If Len(d("systemUuid")) > 0 Then validCount = validCount + 1

    If validCount < 3 Then GoTo TryFallback

    Set CollectHardware = d
    Exit Function

TryFallback:
    On Error GoTo Fail
    Set d = CreateObject("Scripting.Dictionary")
    Dim fallbackCpu As String, fallbackDisk As String, fallbackUuid As String
    Dim fallbackBios As String, fallbackMac As String

    fallbackCpu = CleanHW(Environ$("PROCESSOR_IDENTIFIER"))
    fallbackDisk = CleanHW(GetVolumeSerial())
    fallbackUuid = CleanHW(Environ$("COMPUTERNAME") & "_" & Environ$("USERNAME"))
    fallbackBios = CleanHW(Environ$("COMPUTERNAME"))
    fallbackMac = GetMacAddressFallback()

    d("cpuId") = fallbackCpu
    d("motherboardSerial") = ""
    d("biosSerial") = fallbackBios
    d("diskSerial") = fallbackDisk
    d("macAddress") = fallbackMac
    d("systemUuid") = fallbackUuid

    Dim nonEmptyCount As Integer
    nonEmptyCount = 0
    If Len(d("cpuId")) > 0 Then nonEmptyCount = nonEmptyCount + 1
    If Len(d("biosSerial")) > 0 Then nonEmptyCount = nonEmptyCount + 1
    If Len(d("diskSerial")) > 0 Then nonEmptyCount = nonEmptyCount + 1
    If Len(d("macAddress")) > 0 Then nonEmptyCount = nonEmptyCount + 1
    If Len(d("systemUuid")) > 0 Then nonEmptyCount = nonEmptyCount + 1

    If nonEmptyCount < 3 Then GoTo Fail

    Set CollectHardware = d
    Exit Function

Fail:
    If DEV_ALLOW_BYPASS Then
        Set CollectHardware = BuildFallbackHardware()
    Else
        Set CollectHardware = Nothing
    End If
End Function

Private Function FirstProp(queryResult As Object, propName As String) As String
    On Error Resume Next
    Dim item As Object
    For Each item In queryResult
        FirstProp = CStr(item.Properties_(propName))
        Exit For
    Next item
End Function

Private Function FirstPropMac(svc As Object) As String
    On Error Resume Next
    Dim col As Object, item As Object
    Set col = svc.ExecQuery("Select MACAddress, IPEnabled from Win32_NetworkAdapterConfiguration where IPEnabled = TRUE")
    For Each item In col
        FirstPropMac = CStr(item.MACAddress)
        If Len(FirstPropMac) > 0 Then Exit For
    Next item
End Function

Private Function CleanHW(ByVal s As String) As String
    s = Trim$(s)
    s = Replace$(s, "-", "")
    s = Replace$(s, ":", "")
    CleanHW = UCase$(s)
End Function

Private Function GetVolumeSerial() As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim drv As Object
    Set drv = fso.GetDrive("C:")
    If Not drv Is Nothing Then
        GetVolumeSerial = CStr(drv.SerialNumber)
    End If
End Function

Private Function GetMacAddressFallback() As String
    On Error Resume Next
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    Dim exec As Object
    Set exec = shell.exec("getmac /fo csv /nh")
    Dim output As String
    output = exec.StdOut.ReadAll
    If Len(output) > 0 Then
        Dim lines() As String
        lines = Split(output, vbCrLf)
        If UBound(lines) >= 0 Then
            Dim parts() As String
            parts = Split(lines(0), ",")
            If UBound(parts) >= 0 Then
                GetMacAddressFallback = Replace(parts(0), """", "")
            End If
        End If
    End If
End Function

Private Function EscapeJson(ByVal s As String) As String
    s = Replace$(s, "\", "\\")
    s = Replace$(s, """", "\""")
    EscapeJson = s
End Function

Private Function BuildPayload(ByVal licenseKey As String, hw As Object, ByVal isActivate As Boolean) As String
    Dim comp As String
    comp = """cpuId"":""" & EscapeJson(hw("cpuId")) & """," & _
           """motherboardSerial"":""" & EscapeJson(hw("motherboardSerial")) & """," & _
           """biosSerial"":""" & EscapeJson(hw("biosSerial")) & """," & _
           """diskSerial"":""" & EscapeJson(hw("diskSerial")) & """," & _
           """macAddress"":""" & EscapeJson(hw("macAddress")) & """," & _
           """systemUuid"":""" & EscapeJson(hw("systemUuid")) & """"

    Dim actionKey As String
    If isActivate Then
        actionKey = "license_key"
    Else
        actionKey = "license_key"
    End If

    BuildPayload = "{" & _
        """" & actionKey & """:""" & EscapeJson(licenseKey) & """," & _
        """components"":{" & comp & "}," & _
        """app_id"":""" & APP_ID & """" & _
        "}"
End Function

Private Function HttpPostJson(ByVal url As String, ByVal body As String, ByRef statusCode As Long) As String
    On Error GoTo Fail
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
    statusCode = http.status
    HttpPostJson = CStr(http.responseText)
    Exit Function
Fail:
    statusCode = -1
    HttpPostJson = ""
End Function

Private Function JsonBool(ByVal json As String, ByVal key As String) As Boolean
    Dim pat As String
    pat = """" & key & """:true"
    JsonBool = InStr(1, LCase$(json), LCase$(pat), vbTextCompare) > 0
End Function

Private Function JsonString(ByVal json As String, ByVal key As String) As String
    Dim pat As String, pos As Long, startPos As Long, endPos As Long
    pat = """" & key & """:"""
    pos = InStr(1, json, pat, vbTextCompare)
    If pos = 0 Then Exit Function
    startPos = pos + Len(pat)
    endPos = InStr(startPos, json, """")
    If endPos > startPos Then
        JsonString = Mid$(json, startPos, endPos - startPos)
    End If
End Function

Private Function SafeParseDate(ByVal s As String) As Date
    On Error Resume Next
    SafeParseDate = CDate(s)
End Function

Private Function StatePath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folderPath As String
    folderPath = Environ$("APPDATA") & "\GAFC"
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    StatePath = folderPath & "\audit_tool_license.txt"
End Function

Private Sub SaveState(st As LicenseState)
    On Error Resume Next
    Dim plainData As String
    plainData = "license_key=" & st.licenseKey & vbCrLf & _
                "last_valid_at=" & Format$(st.lastValidAt, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
                "expires_at=" & Format$(st.expiresAt, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
                "grace_until=" & Format$(st.graceUntil, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
                "boot_count=" & st.bootCount & vbCrLf & _
                "last_reason=" & st.lastReason

    Dim checksum As String
    checksum = ComputeChecksum(plainData)
    plainData = plainData & vbCrLf & "checksum=" & checksum

    Dim encrypted As String
    encrypted = EncryptData(plainData)

    Dim f As Integer
    f = FreeFile
    Open StatePath For Output As #f
    ' Ghi dạng mã hóa có prefix để phân biệt với legacy plain-text
    Print #f, "ENC1:" & encrypted
    Close #f
End Sub

Private Function LoadState() As LicenseState
    Dim st As LicenseState
    On Error GoTo Done
    Dim f As Integer
    f = FreeFile
    Open StatePath For Input As #f

    Dim encrypted As String
    encrypted = ""
    Do Until EOF(f)
        Dim line As String
        Line Input #f, line
        encrypted = encrypted & line
    Loop
    Close #f

    Dim plainData As String
    ' Hỗ trợ cả legacy plain-text (không prefix) và bản mã hóa (prefix ENC1)
    If Left$(encrypted, 5) = "ENC1:" Then
        plainData = DecryptData(Mid$(encrypted, 6))
    Else
        plainData = encrypted
    End If

    If Len(plainData) = 0 Then GoTo Done

    Dim savedChecksum As String
    Dim lines() As String
    lines = Split(plainData, vbCrLf)

    Dim i As Long, k As String, v As String, p As Long
    For i = LBound(lines) To UBound(lines)
        p = InStr(1, lines(i), "=")
        If p > 0 Then
            k = Left$(lines(i), p - 1)
            v = Mid$(lines(i), p + 1)
            Select Case k
                Case "license_key": st.licenseKey = v
                Case "last_valid_at": st.lastValidAt = SafeParseDate(v)
                Case "expires_at": st.expiresAt = SafeParseDate(v)
                Case "grace_until": st.graceUntil = SafeParseDate(v)
                Case "boot_count": st.bootCount = CLng(v)
                Case "last_reason": st.lastReason = v
                Case "checksum": savedChecksum = v
            End Select
        End If
    Next i

    Dim dataWithoutChecksum As String
    dataWithoutChecksum = Replace(plainData, vbCrLf & "checksum=" & savedChecksum, "")
    If ComputeChecksum(dataWithoutChecksum) <> savedChecksum Then
        ' Checksum invalid - clear the state
        Dim emptySt As LicenseState
        st = emptySt
    End If

Done:
    LoadState = st
End Function

Private Sub SaveErrorReason(ByVal reason As String)
    Dim st As LicenseState
    st = LoadState()
    st.lastReason = reason
    SaveState st
End Sub

Private Function FriendlyErrorOrFallback(ByVal reasonCode As String, ByVal fallback As String) As String
    If Len(Trim$(reasonCode)) > 0 Then
        FriendlyErrorOrFallback = GetFriendlyErrorMessage(reasonCode)
        If Len(Trim$(FriendlyErrorOrFallback)) = 0 Then FriendlyErrorOrFallback = reasonCode
    Else
        FriendlyErrorOrFallback = fallback
    End If
End Function

Private Function BuildErrorMessage(ByVal status As Long, ByVal resp As String) As String
    Dim reasonCode As String
    reasonCode = JsonString(resp, "reason")
    If Len(Trim$(reasonCode)) > 0 Then
        BuildErrorMessage = FriendlyErrorOrFallback(reasonCode, "")
        Exit Function
    End If
    If status = -1 Then
        BuildErrorMessage = "Khong the ket noi may chu."
    Else
        BuildErrorMessage = "Kich hoat khong thanh cong (HTTP " & status & ")."
    End If
End Function

Public Function GetFriendlyErrorMessage(reason As String) As String
    Select Case reason
        Case "NOT_FOUND": GetFriendlyErrorMessage = "License key khong ton tai"
        Case "EXPIRED": GetFriendlyErrorMessage = "License da het han"
        Case "REVOKED": GetFriendlyErrorMessage = "License da bi thu hoi"
        Case "ALREADY_ACTIVATED_ON_DIFFERENT_HARDWARE": GetFriendlyErrorMessage = "License da kich hoat tren may khac"
        Case "LICENSE_ALREADY_USED": GetFriendlyErrorMessage = "License da duoc su dung"
        Case Else: GetFriendlyErrorMessage = reason
    End Select
End Function

Private Function EncryptData(ByVal data As String) As String
    Dim key As String
    key = GetEncryptionKey()
    Dim result As String
    Dim i As Long
    result = ""
    For i = 1 To Len(data)
        Dim charCode As Integer
        charCode = Asc(Mid(data, i, 1))
        Dim keyChar As Integer
        keyChar = Asc(Mid(key, ((i - 1) Mod Len(key)) + 1, 1))
        result = result & Chr((charCode Xor keyChar) Mod 256)
    Next i
    EncryptData = Base64Encode(result)
End Function

Private Function DecryptData(ByVal encrypted As String) As String
    On Error Resume Next
    Dim data As String
    data = Base64Decode(encrypted)
    If Len(data) = 0 Then Exit Function

    Dim key As String
    key = GetEncryptionKey()
    Dim result As String
    Dim i As Long
    result = ""
    For i = 1 To Len(data)
        Dim charCode As Integer
        charCode = Asc(Mid(data, i, 1))
        Dim keyChar As Integer
        keyChar = Asc(Mid(key, ((i - 1) Mod Len(key)) + 1, 1))
        result = result & Chr((charCode Xor keyChar) Mod 256)
    Next i
    DecryptData = result
End Function

Private Function GetEncryptionKey() As String
    Dim hw As Object
    Set hw = CollectHardware()
    If hw Is Nothing Then
        GetEncryptionKey = Environ$("COMPUTERNAME") & Environ$("USERNAME") & "GAFC2025SALT" & APP_ID
    Else
        GetEncryptionKey = hw("cpuId") & hw("diskSerial") & "GAFC2025SALT" & APP_ID
    End If
End Function

Private Function ComputeChecksum(ByVal data As String) As String
    Dim hash As Long
    hash = 5381
    Dim i As Long
    Dim secret As String
    secret = "GAFC_SECRET_2025"
    data = data & secret
    For i = 1 To Len(data)
        hash = ((hash * 33) Xor Asc(Mid(data, i, 1))) And &H7FFFFFFF
    Next i
    ComputeChecksum = Hex$(hash)
End Function

Private Function ValidateCodeIntegrity() As Boolean
    On Error Resume Next
    ValidateCodeIntegrity = True

    Dim expectedConst As String
    expectedConst = SERVER_URL & APP_ID
    If Len(expectedConst) < 40 Then
        ValidateCodeIntegrity = False
        GoTo MaybeBypass
    End If

    If InStr(1, SERVER_URL, "license-gafc-server") = 0 Then
        ValidateCodeIntegrity = False
        GoTo MaybeBypass
    End If

    If APP_ID <> "audit-tool" Then
        ValidateCodeIntegrity = False
        GoTo MaybeBypass
    End If

    Exit Function

MaybeBypass:
    If DEV_ALLOW_BYPASS Then
        ValidateCodeIntegrity = True
    End If
End Function

Private Function Base64Encode(ByVal text As String) As String
    On Error Resume Next
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Dim node As Object
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = StringToByteArray(text)
    Base64Encode = node.text
End Function

Private Function Base64Decode(ByVal base64 As String) As String
    On Error Resume Next
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Dim node As Object
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.text = base64
    Base64Decode = ByteArrayToString(node.nodeTypedValue)
End Function

Private Function StringToByteArray(ByVal text As String) As Variant
    Dim bytes() As Byte
    ReDim bytes(0 To Len(text) - 1)
    Dim i As Long
    For i = 1 To Len(text)
        bytes(i - 1) = Asc(Mid(text, i, 1))
    Next i
    StringToByteArray = bytes
End Function

Private Function ByteArrayToString(bytes As Variant) As String
    Dim result As String
    Dim i As Long
    result = ""
    For i = LBound(bytes) To UBound(bytes)
        result = result & Chr(bytes(i))
    Next i
    ByteArrayToString = result
End Function

Private Function BuildFallbackHardware() As Object
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("cpuId") = CleanHW(Environ$("PROCESSOR_IDENTIFIER"))
    d("motherboardSerial") = ""
    d("biosSerial") = CleanHW(Environ$("COMPUTERNAME"))
    d("diskSerial") = CleanHW(GetVolumeSerial())
    d("macAddress") = CleanHW(GetMacAddressFallback())
    d("systemUuid") = CleanHW(Environ$("COMPUTERNAME") & "_" & Environ$("USERNAME"))
    Set BuildFallbackHardware = d
End Function
