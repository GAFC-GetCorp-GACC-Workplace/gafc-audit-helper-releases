Attribute VB_Name = "TestChecksum"
Option Explicit

' ===============================================
' HƯỚNG DẪN SỬ DỤNG:
' 1. Import file này vào VBA project
' 2. Mở Immediate Window (Ctrl+G)
' 3. Chạy: RunAllTests
' 4. Xem output để tìm lỗi
' ===============================================

' Chạy tất cả tests một lần
Public Sub RunAllTests()
    Debug.Print "========================================="
    Debug.Print "       FULL DIAGNOSTIC SUITE"
    Debug.Print "========================================="
    Debug.Print ""

    TestSimpleFileWrite
    Debug.Print ""
    Debug.Print String(50, "-")
    Debug.Print ""

    TestBase64
    Debug.Print ""
    Debug.Print String(50, "-")
    Debug.Print ""

    TestActivation
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "      DIAGNOSTIC COMPLETE"
    Debug.Print "========================================="
End Sub

' Test 1: Kiểm tra có thể ghi file cơ bản không
Public Sub TestSimpleFileWrite()
    Debug.Print "=== Testing Simple File Write ==="

    On Error GoTo TestError

    ' Test 1: Check APPDATA path
    Dim appData As String
    appData = Environ$("APPDATA")
    Debug.Print "APPDATA: " & appData

    If Len(appData) = 0 Then
        Debug.Print "✗ ERROR: Cannot get APPDATA path!"
        Exit Sub
    End If

    ' Test 2: Create folder
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderPath As String
    folderPath = appData & "\GAFC"
    Debug.Print "Target folder: " & folderPath

    If Not fso.FolderExists(folderPath) Then
        Debug.Print "Creating folder..."
        fso.CreateFolder folderPath
        Debug.Print "✓ Folder created!"
    Else
        Debug.Print "✓ Folder already exists"
    End If

    ' Test 3: Write simple test file
    Dim testFile As String
    testFile = folderPath & "\test.txt"
    Debug.Print "Writing to: " & testFile

    Dim f As Integer
    f = FreeFile
    Open testFile For Output As #f
    Print #f, "Test content from VBA"
    Close #f

    Debug.Print "✓ File written successfully!"

    ' Test 4: Read it back
    Open testFile For Input As #f
    Dim content As String
    Line Input #f, content
    Close #f

    Debug.Print "Read back: " & content
    Debug.Print "✓ File read successfully!"

    Debug.Print ""
    Debug.Print "=== Test Complete ==="
    Debug.Print "File location: " & testFile
    Exit Sub

TestError:
    Debug.Print "✗ ERROR: " & Err.Description & " (Error " & Err.Number & ")"
    On Error Resume Next
    Close #f
End Sub

' Test 2: Kiểm tra Base64 encoding có hoạt động không
Public Sub TestBase64()
    Debug.Print "=== Testing Base64 Encoding ==="

    On Error GoTo TestError

    Dim testStr As String
    testStr = "Hello World 123"

    Debug.Print "Input: " & testStr
    Debug.Print "Input length: " & Len(testStr)

    ' Test MSXML2
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    If xml Is Nothing Then
        Debug.Print "✗ ERROR: Cannot create MSXML2.DOMDocument"
        Debug.Print "This means MSXML is not available on your system!"
        Debug.Print "You need to install MSXML 6.0 SP2"
        Exit Sub
    End If
    Debug.Print "✓ MSXML2.DOMDocument created successfully"

    ' Test encoding - call directly from modLicenseAudit
    Dim encoded As String
    On Error Resume Next
    Err.Clear
    encoded = modLicenseAudit.Base64Encode(testStr)
    If Err.Number <> 0 Then
        Debug.Print "✗ ERROR calling Base64Encode: " & Err.Description
        Exit Sub
    End If
    On Error GoTo TestError

    Debug.Print "Encoded: " & encoded
    Debug.Print "Encoded length: " & Len(encoded)

    If Len(encoded) = 0 Then
        Debug.Print "✗ ERROR: Encoding returned empty string!"
        Exit Sub
    End If

    ' Test decoding
    Dim decoded As String
    On Error Resume Next
    Err.Clear
    decoded = modLicenseAudit.Base64Decode(encoded)
    If Err.Number <> 0 Then
        Debug.Print "✗ ERROR calling Base64Decode: " & Err.Description
        Exit Sub
    End If
    On Error GoTo TestError

    Debug.Print "Decoded: " & decoded
    Debug.Print "Decoded length: " & Len(decoded)

    If decoded = testStr Then
        Debug.Print "✓ Base64 encode/decode works perfectly!"
    Else
        Debug.Print "✗ ERROR: Decode mismatch!"
        Debug.Print "Expected: " & testStr
        Debug.Print "Got: " & decoded
    End If

    Debug.Print ""
    Debug.Print "=== Test Complete ==="
    Exit Sub

TestError:
    Debug.Print "✗ ERROR: " & Err.Description & " (Error " & Err.Number & ")"
End Sub

' Test 3: Test activation thực tế
Public Sub TestActivation()
    Debug.Print "=== Testing License Activation ==="

    On Error GoTo TestError

    ' Test với fake key (sẽ gây network error và tạo file với grace period)
    Dim errMsg As String
    Dim result As Boolean

    Debug.Print "Calling ActivateLicenseAudit with test key..."
    Debug.Print "(This will fail with network error, which is expected)"
    Debug.Print ""

    result = modLicenseAudit.ActivateLicenseAudit("TEST-KEY-12345-67890-ABCDEF", errMsg)

    Debug.Print "Activation result: " & result
    If Len(errMsg) > 0 Then
        Debug.Print "Error message: " & errMsg
    End If
    Debug.Print ""

    ' Check if file was created
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim stateFile As String
    stateFile = Environ$("APPDATA") & "\GAFC\audit_tool_license.txt"
    Debug.Print "Checking for state file..."
    Debug.Print "Expected location: " & stateFile
    Debug.Print ""

    If fso.FileExists(stateFile) Then
        Debug.Print "✓ STATE FILE WAS CREATED!"
        Debug.Print ""

        ' Read and show first line
        Dim f As Integer
        f = FreeFile
        Open stateFile For Input As #f
        Dim line As String
        Line Input #f, line
        Close #f

        Debug.Print "File preview:"
        Debug.Print "  First 60 chars: " & Left$(line, 60) & "..."
        Debug.Print "  Total length: " & Len(line) & " characters"

        ' Check if it starts with ENC1:
        If Left$(line, 5) = "ENC1:" Then
            Debug.Print "  ✓ Has correct ENC1: prefix"
            If Len(line) > 5 Then
                Debug.Print "  ✓ Contains encrypted data (" & (Len(line) - 5) & " chars)"
            Else
                Debug.Print "  ✗ WARNING: No data after ENC1: prefix!"
            End If
        Else
            Debug.Print "  ✗ WARNING: Missing ENC1: prefix!"
        End If
    Else
        Debug.Print "✗ STATE FILE WAS NOT CREATED!"
        Debug.Print ""
        Debug.Print "Possible reasons:"
        Debug.Print "  1. SaveState function failed silently"
        Debug.Print "  2. Base64 encoding failed (check test above)"
        Debug.Print "  3. No write permissions to APPDATA folder"
        Debug.Print "  4. Antivirus blocked file creation"
        Debug.Print ""
        Debug.Print "Check the Immediate Window for ERROR messages above."
    End If

    Debug.Print ""
    Debug.Print "=== Test Complete ==="
    Exit Sub

TestError:
    Debug.Print "✗ ERROR: " & Err.Description & " (Error " & Err.Number & ")"
    On Error Resume Next
    Close #f
End Sub

' Utility: Show current state file content
Public Sub ShowStateFile()
    Debug.Print "=== Current State File Content ==="

    On Error GoTo ShowError

    Dim stateFile As String
    stateFile = Environ$("APPDATA") & "\GAFC\audit_tool_license.txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(stateFile) Then
        Debug.Print "✗ State file does not exist at: " & stateFile
        Exit Sub
    End If

    Debug.Print "File: " & stateFile
    Debug.Print ""

    Dim f As Integer
    f = FreeFile
    Open stateFile For Input As #f

    Dim lineNum As Integer
    lineNum = 1
    Do Until EOF(f)
        Dim line As String
        Line Input #f, line
        Debug.Print "Line " & lineNum & ": " & Left$(line, 100)
        If Len(line) > 100 Then
            Debug.Print "         ..." & (Len(line) - 100) & " more chars"
        End If
        lineNum = lineNum + 1
    Loop

    Close #f
    Debug.Print ""
    Debug.Print "=== End of File ==="
    Exit Sub

ShowError:
    Debug.Print "✗ ERROR: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

' Utility: Delete state file để test lại từ đầu
Public Sub DeleteStateFile()
    On Error Resume Next

    Dim stateFile As String
    stateFile = Environ$("APPDATA") & "\GAFC\audit_tool_license.txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(stateFile) Then
        fso.DeleteFile stateFile
        Debug.Print "✓ State file deleted: " & stateFile
    Else
        Debug.Print "State file does not exist (already clean)"
    End If
End Sub

' Utility: Verify checksum trong file state (SIMPLE VERSION - no LoadState needed)
Public Sub VerifyStateChecksum()
    Debug.Print "=== Verifying State File Checksum ==="

    On Error GoTo VerifyError

    ' Read raw encrypted file directly
    Dim stateFile As String
    stateFile = Environ$("APPDATA") & "\GAFC\audit_tool_license.txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(stateFile) Then
        Debug.Print "✗ State file does not exist at: " & stateFile
        Exit Sub
    End If

    Debug.Print "Reading state file: " & stateFile
    Debug.Print ""

    Dim f As Integer
    f = FreeFile
    Open stateFile For Input As #f
    Dim rawContent As String
    Line Input #f, rawContent
    Close #f

    Debug.Print "Raw file content (first 100 chars):"
    Debug.Print "  " & Left$(rawContent, 100)
    Debug.Print "  Total length: " & Len(rawContent) & " characters"
    Debug.Print ""

    If Left$(rawContent, 5) = "ENC1:" Then
        Debug.Print "✓ Has correct ENC1: prefix"

        ' Decrypt to check checksum
        Dim encrypted As String
        encrypted = Mid$(rawContent, 6)

        Debug.Print "Encrypted data length: " & Len(encrypted) & " characters"
        Debug.Print ""
        Debug.Print "Decrypting..."

        ' First Base64 decode
        Dim base64Decoded As String
        base64Decoded = modLicenseAudit.Base64Decode(encrypted)

        Debug.Print "After Base64 decode: " & Len(base64Decoded) & " bytes"

        ' Then XOR decrypt using the encryption key
        Dim key As String
        ' Try to get encryption key (same logic as DecryptData)
        key = Environ$("COMPUTERNAME") & Environ$("USERNAME") & "GAFC2025SALT" & "audit-tool"

        Debug.Print "Using decryption key (first 50 chars): " & Left$(key, 50)

        Dim decrypted As String
        decrypted = ""
        Dim i As Long
        For i = 1 To Len(base64Decoded)
            Dim charCode As Integer
            charCode = Asc(Mid(base64Decoded, i, 1))
            Dim keyChar As Integer
            keyChar = Asc(Mid(key, ((i - 1) Mod Len(key)) + 1, 1))
            decrypted = decrypted & Chr((charCode Xor keyChar) Mod 256)
        Next i

        If Len(decrypted) > 0 Then
            Debug.Print "✓ Decryption successful!"
            Debug.Print ""
            Debug.Print "==== DECRYPTED CONTENT ===="
            Debug.Print decrypted
            Debug.Print "==== END DECRYPTED CONTENT ===="
            Debug.Print ""

            ' Check if checksum line exists
            If InStr(1, decrypted, "checksum=") > 0 Then
                ' Extract checksum value
                Dim checksumPos As Long
                checksumPos = InStr(1, decrypted, "checksum=")
                Dim checksumLine As String
                checksumLine = Mid$(decrypted, checksumPos)

                ' Get just the checksum value (everything after "checksum=")
                Dim checksumValue As String
                checksumValue = Mid$(checksumLine, 10) ' Skip "checksum="

                ' Remove any trailing whitespace/newlines
                checksumValue = Trim$(Split(checksumValue, vbCrLf)(0))

                Debug.Print "✓✓✓ FILE HAS CHECKSUM! ✓✓✓"
                Debug.Print "Checksum value: " & checksumValue
            Else
                Debug.Print "✗ WARNING: Checksum line not found in decrypted data"
            End If
        Else
            Debug.Print "✗ ERROR: Could not decrypt data (Base64Decode returned empty)"
        End If
    Else
        Debug.Print "✗ Missing ENC1: prefix (legacy plain-text format)"
        Debug.Print ""
        Debug.Print "File content:"
        Debug.Print rawContent
    End If

    Debug.Print ""
    Debug.Print "=== Verification Complete ==="
    Exit Sub

VerifyError:
    Debug.Print "✗ ERROR: " & Err.Description & " (Error " & Err.Number & ")"
    On Error Resume Next
    Close #f
End Sub

