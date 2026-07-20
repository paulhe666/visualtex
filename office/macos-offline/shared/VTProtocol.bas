Attribute VB_Name = "VTProtocol"
Option Explicit

Public Const VT_PROTOCOL_VERSION As Long = 1
Public Const VT_PLUGIN_VERSION As String = "1.1.0"
Public Const VT_METADATA_PREFIX As String = "visualtex:v1:deflate:"
Public Const VT_PENDING_PREFIX As String = "visualtex:pending:v1:"
Public Const VT_FORMULA_REF_PREFIX As String = "visualtex:formula-ref:v1:"

Private Const VT_MAX_REQUEST_BYTES As Long = 262144
Private Const VT_MAX_METADATA_CHARS As Long = 131072
Private VT_RANDOM_READY As Boolean
Private VT_UUID_COUNTER As Long

Public Function VTSessionRoot() As String
    VTSessionRoot = VTApplicationSupportRoot() & "/OfficeSessions"
End Function

Public Function VTSessionDirectory(ByVal sessionId As String) As String
    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7101, "VisualTeX", "Invalid VisualTeX session id."
    End If
    VTSessionDirectory = VTSessionRoot() & "/" & sessionId
End Function

Public Function VTRequestPath(ByVal sessionId As String) As String
    VTRequestPath = VTSessionDirectory(sessionId) & "/request.json"
End Function

Public Function VTDispatchPath(ByVal sessionId As String) As String
    VTDispatchPath = VTSessionDirectory(sessionId) & "/dispatch.txt"
End Function

Public Function VTActiveSessionPointer(ByVal hostName As String) As String
    Select Case LCase$(hostName)
        Case "word"
            VTActiveSessionPointer = VTSessionRoot() & "/word-active-session.txt"
        Case "powerpoint"
            VTActiveSessionPointer = VTSessionRoot() & "/powerpoint-active-session.txt"
        Case Else
            Err.Raise vbObjectError + 7102, "VisualTeX", "Invalid VisualTeX Office host."
    End Select
End Function

Public Function VTPlaceholderImagePath() As String
    VTPlaceholderImagePath = VTWordApplicationScriptsRoot() & "/VisualTeXPlaceholder.png"
End Function

Public Function VTNewUuidV4() As String
    Dim bytes(0 To 15) As Integer
    Dim index As Long

    If Not VT_RANDOM_READY Then
        Randomize CDbl(Date) + CDbl(Timer) / 86400#
        VT_RANDOM_READY = True
    End If
    If VT_UUID_COUNTER = 2147483647 Then
        VT_UUID_COUNTER = 1
    Else
        VT_UUID_COUNTER = VT_UUID_COUNTER + 1
    End If

    For index = 0 To 15
        bytes(index) = Int(Rnd * 256#)
    Next index
    bytes(15) = bytes(15) Xor (VT_UUID_COUNTER And &HFF)

    bytes(6) = (bytes(6) And &HF) Or &H40
    bytes(8) = (bytes(8) And &H3F) Or &H80

    VTNewUuidV4 = _
        VTHexByte(bytes(0)) & VTHexByte(bytes(1)) & VTHexByte(bytes(2)) & VTHexByte(bytes(3)) & "-" & _
        VTHexByte(bytes(4)) & VTHexByte(bytes(5)) & "-" & _
        VTHexByte(bytes(6)) & VTHexByte(bytes(7)) & "-" & _
        VTHexByte(bytes(8)) & VTHexByte(bytes(9)) & "-" & _
        VTHexByte(bytes(10)) & VTHexByte(bytes(11)) & VTHexByte(bytes(12)) & _
        VTHexByte(bytes(13)) & VTHexByte(bytes(14)) & VTHexByte(bytes(15))

    If Not VTIsCanonicalUuid(VTNewUuidV4) Then
        Err.Raise vbObjectError + 7103, "VisualTeX", "Unable to generate a VisualTeX UUID."
    End If
End Function

Private Function VTHexByte(ByVal value As Integer) As String
    VTHexByte = LCase$(Right$("0" & Hex$(value And &HFF), 2))
End Function

Public Function VTIsCanonicalUuid(ByVal value As String) As Boolean
    Dim index As Long
    Dim current As String

    If Len(value) <> 36 Then Exit Function
    If Mid$(value, 9, 1) <> "-" Or Mid$(value, 14, 1) <> "-" Or _
       Mid$(value, 19, 1) <> "-" Or Mid$(value, 24, 1) <> "-" Then Exit Function
    If LCase$(value) <> value Then Exit Function
    If Mid$(value, 15, 1) <> "4" Then Exit Function
    If InStr(1, "89ab", Mid$(value, 20, 1), vbBinaryCompare) = 0 Then Exit Function

    For index = 1 To Len(value)
        current = Mid$(value, index, 1)
        If index = 9 Or index = 14 Or index = 19 Or index = 24 Then
            If current <> "-" Then Exit Function
        ElseIf InStr(1, "0123456789abcdef", current, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next index
    VTIsCanonicalUuid = True
End Function

Public Function VTPendingMarker(ByVal sessionId As String, ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(sessionId) Or Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7104, "VisualTeX", "Invalid VisualTeX pending marker id."
    End If
    VTPendingMarker = VT_PENDING_PREFIX & sessionId & ":" & formulaId
End Function

Public Function VTJsonString(ByVal value As String) As String
    Dim result As String
    Dim index As Long
    Dim code As Long
    Dim current As String

    If Len(value) > VT_MAX_METADATA_CHARS Then
        Err.Raise vbObjectError + 7105, "VisualTeX", "VisualTeX request text is too large."
    End If

    result = """"
    For index = 1 To Len(value)
        current = Mid$(value, index, 1)
        code = AscW(current)
        Select Case current
            Case "\": result = result & "\\"
            Case """": result = result & "\"""
            Case vbCr: result = result & "\r"
            Case vbLf: result = result & "\n"
            Case vbTab: result = result & "\t"
            Case Else
                If code >= 0 And code < 32 Then
                    result = result & "\u" & Right$("0000" & Hex$(code), 4)
                Else
                    result = result & current
                End If
        End Select
    Next index
    VTJsonString = result & """"
End Function

Public Function VTJsonNullableString(ByVal value As String) As String
    If Len(value) = 0 Then
        VTJsonNullableString = "null"
    Else
        VTJsonNullableString = VTJsonString(value)
    End If
End Function

Public Function VTJsonBoolean(ByVal value As Boolean) As String
    If value Then
        VTJsonBoolean = "true"
    Else
        VTJsonBoolean = "false"
    End If
End Function

Public Function VTJsonNumber(ByVal value As Double) As String
    If value <> value Or Abs(value) > 10000000# Then
        Err.Raise vbObjectError + 7106, "VisualTeX", "Invalid VisualTeX numeric value."
    End If
    VTJsonNumber = Replace$(Trim$(Str$(value)), ",", ".")
End Function

Public Function VTParseInvariantDouble(ByVal value As String) As Double
    Dim text As String
    Dim index As Long
    Dim current As String
    Dim digitCount As Long
    Dim fractionDigitCount As Long
    Dim decimalSeen As Boolean

    text = Trim$(value)
    If Len(text) = 0 Or Len(text) > 64 Then GoTo InvalidNumber

    index = 1
    If Left$(text, 1) = "-" Then index = 2
    If index > Len(text) Then GoTo InvalidNumber

    For index = index To Len(text)
        current = Mid$(text, index, 1)
        If InStr(1, "0123456789", current, vbBinaryCompare) > 0 Then
            digitCount = digitCount + 1
            If decimalSeen Then fractionDigitCount = fractionDigitCount + 1
        ElseIf current = "." And Not decimalSeen Then
            decimalSeen = True
        Else
            GoTo InvalidNumber
        End If
    Next index

    If digitCount = 0 Then GoTo InvalidNumber
    If decimalSeen And fractionDigitCount = 0 Then GoTo InvalidNumber

    On Error GoTo InvalidNumber
    VTParseInvariantDouble = Val(text)
    If VTParseInvariantDouble <> VTParseInvariantDouble Or _
       Abs(VTParseInvariantDouble) > 10000000# Then GoTo InvalidNumber
    Exit Function

InvalidNumber:
    On Error GoTo 0
    Err.Raise vbObjectError + 7123, "VisualTeX", "VisualTeX dispatch contains an invalid invariant number."
End Function

Private Function VTUtf8ByteLength(ByVal value As String) As Long
    Dim bytes() As Byte
    If Len(value) = 0 Then Exit Function
    bytes = VTUtf8Encode(value)
    VTUtf8ByteLength = UBound(bytes) - LBound(bytes) + 1
End Function

Private Function VTUtf8Encode(ByVal value As String) As Byte()
    Dim output() As Byte
    Dim byteCount As Long
    Dim index As Long
    Dim codeUnit As Long
    Dim lowSurrogate As Long
    Dim codePoint As Long

    If Len(value) = 0 Then
        ReDim output(0 To 0)
        VTUtf8Encode = output
        Exit Function
    End If

    ReDim output(0 To Len(value) * 4 - 1)
    index = 1
    Do While index <= Len(value)
        codeUnit = AscW(Mid$(value, index, 1))
        If codeUnit < 0 Then codeUnit = codeUnit + 65536

        If codeUnit >= 55296 And codeUnit <= 56319 Then
            If index = Len(value) Then
                Err.Raise vbObjectError + 7118, "VisualTeX", "VisualTeX text contains an incomplete Unicode surrogate pair."
            End If
            lowSurrogate = AscW(Mid$(value, index + 1, 1))
            If lowSurrogate < 0 Then lowSurrogate = lowSurrogate + 65536
            If lowSurrogate < 56320 Or lowSurrogate > 57343 Then
                Err.Raise vbObjectError + 7118, "VisualTeX", "VisualTeX text contains an invalid Unicode surrogate pair."
            End If
            codePoint = 65536 + (codeUnit - 55296) * 1024 + (lowSurrogate - 56320)
            index = index + 1
        ElseIf codeUnit >= 56320 And codeUnit <= 57343 Then
            Err.Raise vbObjectError + 7118, "VisualTeX", "VisualTeX text contains an unexpected low surrogate."
        Else
            codePoint = codeUnit
        End If

        If codePoint <= 127 Then
            VTAppendUtf8Byte output, byteCount, codePoint
        ElseIf codePoint <= 2047 Then
            VTAppendUtf8Byte output, byteCount, 192 Or (codePoint \ 64)
            VTAppendUtf8Byte output, byteCount, 128 Or (codePoint And 63)
        ElseIf codePoint <= 65535 Then
            VTAppendUtf8Byte output, byteCount, 224 Or (codePoint \ 4096)
            VTAppendUtf8Byte output, byteCount, 128 Or ((codePoint \ 64) And 63)
            VTAppendUtf8Byte output, byteCount, 128 Or (codePoint And 63)
        ElseIf codePoint <= 1114111 Then
            VTAppendUtf8Byte output, byteCount, 240 Or (codePoint \ 262144)
            VTAppendUtf8Byte output, byteCount, 128 Or ((codePoint \ 4096) And 63)
            VTAppendUtf8Byte output, byteCount, 128 Or ((codePoint \ 64) And 63)
            VTAppendUtf8Byte output, byteCount, 128 Or (codePoint And 63)
        Else
            Err.Raise vbObjectError + 7118, "VisualTeX", "VisualTeX text contains an invalid Unicode code point."
        End If
        index = index + 1
    Loop

    ReDim Preserve output(0 To byteCount - 1)
    VTUtf8Encode = output
End Function

Private Sub VTAppendUtf8Byte(ByRef output() As Byte, ByRef byteCount As Long, ByVal value As Long)
    output(byteCount) = CByte(value And 255)
    byteCount = byteCount + 1
End Sub

Public Function VTBase64UrlDecodeUtf8(ByVal value As String) As String
    Dim output() As Byte
    Dim outputCount As Long
    Dim index As Long
    Dim remaining As Long
    Dim firstValue As Long
    Dim secondValue As Long
    Dim thirdValue As Long
    Dim fourthValue As Long

    If Len(value) = 0 Or Len(value) > 2097152 Then
        Err.Raise vbObjectError + 7124, "VisualTeX", "VisualTeX Word LaTeX payload is empty or too large."
    End If
    If Len(value) Mod 4 = 1 Then
        Err.Raise vbObjectError + 7124, "VisualTeX", "VisualTeX Word LaTeX payload has invalid base64url length."
    End If

    ReDim output(0 To ((Len(value) + 3) \ 4) * 3 - 1)
    index = 1
    Do While index <= Len(value)
        remaining = Len(value) - index + 1
        If remaining < 2 Then GoTo InvalidBase64Url
        firstValue = VTBase64UrlValue(Mid$(value, index, 1))
        secondValue = VTBase64UrlValue(Mid$(value, index + 1, 1))
        If firstValue < 0 Or secondValue < 0 Then GoTo InvalidBase64Url

        output(outputCount) = CByte(firstValue * 4 + secondValue \ 16)
        outputCount = outputCount + 1

        If remaining >= 3 Then
            thirdValue = VTBase64UrlValue(Mid$(value, index + 2, 1))
            If thirdValue < 0 Then GoTo InvalidBase64Url
            output(outputCount) = CByte((secondValue And 15) * 16 + thirdValue \ 4)
            outputCount = outputCount + 1
        ElseIf (secondValue And 15) <> 0 Then
            GoTo InvalidBase64Url
        End If

        If remaining >= 4 Then
            fourthValue = VTBase64UrlValue(Mid$(value, index + 3, 1))
            If fourthValue < 0 Then GoTo InvalidBase64Url
            output(outputCount) = CByte((thirdValue And 3) * 64 + fourthValue)
            outputCount = outputCount + 1
        ElseIf remaining = 3 And (thirdValue And 3) <> 0 Then
            GoTo InvalidBase64Url
        End If

        index = index + 4
    Loop

    If outputCount = 0 Then GoTo InvalidBase64Url
    ReDim Preserve output(0 To outputCount - 1)
    VTBase64UrlDecodeUtf8 = VTUtf8Decode(output)
    Exit Function

InvalidBase64Url:
    Err.Raise vbObjectError + 7124, "VisualTeX", "VisualTeX Word LaTeX payload is not valid unpadded base64url."
End Function

Private Function VTBase64UrlEncodeUtf8(ByVal value As String) As String
    Dim bytes() As Byte
    If Len(value) = 0 Then Exit Function
    bytes = VTUtf8Encode(value)
    VTBase64UrlEncodeUtf8 = VTBase64UrlEncodeBytes(bytes)
End Function

Private Function VTBase64UrlEncodeBytes(ByRef bytes() As Byte) As String
    Dim alphabet As String
    Dim result As String
    Dim index As Long
    Dim remaining As Long
    Dim firstByte As Long
    Dim secondByte As Long
    Dim thirdByte As Long

    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
    index = LBound(bytes)
    Do While index <= UBound(bytes)
        remaining = UBound(bytes) - index + 1
        firstByte = CLng(bytes(index))
        If remaining >= 2 Then secondByte = CLng(bytes(index + 1)) Else secondByte = 0
        If remaining >= 3 Then thirdByte = CLng(bytes(index + 2)) Else thirdByte = 0

        result = result & Mid$(alphabet, firstByte \ 4 + 1, 1)
        result = result & Mid$(alphabet, (firstByte And 3) * 16 + secondByte \ 16 + 1, 1)
        If remaining >= 2 Then
            result = result & Mid$(alphabet, (secondByte And 15) * 4 + thirdByte \ 64 + 1, 1)
        End If
        If remaining >= 3 Then
            result = result & Mid$(alphabet, (thirdByte And 63) + 1, 1)
        End If
        index = index + 3
    Loop
    VTBase64UrlEncodeBytes = result
End Function

Private Function VTBase64UrlValue(ByVal value As String) As Long
    Dim code As Long
    If Len(value) <> 1 Then
        VTBase64UrlValue = -1
        Exit Function
    End If
    code = AscW(value)
    Select Case code
        Case 65 To 90: VTBase64UrlValue = code - 65
        Case 97 To 122: VTBase64UrlValue = code - 97 + 26
        Case 48 To 57: VTBase64UrlValue = code - 48 + 52
        Case 45: VTBase64UrlValue = 62
        Case 95: VTBase64UrlValue = 63
        Case Else: VTBase64UrlValue = -1
    End Select
End Function

Private Function VTUtf8Decode(ByRef bytes() As Byte) As String
    Dim index As Long
    Dim lastIndex As Long
    Dim firstByte As Long
    Dim secondByte As Long
    Dim thirdByte As Long
    Dim fourthByte As Long
    Dim codePoint As Long
    Dim result As String

    index = LBound(bytes)
    lastIndex = UBound(bytes)
    Do While index <= lastIndex
        firstByte = CLng(bytes(index))
        If firstByte <= 127 Then
            codePoint = firstByte
        ElseIf firstByte >= 194 And firstByte <= 223 Then
            If index + 1 > lastIndex Then GoTo InvalidUtf8
            secondByte = CLng(bytes(index + 1))
            If secondByte < 128 Or secondByte > 191 Then GoTo InvalidUtf8
            codePoint = (firstByte And 31) * 64 + (secondByte And 63)
            index = index + 1
        ElseIf firstByte >= 224 And firstByte <= 239 Then
            If index + 2 > lastIndex Then GoTo InvalidUtf8
            secondByte = CLng(bytes(index + 1))
            thirdByte = CLng(bytes(index + 2))
            If thirdByte < 128 Or thirdByte > 191 Then GoTo InvalidUtf8
            If firstByte = 224 Then
                If secondByte < 160 Or secondByte > 191 Then GoTo InvalidUtf8
            ElseIf firstByte = 237 Then
                If secondByte < 128 Or secondByte > 159 Then GoTo InvalidUtf8
            ElseIf secondByte < 128 Or secondByte > 191 Then
                GoTo InvalidUtf8
            End If
            codePoint = (firstByte And 15) * 4096 + (secondByte And 63) * 64 + (thirdByte And 63)
            index = index + 2
        ElseIf firstByte >= 240 And firstByte <= 244 Then
            If index + 3 > lastIndex Then GoTo InvalidUtf8
            secondByte = CLng(bytes(index + 1))
            thirdByte = CLng(bytes(index + 2))
            fourthByte = CLng(bytes(index + 3))
            If thirdByte < 128 Or thirdByte > 191 Or fourthByte < 128 Or fourthByte > 191 Then GoTo InvalidUtf8
            If firstByte = 240 Then
                If secondByte < 144 Or secondByte > 191 Then GoTo InvalidUtf8
            ElseIf firstByte = 244 Then
                If secondByte < 128 Or secondByte > 143 Then GoTo InvalidUtf8
            ElseIf secondByte < 128 Or secondByte > 191 Then
                GoTo InvalidUtf8
            End If
            codePoint = (firstByte And 7) * 262144 + (secondByte And 63) * 4096 + _
                        (thirdByte And 63) * 64 + (fourthByte And 63)
            index = index + 3
        Else
            GoTo InvalidUtf8
        End If

        If codePoint <= 65535 Then
            result = result & VTUnicodeCodeUnit(codePoint)
        Else
            codePoint = codePoint - 65536
            result = result & VTUnicodeCodeUnit(55296 + (codePoint \ 1024)) & _
                              VTUnicodeCodeUnit(56320 + (codePoint And 1023))
        End If
        index = index + 1
    Loop
    VTUtf8Decode = result
    Exit Function

InvalidUtf8:
    Err.Raise vbObjectError + 7119, "VisualTeX", "VisualTeX local file is not valid UTF-8."
End Function

Private Function VTUnicodeCodeUnit(ByVal value As Long) As String
    If value < 0 Or value > 65535 Then
        Err.Raise vbObjectError + 7120, "VisualTeX", "VisualTeX decoded an invalid Unicode code unit."
    End If
    If value > 32767 Then
        VTUnicodeCodeUnit = ChrW(value - 65536)
    Else
        VTUnicodeCodeUnit = ChrW(value)
    End If
End Function

Private Function VTFileBridgeScriptName() As String
    Dim hostName As String
    hostName = LCase$(Application.Name)
    If InStr(1, hostName, "powerpoint", vbTextCompare) > 0 Then
        VTFileBridgeScriptName = "VisualTeXPowerPoint.scpt"
    ElseIf InStr(1, hostName, "word", vbTextCompare) > 0 Then
        VTFileBridgeScriptName = "VisualTeXWord.scpt"
    Else
        Err.Raise vbObjectError + 7125, "VisualTeX", "Unable to identify the Office host for the VisualTeX file bridge."
    End If
End Function

Private Function VTFileBridgeCall(ByVal handlerName As String, ByVal parameterValue As String) As String
    Dim response As String
    Dim fields() As String
    Dim detail As String

#If Mac Then
    response = AppleScriptTask(VTFileBridgeScriptName(), handlerName, parameterValue)
#Else
    Err.Raise vbObjectError + 7126, "VisualTeX", "The VisualTeX native Office file bridge is available only on macOS."
#End If

    If Left$(response, 3) = "ok|" Then
        VTFileBridgeCall = Mid$(response, 4)
        Exit Function
    End If

    detail = "The VisualTeX native file bridge returned an invalid response."
    If Left$(response, 6) = "error|" Then
        fields = Split(response, "|")
        If UBound(fields) >= 2 Then detail = fields(2)
        If UBound(fields) >= 1 Then detail = detail & " (AppleScript error " & fields(1) & ")"
    ElseIf Len(response) = 0 Then
        detail = "The VisualTeX native file bridge returned no response. Check that " & VTFileBridgeScriptName() & " is installed and compiled."
    End If
    Err.Raise vbObjectError + 7127, "VisualTeX", detail
End Function

Private Function VTRuntimeRelativePath(ByVal absolutePath As String) As String
    Dim root As String
    Dim prefix As String

    root = VTApplicationSupportRoot()
    prefix = root & "/"
    If absolutePath = root Then
        VTRuntimeRelativePath = "."
    ElseIf Len(absolutePath) > Len(prefix) And Left$(absolutePath, Len(prefix)) = prefix Then
        VTRuntimeRelativePath = Mid$(absolutePath, Len(prefix) + 1)
    Else
        Err.Raise vbObjectError + 7117, "VisualTeX", "VisualTeX rejected a path outside the host runtime directory."
    End If
    If InStr(VTRuntimeRelativePath, "..") > 0 Or InStr(VTRuntimeRelativePath, vbCr) > 0 Or _
       InStr(VTRuntimeRelativePath, vbLf) > 0 Or InStr(VTRuntimeRelativePath, Chr$(0)) > 0 Then
        Err.Raise vbObjectError + 7117, "VisualTeX", "VisualTeX rejected an unsafe local path."
    End If
End Function

Public Sub VTWriteRequest(ByVal sessionId As String, ByVal json As String)
    Dim requestPath As String
    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7107, "VisualTeX", "Invalid VisualTeX session id."
    End If
    If VTUtf8ByteLength(json) > VT_MAX_REQUEST_BYTES Then
        Err.Raise vbObjectError + 7108, "VisualTeX", "VisualTeX request exceeds 256 KiB."
    End If
    ' WriteVisualTeXFile creates the Session parent directory atomically.
    ' Avoid a separate AppleScriptTask round trip solely for mkdir.
    requestPath = VTRequestPath(sessionId)
    VTWriteTextAtomic requestPath, json
End Sub

Public Sub VTWriteTextAtomic(ByVal destination As String, ByVal contents As String)
    Dim relativePath As String
    Dim encodedContents As String

    VTValidateAbsoluteVisualTeXPath destination
    relativePath = VTRuntimeRelativePath(destination)
    encodedContents = VTBase64UrlEncodeUtf8(contents)
    Call VTFileBridgeCall("WriteVisualTeXFile", relativePath & "|" & encodedContents)
End Sub

Public Function VTReadText(ByVal sourcePath As String, Optional ByVal maximumCharacters As Long = 262144) As String
    Dim encodedContents As String
    Dim decodedContents As String

    VTValidateAbsoluteVisualTeXPath sourcePath
    encodedContents = VTFileBridgeCall("ReadVisualTeXFile", VTRuntimeRelativePath(sourcePath))
    If Len(encodedContents) = 0 Then
        decodedContents = ""
    Else
        decodedContents = VTBase64UrlDecodeUtf8(encodedContents)
    End If
    If Len(decodedContents) > maximumCharacters Then
        Err.Raise vbObjectError + 7111, "VisualTeX", "VisualTeX local file exceeds the allowed size."
    End If
    VTReadText = decodedContents
End Function

Public Function VTReadActiveSessionId(ByVal hostName As String) As String
    Dim value As String
    value = Trim$(VTReadText(VTActiveSessionPointer(hostName), 128))
    If Not VTIsCanonicalUuid(value) Then
        Err.Raise vbObjectError + 7112, "VisualTeX", "VisualTeX active Session pointer is invalid."
    End If
    VTReadActiveSessionId = value
End Function

Public Function VTReadDispatch(ByVal sessionId As String) As Object
    Dim dictionary As Object
    Dim contents As String
    Dim rows() As String
    Dim row As Variant
    Dim separator As Long
    Dim key As String
    Dim value As String

    Set dictionary = New Collection
    contents = Replace$(VTReadText(VTDispatchPath(sessionId), 524288), vbCrLf, vbLf)
    contents = Replace$(contents, vbCr, vbLf)
    rows = Split(contents, vbLf)
    For Each row In rows
        If Len(CStr(row)) > 0 Then
            separator = InStr(1, CStr(row), "=", vbBinaryCompare)
            If separator <= 1 Then
                Err.Raise vbObjectError + 7113, "VisualTeX", "VisualTeX dispatch contains an invalid row."
            End If
            key = Left$(CStr(row), separator - 1)
            value = Mid$(CStr(row), separator + 1)
            If VTCollectionHasKey(dictionary, key) Then
                Err.Raise vbObjectError + 7114, "VisualTeX", "VisualTeX dispatch contains a duplicate key."
            End If
            dictionary.Add value, key
        End If
    Next row

    VTRequireDispatchValue dictionary, "protocolVersion"
    VTRequireDispatchValue dictionary, "sessionId"
    VTRequireDispatchValue dictionary, "action"
    VTRequireDispatchValue dictionary, "host"
    If dictionary("protocolVersion") <> CStr(VT_PROTOCOL_VERSION) Or dictionary("sessionId") <> sessionId Then
        Err.Raise vbObjectError + 7115, "VisualTeX", "VisualTeX dispatch identity does not match the active Session."
    End If
    Set VTReadDispatch = dictionary
End Function

Public Sub VTRequireDispatchValue(ByVal dictionary As Object, ByVal key As String)
    If Not VTCollectionHasKey(dictionary, key) Then
        Err.Raise vbObjectError + 7116, "VisualTeX", "VisualTeX dispatch is missing " & key & "."
    End If
    If Len(CStr(dictionary(key))) = 0 Then
        Err.Raise vbObjectError + 7116, "VisualTeX", "VisualTeX dispatch is missing " & key & "."
    End If
End Sub

Public Function VTCollectionHasKey(ByVal collection As Object, ByVal key As String) As Boolean
    Dim value As Variant
    On Error Resume Next
    value = collection(key)
    VTCollectionHasKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Public Sub VTValidateAbsoluteVisualTeXPath(ByVal value As String)
    Dim root As String
    root = VTApplicationSupportRoot() & "/"
    If Len(value) <= Len(root) Or Left$(value, Len(root)) <> root Or _
       InStr(value, "..") > 0 Or InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Or InStr(value, Chr$(0)) > 0 Then
        Err.Raise vbObjectError + 7117, "VisualTeX", "VisualTeX rejected an unsafe local path."
    End If
End Sub

Public Sub VTEnsureDirectory(ByVal directoryPath As String)
    If Len(directoryPath) = 0 Then Exit Sub
    Call VTFileBridgeCall("EnsureVisualTeXDirectory", VTRuntimeRelativePath(directoryPath))
End Sub

Public Function VTParentDirectory(ByVal value As String) As String
    Dim separator As Long
    separator = InStrRev(value, "/")
    If separator <= 1 Then
        VTParentDirectory = "/"
    Else
        VTParentDirectory = Left$(value, separator - 1)
    End If
End Function

Public Function VTPathFileExists(ByVal value As String) As Boolean
    Dim response As String
    Dim handle As Integer

    On Error GoTo MissingFile
    If value = VTPlaceholderImagePath() Then
        handle = FreeFile
        Open value For Binary Access Read As #handle
        Close #handle
        handle = 0
        VTPathFileExists = True
        Exit Function
    End If

    VTValidateAbsoluteVisualTeXPath value
    response = VTFileBridgeCall("VisualTeXFileExists", VTRuntimeRelativePath(value))
    VTPathFileExists = (response = "1")
    Exit Function

MissingFile:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    Err.Clear
    On Error GoTo 0
    VTPathFileExists = False
End Function

Public Function VTProtocolSelfTest() As Boolean
    Dim identifiers As New Collection
    Dim identifier As String
    Dim index As Long
    Dim testPath As String
    Dim sample As String
    Dim parsedNumber As Double

    For index = 1 To 1000
        identifier = VTNewUuidV4()
        If Not VTIsCanonicalUuid(identifier) Or VTCollectionHasKey(identifiers, identifier) Then
            Err.Raise vbObjectError + 7121, "VisualTeX", "VisualTeX UUID self-test failed."
        End If
        identifiers.Add True, identifier
    Next index

    sample = "VisualTeX " & VTUnicodeCodeUnit(20013) & VTUnicodeCodeUnit(25991) & _
             " " & VTUnicodeCodeUnit(960) & " " & VTUnicodeCodeUnit(55357) & VTUnicodeCodeUnit(56832)
    VTEnsureDirectory VTSessionRoot()
    testPath = VTSessionRoot() & "/protocol-self-test.txt"
    VTWriteTextAtomic testPath, sample
    If VTReadText(testPath, 1024) <> sample Then
        Err.Raise vbObjectError + 7122, "VisualTeX", "VisualTeX UTF-8 self-test failed."
    End If
    On Error Resume Next
    Call VTFileBridgeCall("DeleteVisualTeXFile", VTRuntimeRelativePath(testPath))
    On Error GoTo 0

    parsedNumber = VTParseInvariantDouble("-1234.500000")
    If Abs(parsedNumber + 1234.5) > 0.0000001 Then
        Err.Raise vbObjectError + 7124, "VisualTeX", "VisualTeX invariant-number self-test failed."
    End If
    On Error Resume Next
    parsedNumber = VTParseInvariantDouble("12.5invalid")
    If Err.Number = 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError + 7124, "VisualTeX", "VisualTeX invariant-number rejection self-test failed."
    End If
    Err.Clear
    On Error GoTo 0

    VTProtocolSelfTest = True
End Function

Public Sub VTDeleteSessionFiles(ByVal sessionId As String)
    Dim relativeDirectory As String
    If Not VTIsCanonicalUuid(sessionId) Then Exit Sub
    relativeDirectory = "OfficeSessions/" & sessionId
    On Error Resume Next
    Call VTFileBridgeCall("DeleteVisualTeXFile", relativeDirectory & "/request.json")
    Call VTFileBridgeCall("DeleteVisualTeXFile", relativeDirectory & "/dispatch.txt")
    Call VTFileBridgeCall("DeleteVisualTeXFile", relativeDirectory & "/formula.png")
    On Error GoTo 0
End Sub
