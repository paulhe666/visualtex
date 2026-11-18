Attribute VB_Name = "VTLauncher"
Option Explicit

Private VT_OFFICE_RESIDENT_PREWARMED As Boolean
Private Const VT_FAST_OPEN_MAX_REQUEST_BYTES As Long = 65536
Private Const VT_WORD_SANDBOX_HOME_SUFFIX As String = "/Library/Containers/com.microsoft.Word/Data"
Private Const VT_POWERPOINT_SANDBOX_HOME_SUFFIX As String = "/Library/Containers/com.microsoft.Powerpoint/Data"

Public Sub VTLaunchSession(ByVal hostName As String, ByVal sessionId As String)
    Dim scriptName As String
    Dim response As String

    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7300, "VisualTeX", "Invalid VisualTeX Session id."
    End If
    If VT_AGENT_MODE Then VT_AGENT_SESSION_ID = sessionId
    scriptName = VTOfficeLauncherScriptName(hostName)

#If Mac Then
    response = AppleScriptTask(scriptName, "OpenVisualTeXSession", sessionId)
#Else
    Err.Raise vbObjectError + 7302, "VisualTeX", "The VisualTeX offline add-in is available only on macOS."
#End If

    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7303, "VisualTeX", VTAppleScriptErrorMessage(response)
    End If
End Sub

Public Function VTWriteAndLaunchSession( _
    ByVal hostName As String, _
    ByVal sessionId As String, _
    ByVal requestJson As String) As String
    Dim normalizedHost As String
    Dim scriptName As String
    Dim encodedRequest As String
    Dim response As String
    Dim timingDetail As String
    Dim directTimingDetail As String
    Dim operationStage As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed
    operationStage = "normalize-host"
    normalizedHost = VTCanonicalOfficeHost(hostName)
    operationStage = "resolve-script"
    scriptName = VTOfficeLauncherScriptName(normalizedHost)
    operationStage = "validate-request"
    VTValidateRequestPayload sessionId, requestJson
    If VT_AGENT_MODE Then
        VT_AGENT_SESSION_ID = sessionId
        VT_AGENT_REQUEST_JSON = requestJson
        VTWriteAndLaunchSession = "agent"
        Exit Function
    End If
    operationStage = "fast-direct-launch"
#If Mac Then
    If VTTryWriteAndLaunchSessionDirect( _
       normalizedHost, sessionId, requestJson, directTimingDetail) Then
        VTWriteAndLaunchSession = directTimingDetail
        Exit Function
    End If
#End If

    operationStage = "encode-request"
    encodedRequest = VTBase64UrlEncodeUtf8(requestJson)

#If Mac Then
    operationStage = "applescript-task"
    response = AppleScriptTask( _
        scriptName, _
        "WriteAndOpenVisualTeXSession", _
        normalizedHost & "|" & sessionId & "|" & encodedRequest)
#Else
    Err.Raise vbObjectError + 7307, "VisualTeX", "The VisualTeX offline add-in is available only on macOS."
#End If

    operationStage = "validate-response"
    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7308, "VisualTeX", _
            VTAppleScriptErrorMessage(response) & _
            IIf(Len(directTimingDetail) > 0, " (" & directTimingDetail & ")", "")
    End If
#If Mac Then
    VT_OFFICE_RESIDENT_PREWARMED = True
#End If
    timingDetail = Mid$(response, 4)
    If Len(directTimingDetail) > 0 Then
        timingDetail = timingDetail & ";" & directTimingDetail
    End If
    If InStr(1, timingDetail, "totalMs=", vbBinaryCompare) = 0 Or _
       InStr(1, timingDetail, "writeMs=", vbBinaryCompare) = 0 Or _
       InStr(1, timingDetail, "launchMs=", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7309, "VisualTeX", _
            "VisualTeX received an invalid write-and-launch timing response."
    End If
    VTWriteAndLaunchSession = timingDetail
    Exit Function

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, "VisualTeX write and launch", _
        operationStage & ": " & errorDescription
End Function

Public Function VTWriteFormulaRestoreAndLaunchSession( _
    ByVal hostName As String, _
    ByVal sessionId As String, _
    ByVal requestJson As String, _
    ByVal restoreSource As String) As String

    Dim normalizedHost As String
    Dim scriptName As String
    Dim encodedRequest As String
    Dim encodedSource As String
    Dim response As String
    Dim timingDetail As String
    Dim operationStage As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed
    operationStage = "normalize-host"
    normalizedHost = VTCanonicalOfficeHost(hostName)
    If normalizedHost <> "word" Then
        Err.Raise vbObjectError + 7312, "VisualTeX", _
            "Formula restore write-and-launch is available only for Word."
    End If
    operationStage = "resolve-script"
    scriptName = VTOfficeLauncherScriptName(normalizedHost)
    operationStage = "validate-request"
    VTValidateRequestPayload sessionId, requestJson
    If Len(restoreSource) = 0 Then
        Err.Raise vbObjectError + 7312, "VisualTeX", _
            "The formula restore source is empty."
    End If
    If VT_AGENT_MODE Then VT_AGENT_SESSION_ID = sessionId
    operationStage = "encode-payloads"
    encodedRequest = VTBase64UrlEncodeUtf8(requestJson)
    encodedSource = VTBase64UrlEncodeUtf8(restoreSource)

#If Mac Then
    operationStage = "applescript-task"
    response = AppleScriptTask( _
        scriptName, _
        "WriteFormulaRestoreAndOpenVisualTeXSession", _
        normalizedHost & "|" & sessionId & "|" & _
        encodedRequest & "|" & encodedSource)
#Else
    Err.Raise vbObjectError + 7312, "VisualTeX", _
        "The VisualTeX offline add-in is available only on macOS."
#End If

    operationStage = "validate-response"
    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7313, "VisualTeX", _
            VTAppleScriptErrorMessage(response)
    End If
    timingDetail = Mid$(response, 4)
    If InStr(1, timingDetail, "totalMs=", vbBinaryCompare) = 0 Or _
       InStr(1, timingDetail, "writeMs=", vbBinaryCompare) = 0 Or _
       InStr(1, timingDetail, "launchMs=", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7314, "VisualTeX", _
            "VisualTeX received an invalid formula restore timing response."
    End If
    VTWriteFormulaRestoreAndLaunchSession = timingDetail
    Exit Function

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, "VisualTeX formula restore write and launch", _
        operationStage & ": " & errorDescription
End Function

Public Function VTConvertOmmlBatch( _
    ByVal hostName As String, _
    ByVal sessionId As String, _
    Optional ByVal restoreSource As String = "") As String

    Dim normalizedHost As String
    Dim scriptName As String
    Dim encodedSource As String
    Dim response As String
    Dim requestArgument As String
    Dim timingDetail As String
    Dim operationStage As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed
    operationStage = "normalize-host"
    normalizedHost = VTCanonicalOfficeHost(hostName)
    If normalizedHost <> "word" Then
        Err.Raise vbObjectError + 7315, "VisualTeX", _
            "OMML batch conversion is available only for Word."
    End If
    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7315, "VisualTeX", _
            "The OMML batch Session id is invalid."
    End If
    operationStage = "resolve-script"
    scriptName = VTOfficeLauncherScriptName(normalizedHost)
    requestArgument = sessionId
    If Len(restoreSource) > 0 Then
        operationStage = "encode-source"
        encodedSource = VTBase64UrlEncodeUtf8(restoreSource)
        requestArgument = sessionId & "|" & encodedSource
    End If

#If Mac Then
    operationStage = "applescript-task"
    response = AppleScriptTask( _
        scriptName, _
        "ConvertOmmlBatch", _
        requestArgument)
#Else
    Err.Raise vbObjectError + 7315, "VisualTeX", _
        "OMML batch conversion is available only on macOS."
#End If

    operationStage = "validate-response"
    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7316, "VisualTeX", _
            VTAppleScriptErrorMessage(response)
    End If
    timingDetail = Mid$(response, 4)
    If InStr(1, timingDetail, "totalMs=", vbBinaryCompare) = 0 Or _
       InStr(1, timingDetail, "convertMs=", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7317, "VisualTeX", _
            "VisualTeX received an invalid OMML batch timing response."
    End If
    VTConvertOmmlBatch = timingDetail
    Exit Function

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, "VisualTeX OMML batch conversion", _
        operationStage & ": " & errorDescription
End Function

Public Sub VTPrewarmApplication(ByVal hostName As String)
    Dim normalizedHost As String
    Dim scriptName As String
    Dim response As String

    normalizedHost = VTCanonicalOfficeHost(hostName)
    scriptName = VTOfficeLauncherScriptName(normalizedHost)

#If Mac Then
    response = AppleScriptTask( _
        scriptName, "PrewarmVisualTeXApplication", normalizedHost)
#Else
    Err.Raise vbObjectError + 7310, "VisualTeX", "The VisualTeX offline add-in is available only on macOS."
#End If

    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7311, "VisualTeX", VTAppleScriptErrorMessage(response)
    End If
#If Mac Then
    VT_OFFICE_RESIDENT_PREWARMED = True
#End If
End Sub

Public Sub VTOpenApplication(ByVal hostName As String)
    Dim scriptName As String
    Dim response As String

    scriptName = VTOfficeLauncherScriptName(hostName)

#If Mac Then
    response = AppleScriptTask(scriptName, "OpenVisualTeXApplication", "")
#Else
    Err.Raise vbObjectError + 7305, "VisualTeX", "The VisualTeX offline add-in is available only on macOS."
#End If

    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7306, "VisualTeX", VTAppleScriptErrorMessage(response)
    End If
End Sub

Private Function VTTryWriteAndLaunchSessionDirect( _
    ByVal normalizedHost As String, _
    ByVal sessionId As String, _
    ByVal requestJson As String, _
    ByRef timingDetail As String) As Boolean

#If Mac Then
    Dim inboxRoot As String
    Dim requestPath As String
    Dim temporaryPath As String
    Dim requestBytes() As Byte
    Dim requestByteCount As Long
    Dim handle As Integer
    Dim startedAt As Single
    Dim payloadReadyAt As Single
    Dim launchFinishedAt As Single

    startedAt = Timer
    On Error GoTo FastPathFailed
    If Not VT_OFFICE_RESIDENT_PREWARMED Then
        VT_OFFICE_RESIDENT_PREWARMED = VTFastOpenResidentReady(normalizedHost)
    End If
    If Not VT_OFFICE_RESIDENT_PREWARMED Then
        Err.Raise vbObjectError + 7322, "VisualTeX", _
            "The VisualTeX Office resident is not prewarmed."
    End If
    If normalizedHost <> "word" And normalizedHost <> "powerpoint" Then _
        GoTo FastPathFailed
    VTValidateRequestPayload sessionId, requestJson

    ' Only ordinary, small formula-open requests use the sandbox inbox. Batch
    ' imports, redraw/restore operations and oversized metadata retain the
    ' established AppleScriptTask transaction because they have auxiliary files.
    If InStr(1, requestJson, """operation"":""formula""", vbBinaryCompare) = 0 Then _
        GoTo FastPathFailed
    requestByteCount = VTUtf8ByteLength(requestJson)
    If requestByteCount <= 0 Or requestByteCount > VT_FAST_OPEN_MAX_REQUEST_BYTES Then _
        GoTo FastPathFailed

    inboxRoot = VTFastOpenInboxRoot(normalizedHost)
    VTEnsureFastOpenInboxDirectory inboxRoot
    requestPath = inboxRoot & "/" & sessionId & ".json"
    temporaryPath = inboxRoot & "/." & sessionId & ".tmp"
    requestBytes = VTUtf8Encode(requestJson)

    On Error Resume Next
    Kill temporaryPath
    Kill requestPath
    Err.Clear
    On Error GoTo FastPathFailed

    handle = FreeFile
    Open temporaryPath For Binary Access Write As #handle
    Put #handle, , requestBytes
    Close #handle
    handle = 0
    Name temporaryPath As requestPath
    payloadReadyAt = Timer

    ' The prewarmed resident watches this fixed sandbox inbox directly. Do not
    ' launch another process here: Mac VBA Shell can trigger a Terminal permission
    ' prompt, Office-spawned processes cannot use Tauri's /tmp single-instance
    ' socket, and LaunchServices does not reliably emit Reopen for a hidden app.
    ' A recently killed resident can leave a briefly fresh readiness heartbeat,
    ' so require the resident to claim the just-published request before treating
    ' this as a successful fast open. Healthy polling claims it within ~25-50 ms;
    ' otherwise the caller falls back to the proven AppleScriptTask cold path.
    If Not VTFastOpenRequestClaimed(requestPath, payloadReadyAt) Then
        Err.Raise vbObjectError + 7326, "VisualTeX", _
            "The VisualTeX Office resident did not claim the fast-open request."
    End If
    launchFinishedAt = Timer

    timingDetail = _
        "fastPath=inbox-poll;writeMs=" & _
            CStr(CLng(1000# * VTFastLaunchElapsedSeconds( _
                startedAt, payloadReadyAt))) & _
        ";launchMs=" & _
            CStr(CLng(1000# * VTFastLaunchElapsedSeconds( _
                payloadReadyAt, launchFinishedAt))) & _
        ";totalMs=" & _
            CStr(CLng(1000# * VTFastLaunchElapsedSeconds( _
                startedAt, launchFinishedAt)))
    VTTryWriteAndLaunchSessionDirect = True
    Exit Function

FastPathFailed:
    timingDetail = "directError=" & CStr(Err.Number)
    On Error Resume Next
    If handle <> 0 Then Close #handle
    If Len(temporaryPath) > 0 Then Kill temporaryPath
    If Len(requestPath) > 0 Then Kill requestPath
    Err.Clear
    On Error GoTo 0
#End If
End Function

Private Function VTFastOpenResidentReady(ByVal normalizedHost As String) As Boolean
#If Mac Then
    Const MAX_READY_AGE_SECONDS As Long = 3
    Const FUTURE_TOLERANCE_SECONDS As Long = 2
    Dim markerPath As String
    Dim markerAgeSeconds As Long

    On Error GoTo NotReady
    markerPath = VTFastOpenInboxRoot(normalizedHost) & "/resident-ready"
    If Len(Dir$(markerPath, vbNormal)) = 0 Then GoTo NotReady
    markerAgeSeconds = DateDiff("s", FileDateTime(markerPath), Now)
    If markerAgeSeconds < -FUTURE_TOLERANCE_SECONDS Or _
       markerAgeSeconds > MAX_READY_AGE_SECONDS Then GoTo NotReady
    VTFastOpenResidentReady = True
    Exit Function

NotReady:
    Err.Clear
#End If
End Function

Private Function VTFastOpenInboxRoot(ByVal normalizedHost As String) As String
#If Mac Then
    Dim sandboxHome As String
    Dim expectedSuffix As String

    sandboxHome = Environ$("HOME")
    Select Case normalizedHost
        Case "word": expectedSuffix = VT_WORD_SANDBOX_HOME_SUFFIX
        Case "powerpoint": expectedSuffix = VT_POWERPOINT_SANDBOX_HOME_SUFFIX
        Case Else
            Err.Raise vbObjectError + 7323, "VisualTeX", _
                "VisualTeX fast-open host is invalid."
    End Select
    If Len(sandboxHome) <= Len(expectedSuffix) Or _
       Right$(sandboxHome, Len(expectedSuffix)) <> expectedSuffix Then
        Err.Raise vbObjectError + 7324, "VisualTeX", _
            "VisualTeX fast-open is available only inside the Office sandbox."
    End If
    VTFastOpenInboxRoot = _
        sandboxHome & "/Library/Application Support/VisualTeX/FastOpen/" & normalizedHost
#End If
End Function

Private Sub VTEnsureFastOpenInboxDirectory(ByVal inboxRoot As String)
#If Mac Then
    Dim applicationSupport As String
    Dim visualTeXRoot As String
    Dim fastOpenRoot As String

    applicationSupport = VTParentDirectory(VTParentDirectory(VTParentDirectory(inboxRoot)))
    visualTeXRoot = applicationSupport & "/VisualTeX"
    fastOpenRoot = visualTeXRoot & "/FastOpen"
    If Len(Dir$(applicationSupport, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 7325, "VisualTeX", _
            "The Office Application Support directory is unavailable."
    End If
    If Len(Dir$(visualTeXRoot, vbDirectory)) = 0 Then MkDir visualTeXRoot
    If Len(Dir$(fastOpenRoot, vbDirectory)) = 0 Then MkDir fastOpenRoot
    If Len(Dir$(inboxRoot, vbDirectory)) = 0 Then MkDir inboxRoot
#End If
End Sub

Private Function VTFastOpenRequestClaimed( _
    ByVal requestPath As String, _
    ByVal startedAt As Single) As Boolean

#If Mac Then
    Const MAX_CLAIM_WAIT_SECONDS As Double = 0.2
    Dim observedAt As Single

    Do
        If Len(Dir$(requestPath, vbNormal)) = 0 Then
            VTFastOpenRequestClaimed = True
            Exit Function
        End If
        DoEvents
        observedAt = Timer
    Loop While VTFastLaunchElapsedSeconds( _
        startedAt, observedAt) < MAX_CLAIM_WAIT_SECONDS
#End If
End Function

' The numbered-equation package of a manifest (the rows ConvertOmmlBatch of the helper script takes), asked of the resident app
' through the sandbox inbox instead of the helper script: AppleScriptTask costs about half a second a call from Word
' and the script starts a process of the app for it (measured), the resident answers in well under 0.1 s. The request
' <session>.number holds the manifest (Word cannot write into the Session folder, which is in Application Scripts); the
' resident builds the package and answers with <session>.number-done holding the result manifest. "" when the resident is
' not ready, does not answer in time, or fails (the caller then uses ConvertOmmlBatch of the helper script). trace says how far
' it got (for the performance trace).
Public Function VTTryFastConvertOmmlBatch( _
    ByVal sessionId As String, ByVal manifest As String, Optional ByRef trace As String) As String
#If Mac Then
    ' The resident takes a request out of the inbox within a poll interval (25 ms) when it can answer; one it has taken is
    ' answered however long the package takes (a numbering of the whole document is many equations), so the add-in waits for
    ' that answer and does not start the helper script's second conversion beside it.
    Const MAX_CLAIM_WAIT_SECONDS As Double = 0.4
    Const MAX_ANSWER_WAIT_SECONDS As Double = 20#
    ' The first part of the wait does not hand control to Word (DoEvents): its idle work (layout of a long document) took
    ' 0.1 s of a 0.02 s answer. A longer wait does, so that Word stays alive while a large package is built.
    Const SPIN_SECONDS As Double = 0.25
    Dim elapsedSeconds As Double
    Dim inboxRoot As String
    Dim requestPath As String
    Dim temporaryPath As String
    Dim answerPath As String
    Dim startedAt As Single
    Dim handle As Integer
    Dim answer As String
    Dim heartbeat As String
    Dim requestBytes() As Byte

    trace = "start"
    If Not VTIsCanonicalUuid(sessionId) Or Len(manifest) = 0 Then Exit Function
    On Error GoTo NotAvailable
    If Not VTFastOpenResidentReady("word") Then
        trace = trace & ";resident-not-ready"
        Exit Function
    End If
    trace = trace & ";ready"
    inboxRoot = VTFastOpenInboxRoot("word")
    ' only a resident that answers number requests (its heartbeat says so); an older one would leave this waiting
    heartbeat = VTReadInboxText(inboxRoot & "/resident-ready", 4096)
    trace = trace & ";heartbeat=" & CStr(Len(heartbeat))
    If InStr(1, heartbeat, "numbers=1", vbBinaryCompare) = 0 Then
        trace = trace & ";no-numbers"
        Exit Function
    End If
    requestPath = inboxRoot & "/" & sessionId & ".number"
    temporaryPath = inboxRoot & "/." & sessionId & ".number.tmp"
    answerPath = inboxRoot & "/" & sessionId & ".number-done"
    VTDiscardInboxFile answerPath
    VTDiscardInboxFile requestPath
    VTDiscardInboxFile temporaryPath

    requestBytes = VTUtf8Encode(manifest)
    handle = FreeFile
    Open temporaryPath For Binary Access Write As #handle
    Put #handle, , requestBytes
    Close #handle
    handle = 0
    trace = trace & ";written"
    Name temporaryPath As requestPath
    trace = trace & ";published"
    startedAt = Timer
    Do While Len(Dir$(requestPath, vbNormal)) > 0
        If Len(Dir$(answerPath, vbNormal)) > 0 Then Exit Do
        If VTFastLaunchElapsedSeconds(startedAt, Timer) > MAX_CLAIM_WAIT_SECONDS Then
            trace = trace & ";unclaimed"
            GoTo NotAvailable
        End If
    Loop
    trace = trace & ";claimed-after-ms=" & CStr(CLng(1000# * VTFastLaunchElapsedSeconds(startedAt, Timer)))
    Do While Len(Dir$(answerPath, vbNormal)) = 0
        elapsedSeconds = VTFastLaunchElapsedSeconds(startedAt, Timer)
        If elapsedSeconds > SPIN_SECONDS Then DoEvents
        If elapsedSeconds > MAX_ANSWER_WAIT_SECONDS Then
            trace = trace & ";timeout"
            GoTo NotAvailable
        End If
    Loop
    trace = trace & ";answered-after-ms=" & CStr(CLng(1000# * VTFastLaunchElapsedSeconds(startedAt, Timer)))
    answer = VTReadInboxText(answerPath, 65536)
    VTDiscardInboxFile answerPath
    If InStr(1, answer, "operation=nativeNumber", vbBinaryCompare) > 0 Then
        VTTryFastConvertOmmlBatch = answer
        trace = trace & ";ok"
    Else
        trace = trace & ";answer=" & Left$(Replace$(Replace$(answer, vbCr, " "), vbLf, " "), 120)
    End If
    Exit Function

NotAvailable:
    If Err.Number <> 0 Then trace = trace & ";error=" & CStr(Err.Number) & " " & Err.Description
    ' A request nobody answered must not be answered later: the helper script's package would race it.
    If handle <> 0 Then VTCloseFile handle
    VTDiscardInboxFile requestPath
    VTDiscardInboxFile temporaryPath
    VTDiscardInboxFile answerPath
    Err.Clear
#End If
End Function

' Update Equation Format's font check: which of a document's picture formulas the app draws in other fonts than it
' draws in now (src-tauri office/formula_fonts.rs reads their metadata, which VBA cannot inflate). The request
' <session>.fonts in the host's sandbox inbox holds "imageCount" and "image<n>Metadata" rows; the resident answers
' <session>.fonts-done with "stale" (the numbers of those drawn in other fonts) and "fontsEn" (the fonts, Base64URL).
' Only a resident whose heartbeat says fonts=1 knows the fonts; when none runs, the app is started in the background
' first (it publishes them once its windows load). "" when VisualTeX cannot answer.
Public Function VTTryResidentFontCheck( _
    ByVal hostName As String, ByVal sessionId As String, ByVal manifest As String) As String
#If Mac Then
    Const MAX_START_WAIT_SECONDS As Double = 25#
    Const MAX_CLAIM_WAIT_SECONDS As Double = 3#
    Const MAX_ANSWER_WAIT_SECONDS As Double = 60#
    Dim normalizedHost As String
    Dim inboxRoot As String
    Dim requestPath As String
    Dim temporaryPath As String
    Dim answerPath As String
    Dim startedAt As Single
    Dim handle As Integer
    Dim answer As String
    Dim requestBytes() As Byte

    If Not VTIsCanonicalUuid(sessionId) Or Len(manifest) = 0 Then Exit Function
    On Error GoTo NotAvailable
    normalizedHost = VTCanonicalOfficeHost(hostName)
    If Not VTResidentKnowsFormulaFonts(normalizedHost) Then
        VTTryPrewarmApplication normalizedHost
        startedAt = Timer
        Do Until VTResidentKnowsFormulaFonts(normalizedHost)
            If VTFastLaunchElapsedSeconds(startedAt, Timer) > MAX_START_WAIT_SECONDS Then GoTo NotAvailable
            DoEvents
        Loop
    End If
    inboxRoot = VTFastOpenInboxRoot(normalizedHost)
    requestPath = inboxRoot & "/" & sessionId & ".fonts"
    temporaryPath = inboxRoot & "/." & sessionId & ".fonts.tmp"
    answerPath = inboxRoot & "/" & sessionId & ".fonts-done"
    VTDiscardInboxFile answerPath
    VTDiscardInboxFile requestPath
    VTDiscardInboxFile temporaryPath
    requestBytes = VTUtf8Encode(manifest)
    handle = FreeFile
    Open temporaryPath For Binary Access Write As #handle
    Put #handle, , requestBytes
    Close #handle
    handle = 0
    Name temporaryPath As requestPath
    startedAt = Timer
    Do While Len(Dir$(requestPath, vbNormal)) > 0
        If Len(Dir$(answerPath, vbNormal)) > 0 Then Exit Do
        If VTFastLaunchElapsedSeconds(startedAt, Timer) > MAX_CLAIM_WAIT_SECONDS Then GoTo NotAvailable
        DoEvents
    Loop
    Do While Len(Dir$(answerPath, vbNormal)) = 0
        If VTFastLaunchElapsedSeconds(startedAt, Timer) > MAX_ANSWER_WAIT_SECONDS Then GoTo NotAvailable
        DoEvents
    Loop
    answer = VTReadInboxText(answerPath, 1048576)
    VTDiscardInboxFile answerPath
    If InStr(1, answer, "operation=fontCheck", vbBinaryCompare) > 0 Then VTTryResidentFontCheck = answer
    Exit Function

NotAvailable:
    ' A request nobody answered must not be answered later.
    If handle <> 0 Then VTCloseFile handle
    VTDiscardInboxFile requestPath
    VTDiscardInboxFile temporaryPath
    VTDiscardInboxFile answerPath
    Err.Clear
#End If
End Function

' One value of a resident's key=value answer ("" when it has none).
Public Function VTResidentAnswerValue(ByVal answer As String, ByVal key As String) As String
    Dim rows() As String
    Dim index As Long

    rows = Split(Replace$(answer, vbCr, ""), vbLf)
    For index = LBound(rows) To UBound(rows)
        If Left$(rows(index), Len(key) + 1) = key & "=" Then
            VTResidentAnswerValue = Mid$(rows(index), Len(key) + 2)
            Exit Function
        End If
    Next index
End Function

' The resident runs, is fresh and knows the fonts formulas are drawn in now (its heartbeat says fonts=1).
Private Function VTResidentKnowsFormulaFonts(ByVal normalizedHost As String) As Boolean
#If Mac Then
    On Error GoTo NotReady
    If Not VTFastOpenResidentReady(normalizedHost) Then Exit Function
    VTResidentKnowsFormulaFonts = InStr(1, _
        VTReadInboxText(VTFastOpenInboxRoot(normalizedHost) & "/resident-ready", 4096), _
        "fonts=1", vbBinaryCompare) > 0
    Exit Function

NotReady:
    Err.Clear
#End If
End Function

' Starts the resident in the background; a failure only means the check gets no answer.
Private Sub VTTryPrewarmApplication(ByVal normalizedHost As String)
    On Error Resume Next
    VTPrewarmApplication normalizedHost
    Err.Clear
End Sub

' Removes a file of the sandbox inbox if there is one. The empty path of a step not reached yet is not killed: Kill of "" fails
' (error 53 after more than a second in Word for Mac) and, inside an error handler, the failure goes on to the caller as an
' error of its own, which is how a failed resident request once took the whole insertion down.
Private Sub VTDiscardInboxFile(ByVal filePath As String)
#If Mac Then
    If Len(filePath) = 0 Then Exit Sub
    On Error Resume Next
    Kill filePath
    Err.Clear
#End If
End Sub

Private Sub VTCloseFile(ByVal handle As Integer)
    On Error Resume Next
    Close #handle
    Err.Clear
End Sub

' A small file of the sandbox inbox, read as UTF-8 ("" when it cannot be read). VTReadText reads only inside the Session
' folder's tree and would reject these paths.
Private Function VTReadInboxText( _
    ByVal sourcePath As String, ByVal maximumBytes As Long) As String

#If Mac Then
    Dim handle As Integer
    Dim byteCount As Long
    Dim bytes() As Byte

    On Error GoTo ReadFailed
    handle = FreeFile
    Open sourcePath For Binary Access Read As #handle
    byteCount = LOF(handle)
    If byteCount > 0 And byteCount <= maximumBytes Then
        ReDim bytes(0 To byteCount - 1)
        Get #handle, , bytes
        VTReadInboxText = VTUtf8Decode(bytes)
    End If
    Close #handle
    Exit Function

ReadFailed:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    Err.Clear
#End If
End Function

Private Function VTFastLaunchElapsedSeconds( _
    ByVal startedAt As Single, _
    ByVal finishedAt As Single) As Double

    If finishedAt < startedAt Then finishedAt = finishedAt + 86400!
    VTFastLaunchElapsedSeconds = CDbl(finishedAt) - CDbl(startedAt)
End Function

Private Function VTCanonicalOfficeHost(ByVal hostName As String) As String
    Select Case LCase$(Trim$(hostName))
        Case "word", "powerpoint"
            VTCanonicalOfficeHost = LCase$(Trim$(hostName))
        Case Else
            Err.Raise vbObjectError + 7301, "VisualTeX", "Invalid VisualTeX Office host."
    End Select
End Function

Private Function VTOfficeLauncherScriptName(ByVal hostName As String) As String
    Select Case VTCanonicalOfficeHost(hostName)
        Case "word": VTOfficeLauncherScriptName = "VisualTeXWord.scpt"
        Case "powerpoint": VTOfficeLauncherScriptName = "VisualTeXPowerPoint.scpt"
    End Select
End Function

Private Function VTAppleScriptErrorMessage(ByVal response As String) As String
    Dim fields() As String
    If Left$(response, 6) = "error|" Then
        fields = Split(response, "|")
        If UBound(fields) >= 2 Then
            VTAppleScriptErrorMessage = fields(2)
            Exit Function
        End If
    End If
    VTAppleScriptErrorMessage = "VisualTeX could not be opened."
End Function
