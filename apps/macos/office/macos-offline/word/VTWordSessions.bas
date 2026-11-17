Attribute VB_Name = "VTWordSessions"
Option Explicit
Option Private Module

' Formula sessions: creating and editing a formula through the app, committing or cancelling its result, the
' pending and edit bookmarks, and the paragraphs and caret around an inserted formula.

Private Const VT_WORD_BOOKMARK_PREFIX As String = "VT_Pending_"
Private Const VT_WORD_EDIT_BOOKMARK_PREFIX As String = "VT_E_"
Private Const VT_WORD_DOUBLE_CLICK_TRACE_ENABLED As Boolean = False
Private Const VT_WORD_DOUBLE_CLICK_TRACE_SENTINEL As String = _
    "/word-double-click-trace.enabled"
Private VT_WORD_DOUBLE_CLICK_TRACE_INITIALIZED As Boolean
Private VT_WORD_LAST_NATIVE_EDIT_FORMULA_ID As String
Private VT_WORD_LAST_NATIVE_EDIT_AT As Single

Public Function VTTryVisualTeXMetadataShapeAtDoubleClick( _
    ByVal selected As Selection, _
    ByRef resolvedShape As InlineShape) As Boolean

    Dim candidate As InlineShape
    Dim shapeRange As Range
    Dim selectionStart As Long
    Dim selectionEnd As Long

    Set resolvedShape = Nothing
    If selected Is Nothing Then Exit Function
    On Error GoTo NotMatched
    If Not VTTryVisualTeXMetadataShapeAtSelection( _
       selected, candidate) Then Exit Function

    Set shapeRange = candidate.Range.Duplicate
    selectionStart = selected.Range.Start
    selectionEnd = selected.Range.End
    If selectionStart > shapeRange.End + 1 Or _
       selectionEnd < shapeRange.Start - 1 Then Exit Function

    Set resolvedShape = candidate
    VTTryVisualTeXMetadataShapeAtDoubleClick = True
    Exit Function

NotMatched:
    Set resolvedShape = Nothing
    VTTryVisualTeXMetadataShapeAtDoubleClick = False
End Function

Private Function VTNativeEditDispatchDebounced( _
    ByVal formulaId As String) As Boolean

    Dim elapsedSeconds As Double

    If Len(VT_WORD_LAST_NATIVE_EDIT_FORMULA_ID) = 0 Or _
       StrComp(VT_WORD_LAST_NATIVE_EDIT_FORMULA_ID, formulaId, _
           vbBinaryCompare) <> 0 Then Exit Function
    elapsedSeconds = VTTimerElapsedSeconds( _
        VT_WORD_LAST_NATIVE_EDIT_AT, Timer)
    VTNativeEditDispatchDebounced = _
        (elapsedSeconds >= 0# And _
         elapsedSeconds <= VT_WORD_IMAGE_EDIT_DEBOUNCE_SECONDS)
End Function

Public Sub VTWordEditNativeBookmark(ByVal nativeBookmark As Bookmark)
    VTWordOpenNativeSession nativeBookmark
End Sub

Public Function VTWordOpenCopiedNativeSession( _
    ByVal selectedRange As Range, _
    Optional ByVal keepNativeEquation As Boolean = True, _
    Optional ByVal operationName As String = "formula", _
    Optional ByRef launchedSessionId As Variant) As Boolean

    Dim nativeMath As OMath
    Dim targetDocument As Document
    Dim formulaId As String
    Dim sessionId As String
    Dim requestJson As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim fontSizePt As Double
    Dim sourceKey As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim formulaLetterFont As String
    Dim wordMathFontName As String
    Dim nativeSnapshot As String
    Dim snapshotRequest As String
    If selectedRange Is Nothing Then Exit Function
    VTWordPerformanceStartLocal "native-formula-open"
    If Not VTTryResolveSingleNativeMath(selectedRange, nativeMath) Then Exit Function
    VTWordPerformanceMark "native-target-resolved"
    VTRequireWritableWordDocument
    Set targetDocument = nativeMath.Range.Document
    sourceKey = VTWordDocumentIdentityForDocument(targetDocument) & ":" & _
        CStr(nativeMath.Range.Start) & ":" & CStr(nativeMath.Range.End)
    If operationName = "formula" And VTNativeEditDispatchDebounced(sourceKey) Then
        VTWordOpenCopiedNativeSession = True
        Exit Function
    End If
    On Error GoTo OpenFailed
    formulaId = VTNewUuidV4()
    sessionId = VTNewUuidV4()
    VTPrepareLegacyNativeMathForEdit nativeMath
    numbered = Not VTPureNativeNumberField(nativeMath.Range) Is Nothing
    displayMode = "inline"
    If nativeMath.Type = wdOMathDisplay Or numbered Then displayMode = "block"
    fontSizePt = VTPreferredWordFormulaFontSize(nativeMath.Range)
    If operationName = "nativeToImage" Then _
        fontSizePt = VTConversionFontSize(displayMode, fontSizePt)
    wordMathFontName = VTWordMathFontNameFromRange(nativeMath.Range)
    formulaLetterFont = _
        VTVisualTeXLetterFontFromWordFontName(wordMathFontName)
    VTWordPerformanceMark "native-properties-read"
    ' VTWriteTextAtomic creates the session parent. A separate mkdir bridge
    ' costs an AppleScriptTask round trip before every native edit.
    ' The actual native tree is authoritative. Session-only anchors are removed
    ' on apply/cancel; no persistent formula identity or LaTeX cache is needed.
    ' The editor reads the live equation tree, not the styles/theme/font-table
    ' parts Word adds to even a one-equation Flat OPC export. Reuse the lossless
    ' document-part extraction used by native formula restore; retaining its
    ' namespace declarations also preserves every OMath/SEQ element.
    nativeSnapshot = VTFormulaRestoreNativeXml(nativeMath.Range)
    VTWordPerformanceNote "snapshotCharacters", CStr(Len(nativeSnapshot))
    VTWordPerformanceMark "native-snapshot-read"
    requestJson = VTRequestJson(sessionId, VT_WORD_HOST, "edit", formulaId, _
        displayMode, numbered, VTWordDocumentIdentityForDocument(targetDocument), _
        VTWordEditBookmarkName(sessionId), "", "", "", keepNativeEquation, _
        fontSizePt, 0#, 0#, operationName, False, _
        formulaLetterFont, wordMathFontName)
    requestJson = VTV2WithDocumentMathFont(requestJson, targetDocument)
    ' Publish ordinary native edits as one atomic sandbox-inbox request. A
    ' separate snapshot write adds a slow AppleScriptTask before every click.
    ' Large equations keep the established file transport instead of exceeding
    ' the inbox limit. Bound the candidate before JSON escaping a large tree.
    If Len(nativeSnapshot) <= 32768 Then
        snapshotRequest = Left$(requestJson, Len(requestJson) - 1) & _
            ",""nativeEditXml"":" & VTJsonString(nativeSnapshot) & "}"
        If VTUtf8ByteLength(snapshotRequest) <= 65536 Then _
            requestJson = snapshotRequest
    End If
    If requestJson = snapshotRequest And Len(snapshotRequest) > 0 Then
        VTWordPerformanceNote "snapshotTransport", "atomic-edit-request"
    Else
        VTWriteTextAtomic VTSessionDirectory(sessionId) & "/native-edit-original.omml", nativeSnapshot
        VTWordPerformanceNote "snapshotTransport", "session-file"
    End If
    VTWordPerformanceMark "native-snapshot-prepared"
    VTAddWordEditRangeBookmark nativeMath.Range, sessionId
    ' Opening the editor does not change the equation. Create the rollback
    ' document only when Apply actually replaces the current native tree.
    VTWriteAndLaunchSession VT_WORD_HOST, sessionId, requestJson
    VTWordPerformanceMark "native-editor-dispatched"
    VTWordPerformanceFlush sessionId
    If operationName = "formula" Then
        VT_WORD_LAST_NATIVE_EDIT_FORMULA_ID = sourceKey
        VT_WORD_LAST_NATIVE_EDIT_AT = Timer
    End If
    If Not IsMissing(launchedSessionId) Then launchedSessionId = sessionId
    VTWordOpenCopiedNativeSession = True
    Exit Function
OpenFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    VTWordPerformanceFlush sessionId, errorNumber, errorDescription
    On Error Resume Next
    VTDeleteWordEditBookmark targetDocument, sessionId
    VTClosePreparedNativeRollbackDocument
    On Error GoTo 0
    Err.Raise errorNumber, "VisualTeX native Word edit", errorDescription
End Function

Public Sub VTWordOpenNativeSession( _
    ByVal nativeBookmark As Bookmark, _
    Optional ByVal keepNativeEquation As Boolean = True, _
    Optional ByVal operationName As String = "formula")

    Dim documentObject As Document
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim sessionId As String
    Dim requestJson As String
    Dim nativeMath As OMath
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim storedFontSizePt As Double
    Dim observedWordFontSizePt As Double
    Dim bookmarkTrace As String
    Dim launchTiming As String
    Dim debounceArmed As Boolean
    Dim openErrorNumber As Long
    Dim openErrorDescription As String
    Dim formulaLetterFont As String
    Dim wordMathFontName As String

    bookmarkTrace = "nothing"
    If Not nativeBookmark Is Nothing Then bookmarkTrace = nativeBookmark.Name
    VTTraceWordDoubleClick _
        "edit-native-enter", Selection, _
        "bookmark=" & bookmarkTrace
    On Error GoTo OpenFailed

    If nativeBookmark Is Nothing Then
        Err.Raise vbObjectError + 7452, "VisualTeX", "Select one VisualTeX native Word equation."
    End If
    If Not VTTryFormulaIdFromNativeBookmark(nativeBookmark.Name, formulaId) Then
        Err.Raise vbObjectError + 7453, "VisualTeX", "The selected VisualTeX native equation bookmark is invalid."
    End If
    Set documentObject = nativeBookmark.Range.Document
    Set nativeMath = VTNativeMathForBookmark(nativeBookmark)
    If nativeMath Is Nothing Then
        Err.Raise vbObjectError + 7454, "VisualTeX", "The selected VisualTeX bookmark no longer contains exactly one native equation."
    End If
    If Not VTTryReadWordMetadataPayload(documentObject, formulaId, encodedMetadata) Then
        Err.Raise vbObjectError + 7455, "VisualTeX", "The selected native equation is missing VisualTeX edit metadata."
    End If
    If Not VTTryReadWordFormulaFormat( _
        documentObject, formulaId, displayMode, numbered) Then
            Err.Raise vbObjectError + 7456, "VisualTeX", "The selected native equation is missing its VisualTeX display format."
    End If
    VTTraceWordDoubleClick _
        "edit-native-resolved", Selection, _
        "formulaId=" & formulaId & _
        " displayMode=" & displayMode & _
        " numbered=" & CStr(numbered)
    If operationName = "formula" And _
       VTNativeEditDispatchDebounced(formulaId) Then
        VTTraceWordDoubleClick _
            "edit-native-debounced", Selection, "formulaId=" & formulaId
        Exit Sub
    End If
    If operationName = "formula" Then
        VT_WORD_LAST_NATIVE_EDIT_FORMULA_ID = formulaId
        VT_WORD_LAST_NATIVE_EDIT_AT = Timer
        debounceArmed = True
    End If

    ' OMath.Range.Font.Size can report the surrounding paragraph's default
    ' size (commonly 11 pt) even when the visible imported equation and its
    ' identity bookmark retain the requested point size. The bookmark Range is
    ' the same range used by Word's font-size UI and is therefore authoritative.
    fontSizePt = nativeBookmark.Range.Font.Size
    On Error Resume Next
    If VTTryReadWordImageScaleState( _
       documentObject, formulaId, storedFontSizePt, _
       referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
       observedWordFontSizePt) Then
        If Not VTValidWordFormulaFontSize(fontSizePt) Then
            fontSizePt = storedFontSizePt
        End If
    End If
    Err.Clear
    On Error GoTo 0
    If Not VTValidWordFormulaFontSize(fontSizePt) Then
        fontSizePt = nativeMath.Range.Font.Size
    End If
    If Not VTValidWordFormulaFontSize(fontSizePt) Then
        fontSizePt = VTPreferredWordFormulaFontSize(nativeBookmark.Range)
    End If
    If operationName = "nativeToImage" Then _
        fontSizePt = VTConversionFontSize(displayMode, fontSizePt)
    wordMathFontName = VTWordMathFontNameFromRange(nativeMath.Range)
    formulaLetterFont = _
        VTVisualTeXLetterFontFromWordFontName(wordMathFontName)

    sessionId = VTNewUuidV4()
    requestJson = VTRequestJson( _
        sessionId, _
        VT_WORD_HOST, _
        "edit", _
        formulaId, _
        displayMode, _
        numbered, _
        VTWordDocumentIdentity(), _
        nativeBookmark.Name, _
        encodedMetadata, _
        "", _
        "", _
        keepNativeEquation, _
        fontSizePt, _
        referenceWidthPt, _
        referenceHeightPt, _
        operationName, _
        False, _
        formulaLetterFont, _
        wordMathFontName)
    requestJson = VTV2WithDocumentMathFont(requestJson, documentObject)
    ' Creating Word's hidden rollback document costs roughly 0.15 s on macOS.
    ' Prepare only its empty container while the editor opens; Apply still
    ' captures and validates the current OMath immediately before replacement.
    VTPrepareNativeRollbackDocument documentObject
    launchTiming = _
        VTWriteAndLaunchSession(VT_WORD_HOST, sessionId, requestJson)
    VTTraceWordDoubleClick _
        "edit-native-editor-launched", Selection, _
        "sessionId=" & sessionId & " " & launchTiming
    Exit Sub

OpenFailed:
    openErrorNumber = Err.Number
    openErrorDescription = Err.Description
    If debounceArmed And _
       StrComp(VT_WORD_LAST_NATIVE_EDIT_FORMULA_ID, formulaId, _
           vbBinaryCompare) = 0 Then
        VT_WORD_LAST_NATIVE_EDIT_FORMULA_ID = ""
        VT_WORD_LAST_NATIVE_EDIT_AT = 0!
    End If
    Err.Raise openErrorNumber, "VisualTeX Word native edit", _
        openErrorDescription
End Sub

Public Function VTPrepareWordCreateInsertionRange( _
    ByVal requestedRange As Range, _
    ByVal displayMode As String, _
    Optional ByVal nativeEquation As Boolean = False) As Range

    Dim documentObject As Document
    Dim insertionRange As Range
    Dim paragraphRange As Range
    Dim beforeRange As Range
    Dim afterRange As Range
    Dim targetParagraph As Range
    Dim existingLayout As Table
    Dim insertionStart As Long
    Dim contentEnd As Long
    Dim targetStart As Long
    Dim beforeOccupied As Boolean
    Dim afterOccupied As Boolean
    Dim characterValue As String
    Dim typographyAlignment As Long

    If requestedRange Is Nothing Then
        Err.Raise vbObjectError + 7551, "VisualTeX", _
            "The Word formula insertion Range is missing."
    End If
    If displayMode <> "inline" And displayMode <> "block" Then
        Err.Raise vbObjectError + 7451, "VisualTeX", _
            "The VisualTeX Word display mode is invalid."
    End If

    typographyAlignment = requestedRange.Paragraphs(1).Range.ParagraphFormat.BaseLineAlignment
    Set documentObject = requestedRange.Document
    Set insertionRange = requestedRange.Duplicate
    insertionRange.Collapse wdCollapseStart
    If displayMode = "inline" Then
        Set VTPrepareWordCreateInsertionRange = insertionRange
        Exit Function
    End If

    ' A display formula must never be created inside an existing VisualTeX
    ' numbered-layout table. Side cells can look like empty document space, but
    ' they are part of the existing formula scaffold and cannot safely host a
    ' second display formula.
    If VTRangeIsWithinTable(insertionRange) Then
        Set existingLayout = VTNumberedDisplayTableNearRange(insertionRange)
        If Not existingLayout Is Nothing Then
            If Len(VTNumberBookmarkNameForTable(existingLayout)) > 0 Then
                Err.Raise vbObjectError + 7551, "VisualTeX", _
                    "Move the caret outside the existing VisualTeX display formula before inserting another display formula."
            End If
        End If
    End If

    ' A display formula must never reuse a body paragraph that already contains
    ' text, a field, an inline formula, an image, or native OMath. The old path
    ' inserted a one-pixel placeholder at the caret and later cleared that entire
    ' paragraph while building the numbered table, erasing surrounding content.
    ' Split the paragraph first and return a dedicated empty line.
    Set paragraphRange = insertionRange.Paragraphs(1).Range.Duplicate
    insertionStart = insertionRange.Start
    contentEnd = paragraphRange.End
    Do While contentEnd > paragraphRange.Start
        characterValue = documentObject.Range( _
            Start:=contentEnd - 1, End:=contentEnd).Text
        If characterValue = vbCr Or characterValue = Chr$(7) Then
            contentEnd = contentEnd - 1
        Else
            Exit Do
        End If
    Loop
    If insertionStart < paragraphRange.Start Then
        insertionStart = paragraphRange.Start
    End If
    If insertionStart > contentEnd Then insertionStart = contentEnd

    Set beforeRange = documentObject.Range( _
        Start:=paragraphRange.Start, End:=insertionStart)
    Set afterRange = documentObject.Range( _
        Start:=insertionStart, End:=contentEnd)
    beforeOccupied = _
        beforeRange.Fields.Count > 0 Or _
        beforeRange.InlineShapes.Count > 0 Or _
        beforeRange.OMaths.Count > 0 Or _
        VTWordRangeHasMeaningfulText(beforeRange)
    afterOccupied = _
        afterRange.Fields.Count > 0 Or _
        afterRange.InlineShapes.Count > 0 Or _
        afterRange.OMaths.Count > 0 Or _
        VTWordRangeHasMeaningfulText(afterRange)

    targetStart = insertionStart
    If beforeOccupied And afterOccupied Then
        documentObject.Range( _
            Start:=insertionStart, End:=insertionStart).Text = vbCr & vbCr
        targetStart = insertionStart + 1
    ElseIf beforeOccupied Then
        documentObject.Range( _
            Start:=insertionStart, End:=insertionStart).Text = vbCr
        targetStart = insertionStart + 1
    ElseIf afterOccupied Then
        documentObject.Range( _
            Start:=insertionStart, End:=insertionStart).Text = vbCr
        targetStart = insertionStart
    End If

    Set insertionRange = documentObject.Range( _
        Start:=targetStart, End:=targetStart)
    Set targetParagraph = insertionRange.Paragraphs(1).Range.Duplicate
    If targetParagraph.Fields.Count <> 0 Or _
       targetParagraph.InlineShapes.Count <> 0 Or _
       targetParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(targetParagraph) Then
        ' Hidden Equation helper fields and stale paragraph scaffolds can look
        ' completely empty in Word. Insert an identified paragraph at the exact
        ' caret boundary instead of rejecting that visually empty position.
        Set insertionRange = VTCreateDedicatedPlainParagraphAt( _
            documentObject, targetStart)
    ElseIf nativeEquation And Not (beforeOccupied Or afterOccupied) Then
        ' Equation format v2: the user's empty paragraph is reused as is.
    Else
        VTNormalizePlainWordParagraph targetParagraph
    End If
    insertionRange.Paragraphs(1).Range.ParagraphFormat.BaseLineAlignment = typographyAlignment
    Set VTPrepareWordCreateInsertionRange = insertionRange
End Function

Private Function VTCreateDedicatedPlainParagraphAt( _
    ByVal documentObject As Document, _
    ByVal insertionStart As Long) As Range

    Dim insertionRange As Range
    Dim markerRange As Range
    Dim paragraphRange As Range
    Dim markerText As String

    If documentObject Is Nothing Then
        Err.Raise vbObjectError + 7551, "VisualTeX", _
            "The display insertion document is missing."
    End If
    If insertionStart < 0 Or insertionStart > documentObject.Content.End Then
        Err.Raise vbObjectError + 7551, "VisualTeX", _
            "The display insertion position is outside the Word document."
    End If

    markerText = "VisualTeXDisplay_" & _
        Replace$(VTNewUuidV4(), "-", "")
    Set insertionRange = documentObject.Range( _
        Start:=insertionStart, End:=insertionStart)
    insertionRange.Text = markerText & vbCr
    Set markerRange = documentObject.Range( _
        Start:=insertionStart, _
        End:=insertionStart + Len(markerText))
    If markerRange.Text <> markerText Then
        Err.Raise vbObjectError + 7551, "VisualTeX", _
            "Word did not create the dedicated display insertion paragraph."
    End If
    Set paragraphRange = markerRange.Paragraphs(1).Range.Duplicate
    If paragraphRange.Start <> insertionStart Or _
       paragraphRange.Fields.Count <> 0 Or _
       paragraphRange.InlineShapes.Count <> 0 Or _
       paragraphRange.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7551, "VisualTeX", _
            "Word attached the display insertion paragraph to existing content."
    End If

    markerRange.Delete
    Set paragraphRange = documentObject.Range( _
        Start:=insertionStart, End:=insertionStart).Paragraphs(1).Range.Duplicate
    If paragraphRange.Fields.Count <> 0 Or _
       paragraphRange.InlineShapes.Count <> 0 Or _
       paragraphRange.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(paragraphRange) Then
        Err.Raise vbObjectError + 7551, "VisualTeX", _
            "Word did not preserve an empty display insertion paragraph."
    End If
    VTNormalizePlainWordParagraph paragraphRange
    Set VTCreateDedicatedPlainParagraphAt = documentObject.Range( _
        Start:=insertionStart, End:=insertionStart)
End Function

Public Sub VTNormalizePlainWordParagraph(ByVal paragraphRange As Range)
    Dim normalSize As Single

    If paragraphRange Is Nothing Then Exit Sub
    On Error Resume Next
    normalSize = paragraphRange.Document.Styles(wdStyleNormal).Font.Size
    paragraphRange.Style = wdStyleNormal
    On Error GoTo 0
    If normalSize <= 0! Or normalSize > 72! Then normalSize = 12!
    With paragraphRange
        .Font.Position = 0
        .Font.Hidden = False
        .Font.Color = wdColorAutomatic
        .Font.Size = normalSize
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LeftIndent = 0!
            .RightIndent = 0!
            .FirstLineIndent = 0!
            .SpaceBefore = 0!
            .SpaceAfter = 0!
            .LineSpacingRule = wdLineSpaceSingle
            .KeepTogether = False
            .KeepWithNext = False
            .PageBreakBefore = False
            .TabStops.ClearAll
        End With
    End With
End Sub

Private Function VTEnsurePlainContinuationParagraph( _
    ByVal sourceParagraph As Range) As Range

    Dim documentObject As Document
    Dim insertionRange As Range
    Dim probeRange As Range
    Dim markerRange As Range
    Dim continuationParagraph As Range
    Dim continuationStart As Long
    Dim markerStart As Long
    Dim created As Boolean

    If sourceParagraph Is Nothing Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "The display continuation source paragraph is missing."
    End If
    Set documentObject = sourceParagraph.Document
    continuationStart = sourceParagraph.End

    ' A collapsed Range at a paragraph boundary can resolve back into the
    ' preceding helper paragraph on Word for Mac. Probe one actual character
    ' from the following paragraph instead, and reject any backward expansion.
    If continuationStart < documentObject.Content.End Then
        Set probeRange = documentObject.Range( _
            Start:=continuationStart, End:=continuationStart + 1)
        Set continuationParagraph = probeRange.Paragraphs(1).Range.Duplicate
        If continuationParagraph.Start < continuationStart Then
            Set continuationParagraph = Nothing
        End If
    End If

    If continuationParagraph Is Nothing Then
        created = True
    ElseIf continuationParagraph.Information(wdWithInTable) Or _
           continuationParagraph.Fields.Count > 0 Or _
           continuationParagraph.InlineShapes.Count > 0 Or _
           continuationParagraph.OMaths.Count > 0 Then
        created = True
    End If
    If created Then
        ' Equation format v2: created as Windows creates it.
        Set continuationParagraph = VTV2InsertParagraphAfter(sourceParagraph)
        If continuationParagraph Is Nothing Then
            markerStart = sourceParagraph.End - 1
            Set insertionRange = documentObject.Range( _
                Start:=markerStart, End:=markerStart)
            insertionRange.Text = vbCr & ChrW(8288)
            Set markerRange = documentObject.Range( _
                Start:=markerStart + 1, End:=markerStart + 2)
            Set continuationParagraph = _
                markerRange.Paragraphs(1).Range.Duplicate
            markerRange.Delete
            Set probeRange = documentObject.Range( _
                Start:=markerStart + 1, End:=markerStart + 2)
            Set continuationParagraph = probeRange.Paragraphs(1).Range.Duplicate
        End If
    End If

    If continuationParagraph.Start < continuationStart Or _
       continuationParagraph.Information(wdWithInTable) Or _
       continuationParagraph.Fields.Count <> 0 Or _
       continuationParagraph.InlineShapes.Count <> 0 Or _
       continuationParagraph.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "Word did not expose a plain paragraph after the display formula" & _
            " [source=" & CStr(sourceParagraph.Start) & "-" & _
                CStr(sourceParagraph.End) & _
            "; continuation=" & CStr(continuationParagraph.Start) & "-" & _
                CStr(continuationParagraph.End) & "]."
    End If
    ' Equation format v2: Windows leaves the user's following paragraph as it
    ' is (WordFormulaService caret placement after a display formula).
    If created Then
        VTV2ApplyBodyFormatToTypingParagraph sourceParagraph, continuationParagraph
    End If
    Set VTEnsurePlainContinuationParagraph = continuationParagraph
End Function

Public Function VTContinuationRangeAfterDisplayFormula( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String, _
    Optional ByVal resetTypingFont As Boolean = True) As Range

    Dim documentObject As Document
    Dim sourceParagraph As Range
    Dim continuationParagraph As Range
    Dim caretRange As Range
    Dim captionBookmarkName As String

    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "The display formula caret target is missing."
    End If
    Set documentObject = formulaRange.Document
    If formulaRange.Information(wdWithInTable) Then
        captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
        If Not documentObject.Bookmarks.Exists(captionBookmarkName) Then
            Err.Raise vbObjectError + 7553, "VisualTeX", _
                "The numbered display helper paragraph is missing."
        End If
        Set sourceParagraph = documentObject.Bookmarks( _
            captionBookmarkName).Range.Paragraphs(1).Range.Duplicate
    ElseIf formulaRange.OMaths.Count = 1 Or _
           formulaRange.InlineShapes.Count = 1 Then
        captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
        If documentObject.Bookmarks.Exists(captionBookmarkName) Then
            Set sourceParagraph = documentObject.Bookmarks( _
                captionBookmarkName).Range.Paragraphs(1).Range.Duplicate
            If Not VTHelperParagraphOwnsNativeEquationSequence( _
               sourceParagraph) Then
                Set sourceParagraph = _
                    VTWordParagraphContainingFormula(formulaRange)
            End If
        Else
            Set sourceParagraph = VTWordParagraphContainingFormula(formulaRange)
        End If
        If sourceParagraph Is Nothing Then
            Err.Raise vbObjectError + 7553, "VisualTeX", _
                "Word could not resolve the display formula continuation boundary."
        End If
    Else
        Set sourceParagraph = formulaRange.Paragraphs(1).Range.Duplicate
    End If

    Set continuationParagraph = _
        VTEnsurePlainContinuationParagraph(sourceParagraph)
    Set caretRange = continuationParagraph.Duplicate
    caretRange.Collapse wdCollapseStart
    If resetTypingFont Then
        caretRange.Font.Position = 0
        caretRange.Font.Hidden = False
        caretRange.Font.Color = wdColorAutomatic
    End If
    Set VTContinuationRangeAfterDisplayFormula = caretRange.Duplicate
End Function

Public Sub VTPlaceCaretAfterDisplayFormula( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String, _
    Optional ByVal resetTypingFont As Boolean = True)

    Dim caretRange As Range

    Set caretRange = VTContinuationRangeAfterDisplayFormula( _
        formulaRange, formulaId, resetTypingFont)
    caretRange.Select
End Sub

Public Sub VTWordCreate( _
    ByVal displayMode As String, _
    ByVal numbered As Boolean, _
    Optional ByVal nativeEquation As Boolean = False)
    Dim sessionId As String
    Dim formulaId As String
    Dim pendingMarker As String
    Dim insertionRange As Range
    Dim requestJson As String
    Dim requestedFontSizePt As Double
    Dim launchTiming As String
    Dim errorNumber As Long
    Dim errorDescription As String

    VTWordPerformanceStartLocal "formula-open"
    On Error GoTo Failed
    VTRequireWritableWordDocument
    VTWordPerformanceMark "document-ready"
    VTCleanupOrphanedNumberedDisplaySelection Selection.Range
    VTWordPerformanceMark "orphan-cleanup-complete"
    sessionId = VTNewUuidV4()
    formulaId = VTNewUuidV4()
    pendingMarker = VTPendingMarker(sessionId, formulaId)
    VTWordPerformanceMark "identities-ready"
    requestedFontSizePt = VTV2CurrentTypingFontSize()
    VTWordPerformanceMark "font-size-ready"
    Set insertionRange = VTPrepareWordCreateInsertionRange( _
        Selection.Range.Duplicate, displayMode, nativeEquation)
    VTWordPerformanceMark "insertion-range-ready"
    ' Keep create lightweight. The commit/cancel transaction already supports a
    ' pending Bookmark with no InlineShape and resolves the final insertion Range
    ' from that Bookmark. Avoiding a 1×1 PNG here removes Word's synchronous
    ' picture-import cost from the editor-open path.
    VTAddPendingBookmark insertionRange, sessionId
    VTTraceWordSession sessionId, "bookmark-anchor-created", pendingMarker
    Selection.SetRange Start:=insertionRange.Start, End:=insertionRange.Start
    VTWordPerformanceMark "bookmark-anchor-ready"

    requestJson = VTRequestJson( _
        sessionId, _
        VT_WORD_HOST, _
        "create", _
        formulaId, _
        displayMode, _
        numbered, _
        VTWordDocumentIdentity(), _
        pendingMarker, _
        "", _
        pendingMarker, _
        "", _
        nativeEquation, _
        requestedFontSizePt)
    requestJson = VTV2WithDocumentMathFont(requestJson, ActiveDocument)
    VTWordPerformanceMark "request-serialized"
    launchTiming = _
        VTWriteAndLaunchSession(VT_WORD_HOST, sessionId, requestJson)
    VTWordPerformanceNote "launcherTiming", launchTiming
    VTWordPerformanceMark "launcher-returned"
    VTTraceWordSession _
        sessionId, "request-written-and-launched", _
        pendingMarker & " " & launchTiming
    VTTraceWordSession sessionId, "editor-launched", pendingMarker
    VTWordPerformanceMark "formula-open-complete"
    VTWordPerformanceFlush sessionId
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    VTWordPerformanceFlush sessionId, errorNumber, errorDescription
    If Len(sessionId) > 0 Then
        VTDeletePendingBookmark ActiveDocument, sessionId
        VTDeleteSessionFiles sessionId
    End If
    On Error GoTo 0
    VTShowError "Word formula creation", errorNumber, errorDescription
End Sub

Public Sub VTCommitWordDispatch( _
    ByVal sessionId As String, ByVal dispatch As Object, _
    Optional ByVal regressionFailureAfterSourceRemoval As Boolean = False)
    Dim mode As String
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim nativeEquation As Boolean
    Dim nativeDisplayMode As String
    Dim imagePath As String
    Dim vectorDocumentPath As String
    Dim fallbackImagePath As String
    Dim metadata As String
    Dim latexBase64 As String
    Dim ommlBase64 As String
    Dim nativeDocumentPath As String
    Dim pendingMarker As String
    Dim sourceMarker As String
    Dim sourceDocumentId As String
    Dim wordMathFontName As String
    Dim targetDocument As Document
    Dim targetImage As InlineShape
    Dim pendingBookmark As Bookmark
    Dim sourceEditBookmark As Bookmark
    Dim targetFromPendingBookmark As Boolean
    Dim nativeTarget As Bookmark
    Dim targetRange As Range
    Dim originalNativeRange As Range
    Dim originalNativeMath As OMath
    Dim originalNativeStart As Long
    Dim originalNativeBookmarkName As String
    Dim originalNativeBookmarkDeleted As Boolean
    Dim originalNativeBackupDocument As Document
    Dim originalNativeBackupRange As Range
    Dim originalNativeTable As Table
    Dim nativeTableBackupReady As Boolean
    Dim nativeTableStart As Long
    Dim nativeTableFollowingParagraph As Range
    Dim nativeTableBackupSource As Range
    Dim nativeTableReplacementEnd As Range
    Dim restoredNativeTable As Table
    Dim restoredNativeNumber As Range
    Dim needsNativeImageParagraph As Boolean
    Dim nativeTargetReplaced As Boolean
    Dim joinedNative As Boolean
    Dim adjacentNative As OMath
    Dim nativeBookmarkSet As Boolean
    Dim targetIsNative As Boolean
    Dim nativeTargetFromEditBookmark As Boolean
    Dim nativeNumberBookmarks As Collection
    Dim committed As InlineShape
    Dim stagedCandidate As InlineShape
    Dim candidate As InlineShape
    Dim finalImageField As Field
    Dim rollbackMacroField As Field
    Dim insertionRange As Range
    Dim widthPoints As Double
    Dim heightPoints As Double
    Dim baselinePoints As Double
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim inkCenterYRatio As Double
    Dim observedWordFontSizePt As Double
    Dim formulaReference As String
    Dim captionText As String
    Dim insertedNumber As Range
    Dim numberLayoutRange As Range
    Dim numberCreated As Boolean
    Dim nativeEquationRange As Range
    Dim rollbackRange As Range
    Dim rollbackNativeMath As OMath
    Dim previousLatexBase64 As String
    Dim previousOmmlBase64 As String
    Dim previousMetadata As String
    Dim previousDisplayMode As String
    Dim previousNumbered As Boolean
    Dim previousFontSizePt As Double
    Dim previousReferenceWidthPt As Double
    Dim previousReferenceHeightPt As Double
    Dim previousReferenceBaselinePt As Double
    Dim previousInkCenterYRatio As Double
    Dim previousObservedWordFontSizePt As Double
    Dim hadPreviousLatexPayload As Boolean
    Dim hadPreviousOmmlPayload As Boolean
    Dim hadPreviousMetadataPayload As Boolean
    Dim hadPreviousFormat As Boolean
    Dim hadPreviousImageScale As Boolean
    Dim previousImageCanvas As String
    Dim hadPreviousImageCanvas As Boolean
    Dim imageCanvasCaptured As Boolean
    Dim hadPreviousInkCenter As Boolean
    Dim formulaStateStored As Boolean
    Dim pendingPlaceholderRemoved As Boolean
    Dim pendingPlaceholderStart As Long
    Dim committedImageStart As Long
    Dim nativeEquationStart As Long
    Dim restoredPlaceholder As InlineShape
    Dim targetFormulaId As String
    Dim forkCopiedFormula As Boolean
    Dim originalImageBackupDocument As Document
    Dim originalImageBackupRange As Range
    Dim originalImageStart As Long
    Dim originalImageRemoved As Boolean
    Dim sourceWasNumbered As Boolean
    Dim numberingChanged As Boolean
    Dim numberingRollbackRange As Range
    Dim imageNumberBookmarks As Collection
    Dim numberingPrefix As Range
    Dim rollbackPlaceRef As Field
    Dim transactionErrorNumber As Long
    Dim transactionErrorDescription As String
    Dim transactionStage As String
    Dim previousScreenUpdating As Boolean
    Dim screenUpdatingCaptured As Boolean
    Dim internalMutationStarted As Boolean
    Dim v2SequenceName As String
    Dim v2HeadingLevel As Long
    Dim v2Separator As String
    Dim v2LiteralPrefix As String
    Dim v2Automatic As Boolean
    Dim v2Restart As Long
    Dim v2Initial As Long
    Dim v2TransferEquation As Range
    Dim v2NumberBookmarks As Collection
    transactionStage = "validate-document"
    VTWordPerformanceMark "commit-enter"
    committedImageStart = -1
    VTRequireDispatchValue dispatch, "mode"
    VTRequireDispatchValue dispatch, "formulaId"
    VTRequireDispatchValue dispatch, "displayMode"
    VTRequireDispatchValue dispatch, "numbered"
    VTRequireDispatchValue dispatch, "nativeEquation"
    VTRequireDispatchValue dispatch, "metadata"
    VTRequireDispatchValue dispatch, "latexBase64"
    VTRequireDispatchValue dispatch, "ommlBase64"

    transactionStage = "read-dispatch"
    mode = CStr(dispatch("mode"))
    formulaId = CStr(dispatch("formulaId"))
    displayMode = CStr(dispatch("displayMode"))
    numbered = (CStr(dispatch("numbered")) = "1")
    nativeEquation = (CStr(dispatch("nativeEquation")) = "1")
    If nativeEquation Then
        VTRequireDispatchValue dispatch, "nativeDocumentPath"
    Else
        VTRequireDispatchValue dispatch, "imagePath"
        VTRequireDispatchValue dispatch, "vectorDocumentPath"
        VTRequireDispatchValue dispatch, "fallbackImagePath"
    End If
    imagePath = VTDispatchOptional(dispatch, "imagePath")
    vectorDocumentPath = VTDispatchOptional( _
        dispatch, "vectorDocumentPath")
    fallbackImagePath = VTDispatchOptional( _
        dispatch, "fallbackImagePath")
    metadata = CStr(dispatch("metadata"))
    latexBase64 = CStr(dispatch("latexBase64"))
    ommlBase64 = CStr(dispatch("ommlBase64"))
    nativeDocumentPath = VTDispatchOptional(dispatch, "nativeDocumentPath")
    pendingMarker = VTDispatchOptional(dispatch, "pendingMarker")
    sourceMarker = VTDispatchOptional(dispatch, "sourceMarker")
    sourceDocumentId = VTDispatchOptional(dispatch, "sourceDocumentId")
    wordMathFontName = VTDispatchOptional(dispatch, "wordMathFontName")
    forkCopiedFormula = _
        (VTDispatchOptional(dispatch, "forkCopiedFormula") = "1")

    VTTraceWordSession sessionId, "commit-dispatch-read", pendingMarker

    If Len(sourceDocumentId) = 0 Then
        Err.Raise vbObjectError + 7415, "VisualTeX", _
            "The VisualTeX Word source document identity is missing."
    End If
    Set targetDocument = VTFindOpenWordDocumentByIdentity(sourceDocumentId)
    If targetDocument Is Nothing Then
        Err.Raise vbObjectError + 7415, "VisualTeX", _
            "The Word document used to start this VisualTeX operation is no longer open."
    End If
    targetDocument.Activate
    VTRequireWritableWordDocument
    widthPoints = VTDispatchPositiveDouble(dispatch, "widthPoints")
    heightPoints = VTDispatchPositiveDouble(dispatch, "heightPoints")
    baselinePoints = VTDispatchOptionalDouble(dispatch, "baseline", 0#)
    fontSizePt = VTDispatchPositiveDouble(dispatch, "fontSizePt")
    referenceWidthPt = VTDispatchPositiveDouble(dispatch, "referenceWidthPt")
    referenceHeightPt = VTDispatchPositiveDouble(dispatch, "referenceHeightPt")
    referenceBaselinePt = _
        VTDispatchOptionalDouble(dispatch, "referenceBaselinePt", 0#)
    inkCenterYRatio = _
        VTDispatchOptionalDouble(dispatch, "inkCenterYRatio", 0.5)

    If Not VTIsCanonicalUuid(formulaId) Or Not VTIsEncodedMetadata(metadata) Or _
       Not VTIsBase64UrlPayload(latexBase64) Or _
       Not VTIsBase64UrlPayload(ommlBase64) Then
        Err.Raise vbObjectError + 7405, "VisualTeX", "VisualTeX Word result metadata or native-equation payload is invalid."
    End If
    If mode <> "create" And mode <> "edit" Then
        Err.Raise vbObjectError + 7407, "VisualTeX", "The VisualTeX Word result mode is invalid."
    End If
    If displayMode <> "inline" And displayMode <> "block" Then
        Err.Raise vbObjectError + 7451, "VisualTeX", "The VisualTeX Word display mode is invalid."
    End If
    If numbered And displayMode <> "block" Then
        Err.Raise vbObjectError + 7449, "VisualTeX", "Only display formulas can retain a Word equation number."
    End If
    If Not VTValidWordFormulaFontSize(fontSizePt) Or _
       referenceWidthPt <= 0# Or referenceHeightPt <= 0# Or _
       referenceBaselinePt < -256# Or referenceBaselinePt > 0# Or _
       inkCenterYRatio < 0# Or inkCenterYRatio > 1# Then
        Err.Raise vbObjectError + 7498, "VisualTeX", _
            "VisualTeX Word image font-size geometry is invalid."
    End If
    VTWordPerformanceMark "dispatch-validated"

    formulaReference = VTFormulaReference(formulaId, displayMode, numbered)
    captionText = VTEquationCrossReferenceText(latexBase64)
    If nativeEquation Then
        VTValidateAbsoluteVisualTeXPath nativeDocumentPath
        If Not VTPathFileExists(nativeDocumentPath) Then
            Err.Raise vbObjectError + 7406, "VisualTeX", _
                "VisualTeX Word native DOCX result is missing."
        End If
    Else
        VTValidateAbsoluteVisualTeXPath imagePath
        VTValidateAbsoluteVisualTeXPath vectorDocumentPath
        VTValidateAbsoluteVisualTeXPath fallbackImagePath
        If Not VTPathFileExists(imagePath) Then
            Err.Raise vbObjectError + 7406, "VisualTeX", _
                "VisualTeX Word SVG result is missing."
        End If
        If Not VTPathFileExists(vectorDocumentPath) Then
            Err.Raise vbObjectError + 7406, "VisualTeX", _
                "VisualTeX Word SVG staging document is missing."
        End If
        If Not VTPathFileExists(fallbackImagePath) Then
            Err.Raise vbObjectError + 7406, "VisualTeX", _
                "VisualTeX Word PNG compatibility preview is missing."
        End If
    End If

    transactionStage = "resolve-target"
    On Error Resume Next
    If mode = "create" Then
        If targetDocument.Bookmarks.Exists(VTWordBookmarkName(sessionId)) Then
            Set pendingBookmark = targetDocument.Bookmarks(VTWordBookmarkName(sessionId))
            If pendingBookmark.Range.InlineShapes.Count = 1 Then
                Set targetImage = pendingBookmark.Range.InlineShapes(1)
            Else
                targetFromPendingBookmark = True
            End If
        Else
            Set targetImage = VTFindUniqueInlineShape(pendingMarker)
        End If
    ElseIf Left$(sourceMarker, Len(VT_WORD_NATIVE_BOOKMARK_PREFIX)) = _
           VT_WORD_NATIVE_BOOKMARK_PREFIX Then
        If targetDocument.Bookmarks.Exists(sourceMarker) Then
            Set nativeTarget = targetDocument.Bookmarks(sourceMarker)
            If nativeTarget Is Nothing Then
                targetIsNative = False
            Else
                targetIsNative = True
            End If
        End If
    ElseIf Left$(sourceMarker, Len(VT_WORD_EDIT_BOOKMARK_PREFIX)) = _
           VT_WORD_EDIT_BOOKMARK_PREFIX Then
        If targetDocument.Bookmarks.Exists(sourceMarker) Then
            Set sourceEditBookmark = targetDocument.Bookmarks(sourceMarker)
            ' A historical image can still be inside OMath. Its concrete
            ' InlineShape takes precedence over the enclosing native container.
            If sourceEditBookmark.Range.InlineShapes.Count = 1 Then
                Set targetImage = sourceEditBookmark.Range.InlineShapes(1)
            ElseIf sourceEditBookmark.Range.OMaths.Count = 1 Then
                Set nativeTarget = sourceEditBookmark
                targetIsNative = True
                nativeTargetFromEditBookmark = True
            End If
        End If
    Else
        Set targetImage = VTFindUniqueInlineShape(sourceMarker)
    End If
    Err.Clear
    On Error GoTo RollbackCandidate
    VTBeginWordInternalMutation
    internalMutationStarted = True
    previousScreenUpdating = Application.ScreenUpdating
    screenUpdatingCaptured = True
    Application.ScreenUpdating = False
    VTTraceWordSession sessionId, "commit-target-resolved", pendingMarker
    VTWordPerformanceMark "target-resolved"

    transactionStage = "capture-target-range"
    If targetIsNative Then
        If nativeTargetFromEditBookmark Then
            If nativeTarget.Range.OMaths.Count = 1 Then
                Set originalNativeMath = nativeTarget.Range.OMaths(1)
            End If
            targetFormulaId = formulaId
        Else
            Set originalNativeMath = VTNativeMathForBookmark(nativeTarget)
            If Not VTTryFormulaIdFromNativeBookmark( _
               nativeTarget.Name, targetFormulaId) Then
                targetFormulaId = ""
            End If
        End If
        If targetFormulaId <> formulaId Or originalNativeMath Is Nothing Then
            Err.Raise vbObjectError + 7454, "VisualTeX", _
                "The original VisualTeX native equation target is invalid."
        End If
        Set originalNativeRange = originalNativeMath.Range.Duplicate
        sourceWasNumbered = Not VTPureNativeNumberField(originalNativeRange) Is Nothing
        Set nativeNumberBookmarks = VTCaptureNativeNumberBookmarks(originalNativeRange)
        originalNativeStart = originalNativeRange.Start
        originalNativeBookmarkName = nativeTarget.Name
        Set targetRange = originalNativeRange.Duplicate
    ElseIf Not targetImage Is Nothing Then
        Set targetRange = VTVisualTeXImageContainerRange(targetImage)
        sourceWasNumbered = Not VTImagePlaceRef(targetImage.Range) Is Nothing
        If mode = "create" Then pendingPlaceholderStart = targetImage.Range.Start
    ElseIf targetFromPendingBookmark And Not pendingBookmark Is Nothing Then
        Set targetRange = pendingBookmark.Range.Duplicate
    Else
        If nativeEquation Then
            If targetDocument.Bookmarks.Exists(VTNativeFormulaBookmarkName(formulaId)) Then
                VTSetWordLatexPayload targetDocument, formulaId, latexBase64
                VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64
                VTSetWordMetadataPayload targetDocument, formulaId, metadata
                VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
                VTSetWordImageScaleState _
                    targetDocument, formulaId, fontSizePt, referenceWidthPt, _
                    referenceHeightPt, referenceBaselinePt, fontSizePt
                Set originalNativeMath = VTNativeMathForBookmark( _
                    targetDocument.Bookmarks( _
                        VTNativeFormulaBookmarkName(formulaId)))
                If Not originalNativeMath Is Nothing Then
                    VTV2ApplyEquationTypography originalNativeMath.Range, fontSizePt
                End If
                VTDeletePendingBookmark targetDocument, sessionId
                GoTo CommitSucceeded
            End If
        Else
            Set committed = VTFindCommittedInlineShape(metadata, formulaReference)
            If Not committed Is Nothing Then
                VTSetWordLatexPayload targetDocument, formulaId, latexBase64
                VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64
                VTSetWordMetadataPayload targetDocument, formulaId, metadata
                VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
                observedWordFontSizePt = VTInlineShapeWordFontSize(committed)
                If Not VTValidWordFormulaFontSize(observedWordFontSizePt) Then
                    observedWordFontSizePt = fontSizePt
                End If
                VTSetWordImageScaleState _
                    targetDocument, formulaId, fontSizePt, referenceWidthPt, _
                    referenceHeightPt, referenceBaselinePt, _
                    observedWordFontSizePt
                VTSetWordImageInkCenterYRatio _
                    targetDocument, formulaId, inkCenterYRatio
                VTSetWordImageLetterFont targetDocument, formulaId, _
                    VTDispatchOptional(dispatch, "formulaLetterFont")
                VTApplyWordImageFormulaFontSize committed, fontSizePt
                Set committed = _
                    VTEnsureVisualTeXImageMacroButton(committed)
                VTDeletePendingBookmark targetDocument, sessionId
                GoTo CommitSucceeded
            End If
        End If
        Err.Raise vbObjectError + 7426, "VisualTeX", "The original Word formula is missing and no committed VisualTeX result was found."
    End If

    If mode = "edit" And nativeEquation And Not targetImage Is Nothing And _
       VTDispatchOptional(dispatch, "operation") = "imageToNative" And _
       Not forkCopiedFormula Then
        ' The renderer has supplied the missing native staging DOCX. Resume the
        ' same protected conversion used by a warm cache: generic edit commit
        ' backs up native sources only and cannot restore a deleted image.
        transactionStage = "resume-image-to-native-conversion"
        If nativeDocumentPath <> VTNativeWordDocumentPath(formulaId) Then
            Err.Raise vbObjectError + 7406, "VisualTeX", _
                "The rebuilt native DOCX does not belong to the source formula."
        End If
        VTWordConvertInlineShapeToNativeEquation targetImage
        VTDeletePendingBookmark targetDocument, sessionId
        GoTo CommitSucceeded
    End If

    If forkCopiedFormula And mode = "edit" And nativeEquation And _
       Not targetImage Is Nothing And _
       VTDispatchOptional(dispatch, "operation") = "imageToNative" Then
        ' A copied image is being converted under a fresh formula id. Preserve
        ' the exact source container before the generic native commit removes it;
        ' unlike an ordinary native edit, there is no original OMath rollback
        ' document to restore this image if a later layout step fails.
        transactionStage = "backup-forked-image-source"
        originalImageStart = targetRange.Start
        Set originalImageBackupDocument = Documents.Add(Visible:=False)
        Set originalImageBackupRange = originalImageBackupDocument.Content
        originalImageBackupRange.Text = ""
        originalImageBackupRange.Collapse wdCollapseStart
        originalImageBackupRange.FormattedText = targetRange.FormattedText
        If originalImageBackupDocument.InlineShapes.Count <> 1 Then
            Err.Raise vbObjectError + 7472, "VisualTeX", _
                "Word could not back up the copied image before native conversion."
        End If
        Set originalImageBackupRange = _
            VTVisualTeXImageContainerRange( _
                originalImageBackupDocument.InlineShapes(1))
        targetDocument.Activate
    End If

    transactionStage = "read-previous-state"
    hadPreviousLatexPayload = VTTryReadWordLatexPayload( _
        targetDocument, formulaId, previousLatexBase64)
    hadPreviousOmmlPayload = VTTryReadWordOmmlPayload( _
        targetDocument, formulaId, previousOmmlBase64)
    hadPreviousMetadataPayload = VTTryReadWordMetadataPayload( _
        targetDocument, formulaId, previousMetadata)
    hadPreviousFormat = VTTryReadWordFormulaFormat( _
        targetDocument, formulaId, previousDisplayMode, previousNumbered)
    hadPreviousImageScale = VTTryReadWordImageScaleState( _
        targetDocument, formulaId, previousFontSizePt, _
        previousReferenceWidthPt, previousReferenceHeightPt, _
        previousReferenceBaselinePt, previousObservedWordFontSizePt)
    hadPreviousImageCanvas = VTTryGetDocumentVariable( _
        targetDocument, VTWordImageCanvasVariableName(formulaId), previousImageCanvas)
    imageCanvasCaptured = True
    hadPreviousInkCenter = VTTryReadWordImageInkCenterYRatio( _
        targetDocument, formulaId, previousInkCenterYRatio)
    VTWordPerformanceMark "previous-state-read"

    If Not hadPreviousFormat Then previousNumbered = sourceWasNumbered
    numberingChanged = mode = "edit" And _
        previousNumbered <> numbered And displayMode = "block"
    If numberingChanged And Not targetImage Is Nothing Then
        ' Number toggles change tabs, fields and paragraph formatting as well as
        ' the image. Back up that exact paragraph for transactional rollback.
        transactionStage = "backup-numbering-paragraph"
        Set numberingRollbackRange = targetImage.Range.Paragraphs(1).Range.Duplicate
        originalImageStart = numberingRollbackRange.Start
        Set imageNumberBookmarks = VTCaptureImageNumberBookmarks(targetImage, formulaId)
        Set originalImageBackupDocument = VTTakePreparedNativeRollbackDocument()
        If originalImageBackupDocument Is Nothing Then _
            Set originalImageBackupDocument = Documents.Add(Visible:=False)
        Set originalImageBackupRange = originalImageBackupDocument.Content
        originalImageBackupRange.Text = ""
        originalImageBackupRange.Collapse wdCollapseStart
        originalImageBackupRange.FormattedText = numberingRollbackRange.FormattedText
        Set originalImageBackupRange = originalImageBackupDocument.Paragraphs(1).Range.Duplicate
        targetDocument.Activate
    End If

    If numberingChanged And Not nativeEquation And Not targetImage Is Nothing And _
       VTDispatchOptional(dispatch, "numberingOnly") = "1" And _
       previousLatexBase64 = latexBase64 And previousOmmlBase64 = ommlBase64 And _
       hadPreviousImageScale And Abs(previousFontSizePt - fontSizePt) < 0.01 Then
        ' The renderer confirmed that source, fonts and size are unchanged.
        ' Retain the original vector image and mutate only its number scaffold.
        transactionStage = "apply-image-numbering-only"
        Set candidate = targetImage
        originalImageRemoved = True
        candidate.Title = formulaReference
        candidate.AlternativeText = metadata
        VTSetWordLatexPayload targetDocument, formulaId, latexBase64
        VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64
        VTSetWordMetadataPayload targetDocument, formulaId, metadata
        VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
        formulaStateStored = True
        VTApplyImageNumberingEdit candidate, formulaId, numbered
        observedWordFontSizePt = VTInlineShapeWordFontSize(candidate)
        If Not VTValidWordFormulaFontSize(observedWordFontSizePt) Then _
            observedWordFontSizePt = fontSizePt
        VTSetWordImageScaleState targetDocument, formulaId, fontSizePt, _
            referenceWidthPt, referenceHeightPt, referenceBaselinePt, observedWordFontSizePt
        If regressionFailureAfterSourceRemoval Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                "Injected numbering-only failure after final layout."
        End If
        VTSetWordImageOwnerBookmark targetDocument, candidate, formulaId
        candidate.Select
        GoTo CommitSucceeded
    End If

    If nativeEquation Then
        transactionStage = "prepare-native-replacement"
        nativeDisplayMode = displayMode
        ' Unnumbered display staging DOCX now carries a real m:oMathPara, so keep
        ' that native display identity through replacement instead of paying for
        ' an Inline -> Display rebuild after the source transaction settles.
        ' Numbered OMML still enters as inline because its production eqArr/#/REF
        ' layout is built explicitly by VTEnsureNativeEquationNumber below.
        If numbered Then nativeDisplayMode = "inline"

        ' Replacing a native equation by first inserting beside it is unsafe in
        ' Word for Mac: Word can merge both equations and mutate the original
        ' COM Range so a later Delete consumes adjacent body text. Back up the
        ' original OMath in a hidden document and replace its exact Range in
        ' one FormattedText assignment instead.
        If targetIsNative Then
            VTWordPerformanceMark "native-backup-start"
            Set originalNativeBackupDocument = _
                VTTakePreparedNativeRollbackDocument()
            If originalNativeBackupDocument Is Nothing Then
                Set originalNativeBackupDocument = _
                    Documents.Add(Visible:=False)
            End If
            Set originalNativeBackupRange = originalNativeBackupDocument.Content
            originalNativeBackupRange.Text = ""
            originalNativeBackupRange.Collapse wdCollapseStart
            originalNativeBackupRange.FormattedText = originalNativeRange.FormattedText
            If originalNativeBackupDocument.OMaths.Count <> 1 Then
                Err.Raise vbObjectError + 7471, "VisualTeX", _
                    "Word could not back up the original native equation before replacement."
            End If
            Set originalNativeBackupRange = _
                VTNativeMathRoot(originalNativeBackupDocument.OMaths(1)).Range.Duplicate
            targetRange.Document.Activate
            VTWordPerformanceMark "native-backup-complete"
        End If

        If Not targetIsNative And nativeDisplayMode = "inline" Then
            Set adjacentNative = VTAdjacentNativeMath(targetRange)
            If Not adjacentNative Is Nothing Then
                Set originalNativeRange = adjacentNative.Range.Duplicate
                originalNativeStart = originalNativeRange.Start
                Set originalNativeBackupDocument = Documents.Add(Visible:=False)
                Set originalNativeBackupRange = originalNativeBackupDocument.Content
                originalNativeBackupRange.Collapse wdCollapseStart
                originalNativeBackupRange.FormattedText = originalNativeRange.FormattedText
                Set originalNativeBackupRange = VTNativeMathRoot(originalNativeBackupDocument.OMaths(1)).Range.Duplicate
                targetDocument.Activate
            End If
        End If

        If numbered And displayMode = "block" Then
            ' Equation format v2: the app builds the complete numbered equation
            ' with the shared writer (Windows BuildWordNativeNumberedOmml); the
            ' add-in decides only its numbering context and adds no field later.
            transactionStage = "v2-plan-number"
            VTV2PlanCommitNumber targetDocument, targetRange, targetIsNative, _
                v2SequenceName, v2HeadingLevel, v2Separator, v2LiteralPrefix, _
                v2Automatic, v2Restart, v2Initial, v2TransferEquation
            If targetIsNative Then
                If targetRange.OMaths.Count >= 1 Then _
                    Set v2NumberBookmarks = VTCaptureNativeNumberBookmarks( _
                        VTNativeMathRoot(targetRange.OMaths(1)).Range.Duplicate)
            End If
            VTWordPerformanceMark "v2-bookmarks-captured"
            transactionStage = "v2-build-number"
            nativeDocumentPath = VTV2RequestNumberDocx(sessionId, targetDocument, _
                ommlBase64, 1, VTV2NumberItemManifest(1, v2SequenceName, _
                    v2HeadingLevel, v2Separator, v2LiteralPrefix, v2Automatic, _
                    v2Restart, v2Initial))
            VTWordPerformanceMark "v2-number-built"
        End If

        transactionStage = "insert-native-equation"
        Set nativeEquationRange = VTInsertNativeEquationAtRange( _
            targetRange, _
            ommlBase64, _
            nativeDocumentPath, _
            nativeDisplayMode, _
            displayMode = "block", _
            targetIsNative, _
            replacementRollbackRange:=originalNativeBackupRange, _
            joinedNative:=joinedNative, nativeFontSizePt:=fontSizePt)
        nativeEquationStart = nativeEquationRange.Start
        nativeTargetReplaced = targetIsNative
        VTWordPerformanceMark "native-insert-complete"

        transactionStage = "store-native-state"
        targetDocument.Activate
        VTSetWordLatexPayload targetDocument, formulaId, latexBase64
        VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64
        VTSetWordMetadataPayload targetDocument, formulaId, metadata
        VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
        VTSetWordImageScaleState _
            targetDocument, formulaId, fontSizePt, referenceWidthPt, _
            referenceHeightPt, referenceBaselinePt, fontSizePt
        formulaStateStored = True
        VTWordPerformanceMark "native-state-stored"

        ' Delete the source placeholder or image before creating the final
        ' display/number layout. Word can shift OMath and tab Ranges when the
        ' adjacent InlineShape disappears, which previously left formulas off
        ' center and could remove the number during a later edit.
        transactionStage = "remove-native-source"
        If Not targetIsNative And Not targetImage Is Nothing Then
            If mode = "create" Then
                pendingPlaceholderStart = targetImage.Range.Start
                pendingPlaceholderRemoved = True
            End If
            VTDeleteVisualTeXImageContainer targetImage
            If Not originalImageBackupRange Is Nothing Then
                originalImageRemoved = True
            End If
        End If

        transactionStage = "resolve-native-after-source-removal"
        Set nativeEquationRange = VTResolveNativeEquationRange( _
            targetDocument, nativeEquationStart, 16)

        ' Persist an exact identity before numbering or caret movement. Every
        ' later resolve and rollback must use this Bookmark rather than a
        ' nearest-OMath search that can select an existing formula below the
        ' insertion point.
        transactionStage = "bookmark-native-before-layout"
        VTSetNativeFormulaBookmark _
            targetDocument, nativeEquationRange, formulaId
        nativeBookmarkSet = True
        Set originalNativeMath = VTNativeMathForBookmark( _
            targetDocument.Bookmarks(VTNativeFormulaBookmarkName(formulaId)))
        VTWordPerformanceMark "native-source-removed"
        If originalNativeMath Is Nothing Then
            Err.Raise vbObjectError + 7460, "VisualTeX", _
                "Word could not resolve the newly inserted native equation identity."
        End If
        Set nativeEquationRange = originalNativeMath.Range.Duplicate

        If displayMode = "block" Then
            ' Equation format v2: display type, typography and paragraph exactly
            ' as VisualTeX for Windows lays out a display equation.
            transactionStage = "v2-finish-display"
            VTWordPerformanceMark "native-display-promote-start"
            Set nativeEquationRange = _
                VTV2FinishDisplayEquation(nativeEquationRange, fontSizePt)
            VTWordPerformanceMark "native-display-promote-complete"
            If numbered Then
                transactionStage = "v2-finish-number"
                VTV2FinishInsertedNumber targetDocument, nativeEquationRange, _
                    v2TransferEquation
                Set nativeEquationRange = VTResolveNativeEquationRange( _
                    targetDocument, nativeEquationRange.Start, 16)
                VTV2RestoreNumberBookmarks nativeEquationRange, v2NumberBookmarks
            End If
        Else
            transactionStage = "finalize-native-inline"
            Set nativeEquationRange = _
                VTFinalizeInlineNativeEquation(nativeEquationRange)
            transactionStage = "apply-native-font-size"
            If Not joinedNative Then _
                VTV2ApplyEquationTypography nativeEquationRange, fontSizePt
            If Not joinedNative Then _
                VTV2ConfigurePromotedInlineEquation nativeEquationRange
        End If
        VTWordPerformanceMark "native-font-complete"
        Set nativeEquationRange = VTResolveNativeEquationRange( _
            targetDocument, nativeEquationRange.Start, 16)
        ' Equation format v2: the equation is in the document's math font
        ' (ApplyExplicitTransferMathFont in the shared writer); no
        ' per-equation font is applied, as on Windows.
        VTWordPerformanceMark "native-final-resolve-complete"

        If numberingChanged Then
            transactionStage = "reconcile-native-numbering"
            If Not numbered Then
                Set numberingPrefix = nativeEquationRange.Duplicate
                numberingPrefix.Collapse wdCollapseEnd
                VTRetireExternalSequenceHelper targetDocument, formulaId, numberingPrefix
            End If
            VTV2UpdateNumbersFrom targetDocument, nativeEquationRange.Start
        End If

        transactionStage = "bookmark-native-equation"
        VTSetNativeFormulaBookmark targetDocument, nativeEquationRange, formulaId
        VTDeleteWordImageOwnerBookmark targetDocument, formulaId
        nativeBookmarkSet = True
        VTWordPerformanceMark "native-final-bookmark-complete"
        Set originalNativeMath = VTNativeMathForBookmark( _
            targetDocument.Bookmarks(VTNativeFormulaBookmarkName(formulaId)))
        VTWordPerformanceMark "native-final-bookmark-resolve-complete"
        If originalNativeMath Is Nothing Then
            Err.Raise vbObjectError + 7460, "VisualTeX", _
                "Word lost the finalized native equation before structure caching."
        End If
        VTWordPerformanceMark "native-signature-start"
        VTSetWordNativeSignature _
            targetDocument, formulaId, originalNativeMath
        VTWordPerformanceMark "native-signature-complete"
        VTDeletePendingBookmark targetDocument, sessionId
        VTWordPerformanceMark "native-pending-delete-complete"
        VTWordPerformanceMark "native-layout-complete"
        If numberingChanged And regressionFailureAfterSourceRemoval Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                "Injected native numbering failure after final layout."
        End If

        On Error Resume Next
        If Not originalNativeBackupDocument Is Nothing Then
            originalNativeBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
        End If
        On Error GoTo RollbackCandidate
        transactionStage = "place-native-caret"
        On Error Resume Next
        If displayMode = "inline" And mode = "create" Then
            VTPlaceCaretAfterInlineNativeEquation nativeEquationRange
        ElseIf displayMode = "block" And mode = "create" Then
            VTPlaceCaretAfterDisplayFormula nativeEquationRange, formulaId, False
        Else
            ' Editing or direct conversion must not manufacture a U+2060 caret
            ' anchor after an existing formula. Keep the converted formula
            ' selected so the surrounding Word text remains byte-for-byte stable.
            nativeEquationRange.Select
        End If
        If Err.Number <> 0 Then
            Err.Clear
            Set originalNativeMath = VTNativeMathForBookmark( _
                targetDocument.Bookmarks( _
                    VTNativeFormulaBookmarkName(formulaId)))
            If Not originalNativeMath Is Nothing Then
                originalNativeMath.Range.Select
            End If
        End If
        On Error GoTo RollbackCandidate
        VTWordPerformanceMark "native-caret-placed"
        GoTo CommitSucceeded
    End If

    transactionStage = "prepare-image-replacement"
    If targetIsNative And numbered Then
        If originalNativeRange.Information(wdWithInTable) Then
            Set originalNativeTable = originalNativeRange.Tables(1)
            If originalNativeTable.Rows.Count <> 1 Or _
               originalNativeTable.Columns.Count <> 3 Or _
               originalNativeTable.Range.OMaths.Count <> 1 Or _
               originalNativeTable.Range.InlineShapes.Count <> 0 Then
                Err.Raise vbObjectError + 7471, "VisualTeX", _
                    "The numbered native image replacement has an ambiguous table."
            End If
            ' A cell-ending OMath Range rejects image replacement with 5904 on
            ' Mac Word. Keep an exact table/paragraph backup, stage the image in
            ' the following ordinary paragraph, then retire this native table.
            ' Mac Word does not group the cross-document image transfer reliably
            ' into one Undo record, so rollback restores this exact local span.
            nativeTableStart = originalNativeTable.Range.Start
            Set nativeTableBackupSource = originalNativeTable.Range.Duplicate
            Set nativeTableFollowingParagraph = targetDocument.Range( _
                Start:=originalNativeTable.Range.End, _
                End:=originalNativeTable.Range.End + 1).Paragraphs(1).Range.Duplicate
            needsNativeImageParagraph = _
                nativeTableFollowingParagraph.Information(wdWithInTable) Or _
                nativeTableFollowingParagraph.Fields.Count <> 0 Or _
                nativeTableFollowingParagraph.OMaths.Count <> 0 Or _
                nativeTableFollowingParagraph.InlineShapes.Count <> 0 Or _
                VTWordRangeHasMeaningfulText(nativeTableFollowingParagraph)
            If Not needsNativeImageParagraph Then
                nativeTableBackupSource.End = nativeTableFollowingParagraph.End
            End If
            Set originalNativeBackupDocument = Documents.Add(Visible:=False)
            Set originalNativeBackupRange = originalNativeBackupDocument.Content.Duplicate
            originalNativeBackupRange.Collapse wdCollapseStart
            originalNativeBackupRange.FormattedText = nativeTableBackupSource.FormattedText
            If originalNativeBackupDocument.Tables.Count <> 1 Or _
               originalNativeBackupDocument.OMaths.Count <> 1 Then
                Err.Raise vbObjectError + 7471, "VisualTeX", _
                    "Word could not back up the complete numbered native table."
            End If
            Set originalNativeBackupRange = originalNativeBackupDocument.Content.Duplicate
            originalNativeBackupRange.End = originalNativeBackupRange.End - 1
            nativeTableBackupReady = True
            targetDocument.Activate
            Set nativeTableReplacementEnd = nativeTableBackupSource.Duplicate
            nativeTableReplacementEnd.Collapse wdCollapseEnd
            If needsNativeImageParagraph Then
                Set nativeTableFollowingParagraph = _
                    VTInsertDedicatedEquationHelperParagraph(originalNativeTable)
            End If
            Set nativeTableReplacementEnd = nativeTableFollowingParagraph.Duplicate
            nativeTableReplacementEnd.Collapse wdCollapseEnd
            Set targetRange = nativeTableFollowingParagraph.Duplicate
            targetRange.Collapse wdCollapseStart
        End If
    End If

    If targetIsNative And originalNativeTable Is Nothing Then
        VTWordPerformanceMark "image-native-backup-start"
        Set originalNativeBackupDocument = Documents.Add(Visible:=False)
        Set originalNativeBackupRange = originalNativeBackupDocument.Content
        originalNativeBackupRange.Collapse wdCollapseStart
        originalNativeBackupRange.FormattedText = originalNativeRange.FormattedText
        If originalNativeBackupDocument.OMaths.Count <> 1 Then
            Err.Raise vbObjectError + 7471, "VisualTeX", _
                "Word could not back up the original native equation before image replacement."
        End If
        Set originalNativeBackupRange = _
            VTNativeMathRoot(originalNativeBackupDocument.OMaths(1)).Range.Duplicate
        targetRange.Document.Activate
        VTWordPerformanceMark "image-native-backup-complete"
    End If

    Set insertionRange = targetRange.Duplicate
    If Not targetIsNative Then insertionRange.Collapse wdCollapseStart
    targetDocument.Activate
    transactionStage = "insert-image-replacement"
    Set stagedCandidate = VTAddWordFormulaPicture( _
        targetDocument, insertionRange, vectorDocumentPath, _
        fallbackImagePath)
    VTWordPerformanceMark "image-staging-inserted"
    nativeTargetReplaced = targetIsNative
    transactionStage = "detach-image-replacement"
    If targetIsNative Then
        Set candidate = _
            VTDetachWordFormulaPictureFromMath(stagedCandidate)
    Else
        Set candidate = stagedCandidate
    End If
    VTWordPerformanceMark "image-detached"
    candidate.LockAspectRatio = msoFalse
    candidate.Width = CSng(widthPoints)
    candidate.Height = CSng(heightPoints)
    candidate.LockAspectRatio = msoTrue
    candidate.AlternativeText = metadata
    candidate.Title = formulaReference
    ' Record the staging anchor before deleting the old source. Normal image
    ' create/edit resolution is then bounded to this tiny local window instead
    ' of enumerating every InlineShape in a document whose formula count grows.
    committedImageStart = candidate.Range.Start
    ' Mac Word versions differ on whether the native Home font-size box
    ' persists InlineShape.Range.Font.Size. Attempt it, but Width/Height remain
    ' the authoritative visual result and the monitor records what Word reports.
    On Error Resume Next
    candidate.Range.Font.Size = CSng(fontSizePt)
    Err.Clear
    On Error GoTo RollbackCandidate
    If displayMode = "inline" Then
        If baselinePoints > 0# Or baselinePoints < -256# Then
            Err.Raise vbObjectError + 7408, "VisualTeX", "VisualTeX Word baseline is outside the allowed range."
        End If
        VTApplyWordInlineImageBaseline _
            candidate, referenceHeightPt, referenceBaselinePt, fontSizePt
    Else
        candidate.Range.Font.Position = 0
        If Not (targetIsNative And numbered) Then
            candidate.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    End If

    If Abs(candidate.Width - widthPoints) > 0.1 Or Abs(VTWordImageGlyphHeight(candidate) - heightPoints) > 0.1 Or _
       candidate.AlternativeText <> metadata Or candidate.Title <> formulaReference Then
        Err.Raise vbObjectError + 7422, "VisualTeX", "Word did not persist the VisualTeX formula properties."
    End If
    If displayMode = "inline" Then
        If candidate.Range.Font.Position <> _
           VTWordInlineImagePosition( _
               candidate, referenceHeightPt, referenceBaselinePt, fontSizePt) Then
            Err.Raise vbObjectError + 7423, "VisualTeX", "Word did not persist the VisualTeX inline baseline."
        End If
    ElseIf Not (targetIsNative And numbered) Then
        If candidate.Range.ParagraphFormat.Alignment <> wdAlignParagraphCenter Then
            Err.Raise vbObjectError + 7424, "VisualTeX", "Word did not persist the VisualTeX display alignment."
        End If
    End If
    VTWordPerformanceMark "image-properties-ready"

    VTSetWordLatexPayload targetDocument, formulaId, latexBase64
    VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64
    VTSetWordMetadataPayload targetDocument, formulaId, metadata
    VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
    observedWordFontSizePt = VTInlineShapeWordFontSize(candidate)
    If Not VTValidWordFormulaFontSize(observedWordFontSizePt) Then
        observedWordFontSizePt = fontSizePt
    End If
    VTSetWordImageScaleState _
        targetDocument, formulaId, fontSizePt, referenceWidthPt, _
        referenceHeightPt, referenceBaselinePt, observedWordFontSizePt
    VTSetWordImageInkCenterYRatio _
        targetDocument, formulaId, inkCenterYRatio
    VTSetWordImageLetterFont targetDocument, formulaId, _
        VTDispatchOptional(dispatch, "formulaLetterFont")
    formulaStateStored = True
    VTWordPerformanceMark "image-state-stored"

    ' Finalize the paragraph only after the old image/native target has gone.
    ' Otherwise the placeholder participates in the tabbed line and Word shifts
    ' the formula away from the true text-column center when it is deleted.
    transactionStage = "remove-image-source"
    If targetIsNative Then
        If Not originalNativeTable Is Nothing Then
            originalNativeTable.Delete
            ' Removing the table moves the staged paragraph to its old start.
            ' The previous image offset can now be beyond Document.Content.End.
            committedImageStart = nativeTableStart
        End If
        If regressionFailureAfterSourceRemoval Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                "Injected native-to-image layout failure after source deletion."
        End If
        If targetRange.Document.Bookmarks.Exists(originalNativeBookmarkName) Then
            targetRange.Document.Bookmarks(originalNativeBookmarkName).Delete
        End If
        originalNativeBookmarkDeleted = True
    ElseIf Not targetImage Is Nothing Then
        If mode = "create" Then pendingPlaceholderRemoved = True
        VTDeleteVisualTeXImageContainer targetImage
        If numberingChanged Then originalImageRemoved = True
    End If

    transactionStage = "resolve-image-after-source-removal"
    Set candidate = VTFindCommittedInlineShapeNearPosition( _
        targetDocument, committedImageStart, metadata, formulaReference, 64)
    If candidate Is Nothing Then
        ' Compatibility fallback only for host builds that move a freshly
        ' inserted picture unexpectedly far during source deletion. Healthy
        ' create/edit transactions never pay this document-wide scan.
        Set candidate = VTFindCommittedInlineShape(metadata, formulaReference)
    End If
    If candidate Is Nothing Then
        Err.Raise vbObjectError + 7426, "VisualTeX", _
            "Word could not resolve the committed formula image after replacement."
    End If
    committedImageStart = candidate.Range.Start
    Set candidate = VTEnsureVisualTeXImageMacroButton(candidate)
    committedImageStart = candidate.Range.Start
    If displayMode = "inline" Then
        ' Reapply from the final resolved image height after Word has removed the
        ' placeholder/old formula and finished any FormattedText normalization.
        ' This closes the host-version gap where Position was correct on the
        ' staged shape but changed after the surrounding document mutated.
        VTApplyWordInlineImageBaseline _
            candidate, referenceHeightPt, referenceBaselinePt, fontSizePt
    End If
    VTWordPerformanceMark "image-source-removed"

    If displayMode = "block" Then
        If numbered Then
            transactionStage = "normalize-image-number-layout"
            numberCreated = False
            If numberingChanged Then
                Set numberLayoutRange = VTWriteSingleParagraphImageNumber(candidate, formulaId)
            Else
                Set numberLayoutRange = VTEnsureImageEquationNumber( _
                    candidate, heightPoints, formulaId, captionText, numberCreated)
            End If
            If numberCreated Then Set insertedNumber = numberLayoutRange
        Else
            transactionStage = "normalize-image-display-layout"
            If numberingChanged Then
                VTRemoveImageNumberingForEdit candidate, formulaId
                Set numberingPrefix = candidate.Range.Paragraphs(1).Range.Duplicate
                numberingPrefix.End = candidate.Range.Start
                ' The numbered image layout owns a leading tab. Delete only
                ' that exact whitespace, never surrounding ordinary text.
                If numberingPrefix.Text = vbTab Then numberingPrefix.Delete
            End If
            VTNormalizeUnnumberedDisplayParagraph candidate.Range
        End If
        VTWordPerformanceMark "image-number-layout-done"
        transactionStage = "normalize-image-paragraph-spacing"
        VTNormalizeImageDisplayParagraph candidate.Range
        VTWordPerformanceMark "image-spacing-done"
        If numberingChanged Then
            transactionStage = "reconcile-image-numbering"
            VTReconcileNumberingAfterEdit targetDocument, candidate.Range.Start, True
            VTWordPerformanceMark "image-reconciled"
        End If
    End If
    If numberingChanged And regressionFailureAfterSourceRemoval Then
        Err.Raise vbObjectError + 7564, "VisualTeX regression", _
            "Injected numbering edit failure after final layout."
    End If

    transactionStage = "validate-image-native-final"
    Set candidate = VTFindCommittedInlineShapeNearPosition( _
        targetDocument, committedImageStart, metadata, formulaReference, 96)
    If candidate Is Nothing Then
        ' Keep the old global lookup strictly as damaged-host recovery; it is
        ' not part of the normal numbered-image edit complexity.
        Set candidate = VTFindCommittedInlineShapeInDocument( _
            targetDocument, metadata, formulaReference)
    End If
    If candidate Is Nothing Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "Word lost the committed formula picture after final layout."
    End If
    Set finalImageField = VTContainingFieldForInlineShape(candidate)
    If Not finalImageField Is Nothing Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "The committed VisualTeX formula picture is still inside a Word field."
    End If
    VTSetWordImageOwnerBookmark targetDocument, candidate, formulaId
    VTDeletePendingBookmark targetDocument, sessionId
    VTWordPerformanceMark "image-layout-complete"

    On Error Resume Next
    If Not originalNativeBackupDocument Is Nothing Then
        originalNativeBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo RollbackCandidate
    transactionStage = "place-image-caret"
    On Error Resume Next
    If displayMode = "block" And mode = "create" Then
        VTPlaceCaretAfterDisplayFormula candidate.Range, formulaId
    Else
        candidate.Select
    End If
    If Err.Number <> 0 Then
        Err.Clear
        candidate.Select
    End If
    On Error GoTo RollbackCandidate
CommitSucceeded:
    If nativeEquation Then
        VTRestoreNativeNumberBookmarks nativeEquationRange, nativeNumberBookmarks
        VTWordPerformanceMark "commit-number-bookmarks-restored"
        VTFinishPureNativeEquation nativeEquationRange, formulaId, (mode = "create")
        VTWordPerformanceMark "commit-native-finished"
    End If
    VTDeleteWordEditBookmark targetDocument, sessionId
    VTWordPerformanceMark "commit-edit-bookmark-deleted"
    On Error Resume Next
    If Not originalImageBackupDocument Is Nothing Then
        originalImageBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set originalImageBackupDocument = Nothing
    End If
    On Error GoTo 0
    VTWordPerformanceMark "commit-succeeded"
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If screenUpdatingCaptured Then Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

RollbackCandidate:
    transactionErrorNumber = Err.Number
    transactionErrorDescription = Err.Description
    VTWriteWordFailureTrace _
        sessionId, transactionStage, transactionErrorNumber, transactionErrorDescription
    On Error Resume Next
    If nativeTableBackupReady Then
        Set rollbackRange = targetDocument.Range( _
            Start:=nativeTableStart, End:=nativeTableReplacementEnd.Start)
        rollbackRange.FormattedText = originalNativeBackupRange.FormattedText
        Set restoredNativeTable = targetDocument.Range( _
            Start:=nativeTableStart + 1, End:=nativeTableStart + 2).Tables(1)
        VTSetNativeFormulaBookmark targetDocument, _
            VTNativeMathRoot(restoredNativeTable.Cell(1, 2).Range.OMaths(1)).Range.Duplicate, formulaId
        Set restoredNativeNumber = restoredNativeTable.Cell(1, 3).Range.Duplicate
        restoredNativeNumber.End = restoredNativeNumber.End - 2
        targetDocument.Bookmarks.Add _
            Name:=VTEquationNumberBookmarkName(formulaId), Range:=restoredNativeNumber
        GoTo RestorePreviousState
    End If
    If numberingChanged And originalImageRemoved And _
       Not numberingRollbackRange Is Nothing Then
        numberingRollbackRange.FormattedText = originalImageBackupRange.FormattedText
        Set candidate = numberingRollbackRange.InlineShapes(1)
        VTSetWordImageOwnerBookmark targetDocument, candidate, formulaId
        VTRestoreImageNumberBookmarks candidate, imageNumberBookmarks
        If previousNumbered Then
            Set rollbackPlaceRef = VTImagePlaceRef(candidate.Range)
            VTSetEquationNumberBookmarkExact targetDocument, formulaId, _
                targetDocument.Range(VTEquationFieldStart(rollbackPlaceRef), _
                    VTEquationFieldEnd(rollbackPlaceRef))
        End If
        VTReconcileNumberingAfterEdit targetDocument, candidate.Range.Start, True
        GoTo RestorePreviousState
    End If
    If numbered And mode = "create" Then
        ' A failed numbered create (image or native) can already have produced
        ' VT_N_/VT_R_/VT_C_ plus an external SEQ helper before a later validation
        ' raises. Remove that exact formula-specific scaffold first; otherwise a
        ' retry sees the half-created helper and changes behavior.
        VTDeleteEquationNumberScaffold targetDocument, formulaId, False
    End If
    If Not nativeEquation And Not insertedNumber Is Nothing Then insertedNumber.Delete
    If nativeBookmarkSet Then
        If targetRange.Document.Bookmarks.Exists( _
            VTNativeFormulaBookmarkName(formulaId)) Then
            targetRange.Document.Bookmarks( _
                VTNativeFormulaBookmarkName(formulaId)).Delete
        End If
    End If

    If joinedNative And Not originalNativeBackupRange Is Nothing Then
        nativeEquationRange.FormattedText = originalNativeBackupRange.FormattedText
        GoTo RestorePreviousState
    End If
    If targetIsNative And nativeTargetReplaced And _
       Not originalNativeBackupRange Is Nothing Then
        If Not candidate Is Nothing Then
            Set rollbackRange = candidate.Range.Duplicate
        ElseIf Not nativeEquationRange Is Nothing Then
            Set rollbackRange = nativeEquationRange.Duplicate
        Else
            Set rollbackNativeMath = VTNativeMathNearStart( _
                targetRange.Document, originalNativeStart, 8)
            If Not rollbackNativeMath Is Nothing Then
                Set rollbackRange = rollbackNativeMath.Range.Duplicate
            End If
        End If
        If Not rollbackRange Is Nothing Then
            rollbackRange.FormattedText = originalNativeBackupRange.FormattedText
            Set rollbackNativeMath = VTNativeMathNearStart( _
                targetRange.Document, originalNativeStart, 8)
            If Not rollbackNativeMath Is Nothing Then
                If targetRange.Document.Bookmarks.Exists( _
                    originalNativeBookmarkName) Then
                    targetRange.Document.Bookmarks( _
                        originalNativeBookmarkName).Delete
                End If
                targetRange.Document.Bookmarks.Add _
                    Name:=originalNativeBookmarkName, _
                    Range:=rollbackNativeMath.Range.Duplicate
                If numberingChanged Then
                    VTRestoreNativeNumberBookmarks rollbackNativeMath.Range, nativeNumberBookmarks
                    VTReconcileNumberingAfterEdit targetDocument, rollbackNativeMath.Range.Start, False
                End If
            End If
        End If
    Else
        If Not candidate Is Nothing Then
            VTDeleteVisualTeXImageContainer candidate
        End If
        If mode = "create" And committedImageStart >= 0 Then
            Set rollbackMacroField = _
                VTVisualTeXImageMacroButtonFieldNearPosition( _
                    targetDocument, committedImageStart, 16)
            If Not rollbackMacroField Is Nothing Then rollbackMacroField.Delete
        End If
        If Not nativeEquationRange Is Nothing Then nativeEquationRange.Delete
        If originalImageRemoved And Not originalImageBackupRange Is Nothing Then
            Set insertionRange = targetDocument.Range( _
                Start:=originalImageStart, End:=originalImageStart)
            insertionRange.FormattedText = originalImageBackupRange.FormattedText
        End If
        If pendingPlaceholderRemoved Then
            Set insertionRange = targetDocument.Range( _
                Start:=pendingPlaceholderStart, _
                End:=pendingPlaceholderStart)
            Set restoredPlaceholder = targetDocument.InlineShapes.AddPicture( _
                FileName:=VTPlaceholderImagePath(), _
                LinkToFile:=False, _
                SaveWithDocument:=True, _
                Range:=insertionRange)
            restoredPlaceholder.AlternativeText = pendingMarker
            restoredPlaceholder.Title = pendingMarker
            restoredPlaceholder.Width = 1
            restoredPlaceholder.Height = 1
            VTAddPendingBookmark restoredPlaceholder.Range, sessionId
        End If
    End If

RestorePreviousState:
    ' Canvas writes precede the general state commit and can fail during
    ' staging validation. Restore them even when formulaStateStored is False.
    If imageCanvasCaptured Then
        If hadPreviousImageCanvas Then
            VTSetDocumentVariable targetDocument, VTWordImageCanvasVariableName(formulaId), previousImageCanvas
        Else
            VTDeleteDocumentVariable targetDocument, VTWordImageCanvasVariableName(formulaId)
        End If
    End If
    If formulaStateStored Then
        If hadPreviousLatexPayload Then
            VTSetWordLatexPayload targetDocument, formulaId, previousLatexBase64
        Else
            VTDeleteWordLatexPayload targetDocument, formulaId
        End If
        If hadPreviousOmmlPayload Then
            VTSetWordOmmlPayload targetDocument, formulaId, previousOmmlBase64
        Else
            VTDeleteWordOmmlPayload targetDocument, formulaId
        End If
        If hadPreviousMetadataPayload Then
            VTSetWordMetadataPayload targetDocument, formulaId, previousMetadata
        Else
            VTDeleteWordMetadataPayload targetDocument, formulaId
        End If
        If hadPreviousFormat Then
            VTSetWordFormulaFormat _
                targetDocument, formulaId, previousDisplayMode, previousNumbered
        Else
            VTDeleteDocumentVariable targetDocument, VTWordFormatVariableName(formulaId)
        End If
        If hadPreviousImageScale Then
            VTSetWordImageScaleState _
                targetDocument, formulaId, previousFontSizePt, _
                previousReferenceWidthPt, previousReferenceHeightPt, _
                previousReferenceBaselinePt, previousObservedWordFontSizePt
        Else
            VTDeleteDocumentVariable _
                targetDocument, VTWordImageScaleVariableName(formulaId)
        End If
        If hadPreviousInkCenter Then
            VTSetWordImageInkCenterYRatio _
                targetDocument, formulaId, previousInkCenterYRatio
        Else
            VTDeleteDocumentVariable _
                targetDocument, VTWordImageInkCenterVariableName(formulaId)
        End If
    End If
    If targetIsNative And Not nativeTargetReplaced And _
       (originalNativeBookmarkDeleted Or nativeBookmarkSet) Then
        If Not targetRange.Document.Bookmarks.Exists(originalNativeBookmarkName) Then
            targetRange.Document.Bookmarks.Add _
                Name:=originalNativeBookmarkName, Range:=originalNativeRange
        End If
    End If
    If Not originalNativeBackupDocument Is Nothing Then
        originalNativeBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Not originalImageBackupDocument Is Nothing Then
        originalImageBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If screenUpdatingCaptured Then Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Err.Raise transactionErrorNumber, "VisualTeX Word transaction", _
        transactionStage & ": " & transactionErrorDescription
End Sub

Public Sub VTCancelWordDispatch(ByVal sessionId As String, ByVal dispatch As Object)
    Dim pendingMarker As String
    Dim sourceDocumentId As String
    Dim pendingBookmark As Bookmark
    Dim target As InlineShape

    pendingMarker = VTDispatchOptional(dispatch, "pendingMarker")
    sourceDocumentId = VTDispatchOptional(dispatch, "sourceDocumentId")
    ' The document was closed while its editor was open: nothing is left to
    ' clean up, and Word has no ActiveDocument (error 4248).
    If Documents.Count = 0 Then Exit Sub
    If Len(sourceDocumentId) > 0 And sourceDocumentId <> VTWordDocumentIdentity() Then Exit Sub
    If Len(pendingMarker) > 0 Then
        On Error Resume Next
        If ActiveDocument.Bookmarks.Exists(VTWordBookmarkName(sessionId)) Then
            Set pendingBookmark = ActiveDocument.Bookmarks(VTWordBookmarkName(sessionId))
            If pendingBookmark.Range.InlineShapes.Count = 1 Then
                Set target = pendingBookmark.Range.InlineShapes(1)
            End If
        End If
        If target Is Nothing Then Set target = VTFindUniqueInlineShape(pendingMarker)
        If Not target Is Nothing Then target.Delete
        On Error GoTo 0
    End If
    VTDeletePendingBookmark ActiveDocument, sessionId
    VTDeleteWordEditBookmark ActiveDocument, sessionId
    VTClosePreparedNativeRollbackDocument
End Sub

Public Sub VTNormalizeUnnumberedDisplayParagraph( _
    ByVal formulaRange As Range)

    Dim paragraphRange As Range
    Dim typographyAlignment As Long

    If formulaRange Is Nothing Then Exit Sub
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    typographyAlignment = paragraphRange.ParagraphFormat.BaseLineAlignment
    paragraphRange.Style = wdStyleNormal
    With paragraphRange.ParagraphFormat
        .BaseLineAlignment = typographyAlignment
        .Alignment = wdAlignParagraphCenter
        .LeftIndent = 0!
        .RightIndent = 0!
        .FirstLineIndent = 0!
        .TabStops.ClearAll
    End With
    VTApplyUnnumberedImageFormulaVerticalAlignment formulaRange
End Sub

Public Function VTDoubleClickTraceValue(ByVal value As String) As String
    value = Replace$(Replace$(value, vbCr, " "), vbLf, " ")
    value = Replace$(value, vbTab, " ")
    If Len(value) > 240 Then value = Left$(value, 240)
    VTDoubleClickTraceValue = value
End Function

Public Sub VTInitializeWordDoubleClickTrace()
    If VT_WORD_DOUBLE_CLICK_TRACE_INITIALIZED Then Exit Sub
    On Error GoTo TraceDisabled
    VT_WORD_DOUBLE_CLICK_TRACE_INITIALIZED = True
    VT_WORD_DOUBLE_CLICK_TRACE_ACTIVE = _
        VT_WORD_DOUBLE_CLICK_TRACE_ENABLED Or _
        VTPathFileExists( _
            VTApplicationSupportRoot() & _
            VT_WORD_DOUBLE_CLICK_TRACE_SENTINEL)
    Exit Sub

TraceDisabled:
    VT_WORD_DOUBLE_CLICK_TRACE_ACTIVE = False
    VT_WORD_DOUBLE_CLICK_TRACE_INITIALIZED = True
End Sub

Public Function VTWordEditBookmarkName(ByVal sessionId As String) As String
    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7420, "VisualTeX", _
            "VisualTeX cannot create an edit Bookmark for an invalid Session id."
    End If
    VTWordEditBookmarkName = _
        VT_WORD_EDIT_BOOKMARK_PREFIX & Replace$(Left$(sessionId, 24), "-", "")
    If Len(VTWordEditBookmarkName) > 40 Then
        Err.Raise vbObjectError + 7421, "VisualTeX", _
            "VisualTeX generated an edit Bookmark name longer than 40 characters."
    End If
End Function

Public Sub VTAddWordEditBookmark( _
    ByVal formulaShape As InlineShape, _
    ByVal sessionId As String)

    Dim documentObject As Document
    Dim targetRange As Range
    Dim bookmarkName As String

    If formulaShape Is Nothing Then Exit Sub
    Set documentObject = formulaShape.Range.Document
    Set targetRange = VTVisualTeXImageContainerRange(formulaShape)
    bookmarkName = VTWordEditBookmarkName(sessionId)
    On Error Resume Next
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    On Error GoTo 0
    documentObject.Bookmarks.Add Name:=bookmarkName, Range:=targetRange
End Sub

Public Sub VTAddWordEditRangeBookmark( _
    ByVal targetRange As Range, _
    ByVal sessionId As String)

    Dim documentObject As Document
    Dim exactRange As Range
    Dim bookmarkName As String

    If targetRange Is Nothing Then Exit Sub
    Set documentObject = targetRange.Document
    Set exactRange = targetRange.Duplicate
    bookmarkName = VTWordEditBookmarkName(sessionId)
    On Error Resume Next
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    On Error GoTo 0
    documentObject.Bookmarks.Add Name:=bookmarkName, Range:=exactRange
End Sub

Public Sub VTDeleteWordEditBookmark( _
    ByVal documentObject As Document, _
    ByVal sessionId As String)

    Dim bookmarkName As String

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(sessionId) Then Exit Sub
    bookmarkName = VTWordEditBookmarkName(sessionId)
    On Error Resume Next
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    On Error GoTo 0
End Sub

Public Sub VTAddPendingBookmark(ByVal targetRange As Range, ByVal sessionId As String)
    Dim documentObject As Document
    Dim name As String

    If targetRange Is Nothing Then Exit Sub
    Set documentObject = targetRange.Document
    name = VTWordBookmarkName(sessionId)
    On Error Resume Next
    If documentObject.Bookmarks.Exists(name) Then
        documentObject.Bookmarks(name).Delete
    End If
    On Error GoTo 0
    documentObject.Bookmarks.Add Name:=name, Range:=targetRange
End Sub

Public Sub VTDeletePendingBookmark( _
    ByVal documentObject As Document, _
    ByVal sessionId As String)

    Dim name As String

    If documentObject Is Nothing Then Exit Sub
    name = VTWordBookmarkName(sessionId)
    On Error Resume Next
    If documentObject.Bookmarks.Exists(name) Then
        documentObject.Bookmarks(name).Delete
    End If
    On Error GoTo 0
End Sub

Public Function VTWordBookmarkName(ByVal sessionId As String) As String
    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7420, "VisualTeX", "VisualTeX cannot create a Bookmark for an invalid Session id."
    End If
    VTWordBookmarkName = VT_WORD_BOOKMARK_PREFIX & Replace$(Left$(sessionId, 24), "-", "")
    If Len(VTWordBookmarkName) > 40 Then
        Err.Raise vbObjectError + 7421, "VisualTeX", "VisualTeX generated a Word Bookmark name longer than 40 characters."
    End If
End Function

Public Function VTDispatchOptional(ByVal dispatch As Object, ByVal key As String) As String
    If VTCollectionHasKey(dispatch, key) Then VTDispatchOptional = CStr(dispatch(key))
End Function

Public Function VTDispatchPositiveDouble(ByVal dispatch As Object, ByVal key As String) As Double
    VTRequireDispatchValue dispatch, key
    VTDispatchPositiveDouble = VTParseInvariantDouble(CStr(dispatch(key)))
    If VTDispatchPositiveDouble <= 0# Or VTDispatchPositiveDouble > 100000# Then
        Err.Raise vbObjectError + 7414, "VisualTeX", "VisualTeX dispatch contains invalid " & key & "."
    End If
End Function

Public Function VTDispatchOptionalDouble(ByVal dispatch As Object, ByVal key As String, ByVal fallback As Double) As Double
    ' VBA Or evaluates both operands, including a missing Collection lookup.
    If Not VTCollectionHasKey(dispatch, key) Then
        VTDispatchOptionalDouble = fallback
    ElseIf Len(CStr(dispatch(key))) = 0 Then
        VTDispatchOptionalDouble = fallback
    Else
        VTDispatchOptionalDouble = VTParseInvariantDouble(CStr(dispatch(key)))
    End If
End Function

' The number a new display equation at position takes. transferEquation is
' the existing equation whose \r switch moves to the new number (inserting
' before the first number of a series), or Nothing. It is a live Range, so it
' follows the equation when the new one is inserted before it.
Public Sub VTV2PlanInsertion( _
    ByVal doc As Document, _
    ByVal position As Long, _
    ByRef sequenceName As String, _
    ByRef headingLevel As Long, _
    ByRef separator As String, _
    ByRef literalPrefix As String, _
    ByRef automatic As Boolean, _
    ByRef restart As Long, _
    ByRef initial As Long, _
    ByRef transferEquation As Range, _
    Optional ByVal excludeStart As Long = -1)

    Dim entries As Collection
    Dim nearest As Long
    Dim first As Long
    Dim index As Long
    Dim scopeStart As Long
    Dim numberText As String
    Dim nextEntry As Variant

    Set transferEquation = Nothing
    restart = 0
    initial = 1
    literalPrefix = ""
    automatic = False
    VTV2DocumentFormat doc, headingLevel, separator
    VTWordPerformanceMark "v2-plan-format"
    Set entries = VTV2PlanEntries(doc, position, excludeStart)
    VTWordPerformanceMark "v2-plan-entries"
    If entries.Count = 0 Then
        sequenceName = VTNativeEquationLabelName()
    Else
        nearest = 0
        For index = 1 To entries.Count
            If entries(index)(0) < position Then nearest = index
        Next index
        If nearest = 0 Then nearest = 1
        first = nearest
        Do While first > 1
            If entries(first - 1)(3) <> entries(nearest)(3) Or _
               StrComp(entries(first - 1)(1), entries(nearest)(1), vbTextCompare) <> 0 Then Exit Do
            first = first - 1
        Loop
        sequenceName = entries(first)(1)
        If position <= entries(first)(0) Then
            restart = entries(first)(5)
            If entries(first)(6) > 0 Then _
                Set transferEquation = doc.Range(entries(first)(0), entries(first)(8))
            initial = restart
        Else
            initial = entries(first)(5) + 1
        End If
    End If
    If headingLevel > 0 Then
        numberText = VTV2HeadingScope(VTV2PlanHeadings(doc, headingLevel, position), _
            position, headingLevel, scopeStart, automatic)
        If scopeStart < 0 Then automatic = False
        If Not automatic Then literalPrefix = numberText
        ' A new number ahead of the first number of its own chapter or section (the equations before it belong to the chapter
        ' before) is the first of its series: the former first number gives up its restart (\r) to it, as before the first
        ' number of the document. The plan above looked only at the numbers before position.
        If transferEquation Is Nothing And restart = 0 Then
            nextEntry = VTV2FirstEntryAfter(doc, position, excludeStart)
            If Not IsEmpty(nextEntry) Then
                If nextEntry(3) = numberText And nextEntry(6) > 0 And _
                   StrComp(nextEntry(1), sequenceName, vbTextCompare) = 0 Then
                    restart = nextEntry(5)
                    initial = restart
                    Set transferEquation = doc.Range(nextEntry(0), nextEntry(8))
                End If
            End If
        End If
    End If
End Sub

' The typing line after a display formula, created as VisualTeX for Windows
' creates it (Range.InsertParagraphAfter on the formula's paragraph), so the
' formula's own paragraph mark keeps its formatting. Nothing when Word
' inserted nothing (the caller then splits the paragraph itself).
Private Function VTV2InsertParagraphAfter(ByVal sourceParagraph As Range) As Range

    Dim documentObject As Document
    Dim work As Range
    Dim created As Range
    Dim boundary As Long
    Dim contentEndBefore As Long

    Set documentObject = sourceParagraph.Document
    boundary = sourceParagraph.End
    contentEndBefore = documentObject.Content.End
    Set work = sourceParagraph.Duplicate
    On Error Resume Next
    work.InsertParagraphAfter
    If Err.Number <> 0 Or documentObject.Content.End = contentEndBefore Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    Set created = documentObject.Range( _
        Start:=boundary, End:=boundary + 1).Paragraphs(1).Range.Duplicate
    If created.Start <> boundary Or _
       documentObject.Content.End <> contentEndBefore + 1 Or _
       VTWordRangeHasMeaningfulText(created) Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "Word did not create the typing paragraph after the display formula."
    End If
    Set VTV2InsertParagraphAfter = created
End Function

' The local name of a paragraph's style, or "" when Word cannot report one.
Private Function VTV2ParagraphStyleName(ByVal paragraphRange As Range) As String
    On Error Resume Next
    VTV2ParagraphStyleName = paragraphRange.Paragraphs(1).Style.NameLocal
    Err.Clear
End Function

' WordFormulaService.ApplyBodyFormatToTypingParagraph: a paragraph created
' after a display formula copies its centered format, so it takes the style and
' paragraph format of the body paragraph before the formula, or its style's own
' format. Never fails the formula write.
Private Sub VTV2ApplyBodyFormatToTypingParagraph( _
    ByVal displayParagraph As Range, ByVal typingParagraph As Range)

    Dim previousRange As Range
    Dim bodyFormat As ParagraphFormat
    Dim bodyStyleName As String

    On Error GoTo Finished
    If displayParagraph Is Nothing Or typingParagraph Is Nothing Then Exit Sub
    If displayParagraph.Start > 0 Then
        Set previousRange = displayParagraph.Document.Range( _
            Start:=displayParagraph.Start - 1, _
            End:=displayParagraph.Start).Paragraphs(1).Range.Duplicate
        If previousRange.End = displayParagraph.Start And _
           previousRange.Paragraphs(1).OutlineLevel = wdOutlineLevelBodyText And _
           previousRange.OMaths.Count = 0 And _
           previousRange.InlineShapes.Count = 0 And _
           previousRange.Tables.Count = 0 Then
            Set bodyFormat = previousRange.ParagraphFormat.Duplicate
            On Error Resume Next
            bodyStyleName = previousRange.Paragraphs(1).Style.NameLocal
            If Len(bodyStyleName) = 0 Then
                bodyStyleName = displayParagraph.Document.Range( _
                    Start:=previousRange.End - 1, _
                    End:=previousRange.End).Style.NameLocal
            End If
            Err.Clear
            On Error GoTo Finished
            ' Word for Mac splits the next paragraph's first run on any style
            ' assignment, so assign only a different style.
            If Len(bodyStyleName) > 0 Then
                If StrComp(VTV2ParagraphStyleName(typingParagraph), bodyStyleName, vbTextCompare) <> 0 Then _
                    typingParagraph.Style = bodyStyleName
            End If
            typingParagraph.ParagraphFormat = bodyFormat
            Exit Sub
        End If
    End If
    typingParagraph.ParagraphFormat.Reset
Finished:
    Err.Clear
End Sub

' WordFormulaHostLayout.ApplyOmmlLocalTypography: Position 0, the formula's
' size and the same complex-script size on the whole equation.
' VisualTeX for Windows' ReadCurrentTypingFontSize: a new formula follows the
' text where it goes. The size Word reports for the insertion point, else the
' character before it in its paragraph, else the one after it, else the
' paragraph's text; 14 pt when none is a single size from 5 to 200 pt.
Public Function VTV2CurrentTypingFontSize() As Double
    Dim caret As Range
    Dim paragraph As Range
    Dim probe As Range
    Dim size As Double

    VTV2CurrentTypingFontSize = 14#
    On Error Resume Next
    size = Selection.Font.Size
    If size >= 5# And size <= 200# Then GoTo Found
    Set caret = Selection.Range.Duplicate
    If caret Is Nothing Then Exit Function
    Set paragraph = caret.Paragraphs(1).Range.Duplicate
    If paragraph Is Nothing Then Exit Function
    If caret.Start > paragraph.Start Then
        Set probe = caret.Duplicate
        probe.SetRange caret.Start - 1, caret.Start
        size = 0#
        size = probe.Font.Size
        If size >= 5# And size <= 200# Then GoTo Found
    End If
    If caret.Start < paragraph.End - 1 Then
        Set probe = caret.Duplicate
        probe.SetRange caret.Start, caret.Start + 1
        size = 0#
        size = probe.Font.Size
        If size >= 5# And size <= 200# Then GoTo Found
    End If
    If paragraph.End > paragraph.Start Then paragraph.End = paragraph.End - 1
    size = 0#
    size = paragraph.Font.Size
    If size >= 5# And size <= 200# Then GoTo Found
    Exit Function
Found:
    VTV2CurrentTypingFontSize = size
End Function
